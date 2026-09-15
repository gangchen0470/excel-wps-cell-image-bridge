from __future__ import annotations

import json
import hashlib
import math
import posixpath
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from .model import CellImage, FloatingImage, InvalidWorkbookError, UnsupportedWorkbookError

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wps": "http://www.wps.cn/officeDocument/2017/etCustomData",
    "xlrd": "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata",
    "xrvrel": "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    "xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
    "xr3": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
}

for prefix in ("r", "xdr", "a", "mc", "x14ac", "xr", "xr2", "xr3"):
    ET.register_namespace(prefix, NS[prefix])
ET.register_namespace("etc", NS["wps"])
ET.register_namespace("xlrd", NS["xlrd"])
ET.register_namespace("xrvrel", NS["xrvrel"])
ET.register_namespace("", NS["main"])

REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
WPS_CELL_IMAGE_REL = "http://www.wps.cn/officeDocument/2020/cellImage"
RICH_REL_TYPES = {
    "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes",
    "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure",
    "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue",
    "http://schemas.microsoft.com/office/2022/10/relationships/richValueRel",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata",
}
DISPIMG_RE = re.compile(r'DISPIMG\("([^"]+)"', re.IGNORECASE)
CELL_RE = re.compile(r"^([A-Z]+)([1-9][0-9]*)$")
RANGE_RE = re.compile(r"^([A-Z]+[1-9][0-9]*)(?::([A-Z]+[1-9][0-9]*))?$")


def _xml(data: bytes, name: str) -> ET.Element:
    try:
        return ET.fromstring(data)
    except ET.ParseError as exc:
        raise InvalidWorkbookError(f"Invalid XML in {name}: {exc}") from exc


def _resolve(source_part: str, target: str) -> str:
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def _rels_path(part: str) -> str:
    return posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")


def _rel_map(parts: dict[str, bytes], part: str) -> dict[str, tuple[str, str]]:
    path = _rels_path(part)
    if path not in parts:
        return {}
    root = _xml(parts[path], path)
    result = {}
    for rel in root.findall("pr:Relationship", NS):
        result[rel.attrib["Id"]] = (rel.attrib["Type"], _resolve(part, rel.attrib["Target"]))
    return result


def _load(path: Path) -> dict[str, bytes]:
    if not path.is_file():
        raise FileNotFoundError(path)
    if not zipfile.is_zipfile(path):
        raise InvalidWorkbookError(f"Not a valid XLSX ZIP package: {path}")
    with zipfile.ZipFile(path) as archive:
        bad = archive.testzip()
        if bad:
            raise InvalidWorkbookError(f"Corrupt ZIP member: {bad}")
        return {name: archive.read(name) for name in archive.namelist() if not name.endswith("/")}


def _workbook_sheets(parts: dict[str, bytes]) -> list[tuple[str, str]]:
    workbook = "xl/workbook.xml"
    if workbook not in parts:
        raise InvalidWorkbookError("Missing xl/workbook.xml")
    rels = _rel_map(parts, workbook)
    root = _xml(parts[workbook], workbook)
    sheets = []
    for sheet in root.findall("main:sheets/main:sheet", NS):
        rid = sheet.attrib.get(f"{{{NS['r']}}}id")
        if rid in rels:
            sheets.append((sheet.attrib.get("name", ""), rels[rid][1]))
    return sheets


def _content_types(parts: dict[str, bytes]) -> dict[str, str]:
    root = _xml(parts["[Content_Types].xml"], "[Content_Types].xml")
    defaults = {n.attrib["Extension"].lower(): n.attrib["ContentType"] for n in root.findall("ct:Default", NS)}
    overrides = {n.attrib["PartName"].lstrip("/"): n.attrib["ContentType"] for n in root.findall("ct:Override", NS)}
    result = dict(overrides)
    for name in parts:
        result.setdefault(name, defaults.get(Path(name).suffix.lstrip(".").lower(), "application/octet-stream"))
    return result


def _detect_wps(parts: dict[str, bytes], sheets: list[tuple[str, str]], types: dict[str, str]) -> list[CellImage]:
    cell_images = "xl/cellimages.xml"
    if cell_images not in parts:
        return []
    rels = _rel_map(parts, cell_images)
    root = _xml(parts[cell_images], cell_images)
    image_by_id: dict[str, str] = {}
    for pic in root.findall("wps:cellImage/xdr:pic", NS):
        prop = pic.find("xdr:nvPicPr/xdr:cNvPr", NS)
        blip = pic.find("xdr:blipFill/a:blip", NS)
        if prop is None or blip is None:
            continue
        image_id = prop.attrib.get("name", "")
        rid = blip.attrib.get(f"{{{NS['r']}}}embed", "")
        if image_id and rid in rels and rels[rid][0] == REL_IMAGE:
            image_by_id[image_id] = rels[rid][1]

    found = []
    for sheet_name, sheet_path in sheets:
        root = _xml(parts[sheet_path], sheet_path)
        for cell in root.findall(".//main:c", NS):
            formula = cell.findtext("main:f", default="", namespaces=NS)
            match = DISPIMG_RE.search(formula)
            if not match:
                continue
            image_id = match.group(1)
            media = image_by_id.get(image_id)
            if media in parts:
                found.append(CellImage(sheet_name, sheet_path, cell.attrib["r"], media, "wps", image_id, types.get(media)))
    return found


def _detect_excel(parts: dict[str, bytes], sheets: list[tuple[str, str]], types: dict[str, str]) -> list[CellImage]:
    required = ("xl/metadata.xml", "xl/richData/rdrichvalue.xml", "xl/richData/richValueRel.xml")
    if any(name not in parts for name in required):
        return []

    metadata = _xml(parts["xl/metadata.xml"], "xl/metadata.xml")
    rich_indexes = []
    for block in metadata.findall("main:valueMetadata/main:bk", NS):
        rc = block.find("main:rc", NS)
        rich_indexes.append(int(rc.attrib.get("v", "-1")) if rc is not None else -1)

    rel_root = _xml(parts["xl/richData/richValueRel.xml"], "xl/richData/richValueRel.xml")
    rel_ids = [n.attrib.get(f"{{{NS['r']}}}id", "") for n in rel_root.findall("xrvrel:rel", NS)]
    rels = _rel_map(parts, "xl/richData/richValueRel.xml")
    rich_root = _xml(parts["xl/richData/rdrichvalue.xml"], "xl/richData/rdrichvalue.xml")
    rich_media: list[tuple[str, str] | None] = []
    for index, rv in enumerate(rich_root.findall("xlrd:rv", NS)):
        values = rv.findall("xlrd:v", NS)
        try:
            rel_index = int(values[0].text or "-1")
        except (IndexError, ValueError):
            rich_media.append(None)
            continue
        if 0 <= rel_index < len(rel_ids) and rel_ids[rel_index] in rels:
            media = rels[rel_ids[rel_index]][1]
            rich_media.append((str(index), media))
        else:
            rich_media.append(None)

    found = []
    for sheet_name, sheet_path in sheets:
        root = _xml(parts[sheet_path], sheet_path)
        for cell in root.findall(".//main:c", NS):
            try:
                vm_index = int(cell.attrib.get("vm", "0")) - 1
                rich_index = rich_indexes[vm_index]
                item = rich_media[rich_index]
            except (ValueError, IndexError):
                continue
            if item and item[1] in parts:
                found.append(CellImage(sheet_name, sheet_path, cell.attrib["r"], item[1], "excel", item[0], types.get(item[1])))
    return found


def _inspect_parts(parts: dict[str, bytes]) -> list[CellImage]:
    sheets = _workbook_sheets(parts)
    types = _content_types(parts)
    return _detect_excel(parts, sheets, types) + _detect_wps(parts, sheets, types)


def inspect_workbook(path: str | Path) -> list[CellImage]:
    return _inspect_parts(_load(Path(path)))


def _column_index(label: str) -> int:
    result = 0
    for char in label:
        result = result * 26 + ord(char) - 64
    return result - 1


def _cell_position(cell: str) -> tuple[int, int]:
    match = CELL_RE.match(cell.upper())
    if not match:
        raise InvalidWorkbookError(f"Invalid cell reference: {cell}")
    return _column_index(match.group(1)), int(match.group(2)) - 1


def _select_images(images: list[CellImage], sheet: str | None = None, cell_range: str | None = None) -> list[CellImage]:
    bounds = None
    if cell_range:
        match = RANGE_RE.match(cell_range.upper())
        if not match:
            raise ValueError(f"Invalid cell range: {cell_range}")
        start = _cell_position(match.group(1))
        end = _cell_position(match.group(2) or match.group(1))
        bounds = (min(start[0], end[0]), min(start[1], end[1]), max(start[0], end[0]), max(start[1], end[1]))
    selected = []
    for item in images:
        if sheet is not None and item.sheet_name != sheet:
            continue
        if bounds is not None:
            col, row = _cell_position(item.cell)
            if not (bounds[0] <= col <= bounds[2] and bounds[1] <= row <= bounds[3]):
                continue
        selected.append(item)
    return selected


def _image_size(data: bytes, content_type: str | None) -> tuple[int, int]:
    if data.startswith(b"\x89PNG\r\n\x1a\n") and len(data) >= 24:
        return int.from_bytes(data[16:20], "big"), int.from_bytes(data[20:24], "big")
    if data[:6] in {b"GIF87a", b"GIF89a"} and len(data) >= 10:
        return int.from_bytes(data[6:8], "little"), int.from_bytes(data[8:10], "little")
    if data.startswith(b"BM") and len(data) >= 26:
        width = abs(int.from_bytes(data[18:22], "little", signed=True))
        height = abs(int.from_bytes(data[22:26], "little", signed=True))
        if width and height:
            return width, height
    if data.startswith(b"RIFF") and data[8:12] == b"WEBP" and len(data) >= 30:
        chunk = data[12:16]
        if chunk == b"VP8X":
            return 1 + int.from_bytes(data[24:27], "little"), 1 + int.from_bytes(data[27:30], "little")
        if chunk == b"VP8 " and data[23:26] == b"\x9d\x01\x2a":
            return int.from_bytes(data[26:28], "little") & 0x3FFF, int.from_bytes(data[28:30], "little") & 0x3FFF
        if chunk == b"VP8L" and data[20] == 0x2F:
            b1, b2, b3, b4 = data[21:25]
            return 1 + b1 + ((b2 & 0x3F) << 8), 1 + (b2 >> 6) + (b3 << 2) + ((b4 & 0x0F) << 10)
    if data.startswith(b"\xff\xd8"):
        i = 2
        sof_markers = {0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF}
        while i + 9 < len(data):
            if data[i] != 0xFF:
                i += 1
                continue
            marker = data[i + 1]
            if marker in sof_markers:
                return int.from_bytes(data[i + 5:i + 7], "big"), int.from_bytes(data[i + 7:i + 9], "big")
            if i + 4 > len(data):
                break
            i += 2 + int.from_bytes(data[i + 2:i + 4], "big")
    return 1, 1


def _sheet_cell_emu(sheet: ET.Element, col: int, row: int) -> tuple[int, int]:
    sheet_format = sheet.find("main:sheetFormatPr", NS)
    width = float(sheet_format.attrib.get("defaultColWidth", "8.43")) if sheet_format is not None else 8.43
    for node in sheet.findall("main:cols/main:col", NS):
        if int(node.attrib["min"]) - 1 <= col <= int(node.attrib["max"]) - 1:
            width = float(node.attrib.get("width", width))
            break
    height = float(sheet_format.attrib.get("defaultRowHeight", "15")) if sheet_format is not None else 15.0
    row_node = sheet.find(f"main:sheetData/main:row[@r='{row + 1}']", NS)
    if row_node is not None:
        height = float(row_node.attrib.get("ht", height))
    # Use the conservative anchor width shared by Excel and WPS. Excel may
    # visually add cell padding, while WPS includes it inconsistently for
    # custom widths; excluding it guarantees that the picture stays bounded.
    pixels_w = math.floor(((256 * width + math.floor(128 / 7)) / 256) * 7)
    pixels_h = height * 96 / 72
    return max(1, round(pixels_w * 9525)), max(1, round(pixels_h * 9525))


def _absolute_cell_origin(sheet: ET.Element, col: int, row: int) -> tuple[int, int]:
    return (
        sum(_sheet_cell_emu(sheet, index, 0)[0] for index in range(col)),
        sum(_sheet_cell_emu(sheet, 0, index)[1] for index in range(row)),
    )


def _cell_at_point(sheet: ET.Element, x: int, y: int) -> tuple[int, int]:
    col = row = 0
    while col < 16383:
        width = _sheet_cell_emu(sheet, col, 0)[0]
        if x < width:
            break
        x -= width
        col += 1
    while row < 1048575:
        height = _sheet_cell_emu(sheet, 0, row)[1]
        if y < height:
            break
        y -= height
        row += 1
    return col, row


def _cell_ref(col: int, row: int) -> str:
    label = ""
    value = col + 1
    while value:
        value, remainder = divmod(value - 1, 26)
        label = chr(65 + remainder) + label
    return f"{label}{row + 1}"


def _marker_point(sheet: ET.Element, marker: ET.Element) -> tuple[int, int]:
    col = int(marker.findtext("xdr:col", "0", NS))
    row = int(marker.findtext("xdr:row", "0", NS))
    x, y = _absolute_cell_origin(sheet, col, row)
    return x + int(marker.findtext("xdr:colOff", "0", NS)), y + int(marker.findtext("xdr:rowOff", "0", NS))


def _floating_center_cell(sheet: ET.Element, anchor: ET.Element) -> str | None:
    kind = anchor.tag.rsplit("}", 1)[-1]
    if kind == "oneCellAnchor":
        start = anchor.find("xdr:from", NS)
        extent = anchor.find("xdr:ext", NS)
        if start is None or extent is None:
            return None
        x, y = _marker_point(sheet, start)
        x += int(extent.attrib.get("cx", "0")) // 2
        y += int(extent.attrib.get("cy", "0")) // 2
    elif kind == "twoCellAnchor":
        start, end = anchor.find("xdr:from", NS), anchor.find("xdr:to", NS)
        if start is None or end is None:
            return None
        x1, y1 = _marker_point(sheet, start)
        x2, y2 = _marker_point(sheet, end)
        x, y = (x1 + x2) // 2, (y1 + y2) // 2
    else:
        return None
    return _cell_ref(*_cell_at_point(sheet, x, y))


def _detect_floating(parts: dict[str, bytes], sheets: list[tuple[str, str]], types: dict[str, str]) -> list[FloatingImage]:
    found = []
    for sheet_name, sheet_path in sheets:
        sheet = _xml(parts[sheet_path], sheet_path)
        sheet_rels = _rel_map(parts, sheet_path)
        drawing_node = sheet.find("main:drawing", NS)
        if drawing_node is None:
            continue
        drawing_rid = drawing_node.attrib.get(f"{{{NS['r']}}}id", "")
        if drawing_rid not in sheet_rels:
            continue
        drawing_path = sheet_rels[drawing_rid][1]
        if drawing_path not in parts:
            continue
        drawing = _xml(parts[drawing_path], drawing_path)
        drawing_rels = _rel_map(parts, drawing_path)
        for index, anchor in enumerate(list(drawing)):
            cell = _floating_center_cell(sheet, anchor)
            blip = anchor.find(".//a:blip", NS)
            if cell is None or blip is None:
                continue
            rid = blip.attrib.get(f"{{{NS['r']}}}embed", "")
            if rid in drawing_rels and drawing_rels[rid][0] == REL_IMAGE and drawing_rels[rid][1] in parts:
                media = drawing_rels[rid][1]
                found.append(FloatingImage(sheet_name, sheet_path, cell, media, "floating", rid, types.get(media), drawing_path, rid, index))
    return found


def inspect_floating_images(path: str | Path) -> list[FloatingImage]:
    parts = _load(Path(path))
    return _detect_floating(parts, _workbook_sheets(parts), _content_types(parts))


def _fit_contain(cell_w: int, cell_h: int, img_w: int, img_h: int, margin: int) -> tuple[int, int, int, int]:
    """Return bounded width, height and centered offsets in EMUs."""
    margin_emu = max(0, margin) * 9525
    avail_w = max(1, cell_w - 2 * margin_emu)
    avail_h = max(1, cell_h - 2 * margin_emu)
    scale = min(avail_w / max(1, img_w), avail_h / max(1, img_h))
    draw_w = min(avail_w, max(1, math.floor(img_w * scale)))
    draw_h = min(avail_h, max(1, math.floor(img_h * scale)))
    return draw_w, draw_h, (cell_w - draw_w) // 2, (cell_h - draw_h) // 2


def _next_rid(root: ET.Element) -> str:
    used = {n.attrib.get("Id", "") for n in root.findall("pr:Relationship", NS)}
    i = 1
    while f"rId{i}" in used:
        i += 1
    return f"rId{i}"


def _new_anchor(item: CellImage, rel_id: str, image_number: int, sheet_root: ET.Element, media: bytes, margin: int) -> ET.Element:
    col, row = _cell_position(item.cell)
    cell_w, cell_h = _sheet_cell_emu(sheet_root, col, row)
    img_w, img_h = _image_size(media, item.content_type)
    draw_w, draw_h, xoff, yoff = _fit_contain(cell_w, cell_h, img_w, img_h, margin)

    anchor = ET.Element(f"{{{NS['xdr']}}}oneCellAnchor")
    marker = ET.SubElement(anchor, f"{{{NS['xdr']}}}from")
    for name, value in (("col", col), ("colOff", xoff), ("row", row), ("rowOff", yoff)):
        ET.SubElement(marker, f"{{{NS['xdr']}}}{name}").text = str(value)
    ET.SubElement(anchor, f"{{{NS['xdr']}}}ext", {"cx": str(draw_w), "cy": str(draw_h)})
    pic = ET.SubElement(anchor, f"{{{NS['xdr']}}}pic")
    nv = ET.SubElement(pic, f"{{{NS['xdr']}}}nvPicPr")
    ET.SubElement(nv, f"{{{NS['xdr']}}}cNvPr", {"id": str(image_number), "name": f"Picture {image_number}"})
    locks = ET.SubElement(ET.SubElement(nv, f"{{{NS['xdr']}}}cNvPicPr"), f"{{{NS['a']}}}picLocks")
    locks.set("noChangeAspect", "1")
    fill = ET.SubElement(pic, f"{{{NS['xdr']}}}blipFill")
    ET.SubElement(fill, f"{{{NS['a']}}}blip", {f"{{{NS['r']}}}embed": rel_id})
    stretch = ET.SubElement(fill, f"{{{NS['a']}}}stretch")
    ET.SubElement(stretch, f"{{{NS['a']}}}fillRect")
    shape = ET.SubElement(pic, f"{{{NS['xdr']}}}spPr")
    transform = ET.SubElement(shape, f"{{{NS['a']}}}xfrm")
    ET.SubElement(transform, f"{{{NS['a']}}}off", {"x": "0", "y": "0"})
    ET.SubElement(transform, f"{{{NS['a']}}}ext", {"cx": str(draw_w), "cy": str(draw_h)})
    geom = ET.SubElement(shape, f"{{{NS['a']}}}prstGeom", {"prst": "rect"})
    ET.SubElement(geom, f"{{{NS['a']}}}avLst")
    ET.SubElement(anchor, f"{{{NS['xdr']}}}clientData", {"fLocksWithSheet": "0", "fPrintsWithSheet": "1"})
    return anchor


def _serialize(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _normalize_ignorable(root: ET.Element) -> None:
    """Keep mc:Ignorable tokens aligned with namespaces present after serialization."""
    key = f"{{{NS['mc']}}}Ignorable"
    value = root.attrib.get(key)
    if not value:
        return
    # Attribute values are intentionally excluded: a prefix named only by
    # mc:Ignorable is not sufficient to justify retaining it.
    qualified_names = " ".join(
        name for node in root.iter() for name in (node.tag, *node.attrib.keys())
    )
    retained = [token for token in value.split() if token in NS and f"{{{NS[token]}}}" in qualified_names]
    if retained:
        root.set(key, " ".join(retained))
    else:
        root.attrib.pop(key, None)


def _lock_row_height(root: ET.Element, cell_ref: str) -> None:
    """Freeze the row's current effective height without changing its value."""
    _, row_index = _cell_position(cell_ref)
    row_node = root.find(f"main:sheetData/main:row[@r='{row_index + 1}']", NS)
    if row_node is None:
        return
    if "ht" not in row_node.attrib:
        defaults = root.find("main:sheetFormatPr", NS)
        row_node.set("ht", defaults.attrib.get("defaultRowHeight", "15") if defaults is not None else "15")
    row_node.set("customHeight", "1")


def _ensure_content_type(parts: dict[str, bytes], drawing_path: str) -> None:
    root = _xml(parts["[Content_Types].xml"], "[Content_Types].xml")
    part_name = "/" + drawing_path
    if not any(n.attrib.get("PartName") == part_name for n in root.findall("ct:Override", NS)):
        ET.SubElement(root, f"{{{NS['ct']}}}Override", {
            "PartName": part_name,
            "ContentType": "application/vnd.openxmlformats-officedocument.drawing+xml",
        })
    parts["[Content_Types].xml"] = _serialize(root)


def _add_content_override(parts: dict[str, bytes], part_name: str, content_type: str) -> None:
    root = _xml(parts["[Content_Types].xml"], "[Content_Types].xml")
    normalized = "/" + part_name.lstrip("/")
    for node in root.findall("ct:Override", NS):
        if node.attrib.get("PartName") == normalized:
            node.set("ContentType", content_type)
            break
    else:
        ET.SubElement(root, f"{{{NS['ct']}}}Override", {"PartName": normalized, "ContentType": content_type})
    parts["[Content_Types].xml"] = _serialize(root)


def _add_workbook_rel(parts: dict[str, bytes], rel_type: str, target: str) -> str:
    path = "xl/_rels/workbook.xml.rels"
    root = _xml(parts[path], path)
    for rel in root.findall("pr:Relationship", NS):
        if rel.attrib.get("Type") == rel_type:
            return rel.attrib["Id"]
    rid = _next_rid(root)
    ET.SubElement(root, f"{{{NS['pr']}}}Relationship", {"Id": rid, "Type": rel_type, "Target": target})
    parts[path] = _serialize(root)
    return rid


def _convert_sheet(parts: dict[str, bytes], sheet_path: str, items: list[CellImage], margin: int, drawing_no: int) -> None:
    sheet_root = _xml(parts[sheet_path], sheet_path)
    sheet_rels_path = _rels_path(sheet_path)
    if sheet_rels_path in parts:
        sheet_rels = _xml(parts[sheet_rels_path], sheet_rels_path)
    else:
        sheet_rels = ET.Element(f"{{{NS['pr']}}}Relationships")

    drawing_rel = None
    drawing_path = None
    for rel in sheet_rels.findall("pr:Relationship", NS):
        if rel.attrib.get("Type") == REL_DRAWING:
            drawing_rel = rel.attrib["Id"]
            drawing_path = _resolve(sheet_path, rel.attrib["Target"])
            break
    if drawing_path and drawing_path in parts:
        drawing_root = _xml(parts[drawing_path], drawing_path)
        drawing_rels_path = _rels_path(drawing_path)
        drawing_rels = _xml(parts[drawing_rels_path], drawing_rels_path) if drawing_rels_path in parts else ET.Element(f"{{{NS['pr']}}}Relationships")
    else:
        drawing_path = f"xl/drawings/drawing{drawing_no}.xml"
        drawing_rels_path = _rels_path(drawing_path)
        drawing_root = ET.Element(f"{{{NS['xdr']}}}wsDr")
        drawing_rels = ET.Element(f"{{{NS['pr']}}}Relationships")
        drawing_rel = _next_rid(sheet_rels)
        ET.SubElement(sheet_rels, f"{{{NS['pr']}}}Relationship", {
            "Id": drawing_rel, "Type": REL_DRAWING,
            "Target": posixpath.relpath(drawing_path, posixpath.dirname(sheet_path)),
        })
        ET.SubElement(sheet_root, f"{{{NS['main']}}}drawing", {f"{{{NS['r']}}}id": drawing_rel})

    picture_id = len(drawing_root.findall(".//xdr:pic", NS)) + 1
    for item in items:
        rid = _next_rid(drawing_rels)
        ET.SubElement(drawing_rels, f"{{{NS['pr']}}}Relationship", {
            "Id": rid, "Type": REL_IMAGE,
            "Target": posixpath.relpath(item.media_path, posixpath.dirname(drawing_path)),
        })
        drawing_root.append(_new_anchor(item, rid, picture_id, sheet_root, parts[item.media_path], margin))
        picture_id += 1
        cell = sheet_root.find(f".//main:c[@r='{item.cell}']", NS)
        if cell is not None:
            cell.attrib.pop("vm", None)
            for child in list(cell):
                if child.tag in {f"{{{NS['main']}}}f", f"{{{NS['main']}}}v"}:
                    cell.remove(child)
            cell.attrib.pop("t", None)
        _lock_row_height(sheet_root, item.cell)

    _normalize_ignorable(sheet_root)
    parts[sheet_path] = _serialize(sheet_root)
    parts[sheet_rels_path] = _serialize(sheet_rels)
    parts[drawing_path] = _serialize(drawing_root)
    parts[drawing_rels_path] = _serialize(drawing_rels)
    _ensure_content_type(parts, drawing_path)


def _remove_native_indexes(parts: dict[str, bytes], formats: set[str]) -> None:
    workbook_rels_path = "xl/_rels/workbook.xml.rels"
    rels = _xml(parts[workbook_rels_path], workbook_rels_path)
    for rel in list(rels):
        rel_type = rel.attrib.get("Type", "")
        if ("wps" in formats and rel_type == WPS_CELL_IMAGE_REL) or ("excel" in formats and rel_type in RICH_REL_TYPES):
            rels.remove(rel)
    parts[workbook_rels_path] = _serialize(rels)

    removals = set()
    if "wps" in formats:
        removals.update({"xl/cellimages.xml", "xl/_rels/cellimages.xml.rels"})
    if "excel" in formats:
        removals.update(name for name in parts if name == "xl/metadata.xml" or name.startswith("xl/richData/"))
    for name in removals:
        parts.pop(name, None)

    types = _xml(parts["[Content_Types].xml"], "[Content_Types].xml")
    for node in list(types):
        if node.attrib.get("PartName", "").lstrip("/") in removals:
            types.remove(node)
    parts["[Content_Types].xml"] = _serialize(types)


def _set_cell_excel_native(parts: dict[str, bytes], item: CellImage, metadata_index: int) -> None:
    root = _xml(parts[item.sheet_path], item.sheet_path)
    cell = _ensure_cell(root, item.cell)
    cell.set("t", "e")
    cell.set("vm", str(metadata_index + 1))
    for child in list(cell):
        if child.tag in {f"{{{NS['main']}}}f", f"{{{NS['main']}}}v"}:
            cell.remove(child)
    ET.SubElement(cell, f"{{{NS['main']}}}v").text = "#VALUE!"
    _lock_row_height(root, item.cell)
    _normalize_ignorable(root)
    parts[item.sheet_path] = _serialize(root)


def _write_excel_native(parts: dict[str, bytes], images: list[CellImage]) -> None:
    _remove_native_indexes(parts, {"wps"})
    metadata = ET.Element(f"{{{NS['main']}}}metadata")
    metadata_types = ET.SubElement(metadata, f"{{{NS['main']}}}metadataTypes", {"count": "1"})
    ET.SubElement(metadata_types, f"{{{NS['main']}}}metadataType", {
        "name": "XLRICHVALUE", "minSupportedVersion": "120000", "copy": "1", "pasteAll": "1",
        "pasteValues": "1", "merge": "1", "splitFirst": "1", "rowColShift": "1", "clearFormats": "1",
        "clearComments": "1", "assign": "1", "coerce": "1",
    })
    future = ET.SubElement(metadata, f"{{{NS['main']}}}futureMetadata", {"name": "XLRICHVALUE", "count": str(len(images))})
    values = ET.SubElement(metadata, f"{{{NS['main']}}}valueMetadata", {"count": str(len(images))})
    for index, item in enumerate(images):
        bk = ET.SubElement(future, f"{{{NS['main']}}}bk")
        ext_list = ET.SubElement(bk, f"{{{NS['main']}}}extLst")
        ext = ET.SubElement(ext_list, f"{{{NS['main']}}}ext", {"uri": "{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"})
        ET.SubElement(ext, f"{{{NS['xlrd']}}}rvb", {"i": str(index)})
        value_bk = ET.SubElement(values, f"{{{NS['main']}}}bk")
        ET.SubElement(value_bk, f"{{{NS['main']}}}rc", {"t": "1", "v": str(index)})
        _set_cell_excel_native(parts, item, index)
    parts["xl/metadata.xml"] = _serialize(metadata)

    rv_data = ET.Element(f"{{{NS['xlrd']}}}rvData", {"count": str(len(images))})
    rel_list = ET.Element(f"{{{NS['xrvrel']}}}richValueRels")
    rels = ET.Element(f"{{{NS['pr']}}}Relationships")
    for index, item in enumerate(images):
        rv = ET.SubElement(rv_data, f"{{{NS['xlrd']}}}rv", {"s": "0"})
        ET.SubElement(rv, f"{{{NS['xlrd']}}}v").text = str(index)
        ET.SubElement(rv, f"{{{NS['xlrd']}}}v").text = "5"
        rid = f"rId{index + 1}"
        ET.SubElement(rel_list, f"{{{NS['xrvrel']}}}rel", {f"{{{NS['r']}}}id": rid})
        ET.SubElement(rels, f"{{{NS['pr']}}}Relationship", {
            "Id": rid, "Type": REL_IMAGE,
            "Target": posixpath.relpath(item.media_path, "xl/richData"),
        })
    parts["xl/richData/rdrichvalue.xml"] = _serialize(rv_data)
    parts["xl/richData/richValueRel.xml"] = _serialize(rel_list)
    parts["xl/richData/_rels/richValueRel.xml.rels"] = _serialize(rels)

    structures = ET.Element(f"{{{NS['xlrd']}}}rvStructures", {"count": "1"})
    structure = ET.SubElement(structures, f"{{{NS['xlrd']}}}s", {"t": "_localImage"})
    ET.SubElement(structure, f"{{{NS['xlrd']}}}k", {"n": "_rvRel:LocalImageIdentifier", "t": "i"})
    ET.SubElement(structure, f"{{{NS['xlrd']}}}k", {"n": "CalcOrigin", "t": "i"})
    parts["xl/richData/rdrichvaluestructure.xml"] = _serialize(structures)

    rich2 = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"
    types_info = ET.Element(f"{{{rich2}}}rvTypesInfo")
    global_node = ET.SubElement(types_info, f"{{{rich2}}}global")
    key_flags = ET.SubElement(global_node, f"{{{rich2}}}keyFlags")
    flags = {
        "_Self": ("ExcludeFromFile", "ExcludeFromCalcComparison"),
        "_DisplayString": ("ExcludeFromCalcComparison",),
        "_Flags": ("ExcludeFromCalcComparison",),
        "_Format": ("ExcludeFromCalcComparison",),
        "_SubLabel": ("ExcludeFromCalcComparison",),
        "_Attribution": ("ExcludeFromCalcComparison",),
        "_Icon": ("ExcludeFromCalcComparison",),
        "_Display": ("ExcludeFromCalcComparison",),
        "_CanonicalPropertyNames": ("ExcludeFromCalcComparison",),
        "_ClassificationId": ("ExcludeFromCalcComparison",),
    }
    for key_name, flag_names in flags.items():
        key_node = ET.SubElement(key_flags, f"{{{rich2}}}key", {"name": key_name})
        for flag_name in flag_names:
            ET.SubElement(key_node, f"{{{rich2}}}flag", {"name": flag_name, "value": "1"})
    parts["xl/richData/rdRichValueTypes.xml"] = _serialize(types_info)

    relationships = (
        ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata", "metadata.xml"),
        ("http://schemas.microsoft.com/office/2022/10/relationships/richValueRel", "richData/richValueRel.xml"),
        ("http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue", "richData/rdrichvalue.xml"),
        ("http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure", "richData/rdrichvaluestructure.xml"),
        ("http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes", "richData/rdRichValueTypes.xml"),
    )
    for rel_type, target in relationships:
        _add_workbook_rel(parts, rel_type, target)
    for name, content_type in (
        ("xl/metadata.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"),
        ("xl/richData/richValueRel.xml", "application/vnd.ms-excel.richvaluerel+xml"),
        ("xl/richData/rdrichvalue.xml", "application/vnd.ms-excel.rdrichvalue+xml"),
        ("xl/richData/rdrichvaluestructure.xml", "application/vnd.ms-excel.rdrichvaluestructure+xml"),
        ("xl/richData/rdRichValueTypes.xml", "application/vnd.ms-excel.rdrichvaluetypes+xml"),
    ):
        _add_content_override(parts, name, content_type)


def _set_cell_wps_native(parts: dict[str, bytes], item: CellImage, image_id: str) -> None:
    root = _xml(parts[item.sheet_path], item.sheet_path)
    cell = _ensure_cell(root, item.cell)
    cell.attrib.pop("vm", None)
    cell.set("t", "str")
    for child in list(cell):
        if child.tag in {f"{{{NS['main']}}}f", f"{{{NS['main']}}}v"}:
            cell.remove(child)
    formula = f'_xlfn.DISPIMG("{image_id}",1)'
    ET.SubElement(cell, f"{{{NS['main']}}}f").text = formula
    ET.SubElement(cell, f"{{{NS['main']}}}v").text = "=" + formula.removeprefix("_xlfn.")
    _lock_row_height(root, item.cell)
    _normalize_ignorable(root)
    parts[item.sheet_path] = _serialize(root)


def _ensure_cell(root: ET.Element, cell_ref: str) -> ET.Element:
    existing = root.find(f".//main:c[@r='{cell_ref}']", NS)
    if existing is not None:
        return existing
    col, row_index = _cell_position(cell_ref)
    sheet_data = root.find("main:sheetData", NS)
    if sheet_data is None:
        raise InvalidWorkbookError("Worksheet has no sheetData")
    row = sheet_data.find(f"main:row[@r='{row_index + 1}']", NS)
    if row is None:
        row = ET.Element(f"{{{NS['main']}}}row", {"r": str(row_index + 1)})
        inserted = False
        for index, candidate in enumerate(list(sheet_data)):
            if int(candidate.attrib.get("r", "0")) > row_index + 1:
                sheet_data.insert(index, row)
                inserted = True
                break
        if not inserted:
            sheet_data.append(row)
    cell = ET.Element(f"{{{NS['main']}}}c", {"r": cell_ref})
    inserted = False
    for index, candidate in enumerate(list(row)):
        candidate_col, _ = _cell_position(candidate.attrib.get("r", "A1"))
        if candidate_col > col:
            row.insert(index, cell)
            inserted = True
            break
    if not inserted:
        row.append(cell)
    return cell


def _remove_floating_anchors(parts: dict[str, bytes], images: list[FloatingImage]) -> None:
    grouped: dict[str, list[FloatingImage]] = {}
    for item in images:
        grouped.setdefault(item.drawing_path, []).append(item)
    for drawing_path, drawing_images in grouped.items():
        drawing = _xml(parts[drawing_path], drawing_path)
        for item in sorted(drawing_images, key=lambda value: value.anchor_index, reverse=True):
            anchors = list(drawing)
            if 0 <= item.anchor_index < len(anchors):
                drawing.remove(anchors[item.anchor_index])
        rels_path = _rels_path(drawing_path)
        if not list(drawing):
            sheet_path = drawing_images[0].sheet_path
            sheet = _xml(parts[sheet_path], sheet_path)
            sheet_rels_path = _rels_path(sheet_path)
            sheet_rels = _xml(parts[sheet_rels_path], sheet_rels_path)
            removed_ids = set()
            for rel in list(sheet_rels):
                if rel.attrib.get("Type") == REL_DRAWING and _resolve(sheet_path, rel.attrib.get("Target", "")) == drawing_path:
                    removed_ids.add(rel.attrib.get("Id", ""))
                    sheet_rels.remove(rel)
            for node in list(sheet):
                if node.tag == f"{{{NS['main']}}}drawing" and node.attrib.get(f"{{{NS['r']}}}id") in removed_ids:
                    sheet.remove(node)
            parts[sheet_path] = _serialize(sheet)
            parts[sheet_rels_path] = _serialize(sheet_rels)
            parts.pop(drawing_path, None)
            parts.pop(rels_path, None)
            types = _xml(parts["[Content_Types].xml"], "[Content_Types].xml")
            for node in list(types):
                if node.attrib.get("PartName", "").lstrip("/") == drawing_path:
                    types.remove(node)
            parts["[Content_Types].xml"] = _serialize(types)
            continue
        parts[drawing_path] = _serialize(drawing)
        if rels_path not in parts:
            continue
        rels = _xml(parts[rels_path], rels_path)
        remaining_ids = {
            node.attrib.get(f"{{{NS['r']}}}embed", "")
            for node in drawing.findall(".//a:blip", NS)
        }
        for rel in list(rels):
            if rel.attrib.get("Type") == REL_IMAGE and rel.attrib.get("Id") not in remaining_ids:
                rels.remove(rel)
        parts[rels_path] = _serialize(rels)


def _write_wps_native(parts: dict[str, bytes], images: list[CellImage]) -> None:
    _remove_native_indexes(parts, {"excel"})
    root = ET.Element(f"{{{NS['wps']}}}cellImages")
    rels = ET.Element(f"{{{NS['pr']}}}Relationships")
    for index, item in enumerate(images):
        digest = hashlib.sha256(parts[item.media_path]).hexdigest().upper()
        image_id = "ID_" + digest[:32]
        container = ET.SubElement(root, f"{{{NS['wps']}}}cellImage")
        pic = ET.SubElement(container, f"{{{NS['xdr']}}}pic")
        nv = ET.SubElement(pic, f"{{{NS['xdr']}}}nvPicPr")
        ET.SubElement(nv, f"{{{NS['xdr']}}}cNvPr", {"id": str(index + 2), "name": image_id, "descr": digest.lower()[:32]})
        pic_props = ET.SubElement(nv, f"{{{NS['xdr']}}}cNvPicPr")
        ET.SubElement(pic_props, f"{{{NS['a']}}}picLocks", {"noChangeAspect": "1"})
        fill = ET.SubElement(pic, f"{{{NS['xdr']}}}blipFill")
        rid = f"rId{index + 1}"
        ET.SubElement(fill, f"{{{NS['a']}}}blip", {f"{{{NS['r']}}}embed": rid})
        stretch = ET.SubElement(fill, f"{{{NS['a']}}}stretch")
        ET.SubElement(stretch, f"{{{NS['a']}}}fillRect")
        shape = ET.SubElement(pic, f"{{{NS['xdr']}}}spPr")
        transform = ET.SubElement(shape, f"{{{NS['a']}}}xfrm")
        width, height = _image_size(parts[item.media_path], item.content_type)
        ET.SubElement(transform, f"{{{NS['a']}}}off", {"x": "0", "y": "0"})
        ET.SubElement(transform, f"{{{NS['a']}}}ext", {"cx": str(width * 9525), "cy": str(height * 9525)})
        geom = ET.SubElement(shape, f"{{{NS['a']}}}prstGeom", {"prst": "rect"})
        ET.SubElement(geom, f"{{{NS['a']}}}avLst")
        ET.SubElement(rels, f"{{{NS['pr']}}}Relationship", {
            "Id": rid, "Type": REL_IMAGE,
            "Target": posixpath.relpath(item.media_path, "xl"),
        })
        _set_cell_wps_native(parts, item, image_id)
    parts["xl/cellimages.xml"] = _serialize(root)
    parts["xl/_rels/cellimages.xml.rels"] = _serialize(rels)
    _add_workbook_rel(parts, WPS_CELL_IMAGE_REL, "cellimages.xml")
    _add_content_override(parts, "xl/cellimages.xml", "application/vnd.wps-officedocument.cellimage+xml")


def _write(parts: dict[str, bytes], destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    temp = destination.with_name(destination.name + ".tmp")
    try:
        with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, data in parts.items():
                archive.writestr(name, data)
        with zipfile.ZipFile(temp) as archive:
            bad = archive.testzip()
            if bad:
                raise InvalidWorkbookError(f"Generated corrupt ZIP member: {bad}")
        temp.replace(destination)
    finally:
        temp.unlink(missing_ok=True)


def make_compatible_copy(
    source: str | Path,
    destination: str | Path,
    margin: int = 0,
    sheet: str | None = None,
    cell_range: str | None = None,
) -> list[CellImage]:
    source_path, destination_path = Path(source), Path(destination)
    if source_path.resolve() == destination_path.resolve():
        raise ValueError("Refusing to overwrite the source workbook")
    parts = _load(source_path)
    all_images = _inspect_parts(parts)
    if not all_images:
        raise UnsupportedWorkbookError("No supported Excel or WPS cell images found")
    images = _select_images(all_images, sheet, cell_range)
    if not images:
        raise UnsupportedWorkbookError("No supported cell images found in the requested scope")
    grouped: dict[str, list[CellImage]] = {}
    for image in images:
        grouped.setdefault(image.sheet_path, []).append(image)
    existing = [int(m.group(1)) for name in parts for m in [re.match(r"xl/drawings/drawing(\d+)\.xml$", name)] if m]
    drawing_no = max(existing, default=0) + 1
    for sheet_path, sheet_images in grouped.items():
        _convert_sheet(parts, sheet_path, sheet_images, margin, drawing_no)
        drawing_no += 1
    fully_converted_formats = {
        source_format for source_format in {item.source_format for item in images}
        if sum(i.source_format == source_format for i in images) == sum(i.source_format == source_format for i in all_images)
    }
    _remove_native_indexes(parts, fully_converted_formats)
    _write(parts, destination_path)
    return images


def convert_native(source: str | Path, destination: str | Path, target: str) -> list[CellImage]:
    source_path, destination_path = Path(source), Path(destination)
    if source_path.resolve() == destination_path.resolve():
        raise ValueError("Refusing to overwrite the source workbook")
    if target not in {"excel", "wps"}:
        raise ValueError("target must be 'excel' or 'wps'")
    parts = _load(source_path)
    images = _inspect_parts(parts)
    if not images:
        raise UnsupportedWorkbookError("No supported Excel or WPS cell images found")
    if target == "excel":
        _write_excel_native(parts, images)
    else:
        _write_wps_native(parts, images)
    _write(parts, destination_path)
    return images


def convert_floating_to_native(
    source: str | Path,
    destination: str | Path,
    target: str,
    sheet: str,
    cell_range: str,
) -> list[FloatingImage]:
    if not sheet or not cell_range:
        raise ValueError("Floating-to-cell conversion requires both sheet and cell range")
    if target not in {"excel", "wps"}:
        raise ValueError("target must be 'excel' or 'wps'")
    source_path, destination_path = Path(source), Path(destination)
    if source_path.resolve() == destination_path.resolve():
        raise ValueError("Refusing to overwrite the source workbook")
    parts = _load(source_path)
    sheets = _workbook_sheets(parts)
    types = _content_types(parts)
    floating = _detect_floating(parts, sheets, types)
    selected = _select_images(floating, sheet, cell_range)
    if not selected:
        raise UnsupportedWorkbookError("No floating images have center points inside the selected cells")
    existing = _detect_excel(parts, sheets, types) if target == "excel" else _detect_wps(parts, sheets, types)
    target_images: list[CellImage] = [*existing, *selected]
    if target == "excel":
        _write_excel_native(parts, target_images)
    else:
        _write_wps_native(parts, target_images)
    _remove_floating_anchors(parts, selected)
    _write(parts, destination_path)
    return selected


def inspection_json(path: str | Path) -> str:
    images = inspect_workbook(path)
    return json.dumps({"file": str(path), "count": len(images), "images": [i.as_dict() for i in images]}, ensure_ascii=False, indent=2)
