"""Map standalone WPS DISPIMG formulas to embedded media."""
from __future__ import annotations
from .model import CellImage, InvalidWorkbookError
from .xlsx_parser import NS, REL_IMAGE, DISPIMG_RE, _xml, _rel_map

def _detect_wps(parts: dict[str, bytes], sheets: list[tuple[str, str]], types: dict[str, str]) -> list[CellImage]:
    cell_images = "xl/cellimages.xml"
    if cell_images not in parts:
        for _, sheet_path in sheets:
            if b"DISPIMG" in parts[sheet_path].upper():
                raise InvalidWorkbookError("DISPIMG formula found but xl/cellimages.xml is missing")
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
                if "DISPIMG" in formula.upper():
                    raise InvalidWorkbookError(f"Unsupported DISPIMG expression at {sheet_name}!{cell.attrib.get('r')}")
                continue
            image_id = match.group(1)
            media = image_by_id.get(image_id)
            if media not in parts:
                raise InvalidWorkbookError(f"Missing image {image_id} at {sheet_name}!{cell.attrib.get('r')}")
            if media in parts:
                found.append(CellImage(sheet_name, sheet_path, cell.attrib["r"], media, "wps", image_id, types.get(media)))
    return found


