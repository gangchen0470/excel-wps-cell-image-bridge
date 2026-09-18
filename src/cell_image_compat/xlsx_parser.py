"""OOXML ZIP, relationships, content types and worksheet discovery."""
from __future__ import annotations
import posixpath
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from .model import InvalidWorkbookError

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
DISPIMG_RE = re.compile(r'^\s*=?\s*(?:_xlfn\.)?DISPIMG\s*\(\s*"([^"]+)"\s*[,;]\s*1\s*\)\s*$', re.IGNORECASE)
CELL_RE = re.compile(r"^([A-Z]+)([1-9][0-9]*)$")
RANGE_RE = re.compile(r"^([A-Z]+[1-9][0-9]*)(?::([A-Z]+[1-9][0-9]*))?$")


def _xml(data: bytes, name: str) -> ET.Element:
    try:
        return ET.fromstring(data)
    except ET.ParseError as exc:
        raise InvalidWorkbookError(f"Invalid XML in {name}: {exc}") from exc


def _resolve(source_part: str, target: str) -> str:
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target)).lstrip("/")


def _rels_path(part: str) -> str:
    return posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")


def _rel_map(parts: dict[str, bytes], part: str) -> dict[str, tuple[str, str]]:
    path = _rels_path(part)
    if path not in parts:
        return {}
    root = _xml(parts[path], path)
    result = {}
    for rel in root.findall("pr:Relationship", NS):
        if rel.attrib.get("TargetMode") == "External":
            continue
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


