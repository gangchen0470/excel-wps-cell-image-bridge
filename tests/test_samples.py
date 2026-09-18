from __future__ import annotations

import os
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from cell_image_compat import convert_floating_to_native, convert_native, inspect_workbook, make_compatible_copy
from cell_image_compat.core import _fit_contain, _image_size, _select_images, inspect_floating_images
from cell_image_compat.model import CellImage


class SampleWorkbookTests(unittest.TestCase):
    def test_scope_selection_across_sheets_and_ranges(self):
        images = [
            CellImage("Sheet1", "xl/worksheets/sheet1.xml", "B2", "xl/media/1.png", "excel", "0"),
            CellImage("Sheet1", "xl/worksheets/sheet1.xml", "B500", "xl/media/2.png", "excel", "1"),
            CellImage("Sheet1", "xl/worksheets/sheet1.xml", "C2", "xl/media/3.png", "excel", "2"),
            CellImage("Sheet2", "xl/worksheets/sheet2.xml", "B2", "xl/media/4.png", "excel", "3"),
        ]
        selected = _select_images(images, "Sheet1", "B2:B500")
        self.assertEqual([item.media_path for item in selected], ["xl/media/1.png", "xl/media/2.png"])
        self.assertEqual(len(_select_images(images, None, "B2")), 2)
        with self.assertRaises(ValueError):
            _select_images(images, None, "not-a-range")
    def test_image_dimensions_for_common_formats(self):
        png = b"\x89PNG\r\n\x1a\n" + b"\0" * 8 + (1600).to_bytes(4, "big") + (400).to_bytes(4, "big")
        gif = b"GIF89a" + (400).to_bytes(2, "little") + (1600).to_bytes(2, "little")
        bmp = b"BM" + b"\0" * 16 + (800).to_bytes(4, "little", signed=True) + (600).to_bytes(4, "little", signed=True)
        vp8x = b"RIFF" + b"\0" * 4 + b"WEBPVP8X" + b"\0" * 8 + (999).to_bytes(3, "little") + (499).to_bytes(3, "little")
        jpeg = b"\xff\xd8\xff\xc2\x00\x11\x08" + (1200).to_bytes(2, "big") + (300).to_bytes(2, "big") + b"\0" * 8
        self.assertEqual(_image_size(png, "image/png"), (1600, 400))
        self.assertEqual(_image_size(gif, "image/gif"), (400, 1600))
        self.assertEqual(_image_size(bmp, "image/bmp"), (800, 600))
        self.assertEqual(_image_size(vp8x, "image/webp"), (1000, 500))
        self.assertEqual(_image_size(jpeg, "image/jpeg"), (300, 1200))

    def test_contain_fit_for_wide_tall_and_square_images(self):
        cell_w, cell_h = 1000, 800
        cases = (
            (1600, 400, 1000, 250, 0, 275),
            (400, 1600, 200, 800, 400, 0),
            (1000, 1000, 800, 800, 100, 0),
        )
        for img_w, img_h, expected_w, expected_h, expected_x, expected_y in cases:
            result = _fit_contain(cell_w, cell_h, img_w, img_h, 0)
            self.assertEqual(result, (expected_w, expected_h, expected_x, expected_y))
            draw_w, draw_h, xoff, yoff = result
            self.assertLessEqual(draw_w, cell_w)
            self.assertLessEqual(draw_h, cell_h)
            self.assertLessEqual(xoff + draw_w, cell_w)
            self.assertLessEqual(yoff + draw_h, cell_h)

    def test_contain_fit_applies_margin_without_distortion(self):
        cell_w, cell_h = 1000 * 9525, 800 * 9525
        draw_w, draw_h, xoff, yoff = _fit_contain(cell_w, cell_h, 1600, 400, 10)
        self.assertLess(draw_w, cell_w)
        self.assertEqual(draw_w * 400, draw_h * 1600)
        self.assertEqual(xoff * 2 + draw_w, cell_w)
        self.assertIn(yoff * 2 + draw_h, (cell_h - 1, cell_h))

    def _sample(self, variable: str) -> Path:
        if not os.environ.get(variable):
            self.skipTest(f"Set {variable} to run real-client sample checks")
        path = Path(os.environ[variable])
        self.assertTrue(path.is_file(), path)
        return path

    def test_excel_detection_and_compatible_copy(self):
        source = self._sample("EXCEL_SAMPLE")
        images = inspect_workbook(source)
        self.assertEqual([(i.source_format, i.cell) for i in images], [("excel", "A1")])
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "excel-compatible.xlsx"
            make_compatible_copy(source, output)
            self._assert_floating_output(output)

    def test_wps_detection_and_compatible_copy(self):
        source = self._sample("WPS_SAMPLE")
        images = inspect_workbook(source)
        self.assertEqual([(i.source_format, i.cell) for i in images], [("wps", "A1")])
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "wps-compatible.xlsx"
            make_compatible_copy(source, output)
            self._assert_floating_output(output)

    def _assert_floating_output(self, output: Path):
        self.assertTrue(zipfile.is_zipfile(output))
        with zipfile.ZipFile(output) as archive:
            self.assertIsNone(archive.testzip())
            names = set(archive.namelist())
            self.assertIn("xl/drawings/drawing1.xml", names)
            self.assertIn("xl/drawings/_rels/drawing1.xml.rels", names)
            self.assertNotIn("xl/cellimages.xml", names)
            self.assertNotIn("xl/metadata.xml", names)
            sheet = archive.read("xl/worksheets/sheet1.xml")
            self.assertIn(b"drawing", sheet)
            self.assertNotIn(b"DISPIMG", sheet)
            self.assertNotIn(b'vm="1"', sheet)
            self.assertIn(b'customHeight="1"', sheet)
            self.assertNotIn(b"xmlns:ns", sheet)
            if b"Ignorable" in sheet:
                self.assertIn(b'mc:Ignorable="x14ac xr"', sheet)
            drawing = archive.read("xl/drawings/drawing1.xml")
            self.assertIn(b'twoCellAnchor editAs="twoCell"', drawing)
            self.assertNotIn(b"oneCellAnchor", drawing)
            self.assertIn(b"<a:xfrm>", drawing)
            drawing_root = ET.fromstring(drawing)
            dns = {"xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"}
            extent = drawing_root.find("xdr:twoCellAnchor/xdr:pic/xdr:spPr/a:xfrm/a:ext", {**dns, "a": "http://schemas.openxmlformats.org/drawingml/2006/main"})
            self.assertGreater(int(extent.attrib["cx"]), 0)
            self.assertGreater(int(extent.attrib["cy"]), 0)
            # Default Excel/WPS column width is at least 59 px. This ceiling
            # catches accidental use of the renderer-specific +5 px padding.
            if output.name.startswith("excel-compatible"):
                self.assertLessEqual(int(extent.attrib["cx"]), 59 * 9525)

    def test_original_row_and_column_dimensions_are_preserved(self):
        for variable in ("EXCEL_SAMPLE", "WPS_SAMPLE"):
            source = self._sample(variable)
            with tempfile.TemporaryDirectory() as directory:
                output = Path(directory) / "compatible.xlsx"
                make_compatible_copy(source, output)
                with zipfile.ZipFile(source) as before_zip, zipfile.ZipFile(output) as after_zip:
                    before = ET.fromstring(before_zip.read("xl/worksheets/sheet1.xml"))
                    after = ET.fromstring(after_zip.read("xl/worksheets/sheet1.xml"))
                ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
                before_row = before.find("m:sheetData/m:row", ns)
                after_row = after.find("m:sheetData/m:row", ns)
                self.assertEqual(before_row.attrib.get("ht"), after_row.attrib.get("ht"))
                before_cols = [(n.attrib.get("min"), n.attrib.get("max"), n.attrib.get("width")) for n in before.findall("m:cols/m:col", ns)]
                after_cols = [(n.attrib.get("min"), n.attrib.get("max"), n.attrib.get("width")) for n in after.findall("m:cols/m:col", ns)]
                self.assertEqual(before_cols, after_cols)

    def test_wps_to_excel_native_round_trip_detection(self):
        source = self._sample("WPS_SAMPLE")
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "excel-native.xlsx"
            original_hash = self._media_hash(source)
            convert_native(source, output, "excel")
            images = inspect_workbook(output)
            self.assertEqual([(i.source_format, i.cell) for i in images], [("excel", "A1")])
            self.assertEqual(original_hash, self._media_hash(output))
            self._assert_height_preserved_and_locked(source, output)

    def test_excel_to_wps_native_round_trip_detection(self):
        source = self._sample("EXCEL_SAMPLE")
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "wps-native.xlsx"
            original_hash = self._media_hash(source)
            convert_native(source, output, "wps")
            images = inspect_workbook(output)
            self.assertEqual([(i.source_format, i.cell) for i in images], [("wps", "A1")])
            self.assertEqual(original_hash, self._media_hash(output))
            self._assert_height_preserved_and_locked(source, output)

    def test_floating_to_native_requires_selection_and_uses_center_cell(self):
        source = self._sample("EXCEL_SAMPLE")
        with tempfile.TemporaryDirectory() as directory:
            floating = Path(directory) / "floating.xlsx"
            make_compatible_copy(source, floating)
            self.assertEqual([item.cell for item in inspect_floating_images(floating)], ["A1"])
            with self.assertRaises(ValueError):
                convert_floating_to_native(floating, Path(directory) / "invalid.xlsx", "wps", "", "")
            with self.assertRaises(Exception):
                convert_floating_to_native(floating, Path(directory) / "outside.xlsx", "wps", "Sheet1", "B2:B3")
            output = Path(directory) / "wps-cell.xlsx"
            converted = convert_floating_to_native(floating, output, "wps", "Sheet1", "A1:A1")
            self.assertEqual([item.cell for item in converted], ["A1"])
            self.assertEqual(inspect_floating_images(output), [])
            self.assertEqual([(item.source_format, item.cell) for item in inspect_workbook(output)], [("wps", "A1")])
            with zipfile.ZipFile(output) as archive:
                self.assertNotIn("xl/drawings/drawing1.xml", archive.namelist())
                self.assertNotIn("xl/drawings/_rels/drawing1.xml.rels", archive.namelist())

    def _media_hash(self, workbook: Path) -> str:
        import hashlib
        with zipfile.ZipFile(workbook) as archive:
            media = sorted(name for name in archive.namelist() if name.startswith("xl/media/") and not name.endswith("/"))
            return hashlib.sha256(archive.read(media[0])).hexdigest()

    def _assert_height_preserved_and_locked(self, source: Path, output: Path):
        ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        with zipfile.ZipFile(source) as source_zip, zipfile.ZipFile(output) as output_zip:
            before = ET.fromstring(source_zip.read("xl/worksheets/sheet1.xml"))
            after = ET.fromstring(output_zip.read("xl/worksheets/sheet1.xml"))
        before_row = before.find("m:sheetData/m:row", ns)
        after_row = after.find("m:sheetData/m:row", ns)
        self.assertEqual(before_row.attrib.get("ht"), after_row.attrib.get("ht"))
        self.assertEqual(after_row.attrib.get("customHeight"), "1")


if __name__ == "__main__":
    unittest.main()
