import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from fixtures.build_wps import build
from cell_image_compat import inspect_workbook, make_compatible_copy
from cell_image_compat.cli import main
from cell_image_compat.core import NS, _marker_point, _fit_contain
from cell_image_compat.model import InvalidWorkbookError, UnsupportedWorkbookError


class V1Tests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source = build(self.root / 'wps.xlsx')
        self.output = self.root / 'compatible.xlsx'

    def test_end_to_end_merged_and_repeated_images(self):
        original = self.source.read_bytes()
        self.assertEqual([i.cell for i in inspect_workbook(self.source)], ['A1', 'B2', 'D5'])
        self.assertEqual(len(make_compatible_copy(self.source, self.output)), 3)
        self.assertEqual(self.source.read_bytes(), original)
        with zipfile.ZipFile(self.source) as before, zipfile.ZipFile(self.output) as after:
            self.assertIsNone(after.testzip())
            for name in ['xl/media/wide.png', 'xl/media/tall.png', 'custom/preserve.bin']:
                self.assertEqual(before.read(name), after.read(name))
            sheet = ET.fromstring(after.read('xl/worksheets/sheet1.xml'))
            self.assertEqual(sheet.find('main:mergeCells/main:mergeCell', NS).get('ref'), 'B2:C3')
            self.assertNotIn(b'DISPIMG', after.read('xl/worksheets/sheet1.xml').upper())
            self.assertNotIn('xl/cellimages.xml', after.namelist())
            drawing = ET.fromstring(after.read('xl/drawings/drawing1.xml'))
            anchors = drawing.findall('xdr:twoCellAnchor', NS)
            self.assertEqual(len(anchors), 3)
            # Expected physical regions in px, including 2px default padding.
            for anchor, box, ratio in zip(anchors, [(0, 0, 59, 80), (59, 80, 280, 100), (339, 200, 59, 80)], [4, .25, 4]):
                self.assertEqual(anchor.get('editAs'), 'twoCell')
                x1, y1 = _marker_point(sheet, anchor.find('xdr:from', NS))
                x2, y2 = _marker_point(sheet, anchor.find('xdr:to', NS))
                left, top, width, height = [v * 9525 for v in box]
                self.assertGreaterEqual(x1, left + 2 * 9525)
                self.assertGreaterEqual(y1, top + 2 * 9525)
                self.assertLessEqual(x2, left + width - 2 * 9525)
                self.assertLessEqual(y2, top + height - 2 * 9525)
                self.assertAlmostEqual((x2 - x1) / (y2 - y1), ratio, places=4)
                self.assertLessEqual(abs(x1 + x2 - (2 * left + width)), 1)
                self.assertLessEqual(abs(y1 + y2 - (2 * top + height)), 1)
                self.assertEqual(anchor.find('.//a:picLocks', NS).get('noChangeAspect'), '1')
        self.assertEqual(inspect_workbook(self.output), [])

    def test_partial_then_remaining_appends_without_duplicates(self):
        make_compatible_copy(self.source, self.output, sheet='Sheet1', cell_range='A1')
        self.assertEqual([i.cell for i in inspect_workbook(self.output)], ['B2', 'D5'])
        final = self.root / 'final.xlsx'
        make_compatible_copy(self.output, final)
        with zipfile.ZipFile(final) as archive:
            drawing = ET.fromstring(archive.read('xl/drawings/drawing1.xml'))
            ids = [n.get('id') for n in drawing.findall('.//xdr:cNvPr', NS)]
            self.assertEqual(len(ids), 3)
            self.assertEqual(len(set(ids)), 3)
        with self.assertRaises(UnsupportedWorkbookError):
            make_compatible_copy(final, self.root / 'again.xlsx')

    def test_missing_media_does_not_write_output(self):
        build(self.source, missing=True)
        with self.assertRaisesRegex(InvalidWorkbookError, 'Missing image tall'):
            make_compatible_copy(self.source, self.output)
        self.assertFalse(self.output.exists())

    def test_nested_formula_is_not_silently_replaced(self):
        build(self.source, formula='IF(1,DISPIMG("wide",1),0)')
        with self.assertRaisesRegex(InvalidWorkbookError, 'Unsupported DISPIMG'):
            make_compatible_copy(self.source, self.output)
        self.assertFalse(self.output.exists())

    def test_source_overwrite_and_invalid_margin(self):
        with self.assertRaises(ValueError):
            make_compatible_copy(self.source, self.source)
        with self.assertRaises(ValueError):
            make_compatible_copy(self.source, self.output, margin=-1)
        self.assertFalse(self.output.exists())
        with self.assertRaises(InvalidWorkbookError):
            _fit_contain(100, 100, 0, 10, 2)

    def test_cli(self):
        self.assertEqual(main(['compatible', str(self.source), str(self.output)]), 0)
        self.assertTrue(self.output.exists())
