import base64
import importlib.util
import io
import json
import tempfile
import threading
import unittest
import zipfile
from http.client import HTTPConnection
from pathlib import Path
from http.server import ThreadingHTTPServer

from fixtures.build_wps import build
from cell_image_compat import convert_native

spec = importlib.util.spec_from_file_location('plugin_service', Path(__file__).parents[1] / 'plugins/service.py')
service = importlib.util.module_from_spec(spec)
spec.loader.exec_module(service)


class ServiceTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.server = ThreadingHTTPServer(('127.0.0.1', 0), service.Handler)
        cls.worker = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.worker.start()

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        cls.worker.join()

    def setUp(self):
        with tempfile.TemporaryDirectory() as folder:
            source = build(Path(folder) / 'sample.xlsx')
            self.data = source.read_bytes()
            excel = Path(folder) / 'excel.xlsx'
            convert_native(source, excel, 'excel')
            self.excel_data = excel.read_bytes()

    def request(self, operation='inspect', **overrides):
        payload = dict(operation=operation, sourceName='sample.xlsx', workbookBase64=base64.b64encode(self.data).decode())
        payload.update(overrides)
        connection = HTTPConnection('127.0.0.1', self.server.server_port)
        self.addCleanup(connection.close)
        connection.request('POST', '/api/convert', json.dumps(payload), {'Content-Type': 'application/json'})
        response = connection.getresponse()
        return response.status, json.loads(response.read())

    def test_inspect_then_repair_download(self):
        status, result = self.request()
        self.assertEqual(status, 200)
        self.assertEqual(result['count'], 3)
        self.assertEqual((result['wpsCount'], result['excelCount']), (3, 0))
        self.assertNotIn('workbookBase64', result)
        status, result = self.request('compatible')
        self.assertEqual(status, 200)
        self.assertEqual(result['converted'], 3)
        self.assertEqual((result['convertedWps'], result['convertedExcel']), (3, 0))
        self.assertEqual(result['filename'], 'sample_兼容版.xlsx')
        with zipfile.ZipFile(io.BytesIO(base64.b64decode(result['workbookBase64']))) as archive:
            self.assertIsNone(archive.testzip())
            self.assertNotIn('xl/cellimages.xml', archive.namelist())
            self.assertIn(b'twoCellAnchor', archive.read('xl/drawings/drawing1.xml'))

    def test_excel_place_in_cell_inspect_and_convert(self):
        status, result = self.request(workbookBase64=base64.b64encode(self.excel_data).decode())
        self.assertEqual(status, 200)
        self.assertEqual((result['count'], result['wpsCount'], result['excelCount']), (3, 0, 3))
        status, result = self.request('compatible', workbookBase64=base64.b64encode(self.excel_data).decode())
        self.assertEqual(status, 200)
        self.assertEqual((result['converted'], result['convertedWps'], result['convertedExcel']), (3, 0, 3))
        with zipfile.ZipFile(io.BytesIO(base64.b64decode(result['workbookBase64']))) as archive:
            self.assertIsNone(archive.testzip())
            self.assertNotIn('xl/metadata.xml', archive.namelist())
            self.assertIn(b'twoCellAnchor', archive.read('xl/drawings/drawing1.xml'))

    def test_zero_images_and_repeat_repair(self):
        _, converted = self.request('compatible')
        encoded = converted['workbookBase64']
        status, result = self.request(workbookBase64=encoded)
        self.assertEqual((status, result['count']), (200, 0))
        status, result = self.request('compatible', workbookBase64=encoded)
        self.assertEqual(status, 400)
        self.assertIn('没有检测到', result['error'])

    def test_bad_data_and_legacy_operations(self):
        for overrides in [dict(workbookBase64='!'), dict(sourceName='book.xlsm'), dict(operation='floating-native'), dict(workbookBase64='', sourcePath='C:/private.xlsx')]:
            status, result = self.request(**overrides)
            self.assertEqual(status, 400)
            self.assertIn('error', result)

    def test_other_origin_is_rejected(self):
        connection = HTTPConnection('127.0.0.1', self.server.server_port)
        self.addCleanup(connection.close)
        connection.request('POST', '/api/convert', '{}', {'Content-Type': 'application/json', 'Origin': 'https://example.com'})
        response = connection.getresponse()
        self.assertEqual(response.status, 403)
        response.read()
