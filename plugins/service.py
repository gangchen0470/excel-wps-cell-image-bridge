from __future__ import annotations

import argparse
import base64
import json
import tempfile
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

from cell_image_compat import inspect_workbook, make_compatible_copy

ROOT = Path(__file__).parent / "shared"
MAX_BODY = 40 * 1024 * 1024
MAX_WORKBOOK_SIZE = 25 * 1024 * 1024


def process_request(request):
    """Process uploaded bytes only; inspection never modifies the workbook."""
    if not isinstance(request, dict):
        raise ValueError("请求格式不正确")
    operation = request.get("operation")
    if operation not in {"inspect", "compatible"}:
        raise ValueError("不支持的操作")
    name = request.get("sourceName", "workbook.xlsx")
    if not isinstance(name, str) or Path(name).suffix.lower() != ".xlsx":
        raise ValueError("请选择 .xlsx 文件")
    encoded = request.get("workbookBase64")
    if not isinstance(encoded, str) or not encoded:
        raise ValueError("请先选择工作簿文件")
    try:
        data = base64.b64decode(encoded, validate=True)
    except Exception as exc:
        raise ValueError("工作簿数据格式不正确") from exc
    if len(data) > MAX_WORKBOOK_SIZE:
        raise ValueError("文件超过 25 MB，暂不支持网页处理")
    with tempfile.TemporaryDirectory(prefix="cell-image-plugin-") as directory:
        source = Path(directory) / "source.xlsx"
        output = Path(directory) / "output.xlsx"
        source.write_bytes(data)
        images = inspect_workbook(source)
        wps_images = [image for image in images if image.source_format == "wps"]
        excel_images = [image for image in images if image.source_format == "excel"]
        payload = {
            "count": len(images),
            "wpsCount": len(wps_images),
            "excelCount": len(excel_images),
            "images": [image.as_dict() for image in images],
        }
        if operation == "compatible":
            if not images:
                raise ValueError("没有检测到可转换的 WPS DISPIMG 或 Excel Place in Cell 图片")
            converted = make_compatible_copy(source, output)
            payload.update(
                converted=len(converted),
                convertedWps=sum(image.source_format == "wps" for image in converted),
                convertedExcel=sum(image.source_format == "excel" for image in converted),
                filename=f"{Path(name).stem}_兼容版.xlsx",
                workbookBase64=base64.b64encode(output.read_bytes()).decode("ascii"),
            )
        return payload


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def do_POST(self):
        if urlparse(self.path).path != "/api/convert":
            self.send_error(404)
            return
        host = self.headers.get("Host", "")
        origin = self.headers.get("Origin")
        allowed_origins = {f"http://{host}", f"https://{host}"}
        if not host or (origin and origin not in allowed_origins):
            self._json(403, {"error": "仅允许同源页面请求"})
            return
        if self.headers.get_content_type() != "application/json":
            self._json(415, {"error": "请使用 JSON 请求"})
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if not 0 < length <= MAX_BODY:
                raise ValueError("请求为空或文件过大")
            payload = process_request(json.loads(self.rfile.read(length)))
            self._json(200, payload)
        except Exception as exc:
            self._json(400, {"error": str(exc)})

    def _json(self, status, payload):
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--host", default="0.0.0.0")
    parser.add_argument("--port", type=int, default=3000)
    args = parser.parse_args()
    print(f"图片兼容工具：http://{args.host}:{args.port}", flush=True)
    ThreadingHTTPServer((args.host, args.port), Handler).serve_forever()


if __name__ == "__main__":
    main()
