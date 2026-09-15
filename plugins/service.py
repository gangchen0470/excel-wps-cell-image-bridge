from __future__ import annotations

import argparse
import base64
import json
import tempfile
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

from cell_image_compat import convert_floating_to_native, make_compatible_copy


ROOT = Path(__file__).parent / "shared"


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def do_POST(self):
        if urlparse(self.path).path != "/api/convert":
            self.send_error(404)
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            request = json.loads(self.rfile.read(length))
            with tempfile.TemporaryDirectory(prefix="cell-image-plugin-") as directory:
                source = Path(directory) / "source.xlsx"
                output = Path(directory) / "output.xlsx"
                if request.get("workbookBase64"):
                    source.write_bytes(base64.b64decode(request["workbookBase64"], validate=True))
                elif request.get("sourcePath"):
                    source = Path(request["sourcePath"])
                else:
                    raise ValueError("请求中缺少工作簿数据")
                operation = request.get("operation")
                if operation == "compatible":
                    converted = make_compatible_copy(source, output)
                    source_stem = Path(request.get("sourceName") or source.name).stem
                    filename = f"{source_stem}_通用兼容版.xlsx"
                elif operation == "floating-native":
                    target = request.get("target")
                    converted = convert_floating_to_native(source, output, target, request.get("sheet", ""), request.get("range", ""))
                    source_stem = Path(request.get("sourceName") or source.name).stem
                    filename = f"{source_stem}_转{target.upper()}单元格图片.xlsx"
                else:
                    raise ValueError("不支持的转换操作")
                payload = {"converted": len(converted), "filename": filename, "workbookBase64": base64.b64encode(output.read_bytes()).decode("ascii")}
            self._json(200, payload)
        except Exception as exc:
            self._json(400, {"error": str(exc)})

    def _json(self, status, payload):
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, default=3000)
    args = parser.parse_args()
    ThreadingHTTPServer(("127.0.0.1", args.port), Handler).serve_forever()


if __name__ == "__main__":
    main()
