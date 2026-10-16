"""python -m ui.local_app：只监听本机的设备台账转换工具。"""
from __future__ import annotations

import argparse
import json
import logging
import secrets
import tempfile
from dataclasses import asdict
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from threading import Lock, BoundedSemaphore
from urllib.parse import parse_qs, quote, urlsplit

from excel_qc import WorkbookError, convert_workbook
from excel_qc.rules import DEFAULT_MAPPING, MappingError

ASSETS = Path(__file__).with_name("web")
MAX_UPLOAD = 32 * 1024 * 1024
XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


class LocalServer(ThreadingHTTPServer):
    daemon_threads = True

    def __init__(self, port=8765, *, mapping_path=DEFAULT_MAPPING):
        self.outputs = {}
        self.output_lock = Lock()
        self.processing = BoundedSemaphore(2)
        self.mapping_path = Path(mapping_path)
        self.temp = tempfile.TemporaryDirectory(prefix="e-master-converter-")
        super().__init__(("127.0.0.1", port), Handler)

    def server_close(self):
        super().server_close()
        self.temp.cleanup()


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        logging.info(fmt, *args)

    def send_bytes(self, status, data, content_type, filename=None):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header("Content-Security-Policy", "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'; base-uri 'none'")
        if filename:
            self.send_header("Content-Disposition", f"attachment; filename=\"result.xlsx\"; filename*=UTF-8''{quote(filename)}")
        self.end_headers()
        self.wfile.write(data)

    def send_json(self, status, value):
        self.send_bytes(status, json.dumps(value, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8")

    def local_request(self):
        allowed = {f"127.0.0.1:{self.server.server_port}", f"localhost:{self.server.server_port}"}
        if self.headers.get("Host") not in allowed:
            raise PermissionError("请通过本机地址访问")
        if self.headers.get("Origin") and self.headers["Origin"] not in {"http://" + host for host in allowed}:
            raise PermissionError("不接受其他网页发起的请求")

    def do_GET(self):
        try:
            self.local_request()
            path = urlsplit(self.path).path
            if path in {"/", "/app.js", "/style.css"}:
                asset = ASSETS / ("index.html" if path == "/" else path[1:])
                content_type = {".html": "text/html", ".js": "text/javascript", ".css": "text/css"}[asset.suffix]
                self.send_bytes(200, asset.read_bytes(), content_type + "; charset=utf-8")
            elif path.startswith("/api/download/"):
                token = path.removeprefix("/api/download/")
                with self.server.output_lock:
                    output = self.server.outputs.get(token)
                    if output is None:
                        self.send_json(404, {"error": "处理结果已失效，请重新处理"})
                        return
                    data = output[0].read_bytes()
                self.send_bytes(200, data, XLSX_TYPE, output[1])
            elif path == "/favicon.ico":
                self.send_bytes(204, b"", "image/x-icon")
            else:
                self.send_json(404, {"error": "页面不存在"})
        except PermissionError as exc:
            self.send_json(403, {"error": str(exc)})

    def do_POST(self):
        acquired = False
        try:
            self.local_request()
            url = urlsplit(self.path)
            if url.path != "/api/convert":
                self.send_json(404, {"error": "操作不存在"})
                return
            length = int(self.headers.get("Content-Length", "0"))
            if not 0 < length <= MAX_UPLOAD:
                raise WorkbookError("请选择 32 MB 以内的 .xlsx 文件")
            filename = parse_qs(url.query).get("filename", [""])[0].replace("\\", "/").rsplit("/", 1)[-1]
            if not filename or Path(filename).suffix.lower() != ".xlsx":
                raise WorkbookError("仅支持 .xlsx 文件")
            acquired = self.server.processing.acquire(blocking=False)
            if not acquired:
                self.send_json(409, {"error": "工具正在处理其他文件，请稍后重试"})
                return
            self.connection.settimeout(60)
            data = self.rfile.read(length)
            if len(data) != length:
                raise WorkbookError("文件上传不完整，请重新选择文件")
            result = convert_workbook(data, self.server.mapping_path)
            response = {
                "success": result.success,
                "issues": [asdict(issue) for issue in result.issues],
                "sheets": result.sheets,
                "skipped_sheets": result.skipped_sheets,
                "rows": result.rows,
            }
            if result.success:
                token = secrets.token_urlsafe(24)
                output = Path(self.server.temp.name) / (token + ".xlsx")
                result.save(output)
                output_name = Path(filename).stem + "_处理结果.xlsx"
                with self.server.output_lock:
                    # 保持本地临时目录有界，移除最早的下载副本。
                    if len(self.server.outputs) >= 32:
                        oldest = next(iter(self.server.outputs))
                        self.server.outputs.pop(oldest)[0].unlink(missing_ok=True)
                    self.server.outputs[token] = (output, output_name)
                response.update({
                    "download_url": "/api/download/" + token,
                    "filename": output_name,
                    "subcategory_filled": result.subcategory_filled,
                    "dates_converted": result.dates_converted,
                    "dates_cleared": result.dates_cleared,
                    "codes_formatted": result.codes_formatted,
                })
            self.send_json(200, response)
        except PermissionError as exc:
            self.send_json(403, {"error": str(exc)})
        except (WorkbookError, MappingError, ValueError, TimeoutError) as exc:
            self.send_json(400, {"error": str(exc)})
        except Exception:
            logging.exception("Excel 转换失败")
            self.send_json(500, {"error": "文件处理失败，请检查文件后重新执行"})
        finally:
            if acquired:
                self.server.processing.release()


def main():
    parser = argparse.ArgumentParser(description="E-Master 设备台账 Excel 格式转换工具")
    parser.add_argument("--port", type=int, default=8765)
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    with LocalServer(args.port) as server:
        print(f"E-Master 已启动：http://127.0.0.1:{server.server_port}", flush=True)
        print("按 Ctrl+C 停止。原始 Excel 始终只读。", flush=True)
        try:
            server.serve_forever()
        except KeyboardInterrupt:
            pass


if __name__ == "__main__":
    main()
