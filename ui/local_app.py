"""独立本地入口：python -m ui.local_app，仅监听 127.0.0.1。"""
from __future__ import annotations

import argparse
import base64
import binascii
import json
import logging
import secrets
import tempfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from threading import Lock
from urllib.parse import urlsplit

from examples.ui_demo import create_demo
from excel_qc import profile_workbook
from ui.workflows import ROLES, WorkflowError, WorkflowSession, display_value, read_json, user_message

ASSETS = Path(__file__).with_name("web")
MAX_REQUEST = 48 * 1024 * 1024


class LocalServer(ThreadingHTTPServer):
    daemon_threads = True

    def __init__(self, port=8765, root=None):
        super().__init__(("127.0.0.1", port), Handler)
        self.root = Path(root) if root is not None else Path(tempfile.mkdtemp(prefix="e-master-ui-"))
        self.root.mkdir(parents=True, exist_ok=True)
        self.sessions = {}
        self.session_lock = Lock()

    def new_session(self):
        token = secrets.token_urlsafe(32)
        with self.session_lock:
            self.sessions[token] = {flow: WorkflowSession(self.root / token / flow) for flow in ROLES}
        return token


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
        if filename:
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.end_headers()
        self.wfile.write(data)

    def send_json(self, status, value):
        self.send_bytes(status, json.dumps(value, ensure_ascii=False, allow_nan=False).encode("utf-8"),
                        "application/json; charset=utf-8")

    def local_request(self):
        address = f"127.0.0.1:{self.server.server_port}"
        if self.headers.get("Host") not in {address, f"localhost:{self.server.server_port}"}:
            raise WorkflowError("请通过本机回环地址访问")
        origin = self.headers.get("Origin")
        if origin and origin not in {f"http://{address}", f"http://localhost:{self.server.server_port}"}:
            raise WorkflowError("不接受其他网页发起的请求")

    def do_GET(self):
        try:
            self.local_request()
            path = urlsplit(self.path).path
            if path == "/":
                token = self.server.new_session()
                content = (ASSETS / "index.html").read_text(encoding="utf-8").replace("__SESSION_VALUE__", token)
                self.send_bytes(200, content.encode("utf-8"), "text/html; charset=utf-8")
            elif path in {"/app.js", "/style.css"}:
                content_type = "text/javascript; charset=utf-8" if path.endswith(".js") else "text/css; charset=utf-8"
                self.send_bytes(200, (ASSETS / path[1:]).read_bytes(), content_type)
            else:
                self.send_json(404, {"error": "页面不存在"})
        except WorkflowError as exc:
            self.send_json(403, {"error": str(exc)})

    def do_POST(self):
        session = None
        acquired = False
        try:
            self.local_request()
            token = self.headers.get("X-Session-Token", "")
            sessions = self.server.sessions.get(token)
            if sessions is None:
                self.send_json(403, {"error": "会话已失效，请刷新页面并重新选择文件"})
                return
            size = int(self.headers.get("Content-Length", "0"))
            if not 0 < size <= MAX_REQUEST:
                raise WorkflowError("请求过大或为空，单个文件请控制在 32 MB 以内")
            payload = json.loads(self.rfile.read(size))
            if not isinstance(payload, dict) or payload.get("flow") not in ROLES:
                raise WorkflowError("操作流程无效")
            flow = payload["flow"]
            session = sessions[flow]
            acquired = session.lock.acquire(blocking=False)
            if not acquired:
                self.send_json(409, {"error": "当前流程正在处理，请等待完成"})
                return
            action = urlsplit(self.path).path
            if action == "/api/upload":
                session.clear_input(flow, payload["role"])
                try:
                    data = base64.b64decode(payload["data"], validate=True)
                except (binascii.Error, ValueError, TypeError) as exc:
                    raise WorkflowError("上传文件内容无效，请重新选择文件") from exc
                response = session.upload(flow, payload["role"], payload["name"], data)
            elif action == "/api/settings":
                session.configure(payload["settings"])
                response = {"revision": session.revision}
            elif action == "/api/invalidate":
                if "role" in payload:
                    session.clear_input(flow, payload["role"])
                else:
                    session.invalidate()
                response = {"revision": session.revision}
            elif action == "/api/demo":
                session.invalidate()
                session.files = {}
                session.settings = {}
                demo = create_demo(session.root / "demo")
                if payload.get("fail") and flow == "prepare":
                    demo[flow]["source"] = demo["fail_source"]
                session.files = demo[flow]
                session.settings = demo["settings"] if flow == "prepare" else {}
                response = {"revision": session.revision, "settings": session.settings, "files": {}}
                for role, path in session.files.items():
                    preview = display_value(profile_workbook(path)) if path.suffix == ".xlsx" else read_json(path)
                    response["files"][role] = {"name": path.name, "preview": preview}
            elif action == "/api/run":
                if flow == "verify" and payload.get("original_confirmed") is not True:
                    raise WorkflowError("请确认回源基准是最初收集的原始源文档")
                response = session.run(flow, payload.get("revision"))
            elif action == "/api/export":
                output = session.export(flow, payload.get("revision"))
                self.send_bytes(200, output.read_bytes(),
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", output.name)
                return
            else:
                self.send_json(404, {"error": "操作不存在"})
                return
            self.send_json(200, response)
        except Exception as exc:
            # 领域错误直接给中文提示，非预期异常仅在维护终端记录堆栈。
            message = user_message(exc)
            if message.startswith("处理失败"):
                logging.exception("P1-12 本地操作失败")
            self.send_json(400, {"error": message, "revision": session.revision if session else None})
        finally:
            if acquired:
                session.lock.release()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--port", type=int, default=8765)
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    try:
        server = LocalServer(args.port)
    except OSError as exc:
        parser.exit(1, f"本地服务启动失败：{exc}。可用 --port 8766 更换端口。\n")
    print(f"E-Master UI: http://127.0.0.1:{server.server_port}", flush=True)
    print(f"Session files retained at: {server.root}", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
