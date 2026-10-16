import io
import json
from http.client import HTTPConnection
from threading import Thread
from urllib.parse import quote

import pytest
from openpyxl import load_workbook

from tests.fixtures import make_mapping, make_workbook
from ui.local_app import LocalServer


@pytest.fixture
def server(tmp_path):
    app = LocalServer(0, mapping_path=make_mapping(tmp_path / "映射.xlsx"))
    thread = Thread(target=app.serve_forever, daemon=True)
    thread.start()
    yield app
    app.shutdown()
    thread.join()
    app.server_close()


def request(server, method, path, data=None, headers=None):
    connection = HTTPConnection("127.0.0.1", server.server_port, timeout=10)
    connection.request(method, path, body=data, headers=headers or {})
    response = connection.getresponse()
    result = response.status, dict(response.getheaders()), response.read()
    connection.close()
    return result


def test_page_has_only_current_product(server):
    status, headers, content = request(server, "GET", "/")
    text = content.decode()
    assert status == 200 and "设备台账数据标准化工具" in text
    assert all(x not in text for x in ["批量", "JSON 配置", "回源", "登录", "工作流", "方案"])
    assert "设备制造日期" in text and "download-button" in text
    assert "Content-Security-Policy" in headers


def test_upload_process_download_readback(server, tmp_path):
    source = make_workbook(tmp_path / "测试台账.xlsx")
    original = source.read_bytes()
    status, _, content = request(server, "POST", "/api/convert?filename=" + quote(source.name), original)
    result = json.loads(content)
    assert status == 200 and result["success"]
    assert result["rows"] == 3 and result["dates_converted"] == 5
    status, headers, content = request(server, "GET", result["download_url"])
    assert status == 200 and "filename*=UTF-8''" in headers["Content-Disposition"]
    book = load_workbook(io.BytesIO(content))
    assert book["设备台账"]["C2"].value == "D-LXB"
    assert book["设备台账"]["F2"].value == "2026-08-01"
    assert book["设备台账"]["G3"].value is None
    book.close()
    assert source.read_bytes() == original


def test_blocked_request_creates_no_result(server, tmp_path):
    source = make_workbook(tmp_path / "问题.xlsx", bad=True)
    status, _, content = request(server, "POST", "/api/convert?filename=bad.xlsx", source.read_bytes())
    result = json.loads(content)
    assert status == 200 and not result["success"] and len(result["issues"]) == 7
    assert "download_url" not in result and not server.outputs
    assert not list(__import__("pathlib").Path(server.temp.name).glob("*.xlsx"))


@pytest.mark.parametrize("filename,data", [("data.xls", b"123"), ("data.xlsx", b"broken"), ("data.xlsx", b"")])
def test_invalid_upload_has_no_download(server, filename, data):
    status, _, content = request(server, "POST", "/api/convert?filename=" + filename, data)
    assert status == 400 and "error" in json.loads(content) and not server.outputs


@pytest.mark.parametrize("path", ["/api/run", "/api/batch/run", "/api/preparation/state", "/api/export"])
def test_obsolete_apis_are_removed(server, path):
    assert request(server, "POST", path, b"{}")[0] == 404


def test_local_only_request(server):
    assert request(server, "GET", "/", headers={"Host": "example.com"})[0] == 403
    assert request(server, "POST", "/api/convert?filename=a.xlsx", b"123", {"Origin": "https://example.com"})[0] == 403


def test_unknown_download(server):
    assert request(server, "GET", "/api/download/nonexistent")[0] == 404
