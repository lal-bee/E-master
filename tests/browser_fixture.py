"""浏览器验收所用合成样本与下载回读；仅写指定临时目录。"""
import argparse
import json
from pathlib import Path

from openpyxl import Workbook

from tests.fixtures import make_workbook, verify_download


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--create", type=Path)
    parser.add_argument("--verify", type=Path)
    parser.add_argument("--serve", action="store_true")
    args = parser.parse_args()
    if args.create:
        folder = args.create
        folder.mkdir(parents=True, exist_ok=True)
        good = make_workbook(folder / "合成台账.xlsx")
        bad = make_workbook(folder / "全部阻断.xlsx", bad=True)
        invalid = folder / "损坏文件.xlsx"
        invalid.write_bytes(b"not an xlsx")
        hostile = folder / "特殊文本.xlsx"
        book = Workbook()
        book.active.append(["设备小类名称", "设备小类编码"])
        book.active.append(['<img src=x onerror=alert(7)>', None])
        book.save(hostile)
        book.close()
        print(json.dumps({"good": str(good), "bad": str(bad), "invalid": str(invalid),
                          "hostile": str(hostile)}, ensure_ascii=True))
    elif args.verify:
        print(json.dumps(verify_download(args.verify)))
    elif args.serve:
        from threading import Thread
        from ui.local_app import LocalServer
        with LocalServer(0) as server:
            thread = Thread(target=server.serve_forever, daemon=True)
            thread.start()
            print(f"http://127.0.0.1:{server.server_port}", flush=True)
            try:
                input()
            finally:
                server.shutdown()
                thread.join()


if __name__ == "__main__":
    main()
