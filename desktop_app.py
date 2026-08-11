from __future__ import annotations

import ctypes
import threading
import traceback
from pathlib import Path

import webview
from werkzeug.serving import make_server

from web_app import app, cleanup_old_jobs


BASE_DIR = Path(__file__).resolve().parent
LOG_PATH = BASE_DIR / "desktop_app.log"


def run() -> None:
    cleanup_old_jobs()
    server = make_server("127.0.0.1", 0, app, threaded=True)
    port = server.server_port
    server_thread = threading.Thread(target=server.serve_forever, daemon=True)
    server_thread.start()

    webview.create_window(
        "流水核对",
        f"http://127.0.0.1:{port}/",
        width=1240,
        height=820,
        min_size=(900, 650),
        background_color="#f2f5f3",
        text_select=False,
    )
    try:
        webview.start(private_mode=True)
    finally:
        server.shutdown()
        server.server_close()


if __name__ == "__main__":
    try:
        run()
    except Exception:
        LOG_PATH.write_text(traceback.format_exc(), encoding="utf-8")
        ctypes.windll.user32.MessageBoxW(
            0,
            f"流水核对软件启动失败。\n错误记录：{LOG_PATH}",
            "流水核对",
            0x10,
        )
