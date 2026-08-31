# -*- coding: utf-8 -*-
"""打包 exe 与本地运行的统一入口。"""

from __future__ import annotations

import sys
import threading
import webbrowser

from app_paths import data_dir, ensure_user_config, is_frozen


def main() -> None:
    ensure_user_config()
    data_dir().joinpath("imgs").mkdir(exist_ok=True)

    from web_app import app, load_config

    cfg = load_config()
    port = int(cfg.get("web_port", 5050))
    url = f"http://127.0.0.1:{port}"

    print("=" * 50)
    print("HubStudio 批量控制台")
    if is_frozen():
        print(f"程序目录: {data_dir()}")
    print(f"请在浏览器打开: {url}")
    print("关闭本窗口将停止服务。")
    print("=" * 50)

    threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    app.run(host="127.0.0.1", port=port, debug=False, threaded=True)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(0)
