#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
============================================================
  敏感词遮蔽工具 - 本地服务器启动脚本
  功能：启动一个简单的 HTTP 服务器，解决浏览器沙盒限制，
        使前端能通过 fetch 读取同目录下的 ban_word.json
============================================================
使用方法：
  python start_server.py
  或直接运行此文件。
============================================================
"""

import http.server
import socketserver
import webbrowser
import os
import sys
from pathlib import Path

# ============================================================
#  宏定义区域（请根据需求修改以下变量）
# ============================================================

# 服务器监听主机地址（0.0.0.0 表示允许局域网内其他设备访问）
HOST = "127.0.0.1"

# 服务器监听端口（若被占用可更换，如 8080, 8888 等）
PORT = 8000

# 静态文件根目录（默认使用脚本所在目录，也可指定绝对路径）
# 留空或 "." 表示脚本所在目录
WEB_DIR = "."

# 是否在启动后自动打开默认浏览器访问页面
AUTO_OPEN_BROWSER = True

# 页面文件名（位于 WEB_DIR 下的 HTML 文件）
INDEX_FILE = "demo5_屏蔽词.html"

# ============================================================
#  以下为服务器实现，一般无需修改
# ============================================================

class CustomHandler(http.server.SimpleHTTPRequestHandler):
    """自定义请求处理器，可在此添加 CORS 头等扩展"""
    def __init__(self, *args, **kwargs):
        # 设置静态文件目录
        super().__init__(*args, directory=str(Path(WEB_DIR).resolve()), **kwargs)

    def end_headers(self):
        # 添加 CORS 头（可选，解决跨域问题）
        self.send_header("Access-Control-Allow-Origin", "*")
        super().end_headers()

    def log_message(self, format, *args):
        # 自定义日志输出格式，更简洁
        sys.stdout.write(f"[访问] {self.address_string()} - {format % args}\n")

def main():
    # 切换到脚本所在目录，保证相对路径正确
    script_dir = Path(__file__).parent.resolve()
    os.chdir(script_dir)

    # 解析最终的 Web 目录
    web_dir = Path(WEB_DIR).resolve()
    if not web_dir.exists() or not web_dir.is_dir():
        print(f"❌ 错误：指定的静态文件目录不存在 -> {web_dir}")
        sys.exit(1)

    # 检查索引文件是否存在
    index_path = web_dir / INDEX_FILE
    if not index_path.exists():
        print(f"⚠️  警告：索引文件不存在 -> {index_path}，服务器仍会启动，但访问根路径可能 404")
    else:
        print(f"📄 索引文件: {index_path}")

    # 创建 TCP 服务器
    try:
        with socketserver.TCPServer((HOST, PORT), CustomHandler) as httpd:
            server_url = f"http://{HOST}:{PORT}"
            print("=" * 60)
            print(f"  🚀 本地服务器已启动")
            print(f"  地址: {server_url}")
            print(f"  目录: {web_dir}")
            print(f"  按 Ctrl+C 停止服务器")
            print("=" * 60)

            if AUTO_OPEN_BROWSER:
                # 打开浏览器
                webbrowser.open(f"{server_url}/{INDEX_FILE}" if INDEX_FILE != "index.html" else server_url)

            # 开始监听请求
            httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n\n🛑 服务器已安全停止。")
    except OSError as e:
        if e.errno == 48 or "Address already in use" in str(e):
            print(f"❌ 端口 {PORT} 已被占用，请更换宏定义中的 PORT 值后重试。")
        else:
            print(f"❌ 启动失败: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()