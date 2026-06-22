#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
独立预览服务器
==============
在本地启动 HTTP 服务器，用于预览静态网站（如 index.html）。

用法：
    python preview_server.py [目录] [端口] [首页文件]

参数：
    目录      - 网站根目录，默认为当前目录
    端口      - 监听端口，默认 8800（若被占用则自动递增）
    首页文件  - 浏览器打开的页面，默认为 index.html

示例：
    python preview_server.py ./web 8800 index.html
    python preview_server.py           # 使用所有默认值

按 Ctrl+C 停止服务器。
"""

import os
import sys
import socket
import threading
import webbrowser
from http.server import HTTPServer, SimpleHTTPRequestHandler


def find_available_port(start=8800):
    """从 start 开始寻找一个可用的 TCP 端口"""
    port = start
    while True:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(('localhost', port)) != 0:
                return port
            port += 1


def start_preview(root_dir, port=8800, index_page='index.html'):
    """
    启动预览服务器

    :param root_dir:   网站根目录
    :param port:       起始端口号
    :param index_page: 自动打开的页面文件名
    """
    # 确保目录存在
    if not os.path.isdir(root_dir):
        print(f"错误：目录不存在 '{root_dir}'")
        sys.exit(1)

    # 寻找可用端口
    port = find_available_port(port)

    # 自定义请求处理器，抑制日志并设置根目录
    class QuietHandler(SimpleHTTPRequestHandler):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, directory=root_dir, **kwargs)

        def log_message(self, format, *args):
            pass  # 不输出访问日志

    # 创建服务器
    server = HTTPServer(('localhost', port), QuietHandler)
    print(f"预览服务器已启动")
    print(f"本地地址：http://localhost:{port}/{index_page}")
    print("按 Ctrl+C 停止服务器")

    # 在新线程中运行服务器，避免阻塞主线程（但主线程用于等待 Ctrl+C）
    server_thread = threading.Thread(target=server.serve_forever, daemon=True)
    server_thread.start()

    # 自动打开浏览器
    webbrowser.open(f"http://localhost:{port}/{index_page}")

    try:
        # 保持运行直到用户中断
        while True:
            server_thread.join(1)
    except KeyboardInterrupt:
        print("\n正在关闭服务器...")
        server.shutdown()
        print("服务器已停止")


def main():
    # 解析命令行参数
    args = sys.argv[1:]
    root_dir = args[0] if len(args) >= 1 else os.getcwd()
    port = int(args[1]) if len(args) >= 2 else 8800
    index_page = args[2] if len(args) >= 3 else 'index.html'

    start_preview(root_dir=root_dir, port=port, index_page=index_page)


if __name__ == '__main__':
    main()