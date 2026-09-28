# -*- coding: utf-8 -*-
"""
张老师·C语言选择结构智能助教 - 本地部署服务器
双击「启动服务器.bat」或命令行运行：python start_server.py
手机和电脑连同一个 WiFi，手机浏览器访问下方显示的地址即可
"""
import http.server
import socketserver
import socket
import os
import sys
import webbrowser
import threading
import io

# Windows 命令行中文输出编码修复
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

PORT = 8080
HTML_FILE = "小C老师-C语言选择结构智能助教.html"


def get_local_ip():
    """获取本机局域网 IP"""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


class AppHandler(http.server.SimpleHTTPRequestHandler):
    """访问根路径直接显示智能体页面，不用点目录"""

    def do_GET(self):
        if self.path == "/" or self.path == "/index.html":
            self.path = "/" + HTML_FILE
        return super().do_GET()

    def end_headers(self):
        self.send_header("Cache-Control", "no-cache")
        super().end_headers()


def main():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    if not os.path.exists(HTML_FILE):
        print(f"[错误] 找不到文件: {HTML_FILE}")
        print("请确认 start_server.py 和 HTML 文件在同一文件夹内")
        input("按回车退出...")
        sys.exit(1)

    local_ip = get_local_ip()
    local_url = f"http://localhost:{PORT}"
    lan_url = f"http://{local_ip}:{PORT}"

    print()
    print("  ╔══════════════════════════════════════════╗")
    print("  ║   张老师 · C语言选择结构智能助教          ║")
    print("  ║   本地部署服务器已启动                    ║")
    print("  ╠══════════════════════════════════════════╣")
    print("  ║                                          ║")
    print(f"  ║   电脑访问:  {local_url}")
    print(f"  ║   手机访问:  {lan_url}")
    print("  ║              (手机和电脑连同一个WiFi)     ║")
    print("  ║                                          ║")
    print("  ║   手机上建议「添加到主屏幕」当APP用       ║")
    print("  ║   关闭本窗口或按 Ctrl+C 停止服务器       ║")
    print("  ╚══════════════════════════════════════════╝")
    print()

    # 自动在默认浏览器中打开
    threading.Timer(1.0, lambda: webbrowser.open(local_url)).start()

    socketserver.TCPServer.allow_reuse_address = True

    def make_server(port):
        return socketserver.TCPServer(("", port), AppHandler)

    httpd = None
    for try_port in [PORT, 8081, 8888, 9090, 0]:  # 0 = 系统随机分配
        try:
            httpd = make_server(try_port)
            if try_port == 0:
                try_port = httpd.server_address[1]
            if try_port != PORT:
                print(f"  [提示] 端口 {PORT} 被占用，已自动切换到 {try_port}")
                local_url = f"http://localhost:{try_port}"
                lan_url = f"http://{local_ip}:{try_port}"
                print(f"  电脑访问:  {local_url}")
                print(f"  手机访问:  {lan_url}")
            break
        except OSError:
            continue

    if httpd is None:
        print("  [错误] 无法启动服务器，请检查防火墙设置")
        input("\n  按回车退出...")
        sys.exit(1)

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n  服务器已停止")


if __name__ == "__main__":
    main()
