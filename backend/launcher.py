# -*- coding: utf-8 -*-
"""桌面图标逻辑：检测端口 -> 必要时静默拉起服务 -> 打开浏览器
用法：launcher.py          双击场景（拉起+开浏览器）
      launcher.py --autostart  开机自启场景（只拉起，不开浏览器）
"""
import sys, os, socket, subprocess, time

SRV = r"C:\Users\Administrator\WorkBuddy\2026-09-28-16-44-19\server"
PYW = r"E:\Python313\pythonw.exe"
URL = "http://127.0.0.1:8000"

def port_up():
    """socket 直连检测端口，不走代理、不依赖外部命令"""
    s = socket.socket()
    s.settimeout(1.0)
    try:
        s.connect(("127.0.0.1", 8000))
        return True
    except OSError:
        return False
    finally:
        s.close()

def alert(msg):
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, msg, "证照印章管理系统", 0x10)
    except Exception:
        pass

def main():
    autostart = "--autostart" in sys.argv
    if not port_up():
        try:
            subprocess.Popen([PYW, os.path.join(SRV, "server.py")],
                             cwd=SRV,
                             creationflags=0x00000008)  # DETACHED_PROCESS
        except Exception as e:
            alert("启动失败：" + str(e))
            return
        for _ in range(30):
            time.sleep(1)
            if port_up():
                break
        if not port_up():
            alert("服务启动超时，请截图联系技术支持")
            return
    if not autostart:
        try:
            os.startfile(URL)
        except Exception:
            import webbrowser
            webbrowser.open(URL)

if __name__ == "__main__":
    main()