import os
import sys
import winreg
import time
import threading
from pystray import Menu, MenuItem, Icon
from PIL import Image
from pathlib import Path
from ctypes import windll
from win32com.client import Dispatch

# 获取图片位置
def get_image_path(rel_path):
    if hasattr(sys, '_MEIPASS'):
        base = Path(sys._MEIPASS)
    else:
        base = Path(__file__).parent
    return base / rel_path

on_image = get_image_path("red.ico")
off_image = get_image_path("black.ico")

# 设置代理
def set_proxy(enable):
    # 修改注册表
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
        r'Software\Microsoft\Windows\CurrentVersion\Internet Settings',
        0, winreg.KEY_WRITE)
    winreg.SetValueEx(key, 'ProxyEnable', 0, winreg.REG_DWORD, enable)
    winreg.CloseKey(key)
    
    # 刷新系统设置
    internet_set_option = windll.Wininet.InternetSetOptionW
    internet_set_option(0, 39, 0, 0)  # INTERNET_OPTION_SETTINGS_CHANGED
    internet_set_option(0, 37, 0, 0)  # INTERNET_OPTION_REFRESH

# 获取当前代理状态
def get_proxy_status():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
        r'Software\Microsoft\Windows\CurrentVersion\Internet Settings')
    value, _ = winreg.QueryValueEx(key, 'ProxyEnable')
    winreg.CloseKey(key)
    return bool(value)

# 切换代理
def toggle_proxy(icon, item):
    current_status = get_proxy_status()
    set_proxy(0 if current_status else 1)
    icon.title = "代理状态: " + ("已启用" if not current_status else "已关闭")
    icon.icon = Image.open(on_image if not current_status else off_image)
    icon.update_menu()

# 获取开机启动路径
def get_startup_path():
    return os.path.join(
            os.getenv('APPDATA'),
            r'Microsoft\Windows\Start Menu\Programs\Startup'
    )

# 获取当前开机启动状态
def get_startup_status():
    shortcut_path = os.path.join(get_startup_path(), 'SiProxy.lnk')
    return os.path.exists(shortcut_path)

# 切换开机启动
def toggle_startup(icon, item):
    if get_startup_status():
        remove_startup()
    else:
        create_startup()
    icon.update_menu()

def create_startup():
    try:
        shortcut_path = os.path.join(get_startup_path(), 'SiProxy.lnk')
        if not os.path.exists(shortcut_path):
            target = sys.executable
            shell = Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.TargetPath = target
            shortcut.WorkingDirectory = os.path.dirname(target)
            shortcut.save()
        return True
    except Exception as e:
        print(f"创建失败：{e}")
        return False

def remove_startup():
    try:
        shortcut_path = os.path.join(get_startup_path(), 'SiProxy.lnk')
        if os.path.exists(shortcut_path):
            os.remove(shortcut_path)
        return True
    except Exception as e:
        print(f"删除失败：{e}")
        return False

# 图标更新线程
class IconUpdater(threading.Thread):
    def __init__(self, icon):
        super().__init__(daemon=True)
        self.icon = icon
        self.running = True

    def run(self):
        while self.running:
            current_status = get_proxy_status()
            self.icon.title = "代理状态: " + ("已启用" if current_status else "已关闭")
            self.icon.icon = Image.open(on_image if current_status else off_image)
            self.icon.update_menu()
            time.sleep(5)

    def stop(self):
        self.running = False

# 创建任务栏图标
def create_tray_icon():
    image = Image.open(on_image if  get_proxy_status() else off_image)
    
    menu = Menu(
        MenuItem(
            lambda item: "社保代理：开" if get_proxy_status() else "社保代理：关",
            toggle_proxy
        ),
        MenuItem(
            lambda item: "开机启动：开" if get_startup_status() else "开机启动：关",
            toggle_startup
        ),
        MenuItem("退出", lambda: icon.stop())
    )
    
    icon = Icon("proxy", image, "系统代理开关", menu)
    icon.title = "代理状态: " + ("已启用" if get_proxy_status() else "已关闭")

    updater = IconUpdater(icon)
    updater.start()

    icon._on_stop = updater.stop
    icon.run()

if __name__ == "__main__":
    time.sleep(3)
    create_tray_icon()
