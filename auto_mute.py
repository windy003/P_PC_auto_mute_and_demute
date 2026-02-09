"""
Auto-Mute on Idle
当系统空闲超过指定时间后自动静音，有输入操作后恢复音量。
系统托盘图标运行，右键菜单可退出。
"""

import ctypes
import time
import sys
import os
import threading

from dotenv import load_dotenv
from ctypes import Structure, windll, c_uint, sizeof, byref
from PIL import Image
import pystray


class LASTINPUTINFO(Structure):
    _fields_ = [
        ("cbSize", c_uint),
        ("dwTime", c_uint),
    ]


def get_idle_seconds():
    """获取系统空闲秒数（自上次键盘/鼠标/触摸板输入以来）。"""
    lii = LASTINPUTINFO()
    lii.cbSize = sizeof(LASTINPUTINFO)
    windll.user32.GetLastInputInfo(byref(lii))
    tick_count = windll.kernel32.GetTickCount()
    idle_ms = tick_count - lii.dwTime
    return idle_ms / 1000.0


def get_audio_session():
    """获取系统默认音频输出的音量接口。"""
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
    from comtypes import CLSCTX_ALL, cast, POINTER

    speakers = AudioUtilities.GetSpeakers()
    device = getattr(speakers, '_dev', speakers)
    interface = device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    return volume


def main():
    load_dotenv()

    idle_minutes = float(os.getenv("IDLE_MINUTES", "5"))
    idle_threshold = idle_minutes * 60
    print(f"Auto-Mute on Idle 已启动")
    print(f"空闲阈值: {idle_minutes} 分钟 ({idle_threshold:.0f} 秒)")

    volume = get_audio_session()

    muted_by_us = False
    saved_volume = None
    running = True

    # 加载托盘图标
    script_dir = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(script_dir, "2048x2048.png")
    icon_image = Image.open(icon_path)

    def on_quit(icon, item):
        nonlocal running
        running = False
        icon.stop()

    tray_icon = pystray.Icon(
        name="AutoMute",
        icon=icon_image,
        title=f"Auto-Mute on Idle ({idle_minutes}min)",
        menu=pystray.Menu(
            pystray.MenuItem(f"空闲阈值: {idle_minutes} 分钟", None, enabled=False),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("退出(&X)", on_quit),
        ),
    )

    def monitor_loop():
        nonlocal muted_by_us, saved_volume

        # 在此线程中初始化 COM（pycaw 需要）
        import comtypes
        comtypes.CoInitialize()

        while running:
            idle_sec = get_idle_seconds()

            if idle_sec >= idle_threshold and not muted_by_us:
                saved_volume = volume.GetMasterVolumeLevelScalar()
                volume.SetMute(1, None)
                muted_by_us = True
                tray_icon.title = f"Auto-Mute [已静音] 空闲 {idle_sec:.0f}s"
                print(f"[静音] 空闲 {idle_sec:.0f}s，已保存音量 {saved_volume:.0%} 并静音")

            elif idle_sec < idle_threshold and muted_by_us:
                volume.SetMute(0, None)
                if saved_volume is not None:
                    volume.SetMasterVolumeLevelScalar(saved_volume, None)
                muted_by_us = False
                print(f"[恢复] 检测到输入，音量已恢复至 {saved_volume:.0%}")
                tray_icon.title = f"Auto-Mute on Idle ({idle_minutes}min)"
                saved_volume = None

            time.sleep(1)

        # 退出前恢复音量
        if muted_by_us:
            volume.SetMute(0, None)
            if saved_volume is not None:
                volume.SetMasterVolumeLevelScalar(saved_volume, None)
            print("[恢复] 退出前已恢复音量")

        comtypes.CoUninitialize()
        print("已退出")

    # 在后台线程运行监控循环
    monitor_thread = threading.Thread(target=monitor_loop, daemon=True)
    monitor_thread.start()

    # 主线程运行托盘图标（阻塞）
    tray_icon.run()


if __name__ == "__main__":
    main()
