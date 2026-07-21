#!/usr/bin/env python3
import ctypes
import time
import uiautomation as auto

VK_ESCAPE = 0x1B

onenote = auto.WindowControl(searchDepth=1, ClassName="Framework::CFrame")
list = onenote.ListControl()


def search_loop() -> None:
    while True:
        for item in list.GetChildren():
            if not process(item):
                return

        pattern = list.GetScrollPattern()

        old = pattern.VerticalScrollPercent

        pattern.SetScrollPercent(pattern.HorizontalScrollPercent, min(old + 10, 100))

        time.sleep(0.3)

        if pattern.VerticalScrollPercent == old:
            return


def process(item) -> bool:
    item.RightClick()
    menu = auto.MenuControl(searchDepth=2)
    mi = menu.MenuItemControl(Name="コピー")
    mi.Click()
    time.sleep(0.5)
    if bool(ctypes.windll.user32.GetAsyncKeyState(VK_ESCAPE) & 0x8000):
        return False
    return True


search_loop()
