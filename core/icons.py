# -*- coding: utf-8 -*-
"""预览图：把图标源变成 16×16 的小图，显示在列表里。"""
from __future__ import annotations

import os
from pathlib import Path

from PIL import Image

_cache: dict[tuple[str, int, int], Image.Image | None] = {}


def _from_ico(src: str, size: int) -> Image.Image | None:
    try:
        im = Image.open(src)
        im.size = (size, size)
        im.load()
        bg = Image.new("RGB", (size, size), "white")
        rgba = im.convert("RGBA")
        bg.paste(rgba, (0, 0), rgba)
        return bg
    except Exception:
        return None


def _from_pe(src: str, index: int, size: int) -> Image.Image | None:
    """从 exe/dll 里取出图标画到白色画布上。

    类比 C：申请一块内存 DC（相当于 buffer），在上面画完再把像素读出来。
    注意 GDI 对象用完必须销毁，否则每扫一次就漏一点资源。
    """
    import win32gui
    import win32ui

    hicon = None
    # pywin32 正常情况下返回 (large_list, small_list)，但文件里没图标时可能返回 0，
    # 直接解包会 TypeError。所以先判断类型。
    result = win32gui.ExtractIconEx(src, index)
    if isinstance(result, tuple) and len(result) == 2:
        large, small = result
        pool = small or large
        if pool:
            hicon = pool[0]
    if not hicon:
        return None

    screen = win32gui.GetDC(0)
    try:
        sdc = win32ui.CreateDCFromHandle(screen)
        try:
            memdc = sdc.CreateCompatibleDC()
            try:
                bmp = win32ui.CreateBitmap()
                bmp.CreateCompatibleBitmap(sdc, size, size)
                memdc.SelectObject(bmp)
                memdc.FillSolidRect((0, 0, size, size), 0xFFFFFF)
                memdc.DrawIcon((0, 0), hicon)
                bits = bmp.GetBitmapBits(True)
                img = Image.frombuffer("RGB", (size, size), bits, "raw", "BGRX", 0, 1)
                img.load()
                return img
            finally:
                win32gui.DeleteObject(bmp.GetHandle())
                memdc.DeleteDC()
        finally:
            sdc.DeleteDC()
    finally:
        win32gui.ReleaseDC(0, screen)
        if hicon:
            win32gui.DestroyIcon(hicon)


def load_preview(src: str, index: int = 0, size: int = 16) -> Image.Image | None:
    """返回一张 PIL 图片；拿不到就返回 None，界面显示空白。"""
    if not src or not os.path.isfile(src):
        return None
    key = (os.path.normcase(src), index, size)
    if key in _cache:
        return _cache[key]

    ext = Path(src).suffix.lower()
    img = _from_ico(src, size) if ext == ".ico" else _from_pe(src, index, size)
    _cache[key] = img
    return img
