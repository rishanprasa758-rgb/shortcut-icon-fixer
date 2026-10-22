# -*- coding: utf-8 -*-
"""判定层：判断一个快捷方式的图标到底坏没坏，以及能不能修。"""
from __future__ import annotations

import os

from .scan import ShortcutItem

# 状态含义一目了然，界面和 CLI 都用这几个字符串
OK = "ok"
BAD = "bad"
TARGET_MISSING = "target_missing"
SYSTEM = "system"
NOREPO = "norepo"

STATUS_TEXT = {
    OK: "正常",
    BAD: "图标源丢失",
    TARGET_MISSING: "目标不存在",
    SYSTEM: "系统项，跳过",
    NOREPO: "无本地图标源",
}


def is_pe(path: str) -> bool:
    """粗略判断是不是可执行文件：看头两个字节是不是 'MZ'。"""
    try:
        with open(path, "rb") as f:
            return f.read(2) == b"MZ"
    except OSError:
        return False


def icon_file_ok(path: str) -> bool:
    """图标文件算不算可用：存在、且不是个空壳。"""
    try:
        if not os.path.isfile(path):
            return False
        if path.lower().endswith(".ico") and os.path.getsize(path) < 1024:
            return False
        return True
    except OSError:
        return False


def diagnose(item: ShortcutItem) -> ShortcutItem:
    if item.kind == "url":
        if not item.icon_source:
            if item.appid:
                item.status = BAD
                item.note = "Steam 快捷方式没有图标文件（Steam 没生成）"
            else:
                item.status = NOREPO
                item.note = "网页快捷方式，本地没有图标源"
        elif icon_file_ok(item.icon_source):
            item.status = OK
        else:
            item.status = BAD
            item.note = "图标文件不存在或已损坏"
        return item

    # ---- .lnk ----
    # IconLocation 没写路径（形如 ",0"）时，图标来自目标程序本身
    source = item.icon_source or item.target
    if not source:
        item.status = SYSTEM
        item.note = "没有目标路径（UWP/系统项）"
        return item

    if icon_file_ok(source):
        item.status = OK
        return item

    if not os.path.exists(item.target):
        item.status = TARGET_MISSING
        item.note = "程序已卸载或移动，没有图标可找"
        return item

    # 目标还在，图标路径却坏了
    if is_pe(item.target):
        item.status = BAD
        item.note = "图标路径失效，可从目标程序重新取图标"
    else:
        item.status = NOREPO
        item.note = "图标路径失效，且目标不是程序"
    return item
