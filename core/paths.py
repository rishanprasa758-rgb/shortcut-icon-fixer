# -*- coding: utf-8 -*-
"""程序运行期用到的目录。全部放在用户目录下，不需要管理员权限。"""
from __future__ import annotations

import os
from pathlib import Path

BASE = Path(os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))) / "ShortcutIconFixer"
ICONS_DIR = BASE / "icons"
BACKUP_DIR = BASE / "backup"
LOG_DIR = BASE / "logs"


def ensure_dirs() -> None:
    for d in (BASE, ICONS_DIR, BACKUP_DIR, LOG_DIR):
        d.mkdir(parents=True, exist_ok=True)
