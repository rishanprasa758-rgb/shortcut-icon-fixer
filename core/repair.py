# -*- coding: utf-8 -*-
"""修复层：能修的修，不能修的给出人话原因。"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import time
import uuid
from datetime import datetime
from pathlib import Path

from . import paths
from .diagnose import BAD, STATUS_TEXT
from .scan import ShortcutItem, write_url_icon
from .steam import build_icon_for_app

CREATE_NO_WINDOW = 0x08000000


# ---------------------------------------------------------------- 备份

def _backup_dir() -> Path:
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    d = paths.BACKUP_DIR / stamp
    d.mkdir(parents=True, exist_ok=True)
    return d


def backup(paths_to_save: list[Path]) -> Path | None:
    """改动前把原文件存一份。返回备份目录。"""
    existing = [p for p in paths_to_save if p and p.is_file()]
    if not existing:
        return None
    d = _backup_dir()
    manifest = []
    for p in existing:
        try:
            shutil.copy2(p, d / p.name)
            manifest.append({"from": str(p), "to": (d / p.name).name})
        except OSError:
            continue
    (d / "manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    _trim_backups(keep=5)
    return d


def _trim_backups(keep: int) -> None:
    dirs = sorted([d for d in paths.BACKUP_DIR.iterdir() if d.is_dir()], reverse=True)
    for old in dirs[keep:]:
        shutil.rmtree(old, ignore_errors=True)


# ---------------------------------------------------------------- 修复

def _copy_icon(src: Path, dst: Path) -> bool:
    try:
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(src, dst)
        return True
    except OSError:
        return False


def repair_item(item: ShortcutItem, lnk, dry_run: bool = False) -> tuple[bool, str]:
    """修一个快捷方式。返回 (是否成功, 给人看的说明)。"""
    if item.status != BAD:
        return False, STATUS_TEXT.get(item.status, "不需要修")

    if item.kind == "url":
        return _repair_url(item, dry_run)
    return _repair_lnk(item, lnk, dry_run)


def _repair_url(item: ShortcutItem, dry_run: bool) -> tuple[bool, str]:
    if not item.appid:
        return False, "网页快捷方式没有本地图标源"

    if dry_run:
        return True, "[演练] 会从 Steam 游戏提取图标并写回"

    # 先在工作目录生成，避免直接往 System/Program Files 里写
    tmp = paths.ICONS_DIR / f"build_{item.appid}_{uuid.uuid4().hex[:8]}.ico"
    ok, why = build_icon_for_app(item.appid, tmp, hint_icon=item.icon_source)
    if not ok:
        return False, why

    final: Path | None = None
    if item.icon_source:
        original = Path(item.icon_source)
        backup([original])
        if _copy_icon(tmp, original):
            final = original
    if final is None:
        # 原位置写不进去（没权限），改放用户目录，再让快捷方式指过去
        final = paths.ICONS_DIR / f"steam_{item.appid}.ico"
        if not _copy_icon(tmp, final):
            return False, "图标生成成功但写入失败"

    if str(final) != item.icon_source:
        backup([Path(item.path)])
        write_url_icon(item.path, str(final), item.encoding)
        note = f"{why}；无权限写原位置，已改指用户目录"
    else:
        note = why

    tmp.unlink(missing_ok=True)
    item.icon_source = str(final)
    return True, note


def _repair_lnk(item: ShortcutItem, lnk, dry_run: bool) -> tuple[bool, str]:
    target = Path(item.target)
    if not target.is_file():
        return False, "目标程序不存在，无法修复"
    if not item.target.lower().endswith((".exe", ".dll", ".ico")):
        return False, "目标不是程序，没有图标可取"

    new_loc = f"{item.target},0"

    if dry_run:
        return True, f"[演练] 会把图标改指 {item.target},0"

    backup([Path(item.path)])
    try:
        lnk.set_icon(item.path, new_loc)
    except Exception as e:
        text = str(e)
        # 0x80070005 = 拒绝访问；公共开始菜单等在 ProgramData 下的需要管理员
        if "80070005" in text or "-2147024891" in text or "denied" in text.lower():
            return False, "需要管理员权限（点「管理员重试」）"
        return False, f"写入快捷方式失败：{e}"

    # 写完回读校验，防止 Save() 丢了别的属性
    try:
        read_target, _, _ = lnk.read(item.path)
    except Exception:
        read_target = ""
    if read_target and os.path.normcase(read_target) != os.path.normcase(item.target):
        return False, "写入后目标路径异常，已备份原文件"

    item.icon_source = item.target
    return True, "图标已改指目标程序"


# ---------------------------------------------------------------- 刷新显示

def refresh_shell() -> None:
    """让资源管理器重新读图标。不重启 explorer，只是通知一下。"""
    try:
        subprocess.run(
            ["ie4uinit.exe", "-show"], check=False,
            creationflags=CREATE_NO_WINDOW, timeout=20,
        )
    except Exception:
        pass


def rebuild_icon_cache() -> tuple[bool, str]:
    """危险操作：杀掉 explorer、删图标缓存、再拉起来。调用前必须让用户确认。"""
    local = Path(os.environ.get("LOCALAPPDATA", ""))
    targets = [
        *local.glob("Microsoft/Windows/Explorer/iconcache_*.db"),
        local / "IconCache.db",
    ]
    try:
        subprocess.run(["taskkill", "/f", "/im", "explorer.exe"],
                       check=False, creationflags=CREATE_NO_WINDOW)
        time.sleep(1.0)
        removed = 0
        for f in targets:
            try:
                f.unlink()
                removed += 1
            except OSError:
                pass
        subprocess.Popen(["explorer.exe"], creationflags=CREATE_NO_WINDOW)
        return True, f"已删除 {removed} 个缓存文件并重启资源管理器"
    except Exception as e:
        try:
            subprocess.Popen(["explorer.exe"], creationflags=CREATE_NO_WINDOW)
        except Exception:
            pass
        return False, f"重建失败：{e}"
