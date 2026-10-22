# -*- coding: utf-8 -*-
"""扫描层：找到所有快捷方式，并把它们的"图标从哪来"读出来。"""
from __future__ import annotations

import ctypes
import os
import re
import uuid
from ctypes import wintypes
from dataclasses import dataclass, field
from pathlib import Path

import win32com.client

# ---------------------------------------------------------------- 已知文件夹

class _GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", wintypes.DWORD),
        ("Data2", wintypes.WORD),
        ("Data3", wintypes.WORD),
        ("Data4", ctypes.c_ubyte * 8),
    ]


def _guid(text: str) -> _GUID:
    """把 'B4BFCC3A-DB2C-...' 这种字符串变成 Windows 要的结构体。

    类比 C：相当于把一个 16 字节的结构体按小端拆成 4 个字段填进去。
    """
    b = uuid.UUID(text).bytes_le
    return _GUID(
        int.from_bytes(b[0:4], "little"),
        int.from_bytes(b[4:6], "little"),
        int.from_bytes(b[6:8], "little"),
        (ctypes.c_ubyte * 8)(*b[8:16]),
    )


_SHELL32 = ctypes.WinDLL("shell32", use_last_error=True)
_OLE32 = ctypes.WinDLL("ole32", use_last_error=True)
_SHELL32.SHGetKnownFolderPath.argtypes = [
    ctypes.POINTER(_GUID), wintypes.DWORD, wintypes.HANDLE,
    ctypes.POINTER(ctypes.c_wchar_p),
]
_SHELL32.SHGetKnownFolderPath.restype = ctypes.HRESULT
_OLE32.CoTaskMemFree.argtypes = [ctypes.c_void_p]


def known_folder(fid: str) -> str | None:
    """问 Windows 要标准文件夹的真实路径（能正确处理 OneDrive 重定向的桌面）。"""
    try:
        p = ctypes.c_wchar_p()
        _SHELL32.SHGetKnownFolderPath(ctypes.byref(_guid(fid)), 0, None, ctypes.byref(p))
        value = p.value
        _OLE32.CoTaskMemFree(ctypes.cast(p, ctypes.c_void_p))
        return value
    except OSError:
        return None


FOLDER = {
    "desktop": "B4BFCC3A-DB2C-424C-B029-7FE99A87C641",
    "public_desktop": "C4AA340D-F20F-4863-AFEF-F87EF2E6BA25",
    "programs": "A77F5D77-2E2B-44C3-A6A2-ABA601054A51",
    "common_programs": "0139D44E-6AFE-49F2-8690-3DAFCAE6FFB8",
    "roaming": "3EB685DB-65F9-4CF6-A03A-E3EF65729F3D",
}

FILE_ATTRIBUTE_HIDDEN = 0x2
FILE_ATTRIBUTE_SYSTEM = 0x4


# ---------------------------------------------------------------- 数据结构

@dataclass
class ShortcutItem:
    path: str                      # 快捷方式文件本身
    kind: str                      # "url" / "lnk"
    name: str                      # 显示名（去掉后缀）
    place: str                     # 桌面 / 开始菜单 / 任务栏
    target: str = ""               # .lnk 的目标 / .url 的 URL
    icon_source: str = ""          # 当前图标实际取自哪个文件
    icon_index: int = 0            # 图标在文件里的序号
    status: str = "unknown"        # ok / bad / target_missing / system / norepo
    note: str = ""                 # 给人看的原因
    appid: str = ""                # 仅 Steam 游戏
    encoding: str = "utf-8"        # 仅 .url，写回时保持一致
    has_bom: bool = False

    @property
    def repairable(self) -> bool:
        return self.status == "bad"


# ---------------------------------------------------------------- 位置枚举

def _roots() -> list[tuple[str, str, bool]]:
    """返回 [(根目录, 位置标签, 是否递归), ...]"""
    out = []
    pairs = [
        ("desktop", "桌面", False),
        ("public_desktop", "公共桌面", False),
        ("programs", "开始菜单", True),
        ("common_programs", "公共开始菜单", True),
    ]
    for key, label, recursive in pairs:
        p = known_folder(FOLDER[key])
        if p:
            out.append((p, label, recursive))

    roaming = known_folder(FOLDER["roaming"])
    if roaming:
        taskbar = Path(roaming) / "Microsoft" / "Internet Explorer" / "Quick Launch" / "User Pinned" / "Taskbar"
        if taskbar.is_dir():
            out.append((str(taskbar), "任务栏", False))
    return out


def _skip(path: str) -> bool:
    name = os.path.basename(path)
    if name.lower() in ("desktop.ini", "thumbs.db"):
        return True
    if name.startswith("~$"):
        return True
    try:
        attrs = os.stat(path).st_file_attributes
    except OSError:
        return True
    return bool(attrs & (FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_SYSTEM))


def iter_shortcut_files():
    """产出 (快捷方式路径, 位置标签)，已去重、已跳过隐藏/系统文件。"""
    seen = set()
    for root, place, recursive in _roots():
        for dirpath, dirnames, filenames in os.walk(root):
            if not recursive:
                dirnames.clear()          # 只看这一层
            for fn in filenames:
                if not fn.lower().endswith((".lnk", ".url")):
                    continue
                full = os.path.join(dirpath, fn)
                if _skip(full):
                    continue
                # 同一个文件被 junction 指了两遍时只算一次
                key = os.path.normcase(os.path.realpath(full))
                if key in seen:
                    continue
                seen.add(key)
                yield full, place


# ---------------------------------------------------------------- 读取内容

def read_url(path: str) -> tuple[dict, str, bool]:
    """.url 是文本文件。返回 (字段字典, 编码, 是否有 BOM)。

    坑：有的文件是 UTF-8 带 BOM，有的是 GBK，写回时必须原样保持，
    否则中文路径会烂掉。所以这里是"解码一次，把编码也记下来"。
    """
    raw = Path(path).read_bytes()
    has_bom = raw.startswith(b"\xef\xbb\xbf")
    if has_bom:
        enc = "utf-8-sig"
    else:
        enc = "utf-8"
        try:
            raw.decode("utf-8")
        except UnicodeDecodeError:
            enc = "mbcs"

    text = raw.decode(enc, errors="replace")
    fields = {}
    for line in text.splitlines():
        if "=" in line and not line.lstrip().startswith("["):
            k, v = line.split("=", 1)
            fields[k.strip()] = v.strip()
    return fields, enc, has_bom


def write_url_icon(path: str, icon_file: str, encoding: str) -> None:
    """只改 .url 里的 IconFile 一行，其余字节原样保留（换行风格也不动）。"""
    raw = Path(path).read_bytes()
    text = raw.decode(encoding, errors="replace")
    newline = "\r\n" if "\r\n" in text else "\n"

    lines = text.splitlines(keepends=True)
    replaced = False
    for i, line in enumerate(lines):
        if line.lstrip().lower().startswith("iconfile="):
            body = line.rstrip("\r\n")
            ending = line[len(body):]
            lines[i] = f"IconFile={icon_file}{ending}"
            replaced = True
            break

    if not replaced:
        # 没有这一行就追加到末尾；.url 的 IconFile 属于 [InternetShortcut] 段
        if lines and not lines[-1].endswith(("\n", "\r")):
            lines[-1] += newline
        lines.append(f"IconFile={icon_file}{newline}")

    Path(path).write_bytes("".join(lines).encode(encoding, errors="replace"))


def split_icon_location(loc: str) -> tuple[str, int]:
    """把 'C:\\a\\b.exe,3' 拆成 (路径, 3)。

    坑：路径本身可能带逗号（"D:\\My, Games\\x.exe,0"），所以从最后一个逗号切。
    """
    loc = (loc or "").strip().strip('"')
    if not loc:
        return "", 0
    if "," in loc:
        path, _, idx = loc.rpartition(",")
        path = path.strip().strip('"')
        try:
            return path, int(idx)
        except ValueError:
            return loc, 0
    return loc, 0


class LnkReader:
    """用 pywin32 在进程内读 .lnk，比每个文件起一次 PowerShell 快得多。"""

    def __init__(self):
        self.shell = win32com.client.Dispatch("WScript.Shell")

    def read(self, path: str) -> tuple[str, str, int]:
        s = self.shell.CreateShortcut(path)
        icon_path, icon_index = split_icon_location(s.IconLocation)
        return s.TargetPath or "", icon_path, icon_index

    def set_icon(self, path: str, icon_location: str) -> None:
        s = self.shell.CreateShortcut(path)
        s.IconLocation = icon_location
        s.Save()

    def close(self) -> None:
        """主动释放 COM 对象。必须在 CoUninitialize 之前调用，
        否则 Python 回收它时会报 'Win32 exception occurred releasing IUnknown'。"""
        self.shell = None


def read_shortcut(path: str, place: str, lnk: LnkReader) -> ShortcutItem:
    name = Path(path).stem
    if path.lower().endswith(".url"):
        fields, enc, has_bom = read_url(path)
        url = fields.get("URL", "")
        m = re.search(r"steam://rungameid/(\d+)", url)
        return ShortcutItem(
            path=path, kind="url", name=name, place=place,
            target=url,
            icon_source=os.path.expandvars(fields.get("IconFile", "")),
            icon_index=int(fields.get("IconIndex", "0") or 0),
            appid=m.group(1) if m else "",
            encoding=enc, has_bom=has_bom,
        )

    target, icon_path, icon_index = lnk.read(path)
    return ShortcutItem(
        path=path, kind="lnk", name=name, place=place,
        target=target,
        icon_source=os.path.expandvars(icon_path),
        icon_index=icon_index,
    )
