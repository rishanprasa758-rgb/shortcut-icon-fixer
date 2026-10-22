# -*- coding: utf-8 -*-
"""Steam 相关：找到游戏装在哪、从游戏 exe 里把图标抠出来。"""
from __future__ import annotations

import os
import re
from pathlib import Path

DEFAULT_STEAM = Path(r"C:\Program Files (x86)\Steam")

# 这些名字的 exe 基本不是游戏本体，先排除，别把安装器图标抠出来
BAD_WORDS = (
    "unins", "setup", "install", "installer", "launcher", "vcredist",
    "dxsetup", "directx", "crash", "report", "config", "update",
    "redis", "helper", "test",
)


def steam_root() -> Path:
    """先从注册表问 Steam 装在哪；失败再退回默认路径。"""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Valve\Steam") as k:
            value = winreg.QueryValueEx(k, "SteamPath")[0]
            p = Path(str(value).replace("/", "\\"))
            if p.is_dir():
                return p
    except OSError:
        pass
    return DEFAULT_STEAM


def list_libraries(root: Path | None = None) -> list[Path]:
    """读 libraryfolders.vdf，拿到所有 Steam 库（C 盘、D 盘都可能有）。"""
    root = root or steam_root()
    vdf = root / "steamapps" / "libraryfolders.vdf"
    libs = [root]
    if vdf.is_file():
        text = vdf.read_text(encoding="utf-8", errors="ignore")
        for raw in re.findall(r'"path"\s+"([^"]+)"', text):
            p = Path(raw.replace("\\\\", "\\"))
            if p.is_dir() and p not in libs:
                libs.append(p)
    return libs


def find_game_dir(appid: str, root: Path | None = None) -> Path | None:
    """appmanifest_<appid>.acf 里有安装目录名，拼起来就是游戏文件夹。"""
    for lib in list_libraries(root):
        manifest = lib / "steamapps" / f"appmanifest_{appid}.acf"
        if not manifest.is_file():
            continue
        text = manifest.read_text(encoding="utf-8", errors="ignore")
        m = re.search(r'"installdir"\s+"([^"]+)"', text)
        if not m:
            continue
        game = lib / "steamapps" / "common" / m.group(1)
        if game.is_dir():
            return game
    return None


def _norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", s.lower())


def _similarity(a: str, b: str) -> int:
    """粗略相似度：公共前缀越长越像；完全相同额外加分。够排序用了。"""
    common = 0
    for x, y in zip(a, b):
        if x != y:
            break
        common += 1
    return common * 2 + (100 if a == b else 0)


def pick_exe(game_dir: Path) -> list[Path]:
    """按"和目录名像不像 + 文件大不大"给 exe 排序，返回候选列表。"""
    candidates: list[Path] = []
    base = game_dir
    for dirpath, dirnames, filenames in os.walk(base):
        depth = len(Path(dirpath).relative_to(base).parts)
        if depth >= 2:
            dirnames.clear()
        for fn in filenames:
            if not fn.lower().endswith(".exe"):
                continue
            low = fn.lower()
            if any(w in low for w in BAD_WORDS):
                continue
            candidates.append(Path(dirpath) / fn)

    target = _norm(base.name)
    return sorted(
        candidates,
        key=lambda p: (-_similarity(_norm(p.stem), target), -p.stat().st_size),
    )


def extract_from_pe(src: Path, out: Path, index: int = 0) -> bool:
    """从 exe/dll 里抠图标存成 .ico。icoextract 的 get_icon 只收整数索引。"""
    try:
        from icoextract import IconExtractor

        ex = IconExtractor(str(src))
        groups = ex.list_group_icons()
        if not groups:
            return False
        n = index if 0 <= index < len(groups) else 0
        data = ex.get_icon(n)
        raw = data.getvalue() if hasattr(data, "getvalue") else bytes(data)
        if len(raw) < 1000:          # 太小基本是坏的
            return False
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_bytes(raw)
        return True
    except Exception:
        return False


def root_from_icon(icon_path: str) -> Path | None:
    r"""从 <Steam>\steam\games\<hash>.ico 反推 Steam 根目录。

    有些机器的 Steam 装在非默认位置、注册表里也没有记录，这时靠这个兜底。
    """
    try:
        p = Path(icon_path)
        if p.parent.name.lower() == "games" and p.parent.parent.name.lower() == "steam":
            return p.parent.parent.parent
    except (OSError, ValueError):
        pass
    return None


def cover_from_cache(appid: str, out: Path, root: Path | None = None) -> bool:
    """游戏 exe 里没有图标时，退而用 Steam 缓存里的封面拼一个。"""
    try:
        from PIL import Image
    except ImportError:
        return False

    cache = (root or steam_root()) / "appcache" / "librarycache"
    names = (
        f"{appid}_icon.jpg", f"{appid}_icon.png",
        f"{appid}_library_600x900.jpg", f"{appid}_header.jpg",
    )
    for name in names:
        src = cache / name
        if not src.is_file():
            continue
        try:
            im = Image.open(src).convert("RGBA")
        except Exception:
            continue
        side = max(im.size)
        canvas = Image.new("RGBA", (side, side), (0, 0, 0, 0))
        canvas.paste(im, ((side - im.width) // 2, (side - im.height) // 2))
        out.parent.mkdir(parents=True, exist_ok=True)
        canvas.save(out, format="ICO", sizes=[(16, 16), (32, 32), (48, 48), (256, 256)])
        return True
    return False


def build_icon_for_app(appid: str, out: Path, hint_icon: str = "") -> tuple[bool, str]:
    """给 Steam 游戏生成图标文件，返回 (成功, 说明)。

    hint_icon 是快捷方式里记录的图标路径，用来反推 Steam 装在哪，
    这样别人把 Steam 装到奇怪位置也能找到游戏。
    """
    roots: list[Path] = []
    for r in (steam_root(), root_from_icon(hint_icon) if hint_icon else None):
        if r and r not in roots:
            roots.append(r)

    game = None
    for root in roots:
        game = find_game_dir(appid, root)
        if game:
            for exe in pick_exe(game):
                if extract_from_pe(exe, out):
                    return True, f"从 {exe.name} 提取"
            break

    for root in roots:
        if cover_from_cache(appid, out, root):
            return True, "用 Steam 封面生成"
    if game:
        return False, "游戏 exe 里没有图标，Steam 缓存里也没有封面"
    return False, "Steam 里找不到这个游戏（可能没安装）"
