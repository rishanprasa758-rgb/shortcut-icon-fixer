# -*- coding: utf-8 -*-
"""快捷方式图标修复器 —— 命令行入口 + tkinter 界面。

用法：
    python app.py                     打开界面
    python app.py --scan              命令行扫描
    python app.py --scan --json       扫描并以 JSON 输出
    python app.py --fix --dry-run     只报告要做什么，不动文件
    python app.py --fix               修复所有坏的
"""
from __future__ import annotations

import json
import os
import queue
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path

from core import paths
from core.diagnose import BAD, STATUS_TEXT, TARGET_MISSING, diagnose
from core.repair import refresh_shell, rebuild_icon_cache, repair_item
from core.scan import LnkReader, iter_shortcut_files, read_shortcut

# 界面用到的常量：一个"已修复"的临时状态，重新扫描后会被正常状态覆盖
FIXED = "fixed"

BG = "#F5F7FA"
CARD = "#FFFFFF"
BORDER = "#E3E8EF"
TEXT = "#1F2937"
MUTED = "#6B7280"
BLUE = "#2563EB"
BLUE_DARK = "#1D4ED8"
BLUE_DISABLED = "#C7D2FE"
GRAY_BTN = "#9CA3AF"
RED = "#D93025"
ORANGE = "#E8710A"
GREEN = "#13803A"
FONT = "Microsoft YaHei UI"


# ---------------------------------------------------------------- 扫描

def do_scan(log=None) -> tuple[list, LnkReader]:
    """扫一遍所有快捷方式。调用方负责初始化 COM，并在收尾时 lnk.close()。"""
    lnk = LnkReader()
    items = []
    for path, place in iter_shortcut_files():
        try:
            item = read_shortcut(path, place, lnk)
        except Exception:
            continue
        diagnose(item)
        items.append(item)
    return items, lnk


# ---------------------------------------------------------------- 命令行

def cli_scan(as_json: bool) -> int:
    import pythoncom

    pythoncom.CoInitialize()
    try:
        items, lnk = do_scan()
        lnk.close()
    finally:
        pythoncom.CoUninitialize()

    if as_json:
        data = [
            {
                "path": it.path, "name": it.name, "place": it.place,
                "kind": it.kind, "status": it.status,
                "status_text": STATUS_TEXT.get(it.status, it.status),
                "icon_source": it.icon_source, "target": it.target,
                "note": it.note, "appid": it.appid,
            }
            for it in items
        ]
        print(json.dumps(data, ensure_ascii=False, indent=2))
        return 0

    counts: dict[str, int] = {}
    for it in items:
        counts[it.status] = counts.get(it.status, 0) + 1
        if it.status != "ok":
            print(f"{STATUS_TEXT.get(it.status, it.status):<10} {it.place:<6} {it.name}  -> {it.icon_source or it.target}")
    print("-" * 60)
    print("总计", len(items), {STATUS_TEXT.get(k, k): v for k, v in counts.items()})
    return 0


def cli_fix(dry_run: bool, only: str | None) -> int:
    import pythoncom

    pythoncom.CoInitialize()
    try:
        items, lnk = do_scan()
        ok = fail = skip = 0
        need_admin = False
        for it in items:
            if only and os.path.normcase(it.path) != os.path.normcase(only):
                continue
            if it.status != BAD:
                continue
            if not it.repairable:
                skip += 1
                continue
            done, msg = repair_item(it, lnk, dry_run=dry_run)
            print(("OK  " if done else "FAIL") + f" {it.name}: {msg}")
            ok += done
            fail += not done
            need_admin = need_admin or (not done and "管理员" in msg)
        if not dry_run:
            refresh_shell()
        print(f"完成：成功 {ok}，失败 {fail}，跳过 {skip}")
        if need_admin:
            print("有项目需要管理员权限：请用管理员身份运行本命令重试。")
        lnk.close()
        return 0 if fail == 0 else 1
    finally:
        pythoncom.CoUninitialize()


# ---------------------------------------------------------------- 界面

def app_icon_path() -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, "assets", "app.ico")


class App:
    ROW_H = 68
    GAP = 10
    PAD = 16

    def __init__(self) -> None:
        import tkinter as tk
        from tkinter import ttk

        self.tk = tk
        self.root = tk.Tk()
        self.root.title("快捷方式图标修复")
        self.root.geometry("860x660")
        self.root.minsize(680, 520)
        self.root.configure(bg=BG)

        try:
            self.root.iconbitmap(app_icon_path())
        except Exception:
            pass

        self.items: list = []
        self.checked: set[str] = set()
        self.images: dict[str, object] = {}
        self.q: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None
        self.cancel = threading.Event()
        self.counters = {"ok": 0, "fail": 0, "skip": 0}
        self.need_admin = False
        self.show_all = False
        self.scan_finished = False
        self.busy = False
        self.log_visible = False
        self._rows: list[tuple[int, int, str]] = []
        self._resize_job = None
        self._redraw_job = None

        self._build_ui(ttk)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(80, self.poll)
        self.start_scan()

    # ------------------------------------------------------------ 布局

    def _build_ui(self, ttk) -> None:
        tk = self.tk
        import tkinter.font as tkfont
        self.font_name = tkfont.Font(family=FONT, size=10, weight="bold")

        header = tk.Frame(self.root, bg=BG, padx=22)
        header.pack(fill="x", pady=(18, 4))
        tk.Label(header, text="快捷方式图标修复", bg=BG, fg=TEXT,
                 font=(FONT, 15, "bold")).pack(side="left")

        self.btn_toggle = self._flat_button(header, "显示全部", self.toggle_view)
        self.btn_toggle.pack(side="right")

        self.summary = tk.Label(self.root, text="正在扫描…", bg=BG, fg=MUTED,
                                font=(FONT, 10), anchor="w")
        self.summary.pack(fill="x", padx=22)

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Blue.Horizontal.TProgressbar",
                        troughcolor="#E5E7EB", background=BLUE,
                        bordercolor=BG, lightcolor=BLUE, darkcolor=BLUE,
                        thickness=5)
        self.progress = ttk.Progressbar(self.root, style="Blue.Horizontal.TProgressbar",
                                        mode="determinate")

        # 底部操作区（先 pack 底部，剩下的空间留给列表）
        footer = tk.Frame(self.root, bg=BG, padx=22)
        footer.pack(side="bottom", fill="x", pady=(10, 16))

        self.btn_fix = tk.Button(
            footer, text="一键修复 (0)", command=self.start_fix,
            bg=BLUE_DISABLED, fg="white", activebackground=BLUE_DARK,
            activeforeground="white", relief="flat", bd=0,
            font=(FONT, 12, "bold"), cursor="arrow", pady=9,
        )
        self.btn_fix.pack(fill="x")
        self.btn_fix.bind("<Enter>", lambda e: self._hover_primary(True))
        self.btn_fix.bind("<Leave>", lambda e: self._hover_primary(False))

        aux = tk.Frame(footer, bg=BG)
        aux.pack(fill="x", pady=(8, 0))
        self.btn_scan = self._flat_button(aux, "重新扫描", self.start_scan)
        self.btn_cache = self._flat_button(aux, "重建图标缓存", self.on_rebuild_cache)
        self.btn_admin = self._flat_button(aux, "管理员重试", self.on_elevate)
        self.btn_log = self._flat_button(aux, "日志 ▾", self.toggle_log)
        for b in (self.btn_scan, self.btn_cache, self.btn_admin):
            b.pack(side="left", padx=(0, 14))
        self.btn_log.pack(side="right")

        # 日志区（默认隐藏）
        self.log_frame = tk.Frame(self.root, bg=BG, padx=22)
        self.logbox = tk.Text(self.log_frame, height=6, wrap="none", bd=0,
                              bg="#FFFFFF", fg="#374151", font=("Consolas", 9),
                              padx=10, pady=8, highlightthickness=0)
        log_scroll = ttk.Scrollbar(self.log_frame, orient="vertical",
                                   command=self.logbox.yview)
        self.logbox.configure(yscrollcommand=log_scroll.set)
        self.logbox.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")
        self.logbox.configure(state="disabled")

        # 列表
        body = tk.Frame(self.root, bg=BG)
        self.body = body
        body.pack(fill="both", expand=True, padx=(10, 0), pady=(6, 0))
        self.canvas = tk.Canvas(body, bg=BG, highlightthickness=0, bd=0)
        scroll = ttk.Scrollbar(body, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scroll.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        self.canvas.bind("<Configure>", self._on_resize)
        self.canvas.bind("<Button-1>", self._on_click)
        self.canvas.bind_all("<MouseWheel>", self._on_wheel)

    def _flat_button(self, parent, text, command) -> object:
        tk = self.tk
        btn = tk.Button(parent, text=text, command=command, bg=BG, fg="#374151",
                        activebackground="#E5E7EB", activeforeground=TEXT,
                        relief="flat", bd=0, font=(FONT, 9), cursor="hand2",
                        padx=6, pady=3)
        btn.bind("<Enter>", lambda e: btn.configure(bg="#E5E7EB"))
        btn.bind("<Leave>", lambda e: btn.configure(bg=BG))
        return btn

    def _hover_primary(self, on: bool) -> None:
        if str(self.btn_fix["state"]) == "disabled":
            return
        self.btn_fix.configure(bg=BLUE_DARK if on else BLUE)

    # ------------------------------------------------------------ 日志

    def toggle_log(self) -> None:
        if self.log_visible:
            self.log_frame.pack_forget()
            self.btn_log.configure(text="日志 ▾")
        else:
            self.log_frame.pack(side="bottom", fill="x", pady=(0, 8))
            self.btn_log.configure(text="日志 ▴")
        self.log_visible = not self.log_visible

    def log(self, text: str) -> None:
        stamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{stamp}] {text}\n"
        self.logbox.configure(state="normal")
        self.logbox.insert("end", line)
        self.logbox.see("end")
        self.logbox.configure(state="disabled")
        try:
            with open(paths.LOG_DIR / f"run-{datetime.now():%Y%m%d}.log", "a",
                      encoding="utf-8") as f:
                f.write(line)
        except OSError:
            pass

    # ------------------------------------------------------------ 扫描

    def start_scan(self) -> None:
        if self.worker and self.worker.is_alive():
            return
        self.items.clear()
        self.checked.clear()
        self.images.clear()
        self.show_all = False
        self.scan_finished = False
        self.cancel.clear()
        self.canvas.delete("all")
        self._rows = []
        self.summary.configure(text="正在扫描…")
        self._show_progress(True)
        self._set_busy(True)
        self._refresh_buttons()
        self.log("开始扫描…")
        self.worker = threading.Thread(target=self._scan_worker, daemon=True)
        self.worker.start()

    def _scan_worker(self) -> None:
        import pythoncom

        pythoncom.CoInitialize()
        lnk = None
        try:
            lnk = LnkReader()
            files = list(iter_shortcut_files())
            self.q.put(("progress_max", len(files)))
            added = 0
            for i, (path, place) in enumerate(files, 1):
                if self.cancel.is_set():
                    break
                try:
                    item = read_shortcut(path, place, lnk)
                    diagnose(item)
                except Exception:
                    self.q.put(("progress", i))
                    continue
                iid = str(added)
                added += 1
                self.q.put(("item", iid, item))
                try:
                    from core.icons import load_preview
                    img = load_preview(item.icon_source or item.target, item.icon_index, 32)
                except Exception:
                    img = None
                self.q.put(("preview", iid, img))
                self.q.put(("progress", i))
            self.q.put(("scan_done", None))
        except Exception:
            self.q.put(("log", "扫描出错：\n" + traceback.format_exc()))
            self.q.put(("scan_done", None))
        finally:
            if lnk is not None:
                lnk.close()
            pythoncom.CoUninitialize()

    # ------------------------------------------------------------ 修复

    def start_fix(self) -> None:
        if self.worker and self.worker.is_alive():
            return
        targets = [self.items[int(i)] for i in self.checked if i.isdigit()]
        if not targets:
            self.log("没有勾选可修复的项目")
            return
        self.counters = {"ok": 0, "fail": 0, "skip": 0}
        self.need_admin = False
        self._set_busy(True)
        self._show_progress(True, maximum=len(targets))
        self._refresh_buttons()
        self.worker = threading.Thread(target=self._fix_worker, args=(targets,), daemon=True)
        self.worker.start()

    def _fix_worker(self, targets: list) -> None:
        import pythoncom

        pythoncom.CoInitialize()
        lnk = None
        try:
            lnk = LnkReader()
            for n, item in enumerate(targets, 1):
                if self.cancel.is_set():
                    break
                done, msg = repair_item(item, lnk)
                key = "ok" if done else "fail"
                self.counters[key] += 1
                self.q.put(("updated", item, done, msg))
                self.q.put(("progress", n))
                self.q.put(("log", f"{'成功' if done else '失败'}：{item.name} —— {msg}"))
            refresh_shell()
            self.q.put(("fix_done", None))
        except Exception:
            self.q.put(("log", "修复出错：\n" + traceback.format_exc()))
            self.q.put(("fix_done", None))
        finally:
            if lnk is not None:
                lnk.close()
            pythoncom.CoUninitialize()

    # ------------------------------------------------------------ 列表绘制

    @staticmethod
    def _is_problem(item) -> bool:
        return item.status in (BAD, TARGET_MISSING)

    def _visible(self, item) -> bool:
        return self.show_all or self._is_problem(item) or item.status == FIXED

    def toggle_view(self) -> None:
        self.show_all = not self.show_all
        self._refresh_buttons()
        self._redraw_cards()

    def _schedule_redraw(self) -> None:
        if self._redraw_job is not None:
            return
        self._redraw_job = self.root.after(60, self._redraw_now)

    def _redraw_now(self) -> None:
        self._redraw_job = None
        self._redraw_cards()

    def _redraw_cards(self) -> None:
        c = self.canvas
        c.delete("all")
        self._rows = []
        width = c.winfo_width()
        if width <= 1:
            width = 780

        y = 6
        shown = 0
        for idx, item in enumerate(self.items):
            if not self._visible(item):
                continue
            iid = str(idx)
            self._draw_card(y, width, iid, item)
            self._rows.append((y, y + self.ROW_H, iid))
            y += self.ROW_H + self.GAP
            shown += 1

        if shown == 0:
            self._draw_empty(width)
            y = 220

        c.configure(scrollregion=(0, 0, width, y + 10))

    def _draw_card(self, y: int, width: int, iid: str, item) -> None:
        c = self.canvas
        tk = self.tk
        x0, x1 = self.PAD, width - self.PAD
        self._round_rect(x0, y, x1, y + self.ROW_H, 10, CARD, BORDER)

        # 图标
        photo = self.images.get(iid)
        cy = y + self.ROW_H // 2
        if photo is not None:
            c.create_image(x0 + 38, cy, image=photo)
        else:
            c.create_oval(x0 + 20, cy - 16, x0 + 52, cy + 16,
                          fill="#EEF2F7", outline="")
            c.create_text(x0 + 36, cy, text="?", fill="#9CA3AF", font=(FONT, 12, "bold"))

        # 名称（按像素宽度截断，避免换行压到第二行）
        maxpx = max(120, x1 - (x0 + 66) - 70)
        c.create_text(x0 + 66, y + 24, text=self._fit_text(item.name, maxpx),
                      anchor="w", font=self.font_name, fill=TEXT)

        # 位置 · 状态（状态单独上色）
        place_id = c.create_text(x0 + 66, y + 46, text=f"{item.place} · ",
                                 anchor="w", font=(FONT, 9), fill=MUTED)
        bbox = c.bbox(place_id)
        status_x = bbox[2] if bbox else x0 + 100
        label, color = self._status_label(item)
        c.create_text(status_x, y + 46, text=label, anchor="w",
                      font=(FONT, 9), fill=color)

        # 勾选状态
        if item.status == FIXED:
            c.create_text(x1 - 30, cy, text="✓", fill=GREEN, font=(FONT, 15, "bold"))
        elif item.status == BAD:
            checked = iid in self.checked
            c.create_text(x1 - 30, cy, text="☑" if checked else "☐",
                          fill=BLUE if checked else "#B0B7C3",
                          font=(FONT, 15))
        else:
            c.create_text(x1 - 30, cy, text="—", fill="#C4C9D2", font=(FONT, 12))

    @staticmethod
    def _status_label(item) -> tuple[str, str]:
        if item.status == FIXED:
            return "✓ 已修复", GREEN
        if item.status == BAD:
            return "● 图标丢失", RED
        if item.status == TARGET_MISSING:
            return "● 目标不存在", ORANGE
        return "● " + STATUS_TEXT.get(item.status, item.status), MUTED

    def _fit_text(self, text: str, maxpx: int) -> str:
        """名字太长就截断加省略号，保证不换行。"""
        if self.font_name.measure(text) <= maxpx:
            return text
        while text and self.font_name.measure(text + "…") > maxpx:
            text = text[:-1]
        return text + "…" if text else "…"

    def _round_rect(self, x0, y0, x1, y1, r, fill, outline) -> None:
        pts = [x0 + r, y0, x1 - r, y0, x1, y0, x1, y0 + r,
               x1, y1 - r, x1, y1, x1 - r, y1, x0 + r, y1,
               x0, y1, x0, y1 - r, x0, y0 + r, x0, y0]
        self.canvas.create_polygon(pts, smooth=True, splinesteps=16,
                                   fill=fill, outline=outline)

    def _draw_empty(self, width: int) -> None:
        c = self.canvas
        cx = width // 2
        c.create_oval(cx - 24, 80, cx + 24, 128, outline=GREEN, width=3)
        c.create_line(cx - 11, 104, cx - 2, 113, cx + 13, 94,
                      fill=GREEN, width=3, capstyle="round", joinstyle="round")
        c.create_text(cx, 160, text="没有发现问题", font=(FONT, 14, "bold"), fill=TEXT)
        c.create_text(cx, 188, text="桌面和开始菜单的图标都正常",
                      font=(FONT, 9), fill=MUTED)

    # ------------------------------------------------------------ 交互

    def _on_resize(self, event) -> None:
        if self._resize_job is not None:
            self.root.after_cancel(self._resize_job)
        self._resize_job = self.root.after(150, self._redraw_now)

    def _on_wheel(self, event) -> None:
        try:
            widget = self.root.winfo_containing(event.x_root, event.y_root)
        except Exception:
            return
        if widget is self.canvas:
            self.canvas.yview_scroll(-1 * (event.delta // 120), "units")

    def _on_click(self, event) -> None:
        y = self.canvas.canvasy(event.y)
        for y0, y1, iid in self._rows:
            if y0 <= y <= y1:
                self.toggle(iid)
                return

    def toggle(self, iid: str) -> None:
        item = self.items[int(iid)]
        if item.status != BAD:
            return
        if iid in self.checked:
            self.checked.discard(iid)
        else:
            self.checked.add(iid)
        self._refresh_buttons()
        self._redraw_cards()

    def on_rebuild_cache(self) -> None:
        from tkinter import messagebox

        if not messagebox.askyesno(
            "重建图标缓存",
            "这会关闭并重新启动资源管理器（桌面和任务栏会闪一下），\n"
            "并删除图标缓存文件。继续吗？",
        ):
            return
        self._set_busy(True)
        self._refresh_buttons()
        self.log("开始重建图标缓存…")

        def work():
            ok, msg = rebuild_icon_cache()
            self.q.put(("log", msg))
            self.q.put(("cache_done", ok))

        threading.Thread(target=work, daemon=True).start()

    def on_elevate(self) -> None:
        """公共开始菜单里的快捷方式在 ProgramData 下，必须管理员才能改。"""
        from tkinter import messagebox

        if _is_admin():
            messagebox.showinfo("已经是管理员", "当前已经是管理员权限，可以直接重试。")
            return
        if not messagebox.askyesno(
            "需要管理员权限",
            "以管理员身份重新打开本工具吗？\n\n"
            "只会用于修改系统级的快捷方式，Windows 会弹一次 UAC 确认框。",
        ):
            return

        import ctypes
        exe = sys.executable
        if getattr(sys, "frozen", False):
            params = ""
            workdir = os.path.dirname(exe)
        else:
            params = f'"{os.path.abspath(__file__)}"'
            workdir = os.path.dirname(os.path.abspath(__file__))

        r = ctypes.windll.shell32.ShellExecuteW(None, "runas", exe, params, workdir, 1)
        if r <= 32:
            messagebox.showerror("启动失败", "无法以管理员身份启动（UAC 被取消或系统限制）。")
        else:
            self.root.destroy()

    def on_close(self) -> None:
        from tkinter import messagebox

        if self.worker and self.worker.is_alive():
            if not messagebox.askyesno("退出", "任务还没结束，确定退出吗？"):
                return
            self.cancel.set()
        self.root.destroy()

    # ------------------------------------------------------------ 状态

    def _set_busy(self, busy: bool) -> None:
        self.busy = busy
        state = "disabled" if busy else "normal"
        for b in (self.btn_scan, self.btn_cache, self.btn_admin,
                  self.btn_toggle, self.btn_log):
            b.configure(state=state)

    def _show_progress(self, show: bool, maximum: int = 100) -> None:
        """扫描/修复时才显示进度条，平时收起来，免得像一条多余的分隔线。"""
        if show:
            self.progress.configure(maximum=max(maximum, 1), value=0)
            if not self.progress.winfo_ismapped():
                self.progress.pack(fill="x", padx=22, pady=(10, 0), before=self.body)
        else:
            self.progress.configure(value=0)
            self.progress.pack_forget()

    def _refresh_buttons(self) -> None:
        n = len(self.checked)
        self.btn_fix.configure(text=f"一键修复 ({n})")
        if self.busy or n == 0:
            self.btn_fix.configure(state="disabled", bg=BLUE_DISABLED, cursor="arrow")
        else:
            self.btn_fix.configure(state="normal", bg=BLUE, cursor="hand2")
        self.btn_toggle.configure(
            text="只看问题" if self.show_all else f"显示全部 ({len(self.items)})"
        )

    def _update_summary(self) -> None:
        problems = sum(1 for it in self.items if self._is_problem(it))
        repairable = sum(1 for it in self.items if it.status == BAD)
        if not self.scan_finished:
            text = f"正在扫描… 已发现 {problems} 个问题" if self.items else "正在扫描…"
        elif problems:
            tail = "，可以一键修复" if repairable else "，需要手动处理"
            text = f"发现 {problems} 个问题{tail}"
        else:
            text = "没有发现问题"
        self.summary.configure(text=text)
        self._refresh_buttons()

    # ------------------------------------------------------------ 队列消费

    def poll(self) -> None:
        from PIL import ImageTk

        try:
            while True:
                kind, *rest = self.q.get_nowait()

                if kind == "item":
                    iid, item = rest
                    self.items.append(item)
                    if item.status == BAD:
                        self.checked.add(iid)
                    self._update_summary()

                elif kind == "preview":
                    iid, img = rest
                    if img is None:
                        continue
                    self.images[iid] = ImageTk.PhotoImage(img)
                    if self.scan_finished:
                        self._schedule_redraw()

                elif kind == "progress_max":
                    self.progress.configure(maximum=max(rest[0], 1), value=0)

                elif kind == "progress":
                    self.progress.configure(value=rest[0])

                elif kind == "log":
                    self.log(rest[0])

                elif kind == "updated":
                    item, done, msg = rest
                    if not done and "管理员" in msg:
                        self.need_admin = True
                    if done:
                        item.status = FIXED
                        item.note = msg
                        for iid, it in enumerate(self.items):
                            if it is item:
                                self.checked.discard(str(iid))
                                break

                elif kind == "scan_done":
                    self.scan_finished = True
                    self._set_busy(False)
                    self._show_progress(False)
                    self._update_summary()
                    self._redraw_cards()
                    self.canvas.yview_moveto(0)
                    bad = sum(1 for it in self.items if it.status == BAD)
                    self.log(f"扫描完成：共 {len(self.items)} 个，需要修复 {bad} 个")

                elif kind == "fix_done":
                    self._set_busy(False)
                    self._show_progress(False)
                    self._redraw_cards()
                    c = self.counters
                    self.summary.configure(text=f"修复完成：成功 {c['ok']} 个，失败 {c['fail']} 个")
                    self._refresh_buttons()
                    self.log(f"修复结束：成功 {c['ok']}，失败 {c['fail']}")
                    from tkinter import messagebox
                    if self.need_admin:
                        if messagebox.askyesno(
                            "需要管理员权限",
                            f"成功 {c['ok']} 个，失败 {c['fail']} 个。\n\n"
                            "有系统级快捷方式需要管理员权限才能修改。\n"
                            "是否以管理员身份重新打开？",
                        ):
                            self.on_elevate()
                            return
                    else:
                        messagebox.showinfo(
                            "完成",
                            f"成功 {c['ok']} 个，失败 {c['fail']} 个。\n\n"
                            "如果任务栏上的图标没变，需要取消固定再重新固定。",
                        )

                elif kind == "cache_done":
                    self._set_busy(False)
                    self._refresh_buttons()
                    self.log("图标缓存重建完成" if rest[0] else "图标缓存重建失败")

        except queue.Empty:
            pass
        self.root.after(80, self.poll)

    def run(self) -> None:
        self.root.mainloop()


def _is_admin() -> bool:
    try:
        import ctypes
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def run_gui() -> None:
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    paths.ensure_dirs()
    App().run()


def main() -> int:
    # 管道里 stdio 默认按 GBK 编码，中文会变乱码；统一强制 UTF-8
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    paths.ensure_dirs()
    args = sys.argv[1:]
    if "--scan" in args:
        return cli_scan("--json" in args)
    if "--fix" in args:
        only = None
        if "--only" in args:
            only = args[args.index("--only") + 1]
        return cli_fix("--dry-run" in args, only)
    run_gui()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
