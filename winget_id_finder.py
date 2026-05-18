#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WinGet ID Finder – Tìm & gán WinGet ID cho phần mềm trong software_data.json
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess
import threading
import json
import os
import re
import shutil
from datetime import datetime

APP_TITLE = "WinGet ID Finder – Hỗ trợ LicenseChecker"
VERSION = "1.2"
DEFAULT_JSON = "software_data.json"
WINGET_TIMEOUT = 30

# ------------------------------------------------------------
# Hàm tìm đường dẫn winget
# ------------------------------------------------------------
def get_winget_path():
    """Tìm đường dẫn đến winget.exe."""
    # Thử trong PATH
    winget = shutil.which("winget")
    if winget:
        return winget
    # Thử trong App Installer
    local_apps = os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\WindowsApps\winget.exe")
    if os.path.exists(local_apps):
        return local_apps
    # Thử trong Program Files
    prog_files = r"C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*\winget.exe"
    import glob
    matches = glob.glob(prog_files)
    if matches:
        return matches[0]
    return None

# ------------------------------------------------------------
# Tìm WinGet ID cho tên phần mềm
# ------------------------------------------------------------
# ------------------------------------------------------------
# Tìm WinGet ID cho tên phần mềm (SỬA LỖI PARSING CỘT)
# ------------------------------------------------------------
def find_winget_id(software_name: str) -> str:
    """
    Trả về WinGet ID đầu tiên tìm được cho software_name.
    Sử dụng vị trí cột cố định từ dòng tiêu đề để cắt chuỗi chính xác.
    """
    winget = get_winget_path()
    if not winget:
        return ""

    try:
        cmd = [winget, "search", "--name", software_name,
               "--accept-source-agreements",
               "--disable-interactivity"]
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=WINGET_TIMEOUT,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        output = result.stdout
        lines = output.splitlines()

        id_start = -1
        version_start = -1

        for line in lines:
            if not line.strip():
                continue

            # Bước 1: Xác định vị trí các cột dựa vào dòng tiêu đề (Header)
            if "Name" in line and "Id" in line:
                id_start = line.find("Id")
                version_start = line.find("Version")
                continue

            # Bỏ qua dòng gạch ngang ngăn cách
            if line.strip().startswith("-"):
                continue

            # Bước 2: Nếu đã xác định được vị trí cột, tiến hành cắt chuỗi
            if id_start != -1 and version_start != -1:
                # Cắt lấy phần ID dựa theo khoảng vị trí từ cột Id đến cột Version
                p_id = line[id_start:version_start].strip()
                if p_id and '.' in p_id:
                    return p_id  # Trả về ID chính xác (Ví dụ: Google.Chrome)

        return ""
    except Exception:
        return ""

# ------------------------------------------------------------
# Các hàm JSON
# ------------------------------------------------------------
def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def save_json(data, path):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def entries_missing_winget(data):
    """Trả về các mục có free_catalog nhưng winget_id null hoặc rỗng."""
    missing = []
    for item in data:
        fc = item.get("free_catalog")
        if fc and isinstance(fc, dict):
            wid = fc.get("winget_id")
            if not wid or wid == "":
                missing.append(item)
    return missing

# ------------------------------------------------------------
# Giao diện chính (giữ nguyên cấu trúc nhưng dùng các hàm trên)
# ------------------------------------------------------------
class App(tk.Tk):
    PAD = 8

    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE}  v{VERSION}")
        self.geometry("1080x700")
        self.minsize(900, 550)
        self.configure(bg="#ECECEC")

        self.json_path = tk.StringVar(value=DEFAULT_JSON)
        self.data = []
        self._searching = False
        self._cancel = threading.Event()
        self._row_map = {}

        self._setup_styles()
        self._build_ui()

        if os.path.exists(DEFAULT_JSON):
            self._load_file(DEFAULT_JSON)

    def _setup_styles(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        bg = "#ECECEC"
        s.configure(".", background=bg, font=("Segoe UI", 9))
        s.configure("TFrame", background=bg)
        s.configure("TLabel", background=bg)
        s.configure("Accent.TButton", font=("Segoe UI", 9, "bold"), padding=(8, 4))
        s.configure("TButton", font=("Segoe UI", 9), padding=(6, 3))
        s.configure("Treeview", font=("Segoe UI", 9), rowheight=22, background="white")
        s.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"),
                    background="#C8C8C8", foreground="#1A1A1A", relief="flat")
        s.map("Treeview", background=[("selected", "#4A9CC8")])

    def _build_ui(self):
        # Header
        hdr = tk.Frame(self, bg="#3C7FC0", height=38)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        tk.Label(hdr, text=f"  ⚡  {APP_TITLE}",
                 bg="#3C7FC0", fg="white", font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=10, pady=6)
        self.lbl_badge = tk.Label(hdr, text="", bg="#3C7FC0", font=("Segoe UI", 8))
        self.lbl_badge.pack(side=tk.RIGHT, padx=10)

        # Notebook
        nb = ttk.Notebook(self)
        nb.pack(fill=tk.BOTH, expand=True, padx=self.PAD, pady=self.PAD)
        self.tab_batch = ttk.Frame(nb)
        self.tab_single = ttk.Frame(nb)
        nb.add(self.tab_batch, text="  🔄  Cập nhật hàng loạt  ")
        nb.add(self.tab_single, text="  🔍  Tìm kiếm đơn lẻ  ")
        self._build_tab_batch()
        self._build_tab_single()

    # ---------- TAB BATCH ----------
    def _build_tab_batch(self):
        f = self.tab_batch

        # File picker
        fpf = ttk.Frame(f)
        fpf.pack(fill=tk.X, padx=self.PAD, pady=(self.PAD, 2))
        ttk.Label(fpf, text="File JSON:").pack(side=tk.LEFT)
        ttk.Entry(fpf, textvariable=self.json_path, width=52).pack(side=tk.LEFT, padx=6, fill=tk.X, expand=True)
        ttk.Button(fpf, text="📂 Mở...", command=self._browse_json).pack(side=tk.LEFT, padx=2)
        ttk.Button(fpf, text="↺ Nạp", command=lambda: self._load_file(self.json_path.get())).pack(side=tk.LEFT, padx=2)

        # Toolbar
        bar = ttk.Frame(f)
        bar.pack(fill=tk.X, padx=self.PAD, pady=2)
        self.btn_run = ttk.Button(bar, text="⚡  Tìm WinGet ID cho tất cả (còn thiếu)",
                                   command=self._run_batch, style="Accent.TButton")
        self.btn_run.pack(side=tk.LEFT, padx=2)
        self.btn_stop = ttk.Button(bar, text="⛔  Dừng", command=self._stop_batch, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, padx=2)
        ttk.Button(bar, text="💾  Lưu JSON", command=self._save_json).pack(side=tk.LEFT, padx=2)
        ttk.Button(bar, text="📤  Xuất báo cáo", command=self._export_report).pack(side=tk.LEFT, padx=2)

        self.lbl_stat = ttk.Label(bar, text="")
        self.lbl_stat.pack(side=tk.RIGHT, padx=10)

        # Progress
        pbf = ttk.Frame(f)
        pbf.pack(fill=tk.X, padx=self.PAD, pady=(0, 2))
        self.pb = ttk.Progressbar(pbf, mode="determinate")
        self.pb.pack(fill=tk.X, side=tk.LEFT, expand=True)
        self.lbl_pb = ttk.Label(pbf, text="0 / 0", width=8)
        self.lbl_pb.pack(side=tk.LEFT, padx=6)

        # Treeview
        tf = ttk.Frame(f)
        tf.pack(fill=tk.BOTH, expand=True, padx=self.PAD, pady=4)
        cols = ("name", "old_id", "found_id", "status", "category")
        self.tv = ttk.Treeview(tf, columns=cols, show="headings", selectmode="extended")
        hdrs = [("name", "Tên phần mềm", 230), ("old_id", "WinGet ID cũ", 180),
                ("found_id", "WinGet ID tìm được", 200), ("status", "Trạng thái", 110),
                ("category", "Danh mục", 140)]
        for cid, txt, w in hdrs:
            self.tv.heading(cid, text=txt)
            self.tv.column(cid, width=w, minwidth=60)
        self.tv.tag_configure("ok", background="#EAFAF1", foreground="#1D8348")
        self.tv.tag_configure("skipped", background="#EBF5FB", foreground="#1A5276")
        self.tv.tag_configure("notfound", background="#FEF9E7", foreground="#A04000")
        self.tv.tag_configure("error", background="#FDEDEC", foreground="#A93226")
        self.tv.tag_configure("waiting", background="#F8F9FA", foreground="#717D7E")
        vsb = ttk.Scrollbar(tf, orient=tk.VERTICAL, command=self.tv.yview)
        hsb = ttk.Scrollbar(tf, orient=tk.HORIZONTAL, command=self.tv.xview)
        self.tv.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tv.pack(fill=tk.BOTH, expand=True)
        self.tv.bind("<Double-1>", self._tv_dblclick)
        self.tv.bind("<Button-3>", self._tv_rclick)

        # Log
        lf = ttk.LabelFrame(f, text=" Log ")
        lf.pack(fill=tk.X, padx=self.PAD, pady=(0, 4))
        self.log_text = tk.Text(lf, height=4, bg="#0D1117", fg="#C9D1D9",
                                font=("Consolas", 8), wrap=tk.WORD, state=tk.DISABLED, relief=tk.FLAT)
        vsb2 = ttk.Scrollbar(lf, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=vsb2.set)
        vsb2.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(fill=tk.X, padx=2, pady=2)
        self.log_text.tag_configure("ok", foreground="#3FB950")
        self.log_text.tag_configure("warn", foreground="#D29922")
        self.log_text.tag_configure("err", foreground="#F85149")
        self.log_text.tag_configure("info", foreground="#8B949E")

        # Status
        bot = tk.Frame(f, bg="#D4D0C8")
        bot.pack(fill=tk.X)
        self.sv_status = tk.StringVar(value="STATUS : Sẵn sàng.")
        tk.Label(bot, textvariable=self.sv_status, bg="#D4D0C8",
                 font=("Courier New", 8), anchor="w").pack(side=tk.LEFT, padx=6, pady=2)

    # ---------- TAB SINGLE ----------
    def _build_tab_single(self):
        f = self.tab_single
        sf = ttk.Frame(f)
        sf.pack(fill=tk.X, padx=self.PAD, pady=(self.PAD, 4))
        ttk.Label(sf, text="Tên phần mềm:").pack(side=tk.LEFT)
        self.sv_query = tk.StringVar()
        ent = ttk.Entry(sf, textvariable=self.sv_query, width=40, font=("Segoe UI", 10))
        ent.pack(side=tk.LEFT, padx=8, fill=tk.X, expand=True)
        ent.bind("<Return>", lambda _: self._single_search())
        ttk.Button(sf, text="🔍  Tìm kiếm", command=self._single_search,
                   style="Accent.TButton").pack(side=tk.LEFT, padx=2)
        ttk.Button(sf, text="✖  Xoá",
                   command=lambda: (self.sv_query.set(""), self.tv_single.delete(*self.tv_single.get_children()))).pack(side=tk.LEFT, padx=2)

        ttk.Label(f, text="★ Nhấp đúp vào kết quả để sao chép WinGet ID.",
                  foreground="#5D6D7E", font=("Segoe UI", 8)).pack(fill=tk.X, padx=self.PAD, pady=(0, 4))

        tf2 = ttk.Frame(f)
        tf2.pack(fill=tk.BOTH, expand=True, padx=self.PAD, pady=2)
        cols2 = ("name", "id", "version", "source")
        self.tv_single = ttk.Treeview(tf2, columns=cols2, show="headings")
        hdrs2 = [("name", "Tên phần mềm", 300), ("id", "WinGet ID", 250),
                 ("version", "Phiên bản", 100), ("source", "Nguồn", 80)]
        for cid, txt, w in hdrs2:
            self.tv_single.heading(cid, text=txt)
            self.tv_single.column(cid, width=w, minwidth=60)
        vsb3 = ttk.Scrollbar(tf2, orient=tk.VERTICAL, command=self.tv_single.yview)
        self.tv_single.configure(yscrollcommand=vsb3.set)
        vsb3.pack(side=tk.RIGHT, fill=tk.Y)
        self.tv_single.pack(fill=tk.BOTH, expand=True)
        self.tv_single.bind("<Double-1>", self._single_copy)

        # Áp dụng vào JSON
        apf = ttk.LabelFrame(f, text=" ✏️  Áp dụng kết quả vào software_data.json ")
        apf.pack(fill=tk.X, padx=self.PAD, pady=(4, self.PAD))
        r1 = ttk.Frame(apf)
        r1.pack(fill=tk.X, padx=8, pady=4)
        ttk.Label(r1, text="Tên phần mềm trong JSON:").pack(side=tk.LEFT)
        self.sv_apply_name = tk.StringVar()
        ttk.Entry(r1, textvariable=self.sv_apply_name, width=34).pack(side=tk.LEFT, padx=6)
        ttk.Label(r1, text="WinGet ID áp dụng:").pack(side=tk.LEFT, padx=(10,0))
        self.sv_apply_id = tk.StringVar()
        ttk.Entry(r1, textvariable=self.sv_apply_id, width=30).pack(side=tk.LEFT, padx=6)
        ttk.Button(r1, text="✅  Áp dụng", command=self._apply_single, style="Accent.TButton").pack(side=tk.LEFT, padx=4)

        ttk.Label(apf, text="★ Tên phần mềm phải khớp chính xác với field 'name' trong JSON.",
                  foreground="#5D6D7E", font=("Segoe UI", 8)).pack(fill=tk.X, padx=8, pady=(0,4))

        # Status
        bot2 = tk.Frame(f, bg="#D4D0C8")
        bot2.pack(fill=tk.X)
        self.sv_status2 = tk.StringVar(value="STATUS : Nhập tên và nhấn Tìm kiếm.")
        tk.Label(bot2, textvariable=self.sv_status2, bg="#D4D0C8",
                 font=("Courier New", 8), anchor="w").pack(side=tk.LEFT, padx=6, pady=2)

    # ---------- Các hàm xử lý ----------
    def _browse_json(self):
        p = filedialog.askopenfilename(title="Chọn file software_data.json", filetypes=[("JSON","*.json")])
        if p:
            self.json_path.set(p)
            self._load_file(p)

    def _load_file(self, path):
        if not os.path.exists(path):
            self.lbl_badge.config(text=f"❌ Không tìm thấy: {path}", fg="#FFAAAA")
            return
        try:
            self.data = load_json(path)
            self.json_path.set(path)
            missing = entries_missing_winget(self.data)
            total = sum(1 for e in self.data if e.get("free_catalog"))
            self.lbl_badge.config(text=f"✅ Đã nạp {len(self.data)} mục  |  📦 Catalog: {total}  |  ⚠️ Thiếu: {len(missing)}", fg="#A8F0B0")
            self._populate_batch_table()
            self.sv_status.set(f"STATUS : Đã nạp '{os.path.basename(path)}' – {len(missing)} bản ghi cần tìm.")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không đọc được JSON:\n{e}", parent=self)

    def _populate_batch_table(self):
        self.tv.delete(*self.tv.get_children())
        self._row_map.clear()
        for item in self.data:
            fc = item.get("free_catalog")
            if not (fc and isinstance(fc, dict)):
                continue
            wid = fc.get("winget_id", "")
            status = "✅ Đã có" if wid else "⏳ Chờ"
            tag = "skipped" if wid else "waiting"
            iid = self.tv.insert("", tk.END, tags=(tag,), values=(
                item["name"], wid, "", status, item.get("category", "")
            ))
            self._row_map[iid] = item
        total = len(self._row_map)
        has_wid = sum(1 for i in self.tv.get_children() if self.tv.set(i, "old_id"))
        self.lbl_stat.config(text=f"Tổng: {total}  ✅ {has_wid}  ⚠️ {total-has_wid}")

    def _save_json(self):
        if not self.data:
            messagebox.showinfo("Chú ý", "Chưa có dữ liệu để lưu.", parent=self)
            return
        path = self.json_path.get()
        if os.path.exists(path):
            bak = path + f".bak_{datetime.now():%Y%m%d_%H%M%S}"
            shutil.copy2(path, bak)
            self._log(f"Đã sao lưu: {bak}", "info")
        try:
            save_json(self.data, path)
            messagebox.showinfo("Đã lưu", f"✅  Đã lưu vào:\n{path}", parent=self)
            self.sv_status.set(f"STATUS : Đã lưu '{os.path.basename(path)}'.")
        except Exception as e:
            messagebox.showerror("Lỗi", str(e), parent=self)

    def _export_report(self):
        if not self.data:
            return
        path = filedialog.asksaveasfilename(
            title="Lưu báo cáo",
            initialfile=f"winget_report_{datetime.now():%Y%m%d_%H%M}.txt",
            defaultextension=".txt",
            filetypes=[("Text","*.txt")]
        )
        if not path:
            return
        lines = [
            "WINGET ID FINDER – BÁO CÁO",
            f"Ngày: {datetime.now():%d/%m/%Y %H:%M:%S}",
            f"File: {self.json_path.get()}", ""
        ]
        has_id = [e for e in self.data if e.get("free_catalog") and e["free_catalog"].get("winget_id")]
        no_id = [e for e in self.data if e.get("free_catalog") and not e["free_catalog"].get("winget_id")]
        lines += [f"=== ĐÃ CÓ WINGET ID ({len(has_id)}) ==="]
        for e in has_id:
            lines.append(f"  ✅  {e['name']:35s}  {e['free_catalog']['winget_id']}")
        lines += ["", f"=== CHƯA CÓ WINGET ID ({len(no_id)}) ==="]
        for e in no_id:
            lines.append(f"  ⚠️   {e['name']}")
        with open(path, "w", encoding="utf-8") as fp:
            fp.write("\n".join(lines))
        os.startfile(path)

    def _run_batch(self):
        if self._searching:
            return
        if not self.data:
            messagebox.showinfo("Chú ý", "Hãy nạp file JSON trước.", parent=self)
            return
        if not get_winget_path():
            messagebox.showerror("WinGet không khả dụng",
                "Không tìm thấy winget.exe.\nVui lòng cài đặt 'App Installer' từ Microsoft Store.", parent=self)
            return
        missing = entries_missing_winget(self.data)
        if not missing:
            messagebox.showinfo("Thông báo", "✅  Tất cả bản ghi đã có WinGet ID!", parent=self)
            return
        if not messagebox.askyesno("Xác nhận", f"Sẽ tìm WinGet ID cho {len(missing)} phần mềm. Bắt đầu?", parent=self):
            return
        self._cancel.clear()
        self._searching = True
        self.btn_run.configure(state=tk.DISABLED)
        self.btn_stop.configure(state=tk.NORMAL)
        self.pb.configure(maximum=len(missing), value=0)
        threading.Thread(target=self._batch_worker, args=(missing,), daemon=True).start()

    def _batch_worker(self, missing):
        total = len(missing)
        ok = notfound = 0
        for idx, item in enumerate(missing):
            if self._cancel.is_set():
                self.after(0, lambda: self._log("⛔  Đã dừng bởi người dùng.", "warn"))
                break
            name = item["name"]
            self.after(0, lambda n=name, i=idx+1, t=total: (
                self.sv_status.set(f"STATUS : [{i}/{t}] Đang tìm: {n}"),
                self._log(f"▶  [{i}/{t}] Tìm: {n}", "info")
            ))
            found_id = find_winget_id(name)
            if found_id:
                item["free_catalog"]["winget_id"] = found_id
                ok += 1
                self.after(0, lambda n=name, fid=found_id: self._log(f"   ✅ {n}  →  {fid}", "ok"))
                tag = "ok"
                status = "✅ Tìm được"
            else:
                notfound += 1
                self.after(0, lambda n=name: self._log(f"   ⚠️  Không tìm thấy: {n}", "warn"))
                tag = "notfound"
                status = "⚠️ Không tìm thấy"
            # Cập nhật bảng
            self.after(0, lambda n=name, fid=found_id, s=status, t=tag, i=idx+1:
                self._update_table_row(n, found_id, s, t, i, total, ok, notfound))
        self.after(0, lambda: self._batch_done(ok, notfound, total))

    def _update_table_row(self, name, found_id, status, tag, done, total, ok, notfound):
        for iid in self.tv.get_children():
            if self.tv.set(iid, "name") == name:
                self.tv.item(iid, tags=(tag,))
                self.tv.set(iid, "found_id", found_id if found_id else "—")
                self.tv.set(iid, "status", status)
                break
        self.pb.configure(value=done)
        self.lbl_pb.configure(text=f"{done} / {total}")
        self.lbl_stat.configure(text=f"Đang xử lý: {done}/{total}  ✅ {ok}  ⚠️ {notfound}")

    def _batch_done(self, ok, notfound, total):
        self._searching = False
        self.btn_run.configure(state=tk.NORMAL)
        self.btn_stop.configure(state=tk.DISABLED)
        self.pb.configure(value=total)
        self._log(f"\n✅ Hoàn tất: {ok} tìm được  ⚠️ {notfound} không thấy", "ok")
        self.sv_status.set(f"STATUS : Hoàn tất – {ok} tìm được, {notfound} không thấy. Nhấn '💾 Lưu JSON' để ghi lại.")
        if ok > 0:
            messagebox.showinfo("Hoàn tất", f"✅  Tìm được {ok}/{total} WinGet ID.\n\n⚠️  {notfound} phần mềm không tìm thấy.\n\nNhấn '💾 Lưu JSON' để ghi lại.", parent=self)

    def _stop_batch(self):
        self._cancel.set()
        self.btn_stop.configure(state=tk.DISABLED)

    def _log(self, text, tag="info"):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, text + "\n", tag)
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _tv_dblclick(self, event):
        row = self.tv.identify_row(event.y)
        if row:
            found = self.tv.set(row, "found_id")
            if found and found != "—":
                self.clipboard_clear()
                self.clipboard_append(found)
                self.sv_status.set(f"STATUS : Đã sao chép: {found}")

    def _tv_rclick(self, event):
        row = self.tv.identify_row(event.y)
        if not row:
            return
        self.tv.selection_set(row)
        name = self.tv.set(row, "name")
        m = tk.Menu(self, tearoff=0)
        m.add_command(label=f"🔍  Tìm lại '{name[:40]}'",
                      command=lambda: self._search_single_for(name))
        m.add_separator()
        m.add_command(label="✏️  Nhập ID thủ công",
                      command=lambda n=name, r=row: self._manual_edit(n, r))
        m.post(event.x_root, event.y_root)

    def _search_single_for(self, name):
        self.sv_query.set(name)
        self.sv_apply_name.set(name)
        self._single_search()

    def _manual_edit(self, name, iid):
        dlg = tk.Toplevel(self)
        dlg.title(f"Nhập WinGet ID – {name}")
        dlg.geometry("460x140")
        dlg.resizable(False, False)
        dlg.configure(bg="#ECECEC")
        dlg.grab_set()
        dlg.transient(self)
        tk.Label(dlg, text=f"WinGet ID cho '{name}':", bg="#ECECEC",
                 font=("Segoe UI", 9, "bold")).pack(padx=16, pady=(14,4), anchor="w")
        sv = tk.StringVar()
        ent = ttk.Entry(dlg, textvariable=sv, width=50, font=("Consolas", 9))
        ent.pack(padx=16, fill=tk.X)
        ent.focus()
        def apply():
            val = sv.get().strip()
            if val:
                self.tv.set(iid, "found_id", val)
                self.tv.set(iid, "status", "✅ Thủ công")
                self.tv.item(iid, tags=("ok",))
                # Cập nhật data
                for item in self.data:
                    if item["name"] == name:
                        fc = item.get("free_catalog")
                        if fc and isinstance(fc, dict):
                            fc["winget_id"] = val
                        break
            dlg.destroy()
        bf = ttk.Frame(dlg)
        bf.pack(fill=tk.X, padx=16, pady=8)
        ttk.Button(bf, text="✅ Áp dụng", command=apply, style="Accent.TButton").pack(side=tk.LEFT, padx=2)
        ttk.Button(bf, text="Huỷ", command=dlg.destroy).pack(side=tk.LEFT, padx=4)
        ent.bind("<Return>", lambda _: apply())

    # ---------- Tab Single ----------
    def _single_search(self):
        query = self.sv_query.get().strip()
        if not query:
            return
        if not get_winget_path():
            messagebox.showerror("WinGet không khả dụng", "Không tìm thấy winget.exe.", parent=self)
            return
        self.sv_status2.set(f"STATUS : Đang tìm '{query}'...")
        self.tv_single.delete(*self.tv_single.get_children())
        def worker():
            results = self._winget_search_full(query)
            self.after(0, lambda: self._populate_single(results))
        threading.Thread(target=worker, daemon=True).start()

    def _winget_search_full(self, query):
        """Chạy winget search --name query và trả về list kết quả chính xác."""
        winget = get_winget_path()
        if not winget:
            return []
        try:
            cmd = [winget, "search", "--name", query,
                   "--accept-source-agreements", "--disable-interactivity"]
            result = subprocess.run(cmd, capture_output=True, text=True,
                                    encoding="utf-8", errors="replace",
                                    timeout=WINGET_TIMEOUT,
                                    creationflags=subprocess.CREATE_NO_WINDOW)
            output = result.stdout
            lines = output.splitlines()
            results = []

            id_start = -1
            version_start = -1
            source_start = -1

            for line in lines:
                if not line.strip():
                    continue

                # Tìm vị trí index của các cột tiêu đề
                if "Name" in line and "Id" in line:
                    id_start = line.find("Id")
                    version_start = line.find("Version")
                    source_start = line.find("Source")
                    continue

                if line.strip().startswith("-"):
                    continue

                # Tiến hành cắt chuỗi theo vị trí đã định vị
                if id_start != -1 and version_start != -1:
                    p_name = line[:id_start].strip()
                    p_id = line[id_start:version_start].strip()

                    if source_start != -1:
                        p_version = line[version_start:source_start].strip()
                        p_source = line[source_start:].strip()
                    else:
                        p_version = line[version_start:].strip()
                        p_source = "winget"

                    if p_id and '.' in p_id:
                        results.append({
                            'id': p_id,
                            'name': p_name,
                            'version': p_version,
                            'source': p_source
                        })
            return results
        except Exception:
            return []

    def _populate_single(self, results):
        self.tv_single.delete(*self.tv_single.get_children())
        if not results:
            self.sv_status2.set(f"STATUS : Không tìm thấy kết quả cho '{self.sv_query.get()}'.")
            return
        for r in results:
            self.tv_single.insert("", tk.END, values=(r['name'], r['id'], r.get('version',''), r.get('source','')))
        best = results[0]['id'] if results else ""
        self.sv_apply_id.set(best)
        self.sv_apply_name.set(self.sv_query.get())
        self.sv_status2.set(f"STATUS : Tìm thấy {len(results)} kết quả. Nhấp đúp để sao chép ID.")

    def _single_copy(self, event):
        row = self.tv_single.identify_row(event.y)
        if row:
            wid = self.tv_single.set(row, "id")
            if wid:
                self.clipboard_clear()
                self.clipboard_append(wid)
                self.sv_apply_id.set(wid)
                self.sv_status2.set(f"STATUS : Đã sao chép: {wid}")

    def _apply_single(self):
        name = self.sv_apply_name.get().strip()
        new_id = self.sv_apply_id.get().strip()
        if not name or not new_id:
            messagebox.showwarning("Thiếu thông tin", "Hãy nhập đủ 'Tên phần mềm trong JSON' và 'WinGet ID'.", parent=self)
            return
        if not self.data:
            messagebox.showinfo("Chú ý", "Chưa nạp file JSON.", parent=self)
            return
        found_item = None
        for item in self.data:
            if item["name"].lower().strip() == name.lower().strip():
                found_item = item
                break
        if not found_item:
            messagebox.showwarning("Không tìm thấy", f"Không có mục nào tên '{name}' trong JSON.\nTên phải khớp chính xác.", parent=self)
            return
        fc = found_item.get("free_catalog")
        if not fc or not isinstance(fc, dict):
            found_item["free_catalog"] = {"desc": "", "url": "", "winget_id": new_id}
        else:
            old = fc.get("winget_id", "")
            if old and old != new_id:
                if not messagebox.askyesno("Xác nhận ghi đè", f"'{name}' đã có WinGet ID: {old}\n\nGhi đè thành: {new_id}?", parent=self):
                    return
            fc["winget_id"] = new_id
        # Cập nhật bảng batch
        for iid in self.tv.get_children():
            if self.tv.set(iid, "name") == found_item["name"]:
                self.tv.set(iid, "old_id", new_id)
                self.tv.set(iid, "found_id", new_id)
                self.tv.set(iid, "status", "✅ Áp dụng")
                self.tv.item(iid, tags=("ok",))
                break
        self.sv_status2.set(f"STATUS : ✅ Đã cập nhật '{name}' → {new_id}. Nhớ nhấn '💾 Lưu JSON'.")
        messagebox.showinfo("Đã cập nhật", f"✅  '{name}'  →  {new_id}\n\nNhớ nhấn '💾 Lưu JSON' để ghi lại!", parent=self)

if __name__ == "__main__":
    app = App()
    app.mainloop()