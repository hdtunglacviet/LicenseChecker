#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
╔══════════════════════════════════════════════════════════╗
║     KIỂM TRA BẢN QUYỀN PHẦN MỀM  v1.2                    ║
║     Quét - Phân tích - Đề xuất thay thế miễn phí         ║
╚══════════════════════════════════════════════════════════╝
Tác giả  : Huỳnh Đức Tùng
Yêu cầu  : Python 3.8+, Windows
Tuỳ chọn : pip install openpyxl  (để đọc DB từ Excel)
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import winreg
import threading
import webbrowser
import os
import json
import urllib.request
import urllib.error
import urllib.parse
from datetime import datetime
import winget_id_finder
import subprocess
# ── Tuỳ chọn: đọc DB từ Excel ─────────────────────────────
try:
    import openpyxl
    HAS_EXCEL = True
except ImportError:
    HAS_EXCEL = False

# ══════════════════════════════════════════════════════════
#  HẰNG SỐ
# ══════════════════════════════════════════════════════════
APP_TITLE = "Kiểm Tra Bản Quyền Phần Mềm - Huỳnh Đức Tùng"
VERSION   = "1.2"

# ── GitHub – chỉnh 2 dòng này cho đúng repo của bạn ──────
GITHUB_USER       = "HuynhDucTung"          # ← đổi thành username GitHub
GITHUB_REPO       = "LicenseChecker"         # ← đổi thành tên repo
GITHUB_BRANCH     = "main"
GITHUB_JSON_PATH  = "software_data.json"     # đường dẫn file trong repo
# Raw URL đọc JSON
GITHUB_RAW_URL    = (
    f"https://raw.githubusercontent.com/"
    f"{GITHUB_USER}/{GITHUB_REPO}/{GITHUB_BRANCH}/{GITHUB_JSON_PATH}"
)
# URL tạo Issue mới (không cần đăng nhập để mở, nhưng cần đăng nhập để submit)
GITHUB_ISSUES_URL = f"https://github.com/{GITHUB_USER}/{GITHUB_REPO}/issues/new"
# ─────────────────────────────────────────────────────────

LOCAL_JSON  = "software_data.json"
DB_TIMEOUT  = 8          # giây timeout khi tải online

LIC_OSS        = "Mã nguồn mở"
LIC_FREE       = "Miễn phí"
LIC_FREEMIUM   = "Freemium"
LIC_COMMERCIAL = "Thương mại"
LIC_TRIAL      = "Dùng thử"
LIC_UNKNOWN    = "Chưa rõ"

ALL_LICENSES = [LIC_OSS, LIC_FREE, LIC_FREEMIUM, LIC_COMMERCIAL, LIC_TRIAL, LIC_UNKNOWN]

ALL_CATEGORIES = [
    "Ứng dụng văn phòng",
    "Phần mềm Hệ thống",
    "Thiết kế đồ họa",
    "Chỉnh sửa ảnh",
    "Phần mềm Video",
    "Nghe nhạc",
    "Trình duyệt",
    "Chat & Gọi video",
    "Bảo mật",
    "Diệt Virus - Spyware",
    "Phần mềm mạng",
    "Hỗ trợ Download",
    "Phần mềm lập trình",
    "Quản lý Doanh nghiệp",
    "Dữ liệu - File",
    "Quản trị CSDL",
    "Giáo dục - Học tập",
    "Phần mềm cá nhân",
    "Mạng xã hội",
    "Quản lý Email",
    "Drivers - Firmware",
    "Khác",
]

ROW_COLORS = {
    LIC_OSS:        "#EAFAF1",
    LIC_FREE:       "#EBF5FB",
    LIC_FREEMIUM:   "#FEF9E7",
    LIC_COMMERCIAL: "#FDEDEC",
    LIC_TRIAL:      "#FEF5E7",
    LIC_UNKNOWN:    "#F8F9FA",
}

BADGE_COLORS = {
    LIC_OSS:        "#1D8348",
    LIC_FREE:       "#1A5276",
    LIC_FREEMIUM:   "#9A7D0A",
    LIC_COMMERCIAL: "#A93226",
    LIC_TRIAL:      "#A04000",
    LIC_UNKNOWN:    "#717D7E",
}


# ══════════════════════════════════════════════════════════
#  TẢI ONLINE / LOCAL
# ══════════════════════════════════════════════════════════

def fetch_online_db(url: str, timeout: int = DB_TIMEOUT):
    """
    Tải JSON từ GitHub raw URL.
    Trả về (data_list, None) hoặc (None, error_msg).
    """
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "LicenseChecker/1.2"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read().decode("utf-8")
        return json.loads(raw), None
    except urllib.error.URLError as e:
        return None, f"Không kết nối được: {e.reason}"
    except Exception as e:
        return None, str(e)


def parse_software_data(data: list):
    """
    Chuyển list JSON → software_db (dict keyword→info) + free_catalog.
    """
    software_db  = {}
    free_catalog = []
    for item in data:
        for kw in item.get("keywords", []):
            kw_lower = kw.lower().strip()
            if kw_lower:
                software_db[kw_lower] = {
                    "license":  item.get("license",  LIC_UNKNOWN),
                    "category": item.get("category", "Khác"),
                    "alt":      item.get("alt",      ""),
                    "alt_url":  item.get("alt_url",  ""),
                    "note":     item.get("note",     ""),
                }
        fc = item.get("free_catalog")
        if fc and isinstance(fc, dict):
            free_catalog.append({
                "name": item["name"],
                "cat":  item.get("category", "Khác"),
                "lic":  item.get("license",  LIC_UNKNOWN),
                "desc": fc.get("desc", ""),
                "url":  fc.get("url",  ""),
                "winget_id": fc.get("winget_id", None),
            })
    return software_db, free_catalog


def load_software_data(json_path: str = LOCAL_JSON):
    """Đọc JSON cục bộ, trả về (software_db, free_catalog, raw_list)."""
    if not os.path.exists(json_path):
        default = [{"id": "example", "name": "Example Software",
                    "keywords": ["example"], "license": LIC_UNKNOWN,
                    "category": "Khác", "alt": "", "alt_url": "",
                    "note": "Ví dụ.", "free_catalog": None}]
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(default, f, ensure_ascii=False, indent=2)
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    sdb, fc = parse_software_data(data)
    return sdb, fc, data


def save_local_json(data: list, json_path: str = LOCAL_JSON):
    """Ghi danh sách phần mềm ra file JSON cục bộ."""
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ══════════════════════════════════════════════════════════
#  QUÉT REGISTRY WINDOWS
# ══════════════════════════════════════════════════════════

def scan_registry() -> list:
    paths = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_CURRENT_USER,  r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
    ]
    result, seen = [], set()
    for hive, path in paths:
        try:
            key   = winreg.OpenKey(hive, path)
            count = winreg.QueryInfoKey(key)[0]
            for i in range(count):
                try:
                    sk = winreg.OpenKey(key, winreg.EnumKey(key, i))
                    try:
                        name = winreg.QueryValueEx(sk, "DisplayName")[0]
                        if not name or name in seen:
                            continue
                        seen.add(name)
                        def _get(k):
                            try: return winreg.QueryValueEx(sk, k)[0]
                            except: return ""
                        result.append({"name": name, "version": _get("DisplayVersion"),
                                       "publisher": _get("Publisher"), "install_date": _get("InstallDate")})
                    finally:
                        winreg.CloseKey(sk)
                except: pass
            winreg.CloseKey(key)
        except: pass
    return sorted(result, key=lambda x: x["name"].lower())


# ══════════════════════════════════════════════════════════
#  DIALOG THÊM MỚI PHẦN MỀM  (đầy đủ tất cả trường JSON)
# ══════════════════════════════════════════════════════════

class AddSoftwareDialog(tk.Toplevel):
    """
    Form đầy đủ để thêm phần mềm vào software_data.json:
      - Tất cả các trường của cấu trúc JSON
      - Nút tự động tìm WinGet ID
      - Kiểm tra trùng name & keywords trước khi lưu
      - Lưu local hoặc đóng góp lên GitHub Issues
    """

    def __init__(self, parent, raw_software: list, software_db: dict,
                 prefill_name: str = ""):
        super().__init__(parent)
        self.parent       = parent
        self.raw_software = raw_software
        self.software_db  = software_db
        self.result       = None
        self._finding_wg  = False       # đang tìm WinGet ID

        self.title("➕  Thêm / Đóng góp phần mềm vào cơ sở dữ liệu")
        self.geometry("800x800")
        #self.minsize(800, 800)
        self.configure(bg="#ECECEC")
        self.grab_set()
        self.transient(parent)
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)

        self._build(prefill_name)

        # Căn giữa
        self.update_idletasks()
        pw = parent.winfo_rootx() + (parent.winfo_width()  - self.winfo_width())  // 2
        ph = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{pw}+{ph}")

    # ── helper ──────────────────────────────────────────
    def _lbl(self, p, text, required=False):
        f = tk.Frame(p, bg="#ECECEC")
        tk.Label(f, text=text, bg="#ECECEC",
                 font=("Segoe UI", 9, "bold"), anchor="w").pack(side=tk.LEFT)
        if required:
            tk.Label(f, text=" *", bg="#ECECEC", fg="#C0392B",
                     font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)
        return f

    def _entry(self, p, sv, hint="", width=58):
        ttk.Entry(p, textvariable=sv, width=width, font=("Segoe UI", 9)).pack(
            fill=tk.X, padx=14, pady=(0, 1))
        if hint:
            tk.Label(p, text=f"  {hint}", bg="#ECECEC",
                     fg="#888", font=("Segoe UI", 7)).pack(fill=tk.X, padx=14, pady=(0,2))

    # ── Build UI ────────────────────────────────────────
    def _build(self, prefill_name):
        # ── Header ──
        hdr = tk.Frame(self, bg="#3C7FC0", height=36)
        hdr.pack(fill=tk.X); hdr.pack_propagate(False)
        tk.Label(hdr, text="  ➕  Đóng góp phần mềm mới vào cơ sở dữ liệu",
                 bg="#3C7FC0", fg="white",
                 font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=8, pady=4)

        # ── Scrollable body ──
        canvas = tk.Canvas(self, bg="#ECECEC", highlightthickness=0)
        vsb    = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        body = tk.Frame(canvas, bg="#ECECEC")
        body_id = canvas.create_window((0, 0), window=body, anchor="nw")

        def _on_configure(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(body_id, width=canvas.winfo_width())
        body.bind("<Configure>", _on_configure)
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfig(body_id, width=e.width))
        body.bind("<MouseWheel>",
                  lambda e: canvas.yview_scroll(-1*(e.delta//120), "units"))
        canvas.bind("<MouseWheel>",
                    lambda e: canvas.yview_scroll(-1*(e.delta//120), "units"))

        P = 14  # padx chung

        # ═══════════════════════════════
        # NHÓM 1 – Thông tin cơ bản
        # ═══════════════════════════════
        sec1 = tk.LabelFrame(body, text=" 📋  Thông tin cơ bản ",
                             bg="#ECECEC", font=("Segoe UI", 9, "bold"),
                             relief="groove", bd=2)
        sec1.pack(fill=tk.X, padx=P, pady=(10, 6))

        # Tên phần mềm
        self._lbl(sec1,"Tên phần mềm",required=True).pack(fill=tk.X,padx=P,pady=(8,0))
        self.sv_name = tk.StringVar(value=prefill_name)
        ttk.Entry(sec1, textvariable=self.sv_name, width=58,
                  font=("Segoe UI", 9)).pack(fill=tk.X, padx=P, pady=(0,1))
        tk.Label(sec1, text="  Ví dụ: Microsoft Word", bg="#ECECEC",
                 fg="#888", font=("Segoe UI", 7)).pack(fill=tk.X, padx=P, pady=(0,6))

        # Keywords
        self._lbl(sec1,"Keywords nhận dạng  –  cách nhau bởi dấu phẩy",
                  required=True).pack(fill=tk.X, padx=P, pady=(0,0))
        self.sv_kw = tk.StringVar(value=prefill_name.lower().strip())
        # Khi tên thay đổi → tự cập nhật keyword đầu
        self.sv_name.trace("w", self._sync_kw)
        ttk.Entry(sec1, textvariable=self.sv_kw, width=58,
                  font=("Segoe UI", 9)).pack(fill=tk.X, padx=P, pady=(0,1))
        tk.Label(sec1, text="  Ví dụ: microsoft word, ms word, winword",
                 bg="#ECECEC", fg="#888", font=("Segoe UI", 7)).pack(fill=tk.X,padx=P,pady=(0,6))

        # License + Category
        row2 = tk.Frame(sec1, bg="#ECECEC"); row2.pack(fill=tk.X, padx=P, pady=(0,8))

        lf = tk.Frame(row2, bg="#ECECEC"); lf.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._lbl(lf,"Loại bản quyền",required=True).pack(anchor="w")
        self.sv_lic = tk.StringVar(value=LIC_UNKNOWN)
        ttk.Combobox(lf, textvariable=self.sv_lic, values=ALL_LICENSES,
                     state="readonly", width=22).pack(anchor="w", pady=(2,0))

        tk.Frame(row2, bg="#ECECEC", width=16).pack(side=tk.LEFT)

        cf = tk.Frame(row2, bg="#ECECEC"); cf.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._lbl(cf,"Danh mục",required=True).pack(anchor="w")
        self.sv_cat = tk.StringVar(value="Khác")
        ttk.Combobox(cf, textvariable=self.sv_cat, values=ALL_CATEGORIES,
                     state="readonly", width=24).pack(anchor="w", pady=(2,0))

        # ═══════════════════════════════
        # NHÓM 2 – Ghi chú & Thay thế
        # ═══════════════════════════════
        sec2 = tk.LabelFrame(body, text=" 💡  Ghi chú & Phần mềm thay thế ",
                             bg="#ECECEC", font=("Segoe UI", 9, "bold"),
                             relief="groove", bd=2)
        sec2.pack(fill=tk.X, padx=P, pady=(0, 6))

        self._lbl(sec2,"Ghi chú ngắn").pack(fill=tk.X, padx=P, pady=(8,0))
        self.sv_note = tk.StringVar()
        self._entry(sec2, self.sv_note,
                    "Ví dụ: Bộ Office miễn phí mã nguồn mở ✓")

        r3 = tk.Frame(sec2, bg="#ECECEC"); r3.pack(fill=tk.X, padx=P, pady=(4,0))
        af = tk.Frame(r3, bg="#ECECEC"); af.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._lbl(af,"Phần mềm thay thế  (nếu có)").pack(anchor="w")
        self.sv_alt = tk.StringVar()
        ttk.Entry(af, textvariable=self.sv_alt, width=26,
                  font=("Segoe UI", 9)).pack(anchor="w", pady=(2,0))
        tk.Label(af, text="  Ví dụ: LibreOffice", bg="#ECECEC",
                 fg="#888", font=("Segoe UI", 7)).pack(anchor="w")

        tk.Frame(r3, bg="#ECECEC", width=16).pack(side=tk.LEFT)
        uf = tk.Frame(r3, bg="#ECECEC"); uf.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._lbl(uf,"URL tải thay thế  (nếu có)").pack(anchor="w")
        self.sv_alt_url = tk.StringVar()
        ttk.Entry(uf, textvariable=self.sv_alt_url, width=28,
                  font=("Segoe UI", 9)).pack(anchor="w", pady=(2,0))
        tk.Label(uf, text="  Ví dụ: https://libreoffice.org/download/", bg="#ECECEC",
                 fg="#888", font=("Segoe UI", 7)).pack(anchor="w", pady=(0,8))

        # ═══════════════════════════════
        # NHÓM 3 – Free Catalog (Tab 2)
        # ═══════════════════════════════
        sec3 = tk.LabelFrame(body,
                             text=" 📦  Free Catalog  –  Hiển thị ở Tab 'Kho phần mềm miễn phí' ",
                             bg="#ECECEC", font=("Segoe UI", 9, "bold"),
                             relief="groove", bd=2)
        sec3.pack(fill=tk.X, padx=P, pady=(0, 6))

        self.sv_fc_enable = tk.BooleanVar(value=False)
        chk = ttk.Checkbutton(sec3,
            text="Thêm phần mềm này vào Tab Kho miễn phí (free_catalog)",
            variable=self.sv_fc_enable,
            command=self._toggle_fc)
        chk.pack(anchor="w", padx=P, pady=(8,4))

        self._fc_frame = tk.Frame(sec3, bg="#ECECEC")
        self._fc_frame.pack(fill=tk.X, padx=P, pady=(0,4))

        # Mô tả trong catalog
        self._lbl(self._fc_frame,"Mô tả ngắn cho Tab Kho miễn phí").pack(fill=tk.X, pady=(0,0))
        self.sv_fc_desc = tk.StringVar()
        ttk.Entry(self._fc_frame, textvariable=self.sv_fc_desc, width=58,
                  font=("Segoe UI", 9)).pack(fill=tk.X, pady=(0,1))
        tk.Label(self._fc_frame, text="  Ví dụ: Bộ Office hoàn chỉnh thay Microsoft Office",
                 bg="#ECECEC", fg="#888", font=("Segoe UI", 7)).pack(fill=tk.X, pady=(0,6))

        # URL tải (trong catalog)
        self._lbl(self._fc_frame,"URL tải về").pack(fill=tk.X, pady=(0,0))
        self.sv_fc_url = tk.StringVar()
        ttk.Entry(self._fc_frame, textvariable=self.sv_fc_url, width=58,
                  font=("Segoe UI", 9)).pack(fill=tk.X, pady=(0,1))
        tk.Label(self._fc_frame, text="  Ví dụ: https://www.libreoffice.org/download/",
                 bg="#ECECEC", fg="#888", font=("Segoe UI", 7)).pack(fill=tk.X, pady=(0,6))

        # ── WinGet ID + nút tự động tìm ──
        wg_row = tk.Frame(self._fc_frame, bg="#ECECEC")
        wg_row.pack(fill=tk.X, pady=(0,4))

        wgl = tk.Frame(wg_row, bg="#ECECEC"); wgl.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._lbl(wgl,"WinGet ID").pack(anchor="w")
        self.sv_winget_id = tk.StringVar()
        self._entry_winget = ttk.Entry(wgl, textvariable=self.sv_winget_id,
                                        width=36, font=("Consolas", 9))
        self._entry_winget.pack(anchor="w", pady=(2,0))
        tk.Label(wgl, text="  Ví dụ: TheDocumentFoundation.LibreOffice",
                 bg="#ECECEC", fg="#888", font=("Segoe UI", 7)).pack(anchor="w")

        tk.Frame(wg_row, bg="#ECECEC", width=10).pack(side=tk.LEFT)

        wgr = tk.Frame(wg_row, bg="#ECECEC"); wgr.pack(side=tk.LEFT, anchor="center")
        self.btn_find_wg = ttk.Button(
            wgr, text="🔍  Tự động tìm\nWinGet ID",
            command=self.find_winget_id, width=18)
        self.btn_find_wg.pack(pady=(16,0))
        self.lbl_wg_status = tk.Label(
            wgr, text="", bg="#ECECEC",
            fg="#666", font=("Segoe UI", 7), wraplength=140, justify="center")
        self.lbl_wg_status.pack()

        # Ẩn fc_frame ban đầu
        self._toggle_fc()

        # ═══════════════════════════════
        # Thông báo lỗi
        # ═══════════════════════════════
        self.lbl_err = tk.Label(body, text="", bg="#FFFBE6", fg="#B03A2E",
                                 font=("Segoe UI", 8, "bold"), anchor="w",
                                 relief="flat", padx=10, pady=4)
        self.lbl_err.pack(fill=tk.X, padx=P, pady=(0,4))

        # ── Nút hành động (cố định dưới cùng) ──
        btn_frame = tk.Frame(self, bg="#D4D0C8", relief="sunken", bd=1)
        btn_frame.pack(fill=tk.X, side=tk.BOTTOM)

        ttk.Button(btn_frame, text="✅  Lưu vào máy",
                   command=self.on_save,
                   style="Accent.TButton").pack(side=tk.LEFT, padx=10, pady=6)
        ttk.Button(btn_frame, text="🌐  Lưu & Đóng góp GitHub",
                   command=self.on_save_and_contribute).pack(side=tk.LEFT, padx=4, pady=6)
        ttk.Button(btn_frame, text="Huỷ",
                   command=self.on_cancel).pack(side=tk.RIGHT, padx=10, pady=6)
        tk.Label(btn_frame,
                 text="★ 'Đóng góp' mở GitHub Issues – cần đăng nhập để gửi.",
                 bg="#D4D0C8", fg="#555", font=("Segoe UI", 7)).pack(
            side=tk.RIGHT, padx=8)

    # ── Toggle free_catalog frame ────────────────────────
    def _toggle_fc(self):
        if self.sv_fc_enable.get():
            self._fc_frame.pack(fill=tk.X, padx=14, pady=(0,4))
            # Tự điền desc và url từ note / alt_url nếu trống
            if not self.sv_fc_desc.get():
                self.sv_fc_desc.set(self.sv_note.get())
            if not self.sv_fc_url.get():
                self.sv_fc_url.set(self.sv_alt_url.get())
        else:
            self._fc_frame.pack_forget()

    # ── Tự đồng bộ keyword từ tên ───────────────────────
    def _sync_kw(self, *_):
        name = self.sv_name.get().strip().lower()
        current = self.sv_kw.get()
        # Chỉ tự điền nếu người dùng chưa chỉnh sửa thủ công
        if not current or current == getattr(self, "_last_auto_kw", ""):
            self._last_auto_kw = name
            self.sv_kw.set(name)

    # ── Tự động tìm WinGet ID ───────────────────────────
    # ------------------------------------------------------------
    # Tìm WinGet ID cho tên phần mềm (SỬA LỖI PARSING CỘT)
    # ------------------------------------------------------------
    def find_winget_id(software_name: str) -> str:
        """
        Trả về WinGet ID đầu tiên tìm được cho software_name.
        Sử dụng vị trí cột cố định từ dòng tiêu đề để cắt chuỗi chính xác.
        """
        winget = winget_id_finder.get_winget_path()
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
                timeout=winget_id_finder.WINGET_TIMEOUT,
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

    def _on_wg_found(self, name, best_id, results):
        self._finding_wg = False
        self.btn_find_wg.configure(state=tk.NORMAL, text="🔍  Tự động tìm\nWinGet ID")

        if not results:
            self.lbl_wg_status.config(
                text=f"⚠ Không tìm thấy\nkết quả cho '{name}'", fg="#A04000")
            return

        if best_id:
            self.sv_winget_id.set(best_id)
            self.lbl_wg_status.config(
                text=f"✅ Tìm được:\n{best_id}", fg="#1D8348")
        else:
            self.lbl_wg_status.config(
                text=f"⚠ {len(results)} kết quả\nnhưng không khớp", fg="#9A7D0A")

        # Nếu > 1 kết quả, hỏi người dùng chọn
        if len(results) > 1:
            self._show_wg_picker(name, results)

    def _show_wg_picker(self, name, results):
        """Popup chọn WinGet ID từ danh sách kết quả."""
        dlg = tk.Toplevel(self)
        dlg.title(f"Chọn WinGet ID – {name}")
        dlg.geometry("560x360"); dlg.resizable(True, True)
        dlg.configure(bg="#ECECEC"); dlg.grab_set(); dlg.transient(self)

        tk.Label(dlg, text=f"Kết quả tìm kiếm cho: '{name}'",
                 bg="#ECECEC", font=("Segoe UI", 9, "bold")).pack(
            padx=12, pady=(10,4), anchor="w")

        tf = tk.Frame(dlg); tf.pack(fill=tk.BOTH, expand=True, padx=12, pady=4)
        cols = ("name","id","version","source")
        tv   = ttk.Treeview(tf, columns=cols, show="headings", height=12)
        for cid, txt, w in [("name","Tên",220),("id","WinGet ID",200),
                              ("version","Phiên bản",80),("source","Nguồn",60)]:
            tv.heading(cid, text=txt); tv.column(cid, width=w)
        for i, r in enumerate(results[:25]):
            tag = "best" if i == 0 else ""
            tv.insert("", tk.END, tags=(tag,),
                      values=(r.get("name",""), r.get("id",""),
                               r.get("version",""), r.get("source","")))
        tv.tag_configure("best", background="#EAFAF1", foreground="#1D8348")
        vsb = ttk.Scrollbar(tf, orient=tk.VERTICAL, command=tv.yview)
        tv.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y); tv.pack(fill=tk.BOTH, expand=True)

        def apply():
            sel = tv.selection()
            if sel:
                wid = tv.set(sel[0], "id")
                self.sv_winget_id.set(wid)
                self.lbl_wg_status.config(text=f"✅ Đã chọn:\n{wid}", fg="#1D8348")
            dlg.destroy()

        bf = tk.Frame(dlg, bg="#ECECEC"); bf.pack(pady=8)
        ttk.Button(bf, text="✅  Dùng ID này",
                   command=apply, style="Accent.TButton").pack(side=tk.LEFT, padx=6)
        ttk.Button(bf, text="Bỏ qua",
                   command=dlg.destroy).pack(side=tk.LEFT, padx=4)
        tv.bind("<Double-1>", lambda _: apply())

    # ── Kiểm tra trùng ──────────────────────────────────
    def _check_duplicates(self, name: str, keywords: list):
        name_lower = name.lower().strip()
        kws_lower  = [k.lower().strip() for k in keywords if k.strip()]
        for item in self.raw_software:
            if item.get("name","").lower().strip() == name_lower:
                return (f"Tên '{item['name']}' đã có trong CSDL!\n"
                        f"(id: {item.get('id','?')})")
        existing = {kw.lower().strip()
                    for item in self.raw_software
                    for kw in item.get("keywords", [])}
        for kw in kws_lower:
            if kw in existing:
                return (f"Từ khoá '{kw}' đã tồn tại!\n"
                        "Dùng từ khoá khác hoặc cập nhật bản ghi hiện có.")
        return None

    # ── Đọc & validate form ──────────────────────────────
    def _read_form(self):
        name   = self.sv_name.get().strip()
        kw_raw = self.sv_kw.get().strip()

        if not name:
            self.lbl_err.config(text="⚠  Tên phần mềm không được để trống!")
            return None
        if not kw_raw:
            self.lbl_err.config(text="⚠  Phải có ít nhất 1 từ khoá nhận dạng!")
            return None

        keywords = [k.strip() for k in kw_raw.split(",") if k.strip()]
        if not keywords:
            self.lbl_err.config(text="⚠  Từ khoá không hợp lệ!")
            return None

        err = self._check_duplicates(name, keywords)
        if err:
            self.lbl_err.config(text=f"⚠  {err}")
            return None

        # Xây free_catalog
        fc = None
        if self.sv_fc_enable.get():
            fc = {
                "desc":      self.sv_fc_desc.get().strip(),
                "url":       self.sv_fc_url.get().strip(),
                "winget_id": self.sv_winget_id.get().strip(),
            }

        self.lbl_err.config(text="")
        return {
            "id":           name.lower().replace(" ", "-"),
            "name":         name,
            "keywords":     keywords,
            "license":      self.sv_lic.get(),
            "category":     self.sv_cat.get(),
            "alt":          self.sv_alt.get().strip(),
            "alt_url":      self.sv_alt_url.get().strip(),
            "note":         self.sv_note.get().strip(),
            "free_catalog": fc,
        }

    # ── Lưu vào máy ──────────────────────────────────────
    def on_save(self):
        entry = self._read_form()
        if entry is None: return
        self.raw_software.append(entry)
        try:
            save_local_json(self.raw_software, LOCAL_JSON)
        except Exception as e:
            messagebox.showerror("Lỗi ghi file", str(e), parent=self); return
        self.result = entry
        messagebox.showinfo("Đã lưu",
            f"✅  Đã thêm '{entry['name']}' vào software_data.json.\n\n"
            "Nhấn 'Quét lại' ở cửa sổ chính để áp dụng ngay.",
            parent=self)
        self.destroy()

    # ── Lưu & mở GitHub Issue ────────────────────────────
    def on_save_and_contribute(self):
        entry = self._read_form()
        if entry is None: return
        self.raw_software.append(entry)
        try:
            save_local_json(self.raw_software, LOCAL_JSON)
        except Exception as e:
            messagebox.showerror("Lỗi ghi file", str(e), parent=self); return
        self.result = entry

        entry_json = json.dumps(entry, ensure_ascii=False, indent=2)
        title = f"[DB Contribution] Thêm mới: {entry['name']}"
        body  = (
            f"## Đề xuất thêm phần mềm mới\n\n"
            f"**Tên:** {entry['name']}  \n"
            f"**Loại bản quyền:** {entry['license']}  \n"
            f"**Danh mục:** {entry['category']}  \n"
            f"**Keywords:** {', '.join(entry['keywords'])}  \n"
            f"**Thay thế bởi:** {entry.get('alt') or '_(không có)_'}  \n"
            f"**Link thay thế:** {entry.get('alt_url') or '_(không có)_'}  \n"
            f"**Ghi chú:** {entry.get('note') or '_(không có)_'}  \n"
            f"**Free Catalog:** {'Có' if entry.get('free_catalog') else 'Không'}  \n"
            f"**WinGet ID:** {(entry.get('free_catalog') or {}).get('winget_id','_(không có)_')}  \n\n"
            f"### JSON entry\n\n```json\n{entry_json}\n```\n\n"
            f"---\n*Gửi từ LicenseChecker v{VERSION}*"
        )
        params = urllib.parse.urlencode({"title": title, "body": body})
        webbrowser.open(f"{GITHUB_ISSUES_URL}?{params}")
        messagebox.showinfo("Đã lưu & mở GitHub",
            f"✅  Đã lưu '{entry['name']}' vào máy.\n\n"
            "🌐  Trình duyệt vừa mở GitHub Issues.\n"
            "Đăng nhập và nhấn 'Submit new issue' để gửi đóng góp.",
            parent=self)
        self.destroy()

    def on_cancel(self):
        self.result = None
        self.destroy()


# ══════════════════════════════════════════════════════════
#  GIAO DIỆN CHÍNH
# ══════════════════════════════════════════════════════════

class App(tk.Tk):
    PAD = 6

    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE}  –  v{VERSION}")
        self.geometry("1160x700")
        self.minsize(950, 560)
        self.configure(bg="#ECECEC")

        self.all_apps:     list = []
        self.free_checked: set  = set()
        self._scanning:    bool = False
        self._db_source:   str  = "local"   # "online" hoặc "local"

        # ── Tải DB (online → local fallback) ────────────
        self.software_db  = {}
        self.free_catalog = []
        self.raw_software = []
        self._load_db_with_fallback()

        self._setup_styles()
        self._build_ui()
        self.after(400, self.do_scan)

    # ── Tải DB ──────────────────────────────────────────
    def _load_db_with_fallback(self):
        """Thử tải online → fallback local, cập nhật self.*"""
        data_online, err = fetch_online_db(GITHUB_RAW_URL, timeout=DB_TIMEOUT)
        if data_online:
            self.raw_software        = data_online
            self.software_db, self.free_catalog = parse_software_data(data_online)
            self._db_source          = "online"
            # Luôn cache lại local
            try:
                save_local_json(data_online, LOCAL_JSON)
            except Exception:
                pass
        else:
            # Fallback: đọc file cục bộ
            try:
                self.software_db, self.free_catalog, self.raw_software = \
                    load_software_data(LOCAL_JSON)
                self._db_source = "local"
            except Exception as e:
                messagebox.showerror(
                    "Lỗi",
                    f"Không thể tải dữ liệu:\n• Online: {err}\n• Local: {e}\n\nChương trình sẽ thoát."
                )
                self.destroy()

    def _reload_db_background(self):
        """Tải lại DB online trong background thread, cập nhật UI sau."""
        def worker():
            data, err = fetch_online_db(GITHUB_RAW_URL, timeout=DB_TIMEOUT)
            self.after(0, lambda: self._on_db_reload_done(data, err))

        self.sv_status1.set("STATUS : ĐANG TẢI LẠI DATABASE TỪ GITHUB...")
        threading.Thread(target=worker, daemon=True).start()

    def _on_db_reload_done(self, data, err):
        if data:
            self.raw_software = data
            self.software_db, self.free_catalog = parse_software_data(data)
            self._db_source = "online"
            try: save_local_json(data, LOCAL_JSON)
            except Exception: pass
            self._refresh_tab1()
            self._fill_tv2(self.cur_cat2)
            self._update_db_badge()
            self.sv_status1.set("STATUS : Đã tải lại DB từ GitHub thành công.")
        else:
            self.sv_status1.set(f"STATUS : Tải GitHub thất bại ({err}) – đang dùng bản local.")

    # ── Styles ──────────────────────────────────────────
    def _setup_styles(self):
        s = ttk.Style(self); s.theme_use("clam")
        bg = "#ECECEC"
        s.configure(".",              background=bg, font=("Segoe UI", 9))
        s.configure("TFrame",         background=bg)
        s.configure("TLabel",         background=bg)
        s.configure("Bold.TLabel",    background=bg, font=("Segoe UI", 10, "bold"))
        s.configure("Small.TLabel",   background=bg, font=("Segoe UI", 8))
        s.configure("Status.TLabel",  background="#D4D0C8", font=("Courier New", 8),
                    foreground="#1A1A1A", relief="sunken", padding=(4, 2))
        s.configure("TButton",        font=("Segoe UI", 9), padding=(6, 3))
        s.configure("Cat.TButton",    font=("Segoe UI", 9), padding=(8, 3))
        s.configure("CatOn.TButton",  font=("Segoe UI", 9, "bold"), padding=(8, 3))
        s.configure("Accent.TButton", font=("Segoe UI", 9, "bold"), padding=(8, 4))
        s.configure("TNotebook",      background=bg)
        s.configure("TNotebook.Tab",  font=("Segoe UI", 9), padding=(12, 4))
        s.configure("Treeview",       font=("Segoe UI", 9), rowheight=22, background="white")
        s.configure("Treeview.Heading",
                    font=("Segoe UI", 9, "bold"), background="#C8C8C8",
                    foreground="#1A1A1A", relief="flat")
        s.map("Treeview",     background=[("selected", "#4A9CC8")])
        s.map("TNotebook.Tab",background=[("selected", "#FFFFFF")])

    # ── Layout chính ────────────────────────────────────
    def _build_ui(self):
        hdr = tk.Frame(self, bg="#3C7FC0", height=36)
        hdr.pack(fill=tk.X); hdr.pack_propagate(False)
        tk.Label(hdr, text=f"  🔒 {APP_TITLE}  –  Phân tích bản quyền & đề xuất thay thế miễn phí",
                 bg="#3C7FC0", fg="white",
                 font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=8, pady=4)

        # Badge nguồn DB
        self.lbl_db_badge = tk.Label(hdr, text="", bg="#3C7FC0",
                                      font=("Segoe UI", 8, "bold"))
        self.lbl_db_badge.pack(side=tk.RIGHT, padx=8)
        self._update_db_badge()

        nb = ttk.Notebook(self)
        nb.pack(fill=tk.BOTH, expand=True, padx=self.PAD, pady=self.PAD)
        self.tab1 = ttk.Frame(nb); self.tab2 = ttk.Frame(nb)
        nb.add(self.tab1, text="  🔍  Quét & Phân Tích Bản Quyền  ")
        nb.add(self.tab2, text="  📦  Kho Phần Mềm Miễn Phí  ")
        self._build_tab1(); self._build_tab2()

    def _update_db_badge(self):
        if self._db_source == "online":
            self.lbl_db_badge.config(text="🌐 DB: GitHub (online)  ", fg="#A8F0B0")
        else:
            self.lbl_db_badge.config(text="💾 DB: Local (offline)  ", fg="#FFD080")

    # ══════════════════════════════════════════════════════
    #  TAB 1 – QUÉT & PHÂN TÍCH
    # ══════════════════════════════════════════════════════
    def _build_tab1(self):
        f = self.tab1

        bar = ttk.Frame(f); bar.pack(fill=tk.X, padx=self.PAD, pady=(self.PAD, 2))
        ttk.Button(bar, text="⟳  Quét lại",         command=self.do_scan,
                   style="Accent.TButton").pack(side=tk.LEFT, padx=2)
        ttk.Button(bar, text="📊  Xuất báo cáo",    command=self.export_report).pack(side=tk.LEFT, padx=2)
        ttk.Button(bar, text="📂  Nạp DB từ Excel", command=self.load_excel).pack(side=tk.LEFT, padx=2)
        ttk.Button(bar, text="🔄  Tải lại DB",      command=self._reload_db_background).pack(side=tk.LEFT, padx=2)
        ttk.Separator(bar, orient=tk.VERTICAL).pack(side=tk.LEFT, padx=8, fill=tk.Y)
        ttk.Label(bar, text="Tìm:").pack(side=tk.LEFT)
        self.sv_search = tk.StringVar()
        self.sv_search.trace("w", lambda *_: self._refresh_tab1())
        ttk.Entry(bar, textvariable=self.sv_search, width=26).pack(side=tk.LEFT, padx=4)

        leg = ttk.Frame(f); leg.pack(fill=tk.X, padx=self.PAD, pady=1)
        for lic, col in BADGE_COLORS.items():
            tk.Label(leg, text=f"▌ {lic}", fg=col, bg="#ECECEC",
                     font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=5)

        # ── Bộ lọc: Danh mục + Bản quyền (Combobox) ──
        flt = ttk.Frame(f); flt.pack(fill=tk.X, padx=self.PAD, pady=(2, 0))

        ttk.Label(flt, text="Danh mục:").pack(side=tk.LEFT)
        self.cats1 = [
            "Tất cả",
            "Ứng dụng văn phòng", "Phần mềm Hệ thống",
            "Thiết kế đồ họa",    "Chỉnh sửa ảnh",
            "Phần mềm Video",     "Nghe nhạc",
            "Trình duyệt",        "Chat & Gọi video",
            "Bảo mật",            "Diệt Virus - Spyware",
            "Phần mềm mạng",      "Hỗ trợ Download",
            "Phần mềm lập trình", "Quản lý Doanh nghiệp",
            "Dữ liệu - File",     "Quản trị CSDL",
            "Giáo dục - Học tập", "Phần mềm cá nhân",
            "Mạng xã hội",        "Quản lý Email",
            "Drivers - Firmware", "Khác",
        ]
        self.sv_cat1 = tk.StringVar(value="Tất cả")
        self.cur_cat1 = "Tất cả"
        cb_cat1 = ttk.Combobox(flt, textvariable=self.sv_cat1,
                                values=self.cats1, state="readonly", width=24)
        cb_cat1.pack(side=tk.LEFT, padx=(4, 12))
        cb_cat1.bind("<<ComboboxSelected>>",
                     lambda e: self._set_cat1(self.sv_cat1.get()))

        ttk.Label(flt, text="Bản quyền:").pack(side=tk.LEFT)
        self.lic_opts1 = ["Tất cả"] + ALL_LICENSES
        self.sv_lic1   = tk.StringVar(value="Tất cả")
        cb_lic1 = ttk.Combobox(flt, textvariable=self.sv_lic1,
                                values=self.lic_opts1, state="readonly", width=18)
        cb_lic1.pack(side=tk.LEFT, padx=(4, 12))
        cb_lic1.bind("<<ComboboxSelected>>", lambda e: self._refresh_tab1())

        ttk.Button(flt, text="✖ Xoá bộ lọc",
                   command=self._clear_filter1).pack(side=tk.LEFT, padx=2)

        # Thống kê nhanh
        sm = ttk.Frame(f); sm.pack(fill=tk.X, padx=self.PAD, pady=(2, 0))
        self.lbl_ttl  = ttk.Label(sm, text="Tổng: –"); self.lbl_ttl.pack(side=tk.LEFT, padx=8)
        self.lbl_free = tk.Label(sm, text="Miễn phí/OSS: –", bg="#ECECEC",
                                  fg=BADGE_COLORS[LIC_OSS]); self.lbl_free.pack(side=tk.LEFT, padx=6)
        self.lbl_comm = tk.Label(sm, text="Thương mại: –",   bg="#ECECEC",
                                  fg=BADGE_COLORS[LIC_COMMERCIAL]); self.lbl_comm.pack(side=tk.LEFT, padx=6)
        self.lbl_unk  = tk.Label(sm, text="Chưa rõ: –",      bg="#ECECEC",
                                  fg=BADGE_COLORS[LIC_UNKNOWN]); self.lbl_unk.pack(side=tk.LEFT, padx=6)

        tf = ttk.Frame(f); tf.pack(fill=tk.BOTH, expand=True, padx=self.PAD, pady=4)
        cols = ("name", "version", "license", "category", "alt", "note")
        self.tv1 = ttk.Treeview(tf, columns=cols, show="headings", selectmode="extended")
        hdrs = [("name","Tên phần mềm",240), ("version","Phiên bản",80),
                ("license","Loại bản quyền",115), ("category","Danh mục",100),
                ("alt","Phần mềm thay thế",180), ("note","Ghi chú",340)]
        for cid, txt, w in hdrs:
            self.tv1.heading(cid, text=txt, command=lambda c=cid: self._sort_tv1(c))
            self.tv1.column(cid, width=w, minwidth=60)
        for lic, bg in ROW_COLORS.items():
            self.tv1.tag_configure(lic, background=bg)
        vsb = ttk.Scrollbar(tf, orient=tk.VERTICAL,   command=self.tv1.yview)
        hsb = ttk.Scrollbar(tf, orient=tk.HORIZONTAL, command=self.tv1.xview)
        self.tv1.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y); hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tv1.pack(fill=tk.BOTH, expand=True)
        self.tv1.bind("<Double-1>", self._tv1_dblclick)
        self.tv1.bind("<Button-3>", self._tv1_rclick)

        bot = tk.Frame(f, bg="#D4D0C8", bd=1, relief="sunken"); bot.pack(fill=tk.X)
        self.sv_status1 = tk.StringVar(value="STATUS : ĐANG ĐỢI USER.")
        tk.Label(bot, textvariable=self.sv_status1, bg="#D4D0C8",
                 font=("Courier New", 8), anchor="w").pack(side=tk.LEFT, padx=6, pady=2)
        ttk.Button(bot, text="[ Chọn cần chú ý ]",  command=self.select_commercial).pack(side=tk.RIGHT, padx=4, pady=2)
        ttk.Button(bot, text="[ Mở thay thế ]",     command=self.open_alt).pack(side=tk.RIGHT, padx=2, pady=2)

        self._sort_col1, self._sort_rev1 = "name", False

    def _clear_filter1(self):
        self.sv_cat1.set("Tất cả")
        self.sv_lic1.set("Tất cả")
        self.cur_cat1 = "Tất cả"
        self._refresh_tab1()

    def _set_cat1(self, cat, refresh=True):
        self.cur_cat1 = cat
        self.sv_cat1.set(cat)
        if refresh: self._refresh_tab1()

    def _refresh_tab1(self):
        kw   = self.sv_search.get().lower()
        cat  = self.sv_cat1.get()
        lic  = self.sv_lic1.get()
        rows = [
            a for a in self.all_apps
            if (cat == "Tất cả" or a.get("category") == cat)
            and (lic == "Tất cả" or a.get("license") == lic)
            and (not kw or kw in a["name"].lower())
        ]
        self._fill_tv1(rows)

    def _fill_tv1(self, rows):
        self.tv1.delete(*self.tv1.get_children())
        for a in rows:
            lic = a.get("license", LIC_UNKNOWN)
            self.tv1.insert("", tk.END, tags=(lic,), values=(
                a["name"], a.get("version",""), lic,
                a.get("category","Khác"), a.get("alt",""), a.get("note",""),
            ))

    def _sort_tv1(self, col):
        self._sort_rev1 = not self._sort_rev1 if col == self._sort_col1 else False
        self._sort_col1 = col
        rows = [(self.tv1.set(i, col), i) for i in self.tv1.get_children()]
        rows.sort(key=lambda x: x[0].lower(), reverse=self._sort_rev1)
        for pos, (_, i) in enumerate(rows): self.tv1.move(i, "", pos)

    def _tv1_dblclick(self, _):
        sel = self.tv1.selection()
        if not sel: return
        name = str(self.tv1.item(sel[0])["values"][0]).lower()
        for key, info in self.software_db.items():
            if key in name and info.get("alt_url"):
                webbrowser.open(info["alt_url"]); return

    def _tv1_rclick(self, event):
        row = self.tv1.identify_row(event.y)
        if not row: return
        self.tv1.selection_set(row)
        v    = self.tv1.item(row)["values"]
        name = str(v[0])
        lic  = str(v[2]) if len(v) > 2 else LIC_UNKNOWN

        m = tk.Menu(self, tearoff=0)
        m.add_command(
            label="📋  Sao chép tên",
            command=lambda: (self.clipboard_clear(), self.clipboard_append(name))
        )
        m.add_command(
            label="🌐  Tìm phần mềm thay thế trên Google",
            command=lambda: webbrowser.open(
                f"https://www.google.com/search?q=phan+mem+mien+phi+thay+cho+{urllib.parse.quote(name)}")
        )
        m.add_command(
            label="🌐  Tìm bản quyền / giá mua",
            command=lambda: webbrowser.open(
                f"https://www.google.com/search?q={urllib.parse.quote(name)}+license+price")
        )
        if v[4]:  # có alt
            m.add_command(
                label=f"⬇️  Tải thay thế: {v[4]}",
                command=lambda n=name: self._open_alt_url(n)
            )

        # ── Thêm mới / đóng góp ──────────────────────────────
        m.add_separator()
        if lic == LIC_UNKNOWN:
            # Phần mềm chưa có trong DB → nổi bật
            m.add_command(
                label=f"➕  Thêm '{name[:40]}' vào CSDL  ★",
                font=("Segoe UI", 9, "bold"),
                command=lambda n=name: self._open_add_dialog(n)
            )
        else:
            m.add_command(
                label="➕  Thêm phần mềm mới vào CSDL",
                command=lambda: self._open_add_dialog("")
            )

        m.post(event.x_root, event.y_root)

    def _open_alt_url(self, name):
        nl = name.lower()
        for key, info in self.software_db.items():
            if key in nl and info.get("alt_url"):
                webbrowser.open(info["alt_url"]); return

    def _open_add_dialog(self, prefill_name: str = ""):
        """Mở dialog thêm mới phần mềm."""
        dlg = AddSoftwareDialog(
            parent=self,
            raw_software=self.raw_software,
            software_db=self.software_db,
            prefill_name=prefill_name,
        )
        self.wait_window(dlg)
        if dlg.result:
            # Cập nhật software_db ngay lập tức để scan mới nhận ra
            for kw in dlg.result.get("keywords", []):
                self.software_db[kw.lower().strip()] = {
                    "license":  dlg.result["license"],
                    "category": dlg.result["category"],
                    "alt":      dlg.result["alt"],
                    "alt_url":  dlg.result["alt_url"],
                    "note":     dlg.result["note"],
                }
            # Cập nhật free_catalog nếu cần
            fc = dlg.result.get("free_catalog")
            if fc and isinstance(fc, dict):
                self.free_catalog.append({
                    "name": dlg.result["name"],
                    "cat":  dlg.result["category"],
                    "lic":  dlg.result["license"],
                    "desc": fc.get("desc",""),
                    "url":  fc.get("url",""),
                })
            self._refresh_tab1()
            self.sv_status1.set(
                f"STATUS : Đã thêm '{dlg.result['name']}' – nhấn 'Quét lại' để phân tích lại toàn bộ."
            )

    # ── Phân tích bản quyền ──────────────────────────────
    def analyze(self, name: str) -> dict:
        nl = name.lower()
        for key, info in self.software_db.items():
            if key in nl:
                return info
        return {"license": LIC_UNKNOWN, "category": "Khác",
                "alt": "", "alt_url": "", "note": "Chưa có trong cơ sở dữ liệu"}

    # ── Scan ────────────────────────────────────────────
    def do_scan(self):
        if self._scanning: return
        self._scanning = True
        self.sv_status1.set("STATUS : ĐANG QUÉT PHẦN MỀM – VUI LÒNG CHỜ...")
        self.tv1.delete(*self.tv1.get_children())
        threading.Thread(target=self._scan_worker, daemon=True).start()

    def _scan_worker(self):
        apps     = scan_registry()
        analyzed = [{**a, **self.analyze(a["name"])} for a in apps]
        self.all_apps = analyzed
        self.after(0, self._scan_done)

    def _scan_done(self):
        self._scanning = False
        self._refresh_tab1()
        total = len(self.all_apps)
        free  = sum(1 for a in self.all_apps if a.get("license") in (LIC_OSS, LIC_FREE))
        comm  = sum(1 for a in self.all_apps if a.get("license") == LIC_COMMERCIAL)
        unk   = sum(1 for a in self.all_apps if a.get("license") == LIC_UNKNOWN)
        self.lbl_ttl.config(text=f"Tổng: {total}")
        self.lbl_free.config(text=f"Miễn phí/OSS: {free}")
        self.lbl_comm.config(text=f"Thương mại: {comm}")
        self.lbl_unk.config(text=f"Chưa rõ: {unk}")
        src = "GitHub" if self._db_source == "online" else "Local"
        self.sv_status1.set(
            f"STATUS : HOÀN TẤT [{src}] – {total} phần mềm  |  "
            f"Cần bản quyền: {comm}  |  Miễn phí/OSS: {free}  |  Chưa rõ: {unk}"
        )

    def select_commercial(self):
        self.tv1.selection_set()
        for i in self.tv1.get_children():
            v = self.tv1.item(i)["values"]
            if len(v) > 2 and v[2] in (LIC_COMMERCIAL, LIC_TRIAL):
                self.tv1.selection_add(i)
        n = len(self.tv1.selection())
        self.sv_status1.set(f"STATUS : Đã chọn {n} phần mềm Thương mại / Dùng thử.")

    def open_alt(self):
        sel = self.tv1.selection()
        if not sel:
            messagebox.showinfo("Chú ý", "Hãy chọn ít nhất 1 phần mềm.", parent=self); return
        opened = 0
        for item in sel[:5]:
            v  = self.tv1.item(item)["values"]
            nl = str(v[0]).lower()
            for key, info in self.software_db.items():
                if key in nl and info.get("alt_url"):
                    webbrowser.open(info["alt_url"]); opened += 1; break
        if not opened:
            messagebox.showinfo("Thông báo",
                "Không tìm thấy đề xuất thay thế cho mục đã chọn.", parent=self)

    # ── Xuất báo cáo ────────────────────────────────────
    def export_report(self):
        if not self.all_apps:
            messagebox.showinfo("Chú ý", "Chưa có dữ liệu.", parent=self); return
        path = filedialog.asksaveasfilename(
            parent=self, title="Lưu báo cáo",
            initialfile=f"bao_cao_ban_quyen_{datetime.now():%Y%m%d_%H%M}.txt",
            defaultextension=".txt",
            filetypes=[("Text file","*.txt"),("All files","*.*")])
        if not path: return
        groups = {
            "=== PHẦN MỀM THƯƠNG MẠI (CẦN MUA BẢN QUYỀN) ===": LIC_COMMERCIAL,
            "=== PHẦN MỀM DÙNG THỬ (CHƯA CÓ BẢN QUYỀN) ===":   LIC_TRIAL,
            "=== PHẦN MỀM FREEMIUM (CÓ THỂ CẦN NÂNG CẤP) ===":  LIC_FREEMIUM,
            "=== PHẦN MỀM MIỄN PHÍ ===":                          LIC_FREE,
            "=== PHẦN MỀM MÃ NGUỒN MỞ ===":                      LIC_OSS,
            "=== CHƯA XÁC ĐỊNH ===":                              LIC_UNKNOWN,
        }
        src   = "GitHub (online)" if self._db_source == "online" else "Local (offline)"
        lines = [
            "╔══════════════════════════════════════════════════════════╗",
            "║         BÁO CÁO PHÂN TÍCH BẢN QUYỀN PHẦN MỀM           ║",
            "╚══════════════════════════════════════════════════════════╝",
            f"Ngày tạo    : {datetime.now():%d/%m/%Y %H:%M:%S}",
            f"Tổng cộng   : {len(self.all_apps)} phần mềm",
            f"Nguồn DB    : {src}", ""
        ]
        for header, lic in groups.items():
            lst = [a for a in self.all_apps if a.get("license") == lic]
            if not lst: continue
            lines += ["", header]
            for a in lst:
                alt  = f" → Thay thế: {a['alt']}" if a.get("alt") else ""
                note = f" ({a['note']})"          if a.get("note") else ""
                lines.append(f"  • {a['name']}  v{a.get('version','?')}{alt}{note}")
        with open(path, "w", encoding="utf-8") as fp:
            fp.write("\n".join(lines))
        messagebox.showinfo("Hoàn tất", f"Đã lưu báo cáo:\n{path}", parent=self)
        os.startfile(path)

    # ── Nạp DB từ Excel ─────────────────────────────────
    def load_excel(self):
        if not HAS_EXCEL:
            messagebox.showerror("Cần thư viện",
                "Vui lòng cài: pip install openpyxl\n\n"
                "Định dạng (hàng đầu là tiêu đề):\n"
                "  A: TenPhanMem  B: LoaiBanQuyen  C: DanhMuc\n"
                "  D: PhanMemThayChe  E: URL  F: GhiChu", parent=self); return
        path = filedialog.askopenfilename(parent=self, title="Chọn file Excel",
            filetypes=[("Excel","*.xlsx *.xls"),("All","*.*")])
        if not path: return
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws = wb.active; added = 0
            def _col(row, i, default=""):
                return str(row[i]).strip() if len(row) > i and row[i] else default
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row[0]: continue
                name = str(row[0]).strip(); key = name.lower()
                entry = {
                    "id": key.replace(" ","-"),
                    "name": name, "keywords": [key],
                    "license": _col(row,1,LIC_UNKNOWN), "category": _col(row,2,"Khác"),
                    "alt": _col(row,3), "alt_url": _col(row,4),
                    "note": _col(row,5), "free_catalog": None,
                }
                found = False
                for item in self.raw_software:
                    if item.get("name","").lower() == name.lower():
                        item.update({k: entry[k] for k in entry if k != "id"})
                        found = True; break
                if not found:
                    self.raw_software.append(entry)
                self.software_db[key] = {k: entry[k] for k in
                    ("license","category","alt","alt_url","note")}
                added += 1
            wb.close()
            save_local_json(self.raw_software, LOCAL_JSON)
            messagebox.showinfo("Thành công",
                f"Đã nạp {added} bản ghi, lưu vào {LOCAL_JSON}.\n"
                "Nhấn 'Quét lại' để áp dụng.", parent=self)
        except Exception as e:
            messagebox.showerror("Lỗi", str(e), parent=self)

    # ══════════════════════════════════════════════════════
    #  TAB 2 – KHO PHẦN MỀM MIỄN PHÍ
    # ══════════════════════════════════════════════════════
    def _build_tab2(self):
        f = self.tab2
        bar = ttk.Frame(f); bar.pack(fill=tk.X, padx=self.PAD, pady=(self.PAD, 2))
        #ttk.Button(bar, text="⬇  Tải về (mở web)", command=self.download_sel, style="Accent.TButton").pack(side=tk.LEFT, padx=2)
        ttk.Button(bar, text="🔍 Kiểm tra PM đã cài bằng WinGet", command=self.check_installed_winget,
                   style="Accent.TButton").pack(side=tk.LEFT, padx=2)
        ttk.Button(bar, text="⚡  Cài đặt bằng WinGet", command=self.install_selected_winget,
                   style="Accent.TButton").pack(side=tk.LEFT, padx=2)
        ttk.Button(bar, text="☑  Chọn trang hiện tại", command=self.check_all2).pack(side=tk.LEFT, padx=2)
        ttk.Button(bar, text="☐  Bỏ chọn tất cả",      command=self.uncheck_all2).pack(side=tk.LEFT, padx=2)
        self.lbl_sel2 = ttk.Label(bar, text="Đã chọn: 0 phần mềm",
                                   font=("Segoe UI", 9, "bold"))
        self.lbl_sel2.pack(side=tk.RIGHT, padx=10)

        nc = ttk.Frame(f)
        nc.pack(fill=tk.X, padx=self.PAD, pady=0)
        bar = ttk.Frame(f)
        bar.pack(fill=tk.X, padx=self.PAD, pady=(self.PAD, 2))

        ttk.Label(nc, text="★ Nhấn ô [☐] để chọn → 'Tải về' để mở trang tải.  "
                           "Nhấp đúp để mở web.  Tất cả đều miễn phí / mã nguồn mở.",
                  style="Small.TLabel", foreground="#5D6D7E").pack(side=tk.LEFT)

        # Combobox chọn danh mục
        self.cats2 = [
            "Tất cả", "Ứng dụng văn phòng", "Phần mềm Hệ thống",
            "Thiết kế đồ họa", "Chỉnh sửa ảnh", "Phần mềm Video",
            "Nghe nhạc", "Trình duyệt", "Chat & Gọi video",
            "Bảo mật", "Diệt Virus - Spyware", "Phần mềm mạng",
            "Hỗ trợ Download", "Phần mềm lập trình", "Quản lý Doanh nghiệp",
            "Dữ liệu - File", "Quản trị CSDL", "Giáo dục - Học tập",
            "Phần mềm cá nhân", "Mạng xã hội", "Quản lý Email",
            "Drivers - Firmware"
        ]
        self.cur_cat2 = "Tất cả"
        self.sv_cat2 = tk.StringVar(value="Tất cả")
        cb_cat2 = ttk.Combobox(nc, textvariable=self.sv_cat2, values=self.cats2,
                               state="readonly", width=24)
        cb_cat2.pack(side=tk.LEFT, padx=(4, 12))
        cb_cat2.bind("<<ComboboxSelected>>", lambda e: self._set_cat2(self.sv_cat2.get()))

        tf = ttk.Frame(f); tf.pack(fill=tk.BOTH, expand=True, padx=self.PAD, pady=2)
        cols = ("chon","name","cat","lic","desc","url","winget_id")
        self.tv2 = ttk.Treeview(tf, columns=cols, show="headings", selectmode="extended")
        hdrs2 = [("chon","Chọn",52),("name","Tên phần mềm",200),
                 ("cat","Danh mục",120),("lic","Loại",110),
                 ("desc","Mô tả",320),("url","Trang tải",260),("winget_id","winget ID",120)]
        for cid, txt, w in hdrs2:
            self.tv2.heading(cid, text=txt)
            self.tv2.column(cid, width=w, minwidth=40,
                             anchor="center" if cid == "chon" else "w")
        for lic, bg in ROW_COLORS.items():
            self.tv2.tag_configure(lic, background=bg)
        self.tv2.tag_configure("checked", background="#D5F5E3")
        vsb = ttk.Scrollbar(tf, orient=tk.VERTICAL, command=self.tv2.yview)
        self.tv2.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y); self.tv2.pack(fill=tk.BOTH, expand=True)
        self.tv2.bind("<Button-1>",  self._tv2_click)
        self.tv2.bind("<Double-1>",  self._tv2_dblclick)

        bot = tk.Frame(f, bg="#D4D0C8", bd=1, relief="sunken"); bot.pack(fill=tk.X)
        self.sv_status2 = tk.StringVar(value="STATUS : ĐANG ĐỢI USER.")
        tk.Label(bot, textvariable=self.sv_status2, bg="#D4D0C8",
                 font=("Courier New", 8), anchor="w").pack(side=tk.LEFT, padx=6, pady=2)

        self._refresh_tab2()  # gọi lần đầu

    def _set_cat2(self, cat):
        self.cur_cat2 = cat
        self._refresh_tab2()

    def _refresh_tab2(self):
        """Làm mới tab2 dựa trên danh mục đã chọn."""
        cat = self.cur_cat2
        if cat == "Tất cả":
            rows = self.free_catalog
        else:
            rows = [s for s in self.free_catalog if s["cat"] == cat]
        self._fill_tv2(rows)

    def _fill_tv2(self, rows):
        """Đổ dữ liệu vào Treeview của tab2."""
        self.tv2.delete(*self.tv2.get_children())
        seen = set()
        for sw in rows:
            if sw["name"] in seen:
                continue
            seen.add(sw["name"])
            checked = sw["name"] in self.free_checked
            tag = "checked" if checked else sw["lic"]
            icon = "☑" if checked else "☐"
            # Lấy winget_id ra, nếu không có hoặc rỗng thì hiển thị dấu gạch ngang "—"
            self.tv2.insert("", tk.END, iid=f"sw::{sw['name']}",
                            values=(icon, sw["name"], sw["cat"], sw["lic"],
                                    sw["desc"], sw["url"],sw["winget_id"]),
                            tags=(tag,))

    def _tv2_click(self, event):
        if self.tv2.identify_column(event.x) == "#1":
            row = self.tv2.identify_row(event.y)
            if row: self._toggle2(row)

    def _toggle2(self, iid):
        v = self.tv2.item(iid)["values"]
        if str(v[0]) == "✓ Đã cài":
            return  # Bỏ qua hoàn toàn, không xử lý đổi dấu tích chọn
        name = str(v[1])
        if name in self.free_checked:
            self.free_checked.discard(name)
            sw  = next((s for s in self.free_catalog if s["name"] == name), None)
            tag = sw["lic"] if sw else LIC_FREE
            self.tv2.item(iid, values=("☐", name, v[2], v[3], v[4], v[5]), tags=(tag,))
        else:
            self.free_checked.add(name)
            self.tv2.item(iid, values=("☑", name, v[2], v[3], v[4], v[5]), tags=("checked",))
        n = len(self.free_checked)
        self.lbl_sel2.config(text=f"Đã chọn: {n} phần mềm")
        self.sv_status2.set(f"STATUS : Đã chọn {n} phần mềm – nhấn 'Tải về'." if n
                            else "STATUS : Chọn phần mềm muốn tải.")

    def _tv2_dblclick(self, _):
        sel = self.tv2.selection()
        if not sel: return
        v = self.tv2.item(sel[0])["values"]
        if len(v) > 5 and v[5]: webbrowser.open(str(v[5]))

    def check_all2(self):
        for iid in self.tv2.get_children():
            v = self.tv2.item(iid)["values"]
            if str(v[0]) == "✓ Đã cài":
                continue
            name = str(v[1])
            self.free_checked.add(name)
            self.tv2.item(iid, values=("☑",name,v[2],v[3],v[4],v[5]), tags=("checked",))
        n = len(self.free_checked)
        self.lbl_sel2.config(text=f"Đã chọn: {n} phần mềm")
        self.sv_status2.set(f"STATUS : Đã chọn {n} phần mềm.")

    def uncheck_all2(self):
        self.free_checked.clear()
        for iid in self.tv2.get_children():
            v = self.tv2.item(iid)["values"]
            name = str(v[1])
            sw  = next((s for s in self.free_catalog if s["name"] == name), None)
            tag = sw["lic"] if sw else LIC_FREE
            self.tv2.item(iid, values=("☐",name,v[2],v[3],v[4],v[5]), tags=(tag,))
        self.lbl_sel2.config(text="Đã chọn: 0 phần mềm")
        self.sv_status2.set("STATUS : Đã bỏ chọn tất cả.")

    def download_sel(self):
        if not self.free_checked:
            messagebox.showinfo("Chú ý","Chưa chọn phần mềm nào.",parent=self); return
        urls = [(n, s["url"]) for n in self.free_checked
                for s in self.free_catalog if s["name"] == n and s.get("url")]
        if not urls:
            messagebox.showinfo("Chú ý","Không có link tải.",parent=self); return
        preview = "\n".join(f"• {n}" for n,_ in urls[:12])
        if len(urls) > 12: preview += f"\n... và {len(urls)-12} phần mềm khác"
        if not messagebox.askyesno("Xác nhận tải",
            f"Sẽ mở {len(urls)} trang web:\n\n{preview}\n\nTiếp tục?", parent=self): return
        for _, url in urls: webbrowser.open(url)
        self.sv_status2.set(f"STATUS : Đã mở {len(urls)} trang tải phần mềm.")

    def install_selected_winget(self):
        """Cài đặt các phần mềm đã chọn trong tab2 bằng WinGet."""
        selected = [self.tv2.item(iid)["values"] for iid in self.tv2.selection()]
        if not selected:
            # Nếu không có dòng nào được chọn, lấy tất cả các dòng đã check
            selected = [self.tv2.item(iid)["values"] for iid in self.tv2.get_children()
                        if self.tv2.item(iid)["values"][0] == "☑"]
        if not selected:
            messagebox.showinfo("Chú ý", "Hãy chọn ít nhất một phần mềm (bằng cách nhấn vào ô ☐).", parent=self)
            return

        winget_ids = []
        for values in selected:
            # values[1] là tên, tìm trong free_catalog để lấy winget_id
            name = values[1]
            for fc in self.free_catalog:
                if fc["name"] == name and fc.get("winget_id"):
                    winget_ids.append(fc["winget_id"])
                    break
        if not winget_ids:
            messagebox.showinfo("Không hỗ trợ",
                                "Các phần mềm đã chọn không có WinGet ID.\nHãy dùng nút 'Tải về' để tải thủ công.",
                                parent=self)
            return
        self.install_with_winget(winget_ids)

    def is_winget_available(self):
        """Kiểm tra xem winget có trong PATH không."""
        try:
            subprocess.run(["winget", "--version"], capture_output=True, check=True)
            return True
        except (subprocess.SubprocessError, FileNotFoundError):
            return False

    def install_with_winget(self, winget_ids):
        """Cài đặt danh sách các winget_id (lần lượt)."""
        if not winget_ids:
            return
        if not self.is_winget_available():
            messagebox.showerror(
                "WinGet không có sẵn",
                "WinGet (Trình quản lý gói Windows) chưa được cài đặt.\n\n"
                "Vui lòng cài đặt 'App Installer' từ Microsoft Store hoặc tải về từ:\n"
                "https://aka.ms/getwinget",
                parent=self
            )
            return
        # Hỏi xác nhận
        if not messagebox.askyesno(
                "Xác nhận cài đặt",
                f"Sẽ cài đặt {len(winget_ids)} phần mềm bằng WinGet:\n\n" +
                "\n".join(f"• {id_}" for id_ in winget_ids) +
                "\n\nTiến hành?",
                parent=self
        ):
            return
        # Chạy lần lượt (có thể chạy song song nhưng đơn giản là tuần tự)
        for wid in winget_ids:
            try:
                # Chạy lặng lẽ, không hiện cửa sổ console
                subprocess.run(
                    ["winget", "install", "--id", wid, "--silent",
                     "--accept-package-agreements", "--accept-source-agreements"],
                    capture_output=True,
                    check=True
                )
            except subprocess.CalledProcessError as e:
                messagebox.showerror(
                    "Lỗi cài đặt",
                    f"Không thể cài đặt {wid}:\n{e.stderr.decode() if e.stderr else str(e)}",
                    parent=self
                )
            else:
                messagebox.showinfo("Thành công", f"Đã cài đặt {wid}", parent=self)

    def check_installed_winget(self):
        """Khởi chạy luồng kiểm tra ứng dụng đã cài đặt bằng winget"""
        self.sv_status2.set("STATUS: Đang quét danh sách phần mềm đã cài đặt trên hệ thống...")
        threading.Thread(target=self._check_installed_worker, daemon=True).start()

    def _check_installed_worker(self):
        """Hàm chạy ngầm thu thập winget list và đối chiếu dữ liệu"""
        import winget_id_finder
        winget_path = winget_id_finder.get_winget_path() if hasattr(winget_id_finder, 'get_winget_path') else None
        if not winget_path:
            winget_path = "winget"

        installed_ids = set()
        try:
            # Chạy lệnh winget list để lấy tất cả phần mềm quản lý bởi winget
            cmd = [winget_path, "list", "--accept-source-agreements", "--disable-interactivity"]
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            lines = result.stdout.splitlines()
            id_start = -1
            version_start = -1

            for line in lines:
                if not line.strip():
                    continue
                if "Name" in line and "Id" in line:
                    id_start = line.find("Id")
                    version_start = line.find("Version")
                    continue
                if line.strip().startswith("-"):
                    continue

                if id_start != -1 and version_start != -1:
                    p_id = line[id_start:version_start].strip()
                    if p_id:
                        # Thêm vào set (chuyển về chữ thường để so sánh)
                        installed_ids.add(p_id.lower())

        except Exception as e:
            # FIX: Thay self.root.after bằng self.after
            self.after(0, lambda: messagebox.showerror("Lỗi", f"Không thể chạy winget list: {str(e)}", parent=self))
            self.after(0, lambda: self.sv_status2.set("STATUS: Lỗi quét ứng dụng."))
            return

        # FIX: Thay self.root.after bằng self.after
        self.after(0, lambda: self._update_treeview_with_installed(installed_ids))

    def _update_treeview_with_installed(self, installed_ids):
        """Cập nhật trạng thái hiển thị trên Treeview tv2 dựa vào tập hợp ID đã cài"""
        count_disabled = 0

        # Cấu hình tag chữ đậm màu xám cho các dòng đã cài đặt
        self.tv2.tag_configure("installed_disabled", foreground="#888888", font=("Segoe UI", 9, "bold"))

        for iid in self.tv2.get_children():
            v = self.tv2.item(iid)["values"]
            name = str(v[1])  # Tên phần mềm nằm ở index 1

            # Tìm phần mềm tương ứng trong self.free_catalog
            sw = next((s for s in self.free_catalog if s.get("name") == name), None)
            if sw and "winget_id" in sw and sw["winget_id"]:
                w_id = str(sw["winget_id"]).strip().lower()

                # Nếu winget_id nằm trong danh sách máy đã cài
                if w_id in installed_ids:
                    # Giữ lại hiển thị WinGet ID ở cột tương ứng (index 3) thay vì ghi đè bằng Note cũ
                    # v[0]=Trạng thái, v[1]=Tên, v[2]=Loại, v[3]=Ghi chú cũ, v[4]=Mô tả, v[5]=URL
                    # Chúng ta đổi trạng thái v[0] thành "✓ Đã cài" và hiển thị WinGet ID vào vị trí mong muốn
                    self.tv2.item(
                        iid,
                        values=("✓ Đã cài", v[1], v[2], v[3], v[4], v[5],f"ID: {sw['winget_id']}"),
                        tags=("installed_disabled",)
                    )
                    if name in self.free_checked:
                        self.free_checked.remove(name)
                    count_disabled += 1

        self.lbl_sel2.config(text=f"Đã chọn: {len(self.free_checked)} phần mềm")
        self.sv_status2.set(f"STATUS: Đã phát hiện và khóa {count_disabled} phần mềm đã cài trên máy.")
        messagebox.showinfo("Hoàn tất", f"Quét xong! Đã phát hiện và khóa {count_disabled} phần mềm đã có trên máy.",
                            parent=self)
# ══════════════════════════════════════════════════════════
#  KHỞI ĐỘNG
# ══════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = App()
    app.mainloop()