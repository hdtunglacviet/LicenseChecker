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
GITHUB_USER       = "hdtunglacviet"          # ← đổi thành username GitHub
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
#  DIALOG THÊM MỚI PHẦN MỀM
# ══════════════════════════════════════════════════════════

class AddSoftwareDialog(tk.Toplevel):
    """
    Giao diện thêm mới phần mềm vào cơ sở dữ liệu cục bộ
    và đóng góp lên GitHub qua Issue.
    """

    def __init__(self, parent, raw_software: list, software_db: dict,
                 prefill_name: str = ""):
        super().__init__(parent)
        self.parent        = parent
        self.raw_software  = raw_software   # tham chiếu trực tiếp
        self.software_db   = software_db    # tham chiếu trực tiếp
        self.result        = None           # entry mới nếu OK

        self.title("➕  Thêm phần mềm vào cơ sở dữ liệu")
        self.geometry("580x560")
        self.resizable(False, False)
        self.configure(bg="#ECECEC")
        self.grab_set()
        self.transient(parent)

        self._build(prefill_name)
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)

        # Căn giữa màn hình
        self.update_idletasks()
        pw, ph = parent.winfo_rootx(), parent.winfo_rooty()
        w, h = self.winfo_width(), self.winfo_height()
        self.geometry(f"+{pw + (parent.winfo_width()-w)//2}+{ph + (parent.winfo_height()-h)//2}")

    def _lbl(self, parent, text):
        return tk.Label(parent, text=text, bg="#ECECEC",
                        font=("Segoe UI", 9), anchor="w")

    def _build(self, prefill_name):
        padx = 14   # chỉ padx chung; pady đặt riêng từng dòng

        # ── Tiêu đề ──
        hdr = tk.Frame(self, bg="#3C7FC0", height=36)
        hdr.pack(fill=tk.X); hdr.pack_propagate(False)
        tk.Label(hdr, text="  ➕  Đóng góp phần mềm mới vào cơ sở dữ liệu",
                 bg="#3C7FC0", fg="white",
                 font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=8, pady=4)

        body = tk.Frame(self, bg="#ECECEC"); body.pack(fill=tk.BOTH, expand=True)

        # ── Tên phần mềm ──
        self._lbl(body, "Tên phần mềm *").pack(fill=tk.X, padx=padx, pady=(10, 0))
        self.sv_name = tk.StringVar(value=prefill_name)
        ttk.Entry(body, textvariable=self.sv_name, width=60,
                  font=("Segoe UI", 9)).pack(fill=tk.X, padx=padx, pady=(0, 2))
        tk.Label(body, text="  Ví dụ: Microsoft Word", bg="#ECECEC",
                 fg="#888", font=("Segoe UI", 8)).pack(fill=tk.X, padx=padx)

        # ── Keywords ──
        self._lbl(body, "Từ khoá nhận dạng (keywords) *  –  cách nhau bởi dấu phẩy").pack(
            fill=tk.X, padx=padx, pady=(8, 0))
        # Tự điền keyword từ tên
        auto_kw = prefill_name.lower().strip() if prefill_name else ""
        self.sv_kw = tk.StringVar(value=auto_kw)
        ttk.Entry(body, textvariable=self.sv_kw, width=60,
                  font=("Segoe UI", 9)).pack(fill=tk.X, padx=padx, pady=(0, 2))
        tk.Label(body, text="  Ví dụ: microsoft word, ms word, winword",
                 bg="#ECECEC", fg="#888", font=("Segoe UI", 8)).pack(fill=tk.X, padx=padx)

        # ── Hàng license + category ──
        row2 = tk.Frame(body, bg="#ECECEC"); row2.pack(fill=tk.X, padx=padx, pady=(8, 2))

        lc = tk.Frame(row2, bg="#ECECEC"); lc.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._lbl(lc, "Loại bản quyền *").pack(anchor="w")
        self.sv_lic = tk.StringVar(value=LIC_UNKNOWN)
        cb_lic = ttk.Combobox(lc, textvariable=self.sv_lic,
                               values=ALL_LICENSES, state="readonly", width=22)
        cb_lic.pack(anchor="w", pady=(2, 0))

        tk.Frame(row2, bg="#ECECEC", width=16).pack(side=tk.LEFT)

        cc = tk.Frame(row2, bg="#ECECEC"); cc.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._lbl(cc, "Danh mục *").pack(anchor="w")
        self.sv_cat = tk.StringVar(value="Khác")
        cb_cat = ttk.Combobox(cc, textvariable=self.sv_cat,
                               values=ALL_CATEGORIES, state="readonly", width=22)
        cb_cat.pack(anchor="w", pady=(2, 0))

        # ── Phần mềm thay thế ──
        self._lbl(body, "Phần mềm thay thế miễn phí / OSS  (nếu có)").pack(
            fill=tk.X, padx=padx, pady=(8, 0))
        self.sv_alt = tk.StringVar()
        ttk.Entry(body, textvariable=self.sv_alt, width=60).pack(fill=tk.X, padx=padx, pady=(0, 2))
        tk.Label(body, text="  Ví dụ: LibreOffice Writer",
                 bg="#ECECEC", fg="#888", font=("Segoe UI", 8)).pack(fill=tk.X, padx=padx)

        # ── URL tải thay thế ──
        self._lbl(body, "Link tải phần mềm thay thế  (nếu có)").pack(
            fill=tk.X, padx=padx, pady=(8, 0))
        self.sv_url = tk.StringVar()
        ttk.Entry(body, textvariable=self.sv_url, width=60).pack(fill=tk.X, padx=padx, pady=(0, 2))
        tk.Label(body, text="  Ví dụ: https://www.libreoffice.org/download/",
                 bg="#ECECEC", fg="#888", font=("Segoe UI", 8)).pack(fill=tk.X, padx=padx)

        # ── Ghi chú ──
        self._lbl(body, "Ghi chú ngắn").pack(fill=tk.X, padx=padx, pady=(8, 0))
        self.sv_note = tk.StringVar()
        ttk.Entry(body, textvariable=self.sv_note, width=60).pack(fill=tk.X, padx=padx, pady=(0, 4))

        # ── Thông báo lỗi ──
        self.lbl_err = tk.Label(body, text="", bg="#FFFBE6", fg="#B03A2E",
                                 font=("Segoe UI", 8, "bold"), anchor="w",
                                 relief="flat", padx=8)
        self.lbl_err.pack(fill=tk.X, padx=padx, pady=(2, 0))

        # ── Nút hành động ──
        btn_frame = tk.Frame(self, bg="#D4D0C8", relief="sunken", bd=1)
        btn_frame.pack(fill=tk.X, side=tk.BOTTOM)

        ttk.Button(btn_frame, text="✅  Lưu vào máy",
                   command=self.on_save, style="Accent.TButton").pack(
            side=tk.LEFT, padx=10, pady=6)
        ttk.Button(btn_frame, text="🌐  Lưu & Đóng góp lên GitHub",
                   command=self.on_save_and_contribute).pack(
            side=tk.LEFT, padx=4, pady=6)
        ttk.Button(btn_frame, text="Huỷ", command=self.on_cancel).pack(
            side=tk.RIGHT, padx=10, pady=6)

        tk.Label(btn_frame,
                 text="★ 'Đóng góp' sẽ mở GitHub Issue – bạn cần đăng nhập GitHub để gửi.",
                 bg="#D4D0C8", fg="#555", font=("Segoe UI", 7)).pack(
            side=tk.RIGHT, padx=8)

    # ── Kiểm tra trùng lặp ──────────────────────────────
    def _check_duplicates(self, name: str, keywords: list) -> str | None:
        """
        Trả về thông báo lỗi nếu name hoặc keyword bị trùng,
        ngược lại trả về None.
        """
        name_lower = name.lower().strip()
        kws_lower  = [k.lower().strip() for k in keywords if k.strip()]

        # 1. Kiểm tra name trùng với name trong raw_software
        for item in self.raw_software:
            if item.get("name", "").lower().strip() == name_lower:
                return (f"Tên phần mềm '{item['name']}' đã có trong cơ sở dữ liệu!\n"
                        f"(id: {item.get('id', '?')})")

        # 2. Kiểm tra từng keyword
        existing_all_kws = set()
        for item in self.raw_software:
            for kw in item.get("keywords", []):
                existing_all_kws.add(kw.lower().strip())

        for kw in kws_lower:
            if kw in existing_all_kws:
                return (f"Từ khoá '{kw}' đã tồn tại trong cơ sở dữ liệu!\n"
                        "Hãy dùng từ khoá khác hoặc cập nhật bản ghi hiện có.")
        return None

    # ── Đọc & validate form ──────────────────────────────
    def _read_form(self):
        """Đọc dữ liệu form, validate, trả về dict hoặc None nếu lỗi."""
        name = self.sv_name.get().strip()
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

        self.lbl_err.config(text="")
        return {
            "id":       name.lower().replace(" ", "-"),
            "name":     name,
            "keywords": keywords,
            "license":  self.sv_lic.get(),
            "category": self.sv_cat.get(),
            "alt":      self.sv_alt.get().strip(),
            "alt_url":  self.sv_url.get().strip(),
            "note":     self.sv_note.get().strip(),
            "free_catalog": None,
        }

    # ── Lưu vào máy ──────────────────────────────────────
    def on_save(self):
        entry = self._read_form()
        if entry is None:
            return
        self.raw_software.append(entry)
        try:
            save_local_json(self.raw_software, LOCAL_JSON)
        except Exception as e:
            messagebox.showerror("Lỗi ghi file", str(e), parent=self)
            return
        self.result = entry
        messagebox.showinfo(
            "Đã lưu",
            f"✅  Đã thêm '{entry['name']}' vào software_data.json cục bộ.\n\n"
            "Nhấn 'Quét lại' ở cửa sổ chính để áp dụng ngay.",
            parent=self
        )
        self.destroy()

    # ── Lưu & mở GitHub Issue ────────────────────────────
    def on_save_and_contribute(self):
        entry = self._read_form()
        if entry is None:
            return

        # Lưu local trước
        self.raw_software.append(entry)
        try:
            save_local_json(self.raw_software, LOCAL_JSON)
        except Exception as e:
            messagebox.showerror("Lỗi ghi file", str(e), parent=self)
            return
        self.result = entry

        # Chuẩn bị nội dung Issue
        entry_json = json.dumps(entry, ensure_ascii=False, indent=2)
        title  = f"[DB Contribution] Thêm mới: {entry['name']}"
        body   = (
            f"## Đề xuất thêm phần mềm mới\n\n"
            f"**Tên:** {entry['name']}  \n"
            f"**Loại bản quyền:** {entry['license']}  \n"
            f"**Danh mục:** {entry['category']}  \n"
            f"**Keywords:** {', '.join(entry['keywords'])}  \n"
            f"**Thay thế bởi:** {entry['alt'] or '_(không có)_'}  \n"
            f"**Link tải:** {entry['alt_url'] or '_(không có)_'}  \n"
            f"**Ghi chú:** {entry['note'] or '_(không có)_'}  \n\n"
            f"### JSON entry\n\n```json\n{entry_json}\n```\n\n"
            f"---\n*Gửi từ LicenseChecker v{VERSION}*"
        )

        # Mã hoá query string để mở trình duyệt
        params  = urllib.parse.urlencode({"title": title, "body": body})
        url     = f"{GITHUB_ISSUES_URL}?{params}"
        webbrowser.open(url)

        messagebox.showinfo(
            "Đã lưu & mở GitHub",
            f"✅  Đã lưu '{entry['name']}' vào máy.\n\n"
            "🌐  Trình duyệt vừa mở GitHub Issues – hãy đăng nhập\n"
            "và nhấn 'Submit new issue' để gửi đóng góp.",
            parent=self
        )
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

        self.cats1 = [
            "Tất cả",
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
        self.cur_cat1 = "Tất cả"
        cf = ttk.Frame(f); cf.pack(fill=tk.X, padx=self.PAD, pady=2)
        self.cat_btns1 = {}
        for c in self.cats1:
            b = ttk.Button(cf, text=c, style="Cat.TButton",
                           command=lambda x=c: self._set_cat1(x))
            b.pack(side=tk.LEFT, padx=1); self.cat_btns1[c] = b
        self._set_cat1("Tất cả", refresh=False)

        sm = ttk.Frame(f); sm.pack(fill=tk.X, padx=self.PAD, pady=1)
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

    def _set_cat1(self, cat, refresh=True):
        self.cur_cat1 = cat
        for c, b in self.cat_btns1.items():
            b.configure(style="CatOn.TButton" if c == cat else "Cat.TButton")
        if refresh: self._refresh_tab1()

    def _refresh_tab1(self):
        kw  = self.sv_search.get().lower()
        cat = self.cur_cat1
        rows = [a for a in self.all_apps
                if (cat == "Tất cả" or a.get("category") == cat)
                and (not kw or kw in a["name"].lower())]
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
        ttk.Button(bar, text="⬇  Tải về các phần đã chọn",
                   command=self.download_sel, style="Accent.TButton").pack(side=tk.LEFT, padx=2)
        ttk.Button(bar, text="☑  Chọn trang hiện tại", command=self.check_all2).pack(side=tk.LEFT, padx=2)
        ttk.Button(bar, text="☐  Bỏ chọn tất cả",      command=self.uncheck_all2).pack(side=tk.LEFT, padx=2)
        self.lbl_sel2 = ttk.Label(bar, text="Đã chọn: 0 phần mềm",
                                   font=("Segoe UI", 9, "bold"))
        self.lbl_sel2.pack(side=tk.RIGHT, padx=10)

        nc = ttk.Frame(f); nc.pack(fill=tk.X, padx=self.PAD, pady=0)
        ttk.Label(nc, text="★ Nhấn ô [☐] để chọn → 'Tải về' để mở trang tải.  "
                           "Nhấp đúp để mở web.  Tất cả đều miễn phí / mã nguồn mở.",
                  style="Small.TLabel", foreground="#5D6D7E").pack(side=tk.LEFT)

        cats2 = [
            "Tất cả",
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
        ]
        self.cur_cat2 = "Tất cả"
        cf = ttk.Frame(f); cf.pack(fill=tk.X, padx=self.PAD, pady=4)
        self.cat_btns2 = {}
        for c in cats2:
            b = ttk.Button(cf, text=c, style="Cat.TButton",
                           command=lambda x=c: self._set_cat2(x))
            b.pack(side=tk.LEFT, padx=1); self.cat_btns2[c] = b
        self._set_cat2("Tất cả", refresh=False)

        tf = ttk.Frame(f); tf.pack(fill=tk.BOTH, expand=True, padx=self.PAD, pady=2)
        cols = ("chon","name","cat","lic","desc","url")
        self.tv2 = ttk.Treeview(tf, columns=cols, show="headings", selectmode="extended")
        hdrs2 = [("chon","Chọn",52),("name","Tên phần mềm",200),
                 ("cat","Danh mục",120),("lic","Loại",110),
                 ("desc","Mô tả",320),("url","Trang tải",260)]
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

        self._fill_tv2("Tất cả")

    def _set_cat2(self, cat, refresh=True):
        self.cur_cat2 = cat
        for c, b in self.cat_btns2.items():
            b.configure(style="CatOn.TButton" if c == cat else "Cat.TButton")
        if refresh: self._fill_tv2(cat)

    def _fill_tv2(self, cat):
        self.tv2.delete(*self.tv2.get_children())
        rows = self.free_catalog if cat == "Tất cả" else \
               [s for s in self.free_catalog if s["cat"] == cat]
        seen = set()
        for sw in rows:
            if sw["name"] in seen: continue
            seen.add(sw["name"])
            checked = sw["name"] in self.free_checked
            tag  = "checked" if checked else sw["lic"]
            icon = "☑" if checked else "☐"
            self.tv2.insert("", tk.END, iid=f"sw::{sw['name']}",
                             values=(icon, sw["name"], sw["cat"], sw["lic"], sw["desc"], sw["url"]),
                             tags=(tag,))

    def _tv2_click(self, event):
        if self.tv2.identify_column(event.x) == "#1":
            row = self.tv2.identify_row(event.y)
            if row: self._toggle2(row)

    def _toggle2(self, iid):
        v = self.tv2.item(iid)["values"]; name = str(v[1])
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
            v = self.tv2.item(iid)["values"]; name = str(v[1])
            self.free_checked.add(name)
            self.tv2.item(iid, values=("☑",name,v[2],v[3],v[4],v[5]), tags=("checked",))
        n = len(self.free_checked)
        self.lbl_sel2.config(text=f"Đã chọn: {n} phần mềm")
        self.sv_status2.set(f"STATUS : Đã chọn {n} phần mềm.")

    def uncheck_all2(self):
        self.free_checked.clear()
        for iid in self.tv2.get_children():
            v = self.tv2.item(iid)["values"]; name = str(v[1])
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


# ══════════════════════════════════════════════════════════
#  KHỞI ĐỘNG
# ══════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = App()
    app.mainloop()
