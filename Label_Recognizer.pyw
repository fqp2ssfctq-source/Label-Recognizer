"""
Optical Transceiver Label Inspector
====================================
카메라 또는 파일에서 optical transceiver 라벨 이미지를 획득하고
다음 항목을 자동으로 검사합니다:
  1. OCR 텍스트 추출 및 필드 검증 (Model, P/N, S/N, Rev, Date 등)
  2. 바코드 / QR 코드 인식
  3. 인쇄 품질 평가 (흐림, 기울기, 밝기)

의존 패키지:
  pip install opencv-python pillow pytesseract pyzbar numpy openpyxl

tesseract-ocr 설치 필요:
  Windows : https://github.com/UB-Mannheim/tesseract/wiki
  Linux   : sudo apt install tesseract-ocr tesseract-ocr-kor
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import datetime
import json
import os
import re
import csv

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

import cv2
import numpy as np
from PIL import Image, ImageTk
import pytesseract

_TESS_DEFAULT = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
if os.path.exists(_TESS_DEFAULT):
    pytesseract.pytesseract.tesseract_cmd = _TESS_DEFAULT

try:
    from pyzbar import pyzbar
    PYZBAR_OK = True
except ImportError:
    PYZBAR_OK = False

# ──────────────────────────────────────────────
CLR = {
    "bg":        "#1a1d23",
    "panel":     "#22262f",
    "border":    "#2e3340",
    "accent":    "#00c2ff",
    "ok":        "#00e676",
    "ng":        "#ff3d3d",
    "warn":      "#ffaa00",
    "text":      "#e8ecf0",
    "subtext":   "#8899aa",
    "btn":       "#2c3244",
    "btn_hover": "#3a4257",
    "btn_on":    "#1a5f7a",
}

FONT_TITLE  = ("Consolas", 13, "bold")
FONT_BODY   = ("Consolas", 10)
FONT_SMALL  = ("Consolas",  9)
FONT_RESULT = ("Consolas", 11, "bold")


# ──────────────────────────────────────────────
class LabelAnalyzer:
    FIELD_PATTERNS = {
        "Model":    r"(QSFP|SFP|OSFP|CFP|XFP|CXP|DSFP)[\w\-]+",
        "P/N":      r"P/?N[:\s]*([A-Z0-9\-]+)",
        "S/N":      r"S/?N[:\s]*([A-Z0-9]+)",
        "Rev":      r"[Rr][Ee][Vv][:\s]*([A-Za-z0-9\.]+)",
        "Date":     r"(20\d{2}[-/]\d{2}[-/]\d{2}|\d{2}[-/]\d{2}[-/]20\d{2}|\d{4}W\d{2})",
    }

    def __init__(self, thresholds=None):
        self.thr = thresholds or {
            "blur_min":   80.0,
            "tilt_max":    5.0,
            "bright_min":  40,
            "bright_max": 230,
        }

    def preprocess(self, bgr):
        gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        return clahe.apply(gray)

    def check_print_quality(self, bgr):
        gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
        blur_score = cv2.Laplacian(gray, cv2.CV_64F).var()
        blur_ok    = blur_score >= self.thr["blur_min"]
        brightness = int(gray.mean())
        bright_ok  = self.thr["bright_min"] <= brightness <= self.thr["bright_max"]
        edges  = cv2.Canny(gray, 50, 150)
        lines  = cv2.HoughLines(edges, 1, np.pi / 180, 80)
        tilt_deg = 0.0
        if lines is not None:
            angles = []
            for rho, theta in lines[:, 0]:
                a = np.degrees(theta) - 90
                if abs(a) < 45:
                    angles.append(a)
            if angles:
                tilt_deg = float(np.median(angles))
        tilt_ok = abs(tilt_deg) <= self.thr["tilt_max"]
        return {
            "blur_score": round(blur_score, 1),
            "blur_ok":    blur_ok,
            "brightness": brightness,
            "bright_ok":  bright_ok,
            "tilt_deg":   round(tilt_deg, 2),
            "tilt_ok":    tilt_ok,
        }

    def run_ocr(self, bgr):
        _empty = {"raw": "", "fields": {n: "" for n in self.FIELD_PATTERNS}}
        try:
            proc  = self.preprocess(bgr)
            cfg   = "--psm 6 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-/:. "
            raw   = pytesseract.image_to_string(proc, config=cfg)
            fields = {}
            for name, pat in self.FIELD_PATTERNS.items():
                m = re.search(pat, raw)
                fields[name] = m.group(0) if m else ""
            return {"raw": raw.strip(), "fields": fields}
        except pytesseract.TesseractNotFoundError:
            return {"raw": "[Tesseract 미설치 — OCR 비활성]", "fields": _empty["fields"]}
        except Exception:
            return _empty

    def decode_codes(self, bgr):
        if not PYZBAR_OK:
            return []
        results = []
        for obj in pyzbar.decode(bgr):
            results.append({
                "type": obj.type,
                "data": obj.data.decode("utf-8", errors="replace"),
                "rect": list(obj.rect),
            })
        return results

    def analyze(self, bgr):
        quality = self.check_print_quality(bgr)
        ocr     = self.run_ocr(bgr)
        codes   = self.decode_codes(bgr)
        quality_pass = quality["blur_ok"] and quality["bright_ok"] and quality["tilt_ok"]
        ocr_pass     = any(v for v in ocr["fields"].values())
        verdict = "OK" if (quality_pass and ocr_pass) else "NG"
        return {
            "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "verdict":   verdict,
            "quality":   quality,
            "ocr":       ocr,
            "codes":     codes,
        }


# ──────────────────────────────────────────────
class TemplateManager:
    EMPTY = {
        "template_name": "이름없음",
        "barcodes": {
            "PN": {"fixed_value": "", "pattern": "", "required": True},
            "SN": {"pattern": "",                    "required": False},
        },
        "ocr_texts": [],
    }

    def __init__(self):
        self.data = None
        self.path = ""

    @property
    def loaded(self):
        return self.data is not None

    @property
    def name(self):
        return self.data["template_name"] if self.loaded else "템플릿 없음"

    def load(self, path):
        with open(path, encoding="utf-8") as f:
            self.data = json.load(f)
        self.path = path

    def save(self, path):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)
        self.path = path

    def _classify(self, codes):
        pn_cfg = self.data["barcodes"].get("PN", {})
        sn_cfg = self.data["barcodes"].get("SN", {})
        found_pn = found_sn = ""
        for c in codes:
            val = c["data"]
            if pn_cfg.get("fixed_value") and val == pn_cfg["fixed_value"]:
                found_pn = val;  continue
            if pn_cfg.get("pattern") and re.fullmatch(pn_cfg["pattern"], val):
                found_pn = val;  continue
            if sn_cfg.get("pattern") and re.fullmatch(sn_cfg["pattern"], val):
                found_sn = val;  continue
            if not found_pn:
                found_pn = val
            elif not found_sn and val != found_pn:
                found_sn = val
        return found_pn, found_sn

    def compare(self, result):
        if not self.loaded:
            return {}
        comp    = {}
        pn_cfg  = self.data["barcodes"].get("PN", {})
        sn_cfg  = self.data["barcodes"].get("SN", {})
        found_pn, found_sn = self._classify(result.get("codes", []))

        if pn_cfg.get("fixed_value") or pn_cfg.get("required"):
            if not found_pn:
                comp["P/N"] = {"value": "", "ok": False, "reason": "바코드 미인식"}
            elif pn_cfg.get("fixed_value") and found_pn != pn_cfg["fixed_value"]:
                comp["P/N"] = {"value": found_pn, "ok": False,
                               "reason": f"불일치 (기대: {pn_cfg['fixed_value']})"}
            else:
                comp["P/N"] = {"value": found_pn, "ok": True, "reason": ""}

        if found_sn:
            pat    = sn_cfg.get("pattern", "")
            fmt_ok = bool(re.fullmatch(pat, found_sn)) if pat else True
            comp["S/N"] = {"value": found_sn, "ok": fmt_ok,
                           "reason": "" if fmt_ok else "형식 불일치"}
        elif sn_cfg.get("required"):
            comp["S/N"] = {"value": "", "ok": False, "reason": "바코드 미인식"}

        raw_lower = (result["ocr"]["raw"] or "").lower()
        for item in self.data.get("ocr_texts", []):
            label    = item.get("label", "")
            contains = item.get("contains", "")
            if not label and not contains:
                continue
            ok = contains.lower() in raw_lower if contains else True
            comp[label] = {"value": contains, "ok": ok,
                           "reason": "" if ok else f'"{contains}" 미발견'}
        return comp


# ──────────────────────────────────────────────
class App(tk.Tk):
    PREVIEW_W  = 560
    PREVIEW_H  = 360
    _ZOOM_MIN  = 1.0
    _ZOOM_MAX  = 4.0
    _ZOOM_STEP = 0.5
    _CSV_COLS  = [
        "Timestamp", "Verdict",
        "Blur", "Blur_OK", "Brightness", "Bright_OK", "Tilt_deg", "Tilt_OK",
        "Model", "P/N", "S/N", "Rev", "Date", "Barcodes",
        "PN_Value", "PN_OK", "SN_Value", "SN_OK", "Template",
    ]

    def __init__(self):
        super().__init__()
        self.title("Optical Transceiver Label Inspector  v1.0")
        self.configure(bg=CLR["bg"])
        self.resizable(True, True)

        self.analyzer    = LabelAnalyzer()
        self.tmpl        = TemplateManager()
        self.cap         = None
        self.cam_running = False
        self.current_bgr = None
        self._static_bgr = None
        self.history     = []
        self._after_id   = None
        self._flip_h     = False
        self._flip_v     = False
        self._zoom       = 1.0

        _dir = os.path.dirname(os.path.abspath(__file__))
        self._log_path = os.path.join(
            _dir, f"label_log_{datetime.datetime.now().strftime('%Y%m%d')}.csv"
        )
        self._ensure_csv_header()
        self._build_ui()
        self.bind("<space>", lambda e: self._inspect()
                  if not isinstance(self.focus_get(), (tk.Entry, tk.Spinbox)) else None)
        self.bind("<F5>", lambda e: self._inspect())
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── UI ──────────────────────────────────────
    def _build_ui(self):
        self.columnconfigure(0, weight=3)
        self.columnconfigure(1, weight=2)
        self.rowconfigure(1, weight=1)

        # 헤더
        hdr = tk.Frame(self, bg=CLR["panel"], height=48)
        hdr.grid(row=0, column=0, columnspan=2, sticky="ew")
        tk.Label(hdr, text="▣  LABEL INSPECTOR", font=("Consolas", 14, "bold"),
                 fg=CLR["accent"], bg=CLR["panel"]).pack(side="left", padx=16, pady=10)
        self._lbl_clock = tk.Label(hdr, text="", font=FONT_SMALL,
                                   fg=CLR["subtext"], bg=CLR["panel"])
        self._lbl_clock.pack(side="right", padx=16)
        self._lbl_tmpl = tk.Label(hdr, text="[ 템플릿 없음 ]", font=FONT_SMALL,
                                  fg=CLR["warn"], bg=CLR["panel"])
        self._lbl_tmpl.pack(side="right", padx=8)
        self._tick_clock()

        # 좌측 패널
        left = tk.Frame(self, bg=CLR["bg"])
        left.grid(row=1, column=0, sticky="nsew", padx=8, pady=8)
        left.rowconfigure(0, weight=1)
        left.columnconfigure(0, weight=1)

        self._canvas = tk.Canvas(left, width=self.PREVIEW_W, height=self.PREVIEW_H,
                                 bg="#0d1117", highlightthickness=1,
                                 highlightbackground=CLR["border"])
        self._canvas.grid(row=0, column=0, sticky="nsew")
        self._canvas.create_text(self.PREVIEW_W // 2, self.PREVIEW_H // 2,
                                 text="[ 카메라 시작 또는 파일 열기 ]",
                                 fill=CLR["subtext"], font=FONT_BODY, tags="hint")
        self._canvas.bind("<MouseWheel>", self._on_mousewheel)

        # ctrl1: 카메라 제어
        ctrl1 = tk.Frame(left, bg=CLR["bg"])
        ctrl1.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        tk.Label(ctrl1, text="CAM #", font=FONT_SMALL,
                 fg=CLR["subtext"], bg=CLR["bg"]).pack(side="left", padx=(4, 2))
        self._cam_idx = tk.Spinbox(
            ctrl1, from_=0, to=9, width=3,
            bg=CLR["btn"], fg=CLR["text"],
            buttonbackground=CLR["panel"],
            insertbackground=CLR["text"],
            relief="flat", font=FONT_BODY, justify="center",
        )
        self._cam_idx.pack(side="left", padx=(0, 8))
        self._btn_cam  = self._btn(ctrl1, "📷  카메라 시작", self._toggle_camera)
        self._btn_file = self._btn(ctrl1, "📂  파일 열기",   self._open_file)
        self._btn_snap = self._btn(ctrl1, "🔍  검사 실행",   self._inspect, accent=True)
        for w in (self._btn_cam, self._btn_file, self._btn_snap):
            w.pack(side="left", padx=4, pady=4)

        # ctrl2: 템플릿 + 영상 보정
        ctrl2 = tk.Frame(left, bg=CLR["bg"])
        ctrl2.grid(row=2, column=0, sticky="ew", pady=(0, 4))

        self._btn_tmpl = self._btn(ctrl2, "📋  템플릿 로드", self._load_template)
        self._btn_edit = self._btn(ctrl2, "✏  템플릿 편집", self._edit_template)
        self._btn_tmpl.pack(side="left", padx=4, pady=2)
        self._btn_edit.pack(side="left", padx=4, pady=2)

        tk.Frame(ctrl2, bg=CLR["border"], width=1).pack(side="left", fill="y", padx=8, pady=4)

        self._btn_flip_h = self._btn(ctrl2, "↔ 좌우반전", self._toggle_flip_h)
        self._btn_flip_v = self._btn(ctrl2, "↕ 상하반전", self._toggle_flip_v)
        self._btn_flip_h.pack(side="left", padx=4, pady=2)
        self._btn_flip_v.pack(side="left", padx=4, pady=2)

        tk.Frame(ctrl2, bg=CLR["border"], width=1).pack(side="left", fill="y", padx=8, pady=4)

        self._btn(ctrl2, "−", self._zoom_out).pack(side="left", padx=(4, 1), pady=2)
        self._lbl_zoom = tk.Label(ctrl2, text="1.0×", font=FONT_SMALL,
                                  fg=CLR["text"], bg=CLR["btn"],
                                  width=5, cursor="hand2", relief="flat", padx=4, pady=5)
        self._lbl_zoom.pack(side="left", padx=1, pady=2)
        self._lbl_zoom.bind("<Button-1>", lambda e: self._zoom_reset())
        self._btn(ctrl2, "+", self._zoom_in).pack(side="left", padx=(1, 4), pady=2)

        # 우측 패널
        right = tk.Frame(self, bg=CLR["panel"], relief="flat")
        right.grid(row=1, column=1, sticky="nsew", padx=(0, 8), pady=8)
        right.columnconfigure(0, weight=1)
        right.rowconfigure(3, weight=1)

        tk.Label(right, text="INSPECTION RESULT", font=FONT_TITLE,
                 fg=CLR["accent"], bg=CLR["panel"]).grid(row=0, column=0,
                                                          sticky="w", padx=12, pady=(10, 4))
        self._lbl_verdict = tk.Label(right, text="—", font=("Consolas", 32, "bold"),
                                     fg=CLR["subtext"], bg=CLR["panel"])
        self._lbl_verdict.grid(row=1, column=0, pady=(0, 6))

        tk.Frame(right, bg=CLR["border"], height=1).grid(row=2, column=0, sticky="ew", padx=12)

        self._txt = tk.Text(right, bg=CLR["bg"], fg=CLR["text"],
                            font=FONT_BODY, relief="flat", wrap="word",
                            state="disabled", padx=8, pady=8)
        self._txt.grid(row=3, column=0, sticky="nsew", padx=8, pady=8)
        sb = ttk.Scrollbar(right, command=self._txt.yview)
        sb.grid(row=3, column=1, sticky="ns", pady=8)
        self._txt.configure(yscrollcommand=sb.set)
        self._configure_text_tags()

        bot = tk.Frame(right, bg=CLR["panel"])
        bot.grid(row=4, column=0, columnspan=2, sticky="ew", padx=8, pady=(0, 8))
        self._btn(bot, "📊  Excel 내보내기", self._save_result).pack(side="left", padx=4)
        self._btn(bot, "📋  이력 보기",       self._show_history).pack(side="left", padx=4)

        self._statusbar = tk.Label(self, text="준비",
                                   font=FONT_SMALL, fg=CLR["subtext"],
                                   bg=CLR["border"], anchor="w", padx=8)
        self._statusbar.grid(row=2, column=0, columnspan=2, sticky="ew")

    # ── 위젯 헬퍼 ────────────────────────────────
    def _btn(self, parent, text, cmd, accent=False):
        bg = CLR["accent"] if accent else CLR["btn"]
        fg = CLR["bg"]     if accent else CLR["text"]
        b  = tk.Button(parent, text=text, command=cmd, font=FONT_SMALL,
                       bg=bg, fg=fg, activebackground=CLR["btn_hover"],
                       activeforeground=CLR["text"], relief="flat",
                       padx=10, pady=5, cursor="hand2", bd=0)
        b.bind("<Enter>", lambda e: b.configure(bg=CLR["btn_hover"]) if not accent else None)
        b.bind("<Leave>", lambda e: b.configure(bg=bg))
        return b

    def _configure_text_tags(self):
        self._txt.tag_configure("ok",      foreground=CLR["ok"])
        self._txt.tag_configure("ng",      foreground=CLR["ng"])
        self._txt.tag_configure("warn",    foreground=CLR["warn"])
        self._txt.tag_configure("head",    foreground=CLR["accent"], font=FONT_TITLE)
        self._txt.tag_configure("subhead", foreground=CLR["subtext"], font=FONT_SMALL)
        self._txt.tag_configure("value",   foreground=CLR["text"])

    def _tick_clock(self):
        self._lbl_clock.configure(text=datetime.datetime.now().strftime("%Y-%m-%d  %H:%M:%S"))
        self.after(1000, self._tick_clock)

    # ── 반전 / 확대 ──────────────────────────────
    def _apply_flips(self, bgr):
        if self._flip_h:
            bgr = cv2.flip(bgr, 1)
        if self._flip_v:
            bgr = cv2.flip(bgr, 0)
        return bgr

    def _toggle_flip_h(self):
        self._flip_h = not self._flip_h
        self._btn_flip_h.configure(bg=CLR["btn_on"] if self._flip_h else CLR["btn"])
        self._refresh_static()

    def _toggle_flip_v(self):
        self._flip_v = not self._flip_v
        self._btn_flip_v.configure(bg=CLR["btn_on"] if self._flip_v else CLR["btn"])
        self._refresh_static()

    def _refresh_static(self):
        if not self.cam_running and not getattr(self, "_video_running", False):
            if self._static_bgr is not None:
                bgr = self._apply_flips(self._static_bgr.copy())
                self.current_bgr = bgr
                self._show_frame(bgr)

    def _zoom_in(self):
        if self._zoom < self._ZOOM_MAX:
            self._zoom = min(self._ZOOM_MAX, round(self._zoom + self._ZOOM_STEP, 1))
            self._lbl_zoom.configure(text=f"{self._zoom:.1f}×")

    def _zoom_out(self):
        if self._zoom > self._ZOOM_MIN:
            self._zoom = max(self._ZOOM_MIN, round(self._zoom - self._ZOOM_STEP, 1))
            self._lbl_zoom.configure(text=f"{self._zoom:.1f}×")

    def _zoom_reset(self):
        self._zoom = 1.0
        self._lbl_zoom.configure(text="1.0×")

    def _on_mousewheel(self, event):
        if event.delta > 0:
            self._zoom_in()
        else:
            self._zoom_out()

    # ── 카메라 ────────────────────────────────────
    def _toggle_camera(self):
        if self.cam_running:
            self._stop_camera()
        else:
            self._start_camera()

    def _start_camera(self):
        idx = int(self._cam_idx.get())
        self._static_bgr = None
        self._btn_cam.configure(state="disabled", text="연결 중...")
        self._status(f"CAM {idx} 연결 중...")
        threading.Thread(target=self._open_camera_bg, args=(idx,), daemon=True).start()

    @staticmethod
    def _fourcc_str(cap):
        v = int(cap.get(cv2.CAP_PROP_FOURCC))
        return "".join(chr((v >> 8 * i) & 0xFF) for i in range(4)).strip("\x00")

    def _open_camera_bg(self, idx):
        cap = None

        def _try(backend, w, h):
            result = [None]
            def worker():
                try:
                    c = cv2.VideoCapture(idx, backend)
                    if not c.isOpened():
                        return
                    c.set(cv2.CAP_PROP_FRAME_WIDTH, w)
                    c.set(cv2.CAP_PROP_FRAME_HEIGHT, h)
                    c.set(cv2.CAP_PROP_AUTO_EXPOSURE, 1)
                    ret, frame = c.read()
                    if ret and frame is not None:
                        result[0] = c
                    else:
                        c.release()
                except Exception:
                    pass
            t = threading.Thread(target=worker, daemon=True)
            t.start()
            t.join(timeout=6.0)
            return result[0]

        attempts = [
            ("DSHOW",    cv2.CAP_DSHOW,  640, 480),
            ("MSMF/VGA", cv2.CAP_MSMF,  640, 480),
            ("MSMF/HD",  cv2.CAP_MSMF, 1280, 720),
            ("AUTO",     0,              640, 480),
        ]
        for bname, backend, w, h in attempts:
            self.after(0, lambda s=f"CAM {idx} 연결 시도 중... [{bname}]": self._status(s))
            c = _try(backend, w, h)
            if c is not None:
                cap = c
                break

        self.after(0, lambda: self._on_camera_opened(cap, idx))

    def _on_camera_opened(self, cap, idx):
        self._btn_cam.configure(state="normal")
        if cap is None:
            self._btn_cam.configure(text="📷  카메라 시작")
            messagebox.showerror("오류",
                f"카메라 {idx}를 열 수 없습니다.\n\n"
                "확인 사항:\n"
                "  1. 카메라를 사용 중인 다른 앱을 종료 후 재시도하세요.\n"
                "  2. 카메라 연결(USB) 상태를 확인하세요.\n"
                "  3. CAM # 번호가 맞는지 확인하세요.")
            self._status(f"CAM {idx} 열기 실패")
            return
        self.cap           = cap
        self.cam_running   = True
        self._latest_frame = None
        self._frame_count  = 0
        self._btn_cam.configure(text="⏹  카메라 중지")
        fc = self._fourcc_str(cap)
        self._status(f"카메라 스트리밍 중... [FOURCC:{fc}]")
        threading.Thread(target=self._cam_reader, daemon=True).start()
        self.after(33, self._cam_display)

    def _stop_camera(self):
        self.cam_running   = False
        self._latest_frame = None
        if self._after_id:
            self.after_cancel(self._after_id)
        if self.cap:
            self.cap.release()
            self.cap = None
        self._btn_cam.configure(state="normal", text="📷  카메라 시작")
        self._status("카메라 중지됨")

    def _cam_reader(self):
        warmup = 5
        fail   = 0
        try:
            while self.cam_running and self.cap is not None:
                ret, frame = self.cap.read()
                if ret and frame is not None:
                    if warmup > 0:
                        warmup -= 1
                        continue
                    self._latest_frame = frame
                    self._frame_count  = getattr(self, "_frame_count", 0) + 1
                    fail = 0
                else:
                    fail += 1
                    if fail == 10:
                        self.after(0, lambda: self._status(
                            f"CAM {self._cam_idx.get()} — 프레임 없음. 카메라 재연결 또는 CAM # 확인"))
        except Exception as e:
            self.after(0, lambda: self._status(f"카메라 읽기 오류: {e}"))

    def _cam_display(self):
        if not self.cam_running:
            return
        frame = self._latest_frame
        if frame is not None:
            frame = self._apply_flips(frame)
            self.current_bgr   = frame
            self._latest_frame = None
            try:
                self._show_frame(frame)
                h, w = frame.shape[:2]
                fc   = getattr(self, "_frame_count", 0)
                self._status(f"스트리밍 중... {w}×{h}  밝기:{int(frame.mean())}  프레임:{fc}")
            except Exception as e:
                self._status(f"표시 오류: {e}")
        self._after_id = self.after(33, self._cam_display)

    # ── 파일 열기 ────────────────────────────────
    def _open_file(self):
        path = filedialog.askopenfilename(
            title="파일 선택",
            filetypes=[
                ("이미지/동영상", "*.jpg *.jpeg *.png *.bmp *.tiff *.tif *.mp4 *.avi *.mov *.mkv"),
                ("이미지",       "*.jpg *.jpeg *.png *.bmp *.tiff *.tif"),
                ("동영상",       "*.mp4 *.avi *.mov *.mkv"),
                ("모든 파일",    "*.*"),
            ]
        )
        if not path:
            return
        ext = os.path.splitext(path)[1].lower()
        if ext in (".mp4", ".avi", ".mov", ".mkv"):
            self._open_video(path)
        else:
            bgr = cv2.imdecode(np.fromfile(path, dtype=np.uint8), cv2.IMREAD_COLOR)
            if bgr is None:
                messagebox.showerror("오류", "이미지를 읽을 수 없습니다.")
                return
            self._stop_camera()
            self._stop_video()
            self._static_bgr = bgr.copy()
            bgr = self._apply_flips(bgr)
            self.current_bgr = bgr
            self._show_frame(bgr)
            self._status(f"이미지 로드: {os.path.basename(path)}")

    def _open_video(self, path):
        vcap = cv2.VideoCapture(path)
        if not vcap.isOpened():
            messagebox.showerror("오류", "동영상 파일을 열 수 없습니다.")
            return
        self._stop_camera()
        self._stop_video()
        self._static_bgr    = None
        self._vcap          = vcap
        self._video_running = True
        total = int(vcap.get(cv2.CAP_PROP_FRAME_COUNT))
        fps   = vcap.get(cv2.CAP_PROP_FPS) or 30
        self._video_delay   = max(15, int(1000 / fps))
        self._status(f"동영상: {os.path.basename(path)}  ({total}프레임, {fps:.1f}fps)")
        self._video_loop()

    def _stop_video(self):
        self._video_running = False
        if getattr(self, "_video_after_id", None):
            self.after_cancel(self._video_after_id)
            self._video_after_id = None
        if getattr(self, "_vcap", None):
            self._vcap.release()
            self._vcap = None

    def _video_loop(self):
        if not getattr(self, "_video_running", False) or self._vcap is None:
            return
        ret, frame = self._vcap.read()
        if ret and frame is not None:
            frame = self._apply_flips(frame)
            self.current_bgr = frame
            self._show_frame(frame)
            pos   = int(self._vcap.get(cv2.CAP_PROP_POS_FRAMES))
            total = int(self._vcap.get(cv2.CAP_PROP_FRAME_COUNT))
            self._status(f"동영상 재생 중... ({pos}/{total}프레임)")
        else:
            self._vcap.set(cv2.CAP_PROP_POS_FRAMES, 0)
        self._video_after_id = self.after(self._video_delay, self._video_loop)

    # ── 프레임 표시 ──────────────────────────────
    def _show_frame(self, bgr, overlays=None):
        if bgr.ndim == 2:
            bgr = cv2.cvtColor(bgr, cv2.COLOR_GRAY2BGR)
        elif bgr.ndim == 3 and bgr.shape[2] == 1:
            bgr = cv2.cvtColor(bgr.squeeze(), cv2.COLOR_GRAY2BGR)
        elif bgr.ndim == 3 and bgr.shape[2] == 4:
            bgr = bgr[:, :, :3]

        rgb = cv2.cvtColor(bgr.copy(), cv2.COLOR_BGR2RGB)

        if overlays:
            for ov in overlays:
                rect = ov.get("rect")
                if rect:
                    x, y, rw, rh = rect
                    cv2.rectangle(rgb, (x, y), (x + rw, y + rh), (0, 230, 100), 2)
                    cv2.putText(rgb, ov.get("type", ""), (x, y - 6),
                                cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 230, 100), 1)

        # 확대: 중앙 크롭
        if self._zoom > 1.0:
            fh, fw = rgb.shape[:2]
            zh, zw = int(fh / self._zoom), int(fw / self._zoom)
            y0 = (fh - zh) // 2
            x0 = (fw - zw) // 2
            rgb = rgb[y0:y0 + zh, x0:x0 + zw]

        cw = self._canvas.winfo_width()
        ch = self._canvas.winfo_height()
        if cw <= 1:  cw = self.PREVIEW_W
        if ch <= 1:  ch = self.PREVIEW_H
        fh, fw = rgb.shape[:2]
        scale  = min(cw / fw, ch / fh)
        nw, nh = int(fw * scale), int(fh * scale)
        rgb    = cv2.resize(rgb, (nw, nh), interpolation=cv2.INTER_AREA)

        img = ImageTk.PhotoImage(Image.fromarray(rgb))
        self._photo_ref = img
        self._canvas.delete("hint")
        self._canvas.delete("frame")
        self._canvas.create_image(cw // 2, ch // 2, image=img, anchor="center", tags="frame")
        self._canvas.update_idletasks()

    # ── 검사 ─────────────────────────────────────
    def _inspect(self):
        if self.current_bgr is None:
            messagebox.showwarning("경고", "카메라를 시작하거나 이미지 파일을 열어주세요.")
            return
        self._btn_snap.configure(state="disabled", text="분석 중...")
        self._status("검사 진행 중...")
        frame = self.current_bgr.copy()

        def worker():
            try:
                result = self.analyzer.analyze(frame)
                self.after(0, lambda: self._on_result(result, frame))
            except Exception as e:
                import traceback
                self.after(0, lambda: self._on_error(traceback.format_exc()))

        threading.Thread(target=worker, daemon=True).start()

    def _ensure_csv_header(self):
        if not os.path.exists(self._log_path):
            with open(self._log_path, "w", newline="", encoding="utf-8-sig") as f:
                csv.writer(f).writerow(self._CSV_COLS)

    def _append_csv_log(self, r):
        q     = r["quality"]
        flds  = r["ocr"]["fields"]
        codes = " | ".join(f"[{c['type']}]{c['data']}" for c in r.get("codes", []))
        comp  = r.get("comparison", {})
        pn    = comp.get("P/N", {})
        sn    = comp.get("S/N", {})
        row   = [
            r["timestamp"], r["verdict"],
            q["blur_score"], "OK" if q["blur_ok"] else "NG",
            q["brightness"], "OK" if q["bright_ok"] else "NG",
            q["tilt_deg"],   "OK" if q["tilt_ok"]   else "NG",
            flds.get("Model", ""), flds.get("P/N", ""), flds.get("S/N", ""),
            flds.get("Rev",   ""), flds.get("Date", ""), codes,
            pn.get("value", ""), ("OK" if pn["ok"] else "NG") if pn else "",
            sn.get("value", ""), ("OK" if sn["ok"] else "NG") if sn else "",
            self.tmpl.name if self.tmpl.loaded else "",
        ]
        with open(self._log_path, "a", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerow(row)

    def _on_error(self, msg):
        self._btn_snap.configure(state="normal", text="🔍  검사 실행")
        self._status("❌ 오류 발생")
        self._lbl_verdict.configure(text="[ ERR ]", fg=CLR["warn"])
        self._txt.configure(state="normal")
        self._txt.delete("1.0", "end")
        self._txt.insert("end", "오류 발생:\n\n", "head")
        self._txt.insert("end", msg, "ng")
        self._txt.configure(state="disabled")

    def _on_result(self, result, frame):
        self._btn_snap.configure(state="normal", text="🔍  검사 실행")
        comparison = self.tmpl.compare(result) if self.tmpl.loaded else {}
        result["comparison"] = comparison
        if comparison and not all(v["ok"] for v in comparison.values()):
            result["verdict"] = "NG"
        self.history.append(result)
        self._render_result(result)
        self._show_frame(frame, overlays=result.get("codes"))
        verdict = result["verdict"]
        self._lbl_verdict.configure(text=f"[ {verdict} ]",
                                    fg=CLR["ok"] if verdict == "OK" else CLR["ng"])
        self._append_csv_log(result)
        self._status(
            f"검사 완료 — {result['timestamp']}  →  {verdict}"
            f"  |  로그: {os.path.basename(self._log_path)}"
        )

    # ── 결과 렌더링 ──────────────────────────────
    def _render_result(self, r):
        self._txt.configure(state="normal")
        self._txt.delete("1.0", "end")

        def w(text, tag="value"):
            self._txt.insert("end", text, tag)

        def ok_ng(flag, yes="OK", no="NG"):
            w(f"  [{yes if flag else no}]\n", "ok" if flag else "ng")

        q = r["quality"]
        w("■ 인쇄 품질\n", "head")
        w(f"  흐림 지수  : {q['blur_score']:>8.1f}  ", "subhead");  ok_ng(q["blur_ok"], "CLEAR", "BLUR")
        w(f"  밝기       : {q['brightness']:>8d}  ",   "subhead");  ok_ng(q["bright_ok"])
        w(f"  기울기     : {q['tilt_deg']:>+8.2f}°  ", "subhead");  ok_ng(q["tilt_ok"], "STRAIGHT", "TILTED")
        w("\n")

        w("■ OCR 인식 결과\n", "head")
        for name, val in r["ocr"]["fields"].items():
            w(f"  {name:<8}: ", "subhead")
            w(f"{val if val else '(미인식)'}\n", "ok" if val else "ng")
        w("\n  ─ RAW TEXT ─\n", "subhead")
        raw = r["ocr"]["raw"]
        w(f"  {raw[:300]}{'…' if len(raw) > 300 else ''}\n\n", "value")

        w("■ 바코드 / QR 코드\n", "head")
        if r["codes"]:
            for c in r["codes"]:
                w(f"  [{c['type']}] ", "ok");  w(f"{c['data']}\n", "value")
        else:
            w("  pyzbar 미설치 — 바코드 검사 불가\n" if not PYZBAR_OK else "  코드 없음\n",
              "warn" if not PYZBAR_OK else "subhead")

        comp = r.get("comparison", {})
        if comp:
            w("\n■ 템플릿 비교\n", "head")
            for fname, info in comp.items():
                w(f"  {fname:<8}: ", "subhead")
                w(f"{info['value'] if info['value'] else '(없음)'}", "ok" if info["ok"] else "ng")
                if info["reason"]:
                    w(f"  ← {info['reason']}", "warn")
                w("\n")

        self._txt.configure(state="disabled")

    # ── 저장 ─────────────────────────────────────
    def _save_result(self):
        if not self.history:
            messagebox.showinfo("알림", "저장할 결과가 없습니다.")
            return
        if OPENPYXL_OK:
            default_ext = ".xlsx"
            ftypes = [("Excel 통합 문서", "*.xlsx"), ("CSV", "*.csv"), ("모든 파일", "*.*")]
        else:
            default_ext = ".csv"
            ftypes = [("CSV", "*.csv"), ("모든 파일", "*.*")]
        path = filedialog.asksaveasfilename(
            defaultextension=default_ext, filetypes=ftypes,
            initialfile=f"label_export_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}{default_ext}"
        )
        if not path:
            return
        if path.lower().endswith(".csv"):
            self._export_csv(path)
        else:
            self._export_xlsx(path)

    def _export_csv(self, path):
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            wr = csv.writer(f)
            wr.writerow(self._CSV_COLS)
            for r in self.history:
                q     = r["quality"];  flds = r["ocr"]["fields"]
                codes = " | ".join(f"[{c['type']}]{c['data']}" for c in r.get("codes", []))
                comp  = r.get("comparison", {})
                pn    = comp.get("P/N", {})
                sn    = comp.get("S/N", {})
                wr.writerow([
                    r["timestamp"], r["verdict"],
                    q["blur_score"], "OK" if q["blur_ok"] else "NG",
                    q["brightness"], "OK" if q["bright_ok"] else "NG",
                    q["tilt_deg"],   "OK" if q["tilt_ok"]  else "NG",
                    flds.get("Model", ""), flds.get("P/N", ""), flds.get("S/N", ""),
                    flds.get("Rev",   ""), flds.get("Date", ""), codes,
                    pn.get("value", ""), ("OK" if pn["ok"] else "NG") if pn else "",
                    sn.get("value", ""), ("OK" if sn["ok"] else "NG") if sn else "",
                    self.tmpl.name if self.tmpl.loaded else "",
                ])
        self._status(f"CSV 저장 완료: {os.path.basename(path)}")

    def _export_xlsx(self, path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Label Log"
        hdr_fill = PatternFill("solid", fgColor="1A3A5C")
        hdr_font = Font(bold=True, color="FFFFFF", name="Consolas", size=10)
        ok_fill  = PatternFill("solid", fgColor="C8F0D4")
        ng_fill  = PatternFill("solid", fgColor="FFD0D0")
        center   = Alignment(horizontal="center", vertical="center")

        for col, name in enumerate(self._CSV_COLS, 1):
            c = ws.cell(row=1, column=col, value=name)
            c.fill = hdr_fill;  c.font = hdr_font;  c.alignment = center

        for ri, r in enumerate(self.history, 2):
            q     = r["quality"];  flds = r["ocr"]["fields"]
            codes = " | ".join(f"[{c['type']}]{c['data']}" for c in r.get("codes", []))
            comp  = r.get("comparison", {})
            pn    = comp.get("P/N", {})
            sn    = comp.get("S/N", {})
            vals  = [
                r["timestamp"], r["verdict"],
                q["blur_score"], "OK" if q["blur_ok"] else "NG",
                q["brightness"], "OK" if q["bright_ok"] else "NG",
                q["tilt_deg"],   "OK" if q["tilt_ok"]  else "NG",
                flds.get("Model", ""), flds.get("P/N", ""), flds.get("S/N", ""),
                flds.get("Rev",   ""), flds.get("Date", ""), codes,
                pn.get("value", ""), ("OK" if pn["ok"] else "NG") if pn else "",
                sn.get("value", ""), ("OK" if sn["ok"] else "NG") if sn else "",
                self.tmpl.name if self.tmpl.loaded else "",
            ]
            for ci, val in enumerate(vals, 1):
                cell = ws.cell(row=ri, column=ci, value=val)
                if ci in (2, 16, 18):
                    cell.fill = ok_fill if val == "OK" else ng_fill
                    cell.font = Font(bold=True, name="Consolas", size=10)
                    cell.alignment = center
                elif ci in (4, 6, 8):
                    cell.fill = ok_fill if val == "OK" else ng_fill
                    cell.alignment = center

        for col in ws.columns:
            max_len = max((len(str(cell.value or "")) for cell in col), default=8)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 40)
        ws.freeze_panes = "A2"
        ws.row_dimensions[1].height = 20
        wb.save(path)
        self._status(f"Excel 저장 완료: {os.path.basename(path)}")

    # ── 템플릿 ───────────────────────────────────
    def _load_template(self):
        path = filedialog.askopenfilename(
            title="템플릿 파일 선택",
            filetypes=[("JSON 템플릿", "*.json"), ("모든 파일", "*.*")]
        )
        if not path:
            return
        try:
            self.tmpl.load(path)
            self._lbl_tmpl.configure(text=f"[ {self.tmpl.name} ]", fg=CLR["ok"])
            self._status(f"템플릿 로드: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("오류", f"템플릿 로드 실패:\n{e}")

    def _edit_template(self):
        data = self.tmpl.data or json.loads(json.dumps(TemplateManager.EMPTY))
        win  = tk.Toplevel(self)
        win.title("템플릿 편집")
        win.configure(bg=CLR["bg"])
        win.geometry("500x500")
        win.grab_set()

        f = tk.Frame(win, bg=CLR["bg"])
        f.pack(fill="both", expand=True, padx=14, pady=12)
        f.columnconfigure(1, weight=1)

        def lbl(text, row):
            tk.Label(f, text=text, font=FONT_SMALL, fg=CLR["subtext"],
                     bg=CLR["bg"], anchor="w").grid(row=row, column=0,
                     sticky="w", padx=(0, 8), pady=3)

        def ent(row, val="", width=32):
            e = tk.Entry(f, font=FONT_BODY, bg=CLR["btn"], fg=CLR["text"],
                         insertbackground=CLR["text"], relief="flat", width=width)
            e.insert(0, val)
            e.grid(row=row, column=1, sticky="ew", pady=3)
            return e

        def section(text, row):
            tk.Label(f, text=text, font=FONT_TITLE, fg=CLR["accent"],
                     bg=CLR["bg"]).grid(row=row, column=0, columnspan=2,
                     sticky="w", pady=(10, 3))

        section("▣ 기본 정보", 0)
        lbl("템플릿 이름", 1);  e_name   = ent(1, data.get("template_name", ""))
        section("▣ 바코드", 2)
        lbl("P/N 고정값",  3);  e_pn_val = ent(3, data["barcodes"]["PN"].get("fixed_value", ""))
        lbl("P/N 정규식",  4);  e_pn_pat = ent(4, data["barcodes"]["PN"].get("pattern",     ""))
        lbl("S/N 정규식",  5);  e_sn_pat = ent(5, data["barcodes"]["SN"].get("pattern",     ""))
        section("▣ OCR 텍스트 포함 검사", 6)

        ocr_frame = tk.Frame(f, bg=CLR["bg"])
        ocr_frame.grid(row=7, column=0, columnspan=2, sticky="ew")
        ocr_frame.columnconfigure(1, weight=1)
        ocr_frame.columnconfigure(3, weight=1)
        ocr_rows = []

        def add_ocr_row(label="", contains=""):
            r = len(ocr_rows)
            tk.Label(ocr_frame, text="항목명", font=FONT_SMALL,
                     fg=CLR["subtext"], bg=CLR["bg"]).grid(row=r, column=0, padx=(0, 4))
            el = tk.Entry(ocr_frame, font=FONT_BODY, bg=CLR["btn"], fg=CLR["text"],
                          insertbackground=CLR["text"], relief="flat", width=10)
            el.insert(0, label);  el.grid(row=r, column=1, sticky="ew", padx=4, pady=2)
            tk.Label(ocr_frame, text="포함 텍스트", font=FONT_SMALL,
                     fg=CLR["subtext"], bg=CLR["bg"]).grid(row=r, column=2, padx=(8, 4))
            ev = tk.Entry(ocr_frame, font=FONT_BODY, bg=CLR["btn"], fg=CLR["text"],
                          insertbackground=CLR["text"], relief="flat", width=16)
            ev.insert(0, contains);  ev.grid(row=r, column=3, sticky="ew", padx=4, pady=2)
            ocr_rows.append((el, ev))

        for item in data.get("ocr_texts", []):
            add_ocr_row(item.get("label", ""), item.get("contains", ""))
        if not ocr_rows:
            add_ocr_row()

        self._btn(f, "+ 행 추가", add_ocr_row).grid(
            row=8, column=0, columnspan=2, sticky="w", pady=(4, 0))

        bot = tk.Frame(win, bg=CLR["bg"])
        bot.pack(fill="x", padx=14, pady=8)

        def save_and_close():
            new_data = {
                "template_name": e_name.get().strip() or "이름없음",
                "barcodes": {
                    "PN": {"fixed_value": e_pn_val.get().strip(),
                           "pattern":     e_pn_pat.get().strip(),
                           "required":    True},
                    "SN": {"pattern":  e_sn_pat.get().strip(),
                           "required": False},
                },
                "ocr_texts": [
                    {"label": el.get().strip(), "contains": ev.get().strip()}
                    for el, ev in ocr_rows
                    if el.get().strip() or ev.get().strip()
                ],
            }
            path = filedialog.asksaveasfilename(
                defaultextension=".json",
                filetypes=[("JSON 템플릿", "*.json")],
                initialfile=f"{new_data['template_name']}.json",
                title="템플릿 저장",
            )
            if path:
                self.tmpl.data = new_data
                self.tmpl.save(path)
                self._lbl_tmpl.configure(text=f"[ {new_data['template_name']} ]", fg=CLR["ok"])
                self._status(f"템플릿 저장: {os.path.basename(path)}")
                win.destroy()

        self._btn(bot, "💾  저장", save_and_close, accent=True).pack(side="left", padx=4)
        self._btn(bot, "닫기",     win.destroy).pack(side="left", padx=4)

    # ── 이력 ─────────────────────────────────────
    def _show_history(self):
        if not self.history:
            messagebox.showinfo("이력", "검사 이력이 없습니다.")
            return
        win = tk.Toplevel(self)
        win.title("검사 이력")
        win.configure(bg=CLR["bg"])
        win.geometry("700x440")

        total  = len(self.history)
        ok_cnt = sum(1 for h in self.history if h["verdict"] == "OK")
        stat_fr = tk.Frame(win, bg=CLR["panel"])
        stat_fr.pack(fill="x", padx=8, pady=(8, 0))
        tk.Label(stat_fr, text=f"총 {total}건", font=FONT_SMALL,
                 fg=CLR["text"],  bg=CLR["panel"]).pack(side="left", padx=10, pady=4)
        tk.Label(stat_fr, text=f"OK {ok_cnt}", font=FONT_SMALL,
                 fg=CLR["ok"],   bg=CLR["panel"]).pack(side="left", padx=6)
        tk.Label(stat_fr, text=f"NG {total - ok_cnt}", font=FONT_SMALL,
                 fg=CLR["ng"],   bg=CLR["panel"]).pack(side="left", padx=6)

        tree_fr = tk.Frame(win, bg=CLR["bg"])
        tree_fr.pack(fill="both", expand=True, padx=8, pady=4)
        cols = ("시각", "판정", "흐림", "밝기", "기울기(°)", "Model", "P/N", "S/N")
        tv   = ttk.Treeview(tree_fr, columns=cols, show="headings")
        sb   = ttk.Scrollbar(tree_fr, command=tv.yview)
        tv.configure(yscrollcommand=sb.set)
        for c in cols:
            tv.heading(c, text=c)
            tv.column(c, width=80, anchor="center")
        tv.column("시각",  width=140)
        tv.column("Model", width=110)
        tv.column("P/N",   width=100)

        for h in reversed(self.history):
            q = h["quality"];  fld = h["ocr"]["fields"]
            tv.insert("", "end", values=(
                h["timestamp"], h["verdict"],
                q["blur_score"], q["brightness"], q["tilt_deg"],
                fld.get("Model", ""), fld.get("P/N", ""), fld.get("S/N", ""),
            ), tags=(h["verdict"],))

        tv.tag_configure("OK", foreground=CLR["ok"])
        tv.tag_configure("NG", foreground=CLR["ng"])
        sb.pack(side="right", fill="y")
        tv.pack(side="left",  fill="both", expand=True)
        self._btn(win, "닫기", win.destroy).pack(pady=(0, 8))

    # ── 상태 바 / 종료 ────────────────────────────
    def _status(self, msg):
        self._statusbar.configure(text=f"  {msg}")

    def _on_close(self):
        self._stop_camera()
        self._stop_video()
        self.destroy()


# ──────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.geometry("1100x680")
    app.mainloop()
