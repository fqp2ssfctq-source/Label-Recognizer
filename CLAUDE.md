# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 프로젝트 개요

Optical Transceiver 라벨 자동 검사 GUI. 카메라 또는 이미지/동영상 파일에서 라벨을 획득하고, OCR·바코드·인쇄품질을 판정하여 CSV/XLSX로 기록한다.

- 파일: `Label_Recognizer.pyw` (단일 파일, ~1200줄)
- 버전: v1.0 (2026-06-08)

## 실행 방법

```
pip install opencv-python pillow pytesseract pyzbar numpy openpyxl
# Tesseract-OCR 설치: https://github.com/UB-Mannheim/tesseract/wiki
python Label_Recognizer.pyw
```

Tesseract는 `C:\Program Files\Tesseract-OCR\tesseract.exe`에 있으면 자동 감지된다.

## 코드 구조

### 3개 클래스

| 클래스 | 역할 |
|--------|------|
| `LabelAnalyzer` | 이미지 1장 분석: 인쇄품질 + OCR + 바코드 → `dict` 반환 |
| `TemplateManager` | JSON 템플릿 로드/저장, `compare()` 로 P/N·S/N 검증 |
| `App(tk.Tk)` | 전체 GUI, 카메라 스트리밍, 이력 관리, CSV/XLSX 내보내기 |

### 결과 dict 구조 (`LabelAnalyzer.analyze()` 반환값)
```python
{
  "timestamp": "2026-06-08 14:00:00",
  "verdict":   "OK" | "NG",
  "quality":   {"blur_score", "blur_ok", "brightness", "bright_ok", "tilt_deg", "tilt_ok"},
  "ocr":       {"raw": str, "fields": {"Model","P/N","S/N","Rev","Date"}},
  "codes":     [{"type","data","rect"}, ...],
  "comparison": {"P/N": {"ok","value","reason"}, "S/N": {...}}  # 템플릿 로드 시만
}
```

### CSV/XLSX 컬럼 19개 (`_CSV_COLS`)
`Timestamp, Verdict, Blur, Blur_OK, Bright, Bright_OK, Tilt, Tilt_OK, Model, PN_OCR, SN_OCR, Rev, Date, Barcodes, PN_Value, PN_OK, SN_Value, SN_OK, Template`

컬럼 인덱스(1-based): Verdict=2, PN_OK=16, SN_OK=18 → XLSX 컬러링 대상

## 카메라 관련 특이사항

### FOURCC 전략
`_open_camera_bg()`: CAP_DSHOW → 기본 → CAP_MSMF 순서로 시도, MJPEG FOURCC 우선 설정.
일부 카메라는 MJPEG 미지원 시 set()이 무시되고 기본 포맷 유지됨 (정상 동작).

### YUY2 줄무늬 자동 보정
카메라가 BGR 버퍼에 YUY2 raw 데이터를 넣어주는 경우 자동 감지·변환:
- `_is_striped()`: 인접 픽셀 음의 상관(corr < -0.3) → YUY2 판정
- `_fix_yuyv()`: 오프셋 S0/SM/SE 및 WC/WR 방식 순차 시도
- GUI의 "YUY2 강제" 체크박스로 수동 활성화 가능

### 스레딩 구조
- `_cam_reader()`: 백그라운드 스레드 — `_latest_frame` 에 최신 프레임 저장
- `_cam_display()`: 메인 스레드, `after(33ms)` 루프 — `_latest_frame` 읽어 표시
- `_show_frame()`: BGR 채널 정규화 (gray/1ch/4ch → 3ch BGR) 후 Tkinter 표시

### 카메라 디버깅 진단
상태 바 메시지로 판단:
- `프레임:0` → 카메라가 실제 프레임을 전달하지 않음 (CAP_DSHOW 문제)
- `표시 오류:` → `_show_frame()`에서 예외 (채널/형식 문제)
- `진단` 버튼 → 현재 프레임 픽셀값을 `diag.txt`에 저장 (YUY2 여부 판별용)

## 색상 팔레트 (산업용 다크 테마)

```python
CLR = {
    "bg": "#1a1d23", "panel": "#22262f", "border": "#2e3340",
    "accent": "#00c2ff", "ok": "#00e676", "ng": "#ff3d3d",
    "warn": "#ffaa00", "text": "#e8ecf0", "subtext": "#8899aa",
    "btn": "#2c3244", "btn_hover": "#3a4257",
}
FONT_TITLE  = ("Consolas", 13, "bold")
FONT_BODY   = ("Consolas", 10)
FONT_SMALL  = ("Consolas",  9)
FONT_RESULT = ("Consolas", 11, "bold")
```

## 템플릿 JSON 형식

```json
{
  "template_name": "제품명",
  "barcodes": {
    "PN": {"fixed_value": "ABC-123", "pattern": "", "required": true},
    "SN": {"pattern": "[A-Z0-9]{10}", "required": false}
  },
  "ocr_texts": [
    {"label": "제조국", "contains": "Singapore"}
  ]
}
```

## 자동 로그

검사 완료 시 `label_log_YYYYMMDD.csv`에 자동 기록 (`_auto_log()`).
