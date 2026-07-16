"""
카메라 설정 테스트 — 어떤 backend/FOURCC/해상도 조합이 동작하는지 확인
실행: python cam_test.py
"""
import cv2, time

IDX  = 0        # 0과 1 둘 다 시도 — MSMF/DSHOW 인덱스 다를 수 있음
BACKENDS = [
    (cv2.CAP_DSHOW, "DSHOW"),
    (cv2.CAP_MSMF,  "MSMF"),
]
FOURCCS = [
    (None,   "없음"),
    ("MJPG", "MJPG"),
    ("YUY2", "YUY2"),
]
RESOLUTIONS = [
    (640,  480),
    (1280, 720),
    (1920, 1080),
]

def test(bk, bk_name, fourcc_str, w, h):
    try:
        c = cv2.VideoCapture(IDX, bk)
        if not c.isOpened():
            print(f"  [{bk_name}] {fourcc_str or '없음':4s} {w}×{h:4d}  →  열기 실패")
            return
        if fourcc_str:
            c.set(cv2.CAP_PROP_FOURCC, cv2.VideoWriter_fourcc(*fourcc_str))
        t0 = time.time()
        c.set(cv2.CAP_PROP_FRAME_WIDTH,  w)
        c.set(cv2.CAP_PROP_FRAME_HEIGHT, h)
        t1 = time.time()
        # 실제 해상도
        rw = int(c.get(cv2.CAP_PROP_FRAME_WIDTH))
        rh = int(c.get(cv2.CAP_PROP_FRAME_HEIGHT))
        fv = int(c.get(cv2.CAP_PROP_FOURCC))
        fc = "".join(chr((fv>>(8*i))&0xFF) for i in range(4)).strip('\x00') or "?"
        # 첫 프레임
        t2 = time.time()
        ret, frm = c.read()
        t3 = time.time()
        c.release()
        if not ret or frm is None:
            status = "read 실패"
        else:
            mean = frm.mean()
            status = f"OK  mean={mean:.1f}  실제:{rw}×{rh}  포맷:{fc}  set={t1-t0:.2f}s  read={t3-t2:.2f}s"
        print(f"  [{bk_name}] {fourcc_str or '없음':4s} {w}×{h:4d}  →  {status}")
    except Exception as e:
        print(f"  [{bk_name}] {fourcc_str or '없음':4s} {w}×{h:4d}  →  예외: {e}")

print(f"\nIPEVO V4K 카메라 테스트 (index={IDX})\n{'='*60}")
for bk, bk_name in BACKENDS:
    print(f"\n[{bk_name}]")
    for fourcc_str, _ in FOURCCS:
        for w, h in RESOLUTIONS:
            test(bk, bk_name, fourcc_str, w, h)

print("\n완료")
