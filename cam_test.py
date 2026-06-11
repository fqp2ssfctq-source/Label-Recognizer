import cv2, time, struct
print("Python:", struct.calcsize("P")*8, "bit")
print("LogitechCamera.exe가 실행 중인 상태에서 테스트합니다.")
print()

# LogitechCamera.exe가 실행 중인 상태에서 동시 접근 시도
for idx in [1, 2]:
    cap = cv2.VideoCapture(idx, cv2.CAP_DSHOW)
    if not cap.isOpened():
        print(f"CAM{idx}: 열기 실패")
        continue
    fcc = int(cap.get(cv2.CAP_PROP_FOURCC)).to_bytes(4,'little').decode('ascii','replace').strip()
    start = time.time()
    best, count = 0, 0
    while time.time() - start < 3.0:
        ret, frame = cap.read()
        if ret and frame is not None:
            m = float(frame.mean())
            count += 1
            if m > best: best = m
            if m > 20:
                cv2.imwrite(f"cam{idx}_logi.jpg", frame)
                print(f"CAM{idx}: 유효! mean={m:.1f}  frames={count}  FOURCC={fcc}")
                break
    if best <= 20:
        print(f"CAM{idx}: frames={count}  fps={count/3:.1f}  best_mean={best:.3f}  FOURCC={fcc}")
    cap.release()
