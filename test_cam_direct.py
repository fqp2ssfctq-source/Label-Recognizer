"""
OBS 없이 GENERAL-UVC 카메라 직접 접근 테스트
- DSHOW 백엔드로 각 인덱스 시도
- 1280x720 해상도 설정 시도
"""
import cv2
import time

def test_camera(idx, backend_name, backend):
    print(f"\n--- CAM {idx} ({backend_name}) ---")
    cap = cv2.VideoCapture(idx, backend)
    if not cap.isOpened():
        print("  열기 실패")
        return False

    # 1280x720 설정 시도
    cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)
    cap.set(cv2.CAP_PROP_FOURCC, cv2.VideoWriter_fourcc(*'MJPG'))

    w = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
    h = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
    fps = cap.get(cv2.CAP_PROP_FPS)
    fourcc_val = int(cap.get(cv2.CAP_PROP_FOURCC))
    fourcc_str = "".join([chr((fourcc_val >> 8*i) & 0xFF) for i in range(4)])
    print(f"  설정 후 해상도: {w}x{h}, FPS: {fps}, FOURCC: {fourcc_str}")

    # 프레임 읽기 시도
    for attempt in range(10):
        ret, frame = cap.read()
        if ret and frame is not None:
            print(f"  ✓ 프레임 수신! 실제 크기: {frame.shape[1]}x{frame.shape[0]}")
            cap.release()
            return True
        time.sleep(0.1)

    print("  ✗ 프레임 없음")
    cap.release()
    return False

print("=== 카메라 직접 접근 테스트 (OBS 없이) ===\n")

# DSHOW로 인덱스 0~3 테스트
for i in range(4):
    test_camera(i, "DSHOW", cv2.CAP_DSHOW)

print("\n=== 테스트 완료 ===")
