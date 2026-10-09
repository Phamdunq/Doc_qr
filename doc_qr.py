# qr_decoder.py
# ============================================================
# QR DECODER V3 - CCCD / ID PHOTO
# ============================================================
#
# Muc tieu:
# - Doc QR tu anh CCCD chup bang dien thoai
# - Ho tro anh toi, mo, nghieng, nho, co nen phuc tap
# - Tu tao nhieu vung nghi ngo QR
# - Xu ly sang/toi, CLAHE, denoise, sharpen, threshold
# - Thu nhieu goc xoay
# - ZXing-C++ la decoder chinh
# - Ho tro them WeChatQRCode neu may da co model
#
# Cai dat toi thieu:
#   py -m pip install opencv-python numpy zxing-cpp
#
# Khuyen nghi neu muon dung them WeChatQRCode:
#   py -m pip install opencv-contrib-python
#
# Cach chay:
#   py qr_decoder.py "D:\CCCD\001.jpg"
#
# Ket qua:
#   STATUS | SUCCESS
#   QR_DATA | ...
#
# ============================================================

import sys
from pathlib import Path
from typing import List, Tuple, Optional

import cv2
import numpy as np

try:
    import zxingcpp
except ImportError:
    zxingcpp = None


# ------------------------------------------------------------
# CẤU HÌNH
# ------------------------------------------------------------

# Tạo ảnh debug khi không đọc được.
# True = lưu các crop tốt nhất vào thư mục debug_qr
SAVE_DEBUG = True

# Số lượng candidate tối đa đem đi decode.
MAX_CANDIDATES = 100

# Các tỉ lệ crop phía trên bên phải của ảnh.
# Dùng cho trường hợp ảnh CCCD giống ảnh mẫu.
TOP_RIGHT_CROPS = [
    (0.52, 0.18, 0.98, 0.58),
    (0.58, 0.20, 0.96, 0.52),
    (0.62, 0.22, 0.94, 0.48),
    (0.48, 0.12, 0.99, 0.65),
]


# ------------------------------------------------------------
# ĐỌC ẢNH
# ------------------------------------------------------------

def load_image(path: Path) -> Optional[np.ndarray]:
    try:
        raw = np.fromfile(str(path), dtype=np.uint8)
        return cv2.imdecode(raw, cv2.IMREAD_COLOR)
    except Exception:
        return None


# ------------------------------------------------------------
# CHUẨN HÓA ẢNH
# ------------------------------------------------------------

def resize_for_processing(img: np.ndarray) -> np.ndarray:
    """
    Giữ ảnh đủ chi tiết nhưng tránh ảnh quá lớn làm chậm chương trình.
    """
    h, w = img.shape[:2]
    longest = max(h, w)

    # Nếu ảnh quá lớn, giảm về khoảng 3000 px.
    if longest > 3200:
        scale = 3200.0 / longest
        img = cv2.resize(
            img,
            None,
            fx=scale,
            fy=scale,
            interpolation=cv2.INTER_AREA,
        )

    return img


def normalize_lighting(gray: np.ndarray) -> np.ndarray:
    """
    Cân bằng sáng + tăng tương phản cục bộ.
    """
    clahe = cv2.createCLAHE(
        clipLimit=3.0,
        tileGridSize=(8, 8)
    )
    return clahe.apply(gray)


# ------------------------------------------------------------
# TẠO CROP NGHI NGỜ QR
# ------------------------------------------------------------

def add_candidate(candidates, name, img):
    if img is None or img.size == 0:
        return

    h, w = img.shape[:2]

    # Bỏ những ảnh quá nhỏ.
    if h < 80 or w < 80:
        return

    candidates.append((name, img))


def crop_top_right_regions(img: np.ndarray, candidates):
    """
    CCCD mặt trước của Việt Nam thường có QR ở khu vực phía trên bên phải.
    Đây là chiến lược quan trọng cho ảnh giống mẫu người dùng gửi.
    """
    h, w = img.shape[:2]

    for i, (x1, y1, x2, y2) in enumerate(TOP_RIGHT_CROPS):
        crop = img[
            int(h * y1):int(h * y2),
            int(w * x1):int(w * x2)
        ]
        add_candidate(candidates, f"top_right_{i+1}", crop)


def crop_grid_regions(img: np.ndarray, candidates):
    """
    Chia ảnh thành các vùng chồng lấn để QR không bị bỏ sót
    khi ảnh có bố cục khác.
    """
    h, w = img.shape[:2]

    # Grid 3x3 với overlap.
    xs = [(0.00, 0.55), (0.225, 0.775), (0.45, 1.00)]
    ys = [(0.00, 0.55), (0.225, 0.775), (0.45, 1.00)]

    index = 0
    for y1, y2 in ys:
        for x1, x2 in xs:
            index += 1
            crop = img[
                int(h * y1):int(h * y2),
                int(w * x1):int(w * x2)
            ]
            add_candidate(candidates, f"grid_{index}", crop)


def find_qr_like_contours(img: np.ndarray, candidates):
    """
    Tìm vùng có cấu trúc QR bằng contour.
    Không phụ thuộc tuyệt đối vào tọa độ.
    """
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Làm nổi vùng đen/trắng.
    gray = cv2.GaussianBlur(gray, (3, 3), 0)

    for threshold_type in [
        cv2.THRESH_BINARY,
        cv2.THRESH_BINARY_INV,
        cv2.THRESH_BINARY + cv2.THRESH_OTSU,
        cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU,
    ]:
        try:
            if threshold_type & cv2.THRESH_OTSU:
                _, binary = cv2.threshold(
                    gray, 0, 255, threshold_type
                )
            else:
                _, binary = cv2.threshold(
                    gray, 145, 255, threshold_type
                )

            contours, _ = cv2.findContours(
                binary,
                cv2.RETR_LIST,
                cv2.CHAIN_APPROX_SIMPLE
            )

            h, w = gray.shape

            for idx, contour in enumerate(contours):
                x, y, cw, ch = cv2.boundingRect(contour)

                area = cw * ch
                if area < (w * h * 0.002):
                    continue
                if area > (w * h * 0.35):
                    continue

                ratio = cw / float(ch)

                # QR thường gần vuông.
                if 0.55 <= ratio <= 1.8:
                    pad_x = int(cw * 0.20)
                    pad_y = int(ch * 0.20)

                    x1 = max(0, x - pad_x)
                    y1 = max(0, y - pad_y)
                    x2 = min(w, x + cw + pad_x)
                    y2 = min(h, y + ch + pad_y)

                    crop = img[y1:y2, x1:x2]

                    add_candidate(
                        candidates,
                        f"contour_{threshold_type}_{idx}",
                        crop
                    )

        except Exception:
            pass


# ------------------------------------------------------------
# TẠO CÁC PHIÊN BẢN XỬ LÝ ẢNH
# ------------------------------------------------------------

def preprocessing_variants(img: np.ndarray):
    """
    Một crop -> nhiều phiên bản:
    - gốc
    - grayscale
    - CLAHE
    - gamma sáng/tối
    - denoise
    - sharpen
    - Otsu
    - adaptive threshold
    """
    result = []

    if img is None or img.size == 0:
        return result

    # Resize crop.
    h, w = img.shape[:2]

    # QR nhỏ -> phóng to mạnh.
    longest = max(h, w)

    if longest < 800:
        scales = [3.0, 4.0, 5.0]
    elif longest < 1500:
        scales = [2.0, 3.0, 4.0]
    else:
        scales = [1.5, 2.0]

    # BGR
    result.append(("color", img))

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    result.append(("gray", gray))

    # Gamma correction
    def gamma_correct(source, gamma):
        inv = 1.0 / gamma
        table = np.array([
            ((i / 255.0) ** inv) * 255
            for i in np.arange(256)
        ]).astype("uint8")
        return cv2.LUT(source, table)

    # 0.65 = làm sáng vùng tối
    bright = gamma_correct(gray, 0.65)
    result.append(("bright", bright))

    # 1.5 = làm rõ vùng sáng/tương phản
    dark = gamma_correct(gray, 1.5)
    result.append(("dark", dark))

    clahe = normalize_lighting(gray)
    result.append(("clahe", clahe))

    # Denoise nhẹ.
    denoise = cv2.fastNlMeansDenoising(
        clahe,
        None,
        7,
        7,
        21
    )
    result.append(("denoise", denoise))

    # Sharpen.
    blur = cv2.GaussianBlur(denoise, (0, 0), 2.0)
    sharpen = cv2.addWeighted(
        denoise,
        2.0,
        blur,
        -1.0,
        0
    )
    result.append(("sharpen", sharpen))

    for scale in scales:
        for name, source in list(result):
            enlarged = cv2.resize(
                source,
                None,
                fx=scale,
                fy=scale,
                interpolation=cv2.INTER_CUBIC
            )

            yield f"{name}_x{scale}", enlarged

            # Otsu
            if len(enlarged.shape) == 2:
                _, otsu = cv2.threshold(
                    enlarged,
                    0,
                    255,
                    cv2.THRESH_BINARY + cv2.THRESH_OTSU
                )
                yield f"{name}_otsu_x{scale}", otsu

                # Inverse Otsu
                _, otsu_inv = cv2.threshold(
                    enlarged,
                    0,
                    255,
                    cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU
                )
                yield f"{name}_otsu_inv_x{scale}", otsu_inv

                # Adaptive
                adaptive = cv2.adaptiveThreshold(
                    enlarged,
                    255,
                    cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                    cv2.THRESH_BINARY,
                    31,
                    5
                )
                yield f"{name}_adaptive_x{scale}", adaptive


# ------------------------------------------------------------
# XOAY ẢNH
# ------------------------------------------------------------

def rotated_images(img):
    yield "0", img

    yield "90", cv2.rotate(
        img,
        cv2.ROTATE_90_CLOCKWISE
    )

    yield "180", cv2.rotate(
        img,
        cv2.ROTATE_180
    )

    yield "270", cv2.rotate(
        img,
        cv2.ROTATE_90_COUNTERCLOCKWISE
    )


# ------------------------------------------------------------
# ZXING
# ------------------------------------------------------------

def zxing_decode(img):
    if zxingcpp is None:
        return None

    try:
        results = zxingcpp.read_barcodes(
            img,
            formats=zxingcpp.BarcodeFormat.QRCode,
            try_rotate=True,
            try_downscale=True,
        )

        for result in results:
            if result.text:
                return result.text

    except Exception:
        return None

    return None


# ------------------------------------------------------------
# OPENCV QR
# ------------------------------------------------------------

def opencv_decode(img):
    detector = cv2.QRCodeDetector()

    try:
        data, points, _ = detector.detectAndDecode(img)

        if data:
            return data

    except Exception:
        pass

    # Multi QR
    try:
        result = detector.detectAndDecodeMulti(img)

        if len(result) == 4:
            ok, decoded_info, _, _ = result

            if ok and decoded_info:
                values = [
                    x for x in decoded_info if x
                ]

                if values:
                    return "\n".join(values)

    except Exception:
        pass

    return None


# ------------------------------------------------------------
# WECHAT QR - OPTIONAL
# ------------------------------------------------------------

def wechat_decode(img):
    """
    Dùng WeChatQRCode nếu opencv-contrib + model files có sẵn.

    Đặt 4 file model cạnh qr_decoder.py:
        detect.prototxt
        detect.caffemodel
        sr.prototxt
        sr.caffemodel

    Nếu chưa có model -> tự bỏ qua.
    """
    try:
        if not hasattr(cv2, "wechat_qrcode_WeChatQRCode"):
            return None

        base = Path(__file__).resolve().parent

        detector_model = base / "detect.prototxt"
        detector_weights = base / "detect.caffemodel"
        sr_model = base / "sr.prototxt"
        sr_weights = base / "sr.caffemodel"

        if not all([
            detector_model.exists(),
            detector_weights.exists(),
            sr_model.exists(),
            sr_weights.exists(),
        ]):
            return None

        detector = cv2.wechat_qrcode_WeChatQRCode(
            str(detector_model),
            str(detector_weights),
            str(sr_model),
            str(sr_weights),
        )

        results, _ = detector.detectAndDecode(img)

        if results:
            values = [x for x in results if x]
            if values:
                return "\n".join(values)

    except Exception:
        return None

    return None


# ------------------------------------------------------------
# ĐÁNH GIÁ CANDIDATE
# ------------------------------------------------------------

def prepare_candidates(img):
    candidates = []

    # 1. Ảnh gốc
    add_candidate(candidates, "full", img)

    # 2. QR theo bố cục CCCD
    crop_top_right_regions(img, candidates)

    # 3. Grid
    crop_grid_regions(img, candidates)

    # 4. Tìm contour
    find_qr_like_contours(img, candidates)

    # Giới hạn để không chạy vô hạn.
    return candidates[:MAX_CANDIDATES]


# ------------------------------------------------------------
# DEBUG
# ------------------------------------------------------------

def save_debug_images(candidates, image_path):
    if not SAVE_DEBUG:
        return

    debug_dir = image_path.parent / "debug_qr"
    debug_dir.mkdir(exist_ok=True)

    # Chỉ lưu tối đa 20 candidate đầu.
    for i, (name, img) in enumerate(candidates[:20]):
        try:
            safe_name = "".join(
                c if c.isalnum() or c in "-_" else "_"
                for c in name
            )

            output = debug_dir / f"{i:02d}_{safe_name}.jpg"

            cv2.imwrite(str(output), img)

        except Exception:
            pass


# ------------------------------------------------------------
# DECODER CHÍNH
# ------------------------------------------------------------

def decode_image(img, image_path):
    candidates = prepare_candidates(img)

    save_debug_images(candidates, image_path)

    print(f"CANDIDATES | {len(candidates)}")

    attempts = 0

    # --------------------------------------------------------
    # PHASE 1:
    # Ưu tiên vùng top-right vì ảnh CCCD mẫu có QR ở đó.
    # --------------------------------------------------------

    priority = []
    normal = []

    for item in candidates:
        if item[0].startswith("top_right"):
            priority.append(item)
        else:
            normal.append(item)

    ordered = priority + normal

    # --------------------------------------------------------
    # Mỗi candidate -> nhiều preprocessing -> nhiều decoder
    # --------------------------------------------------------

    for candidate_name, candidate in ordered:

        # Thử WeChat trên crop màu trước.
        data = wechat_decode(candidate)

        if data:
            return data, f"WeChat/{candidate_name}"

        # Thử ZXing trực tiếp.
        data = zxing_decode(candidate)

        if data:
            return data, f"ZXing/{candidate_name}"

        # Thử OpenCV trực tiếp.
        data = opencv_decode(candidate)

        if data:
            return data, f"OpenCV/{candidate_name}"

        # Preprocessing
        for prep_name, prepared in preprocessing_variants(candidate):

            attempts += 1

            # WeChat
            data = wechat_decode(prepared)
            if data:
                return data, (
                    f"WeChat/{candidate_name}/{prep_name}"
                )

            # ZXing
            data = zxing_decode(prepared)
            if data:
                return data, (
                    f"ZXing/{candidate_name}/{prep_name}"
                )

            # OpenCV
            data = opencv_decode(prepared)
            if data:
                return data, (
                    f"OpenCV/{candidate_name}/{prep_name}"
                )

            # Xoay các phiên bản quan trọng.
            if prep_name.endswith("_x3.0") or prep_name.endswith("_x4.0"):
                for angle, rotated in rotated_images(prepared):

                    if angle == "0":
                        continue

                    data = zxing_decode(rotated)
                    if data:
                        return data, (
                            f"ZXing/{candidate_name}/"
                            f"{prep_name}/rotate_{angle}"
                        )

                    data = opencv_decode(rotated)
                    if data:
                        return data, (
                            f"OpenCV/{candidate_name}/"
                            f"{prep_name}/rotate_{angle}"
                        )

    return None, None


# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------

def main():

    if len(sys.argv) != 2:
        print(
            'ERROR | Cú pháp: '
            'py qr_decoder.py "duong_dan_anh"'
        )
        sys.exit(2)

    image_path = Path(
        sys.argv[1].strip('"')
    )

    if not image_path.exists():
        print(
            f"ERROR | Không tìm thấy file: "
            f"{image_path}"
        )
        sys.exit(3)

    if not image_path.is_file():
        print(
            f"ERROR | Đường dẫn không phải file: "
            f"{image_path}"
        )
        sys.exit(4)

    img = load_image(image_path)

    if img is None:
        print(
            "ERROR | Không thể đọc ảnh hoặc "
            "ảnh không hợp lệ"
        )
        sys.exit(5)

    original_h, original_w = img.shape[:2]

    img = resize_for_processing(img)

    h, w = img.shape[:2]

    print("=" * 60)
    print("QR DECODER V3")
    print("=" * 60)
    print(f"IMAGE     | {image_path}")
    print(
        f"SIZE      | {original_w}x{original_h}"
    )
    print(
        f"PROCESS   | {w}x{h}"
    )

    if zxingcpp is not None:
        print("ZXING     | AVAILABLE")
    else:
        print(
            "ZXING     | NOT INSTALLED "
            "(py -m pip install zxing-cpp)"
        )

    if hasattr(cv2, "wechat_qrcode_WeChatQRCode"):
        print("WECHAT    | OPENCV CONTRIB AVAILABLE")
    else:
        print("WECHAT    | OPTIONAL / NOT AVAILABLE")

    print("STATUS    | SCANNING")
    print("=" * 60)

    data, method = decode_image(
        img,
        image_path
    )

    if data:
        print("=" * 60)
        print("STATUS    | SUCCESS")
        print(f"METHOD    | {method}")
        print("QR_DATA   | " + data)
        print("=" * 60)

        sys.exit(0)

    print("=" * 60)
    print("STATUS    | NOT_FOUND")
    print(
        "MESSAGE   | Không tìm thấy hoặc không "
        "giải mã được QR trong ảnh"
    )

    if SAVE_DEBUG:
        print(
            f"DEBUG     | {image_path.parent / 'debug_qr'}"
        )

    print("=" * 60)

    sys.exit(1)


if __name__ == "__main__":
    
    import sys
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    main()
