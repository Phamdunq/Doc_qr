import cv2
import os
import re
import shutil
import sys
import tempfile
import unicodedata
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from zipfile import BadZipFile

# ============================================================
# CHƯƠNG TRÌNH ĐỌC QR CCCD TỪ NHIỀU ẢNH
#
# Cấu trúc thư mục:
#   chuong_trinh_doc_qr/
#       doc_qr_cccd.py
#       anh_cccd/                 <- bỏ ảnh vào đây
#       anh_khong_doc_duoc_qr/    <- ảnh không đọc được sẽ được chuyển vào đây
#
# Kết quả:
#   ket_qua_qr_cccd.xlsx
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "anh_cccd"
FAILED_DIR = BASE_DIR / "anh_khong_doc_duoc_qr"
OUTPUT_XLSX = BASE_DIR / "ket_qua_qr_cccd.xlsx"
MAPPING_XLSX = BASE_DIR / "BangChuyendoiĐVHCmoi_cu_final.xlsx"

IMAGE_EXTENSIONS = {
    ".jpg", ".jpeg", ".png", ".bmp", ".webp", ".tif", ".tiff"
}


def normalize_admin_name(value) -> str:
    """Chuẩn hóa tên đơn vị để so khớp không phân biệt dấu/chức danh."""
    text = str(value or "").strip()
    text = re.sub(r"\s*\([^)]*\)", "", text)
    text = unicodedata.normalize("NFD", text.casefold().replace("đ", "d"))
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = re.sub(
        r"^(?:thanh pho|thi xa|thi tran|tinh|huyen|quan|xa|phuong|dac khu)\s+",
        "",
        text,
    )
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def display_admin_name(value) -> str:
    """Bỏ mã số trong ngoặc khỏi tên đơn vị dùng để hiển thị."""
    return re.sub(r"\s*\(\d+\)\s*$", "", str(value or "")).strip()


def load_address_mapping(mapping_path: Path):
    """Đọc ánh xạ xã cũ -> xã mới từ sheet không gộp ô."""
    workbook = load_workbook(mapping_path, read_only=True, data_only=True)
    try:
        sheet_name = "Tổng hợp_không merge "
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f"Không tìm thấy sheet '{sheet_name}'")

        worksheet = workbook[sheet_name]
        entries = []
        for row in worksheet.iter_rows(min_row=4, values_only=True):
            values = list(row)
            if len(values) < 8:
                continue

            new_province, new_commune, new_code = values[:3]
            old_commune, old_code, _, old_district, old_province = values[3:8]
            if not all((new_province, new_commune, new_code, old_commune)):
                continue

            entries.append({
                "old_commune": normalize_admin_name(old_commune),
                "old_code": str(old_code or "").strip(),
                "old_district": normalize_admin_name(old_district),
                "old_province": normalize_admin_name(old_province),
                "new_province": display_admin_name(new_province),
                "new_commune": display_admin_name(new_commune),
                "new_code": str(new_code).strip(),
            })
        return entries
    finally:
        workbook.close()


def convert_cccd_address(address: str, mapping_entries):
    """
    Đổi địa chỉ cũ khi địa chỉ xã/huyện/tỉnh khớp duy nhất trong bảng.
    Trả về địa chỉ mới, mã xã mới và trạng thái chuyển đổi.
    """
    parts = [part.strip() for part in str(address or "").split(",") if part.strip()]
    normalized_parts = [normalize_admin_name(part) for part in parts]
    commune_names = {entry["old_commune"] for entry in mapping_entries}
    matched_communes = [
        name for name in normalized_parts if name and name in commune_names
    ]

    if not matched_communes:
        return "", "", "Không nhận diện được xã cũ"

    old_commune = matched_communes[0]
    candidates = [
        entry for entry in mapping_entries
        if entry["old_commune"] == old_commune
    ]

    district_names = {entry["old_district"] for entry in candidates}
    province_names = {entry["old_province"] for entry in candidates}
    address_names = set(normalized_parts)

    matching_districts = district_names.intersection(address_names)
    if matching_districts:
        candidates = [
            entry for entry in candidates
            if entry["old_district"] in matching_districts
        ]

    matching_provinces = province_names.intersection(address_names)
    if matching_provinces:
        candidates = [
            entry for entry in candidates
            if entry["old_province"] in matching_provinces
        ]

    old_codes = {entry["old_code"] for entry in candidates if entry["old_code"]}
    if len(old_codes) != 1:
        return "", "", "Không xác định duy nhất được mã xã cũ"
    old_code = old_codes.pop()
    candidates = [
        entry for entry in candidates if entry["old_code"] == old_code
    ]

    destinations = {
        (entry["new_commune"], entry["new_code"], entry["new_province"])
        for entry in candidates
    }
    if not destinations:
        return "", "", "Không tìm thấy xã trong bảng quy đổi"
    if len(destinations) != 1:
        return "", "", "Xã cũ có nhiều xã mới; cần xác minh thêm"

    new_commune, new_code, new_province = destinations.pop()
    new_address = ", ".join((new_commune, new_province))
    return new_address, new_code, "Đã chuyển đổi"


def parse_cccd_qr(data: str):
    """
    QR CCCD thường có dạng:
    CCCD||Họ tên|Ngày sinh|Giới tính|Địa chỉ|Ngày cấp

    Trả về 6 trường chính.
    Nếu định dạng khác, vẫn giữ toàn bộ dữ liệu ở cột QR_Goc.
    """
    data = data.strip().replace("\x00", "")

    parts = data.split("|")

    # Một số QR CCCD có || sau số CCCD.
    if len(parts) >= 7 and parts[1] == "":
        cccd = parts[0]
        ho_ten = parts[2]
        ngay_sinh = parts[3]
        gioi_tinh = parts[4]
        dia_chi = parts[5]
        ngay_cap = parts[6]
        return cccd, ho_ten, ngay_sinh, gioi_tinh, dia_chi, ngay_cap

    # Trường hợp dữ liệu dùng đúng 6 phần.
    if len(parts) >= 6:
        return (
            parts[0], parts[1], parts[2],
            parts[3], parts[4], parts[5]
        )

    return "", "", "", "", "", ""


def rename_scanned_image(image_path: Path, ho_ten: str) -> Path:
    """Đổi tên ảnh theo họ tên và trạng thái mà không ghi đè file có sẵn."""
    return rename_image_with_marker(image_path, ho_ten, "Đã quét")


def rename_duplicate_image(image_path: Path, ho_ten: str) -> Path:
    """Đánh dấu ảnh có số CCCD đã tồn tại."""
    return rename_image_with_marker(image_path, ho_ten, "Trùng")


def rename_image_with_marker(
    image_path: Path,
    ho_ten: str,
    marker: str,
) -> Path:
    name = ho_ten.strip() or image_path.stem
    safe_name = "".join(
        "_" if ord(char) < 32 or char in '<>:"/\\|?*' else char
        for char in name
    )
    safe_name = " ".join(safe_name.split()).strip(" .")

    if not safe_name:
        raise ValueError("QR không có họ tên hợp lệ để đặt tên file")

    if safe_name.split(".", 1)[0].upper() in {
        "CON", "PRN", "AUX", "NUL",
        *(f"COM{i}" for i in range(1, 10)),
        *(f"LPT{i}" for i in range(1, 10)),
    }:
        safe_name = f"_{safe_name}"

    base_name = f"{safe_name} - {marker}"
    target = image_path.with_name(f"{base_name}{image_path.suffix}")
    counter = 2

    while target.exists() and target != image_path:
        target = image_path.with_name(
            f"{base_name} ({counter}){image_path.suffix}"
        )
        counter += 1

    if target != image_path:
        image_path.rename(target)

    return target


def is_processed_image(image_path: Path) -> bool:
    """Nhận diện ảnh đã có trạng thái để không quét lại."""
    return bool(re.search(
        r" - (?:đã quét|không quét được|trùng)(?: \(\d+\))?$",
        image_path.stem,
        flags=re.IGNORECASE,
    ))


def move_failed_image(image_path: Path, failed_dir: Path) -> Path:
    """Đánh dấu và chuyển ảnh lỗi khỏi thư mục đầu vào, không ghi đè file."""
    base_name = f"{image_path.stem} - Không quét được"
    destination = failed_dir / f"{base_name}{image_path.suffix}"
    counter = 2

    while destination.exists():
        destination = failed_dir / (
            f"{base_name} ({counter}){image_path.suffix}"
        )
        counter += 1

    return Path(shutil.move(str(image_path), str(destination)))


def decode_once(detector, image):
    """Thử đọc QR trực tiếp từ một ảnh."""
    try:
        data, _, _ = detector.detectAndDecode(image)
        if data:
            return data
    except Exception:
        pass
    return None


def decode_qr_robust(image):
    """
    Đọc QR qua nhiều phiên bản ảnh để tăng khả năng đọc:
    - ảnh gốc
    - ảnh xám
    - tăng tương phản
    - threshold OTSU
    - threshold adaptive
    - làm nét
    - phóng to / thu nhỏ
    - xoay 90/180/270 độ

    Không phụ thuộc vào kích thước ảnh.
    """
    detector = cv2.QRCodeDetector()

    h, w = image.shape[:2]

    # Giới hạn kích thước để tránh xử lý ảnh quá lớn.
    # Ảnh nhỏ sẽ được phóng lên; ảnh lớn sẽ được thu xuống.
    scales = [1.0]

    if min(h, w) < 1000:
        scales.extend([1.5, 2.0, 3.0])
    elif max(h, w) > 2500:
        scales.extend([0.75, 0.5])

    candidates = []

    for scale in scales:
        if scale == 1.0:
            img = image
        else:
            img = cv2.resize(
                image, None, fx=scale, fy=scale,
                interpolation=cv2.INTER_CUBIC if scale > 1 else cv2.INTER_AREA
            )

        candidates.append(img)

        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        candidates.append(gray)

        # Tăng tương phản cục bộ.
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        enhanced = clahe.apply(gray)
        candidates.append(enhanced)

        # OTSU.
        _, otsu = cv2.threshold(
            enhanced, 0, 255,
            cv2.THRESH_BINARY + cv2.THRESH_OTSU
        )
        candidates.append(otsu)

        # Adaptive threshold.
        adaptive = cv2.adaptiveThreshold(
            enhanced, 255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY,
            31, 7
        )
        candidates.append(adaptive)

        # Làm nét nhẹ.
        blur = cv2.GaussianBlur(gray, (0, 0), 1.2)
        sharp = cv2.addWeighted(gray, 1.7, blur, -0.7, 0)
        candidates.append(sharp)

    # Thử ảnh gốc trước.
    for candidate in candidates:
        data = decode_once(detector, candidate)
        if data:
            return data

    # Thử xoay ảnh trong trường hợp ảnh bị xoay.
    for angle in (90, 180, 270):
        if angle == 90:
            rotated = cv2.rotate(image, cv2.ROTATE_90_CLOCKWISE)
        elif angle == 180:
            rotated = cv2.rotate(image, cv2.ROTATE_180)
        else:
            rotated = cv2.rotate(image, cv2.ROTATE_90_COUNTERCLOCKWISE)

        data = decode_once(detector, rotated)
        if data:
            return data

        gray = cv2.cvtColor(rotated, cv2.COLOR_BGR2GRAY)
        data = decode_once(detector, gray)
        if data:
            return data

    return None


DATA_HEADERS = [
        "STT",
        "Ten anh",
        "So CCCD",
        "Ho va ten",
        "Ngay sinh",
        "Gioi tinh",
        "Dia chi",
        "Dia chi moi",
        "Ma xa moi",
        "Trang thai chuyen doi",
        "Ngay cap",
        "QR_Goc",
        "Trang thai"
]
FAILED_HEADERS = ["STT", "Ten anh", "Ly do", "Duong dan anh da copy"]


def load_existing_cccds(output_path: Path):
    """Đọc các số CCCD đã có để ngăn thêm bản ghi trùng."""
    if not output_path.exists():
        return set()

    workbook = load_workbook(output_path, read_only=True, data_only=True)
    try:
        if "Du lieu QR" not in workbook.sheetnames:
            raise ValueError("Excel hiện có thiếu sheet 'Du lieu QR'")

        worksheet = workbook["Du lieu QR"]
        headers = next(
            worksheet.iter_rows(min_row=1, max_row=1, values_only=True),
            (),
        )
        if tuple(headers) != tuple(DATA_HEADERS):
            raise ValueError(
                "Cấu trúc sheet 'Du lieu QR' không khớp với chương trình; "
                "không thể kiểm tra trùng an toàn"
            )

        cccd_column = DATA_HEADERS.index("So CCCD")
        return {
            str(row[cccd_column]).strip()
            for row in worksheet.iter_rows(min_row=2, values_only=True)
            if len(row) > cccd_column and row[cccd_column] not in (None, "")
        }
    finally:
        workbook.close()


def create_excel(rows, failed_files):
    """Thêm kết quả mới vào workbook hiện có, không thay các dòng cũ."""
    if OUTPUT_XLSX.exists():
        wb = load_workbook(OUTPUT_XLSX)
        if "Du lieu QR" not in wb.sheetnames or "Anh khong doc duoc" not in wb.sheetnames:
            wb.close()
            raise ValueError("Excel hiện có thiếu sheet kết quả cần thiết")
        ws = wb["Du lieu QR"]
        ws2 = wb["Anh khong doc duoc"]
        if tuple(cell.value for cell in ws[1]) != tuple(DATA_HEADERS):
            wb.close()
            raise ValueError("Cấu trúc sheet 'Du lieu QR' không khớp")
        if tuple(cell.value for cell in ws2[1]) != tuple(FAILED_HEADERS):
            wb.close()
            raise ValueError("Cấu trúc sheet ảnh lỗi không khớp")
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Du lieu QR"
        ws.append(DATA_HEADERS)
        ws2 = wb.create_sheet("Anh khong doc duoc")
        ws2.append(FAILED_HEADERS)

    for row in rows:
        ws.append([
            ws.max_row,
            row["Ten anh"],
            row["So CCCD"],
            row["Ho va ten"],
            row["Ngay sinh"],
            row["Gioi tinh"],
            row["Dia chi"],
            row["Dia chi moi"],
            row["Ma xa moi"],
            row["Trang thai chuyen doi"],
            row["Ngay cap"],
            row["QR_Goc"],
            row["Trang thai"]
        ])

    for item in failed_files:
        ws2.append([
            ws2.max_row,
            item["Ten anh"],
            item["Ly do"],
            item["Duong dan"]
        ])

    # Định dạng.
    for sheet in (ws, ws2):
        for cell in sheet[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="D9EAF7")
            cell.alignment = Alignment(horizontal="center", vertical="center")

        sheet.freeze_panes = "A2"

        for column_cells in sheet.columns:
            max_len = max(
                (len(str(cell.value)) for cell in column_cells if cell.value is not None),
                default=0,
            )
            column_letter = column_cells[0].column_letter
            current_width = sheet.column_dimensions[column_letter].width or 0
            sheet.column_dimensions[column_letter].width = min(
                max(current_width, max(max_len + 2, 12)),
                70,
            )

    # Cột địa chỉ / QR gốc cho rộng hơn.
    ws.column_dimensions["G"].width = 45
    ws.column_dimensions["H"].width = 55
    ws.column_dimensions["L"].width = 70
    ws2.column_dimensions["D"].width = 45

    temporary_path = None
    try:
        with tempfile.NamedTemporaryFile(
            dir=OUTPUT_XLSX.parent,
            prefix=f".{OUTPUT_XLSX.stem}.",
            suffix=".tmp.xlsx",
            delete=False,
        ) as temporary_file:
            temporary_path = Path(temporary_file.name)

        wb.save(temporary_path)
        os.replace(temporary_path, OUTPUT_XLSX)
        temporary_path = None
    finally:
        wb.close()
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)


def main():
    try:
        mapping_entries = load_address_mapping(MAPPING_XLSX)
    except (OSError, ValueError, BadZipFile) as exc:
        print(f"KHONG DOC DUOC BANG QUY DOI: {exc}")
        input("Nhan Enter de thoat...")
        return

    if not mapping_entries:
        print(f"BANG QUY DOI KHONG CO DU LIEU: {MAPPING_XLSX}")
        input("Nhan Enter de thoat...")
        return

    try:
        existing_cccds = load_existing_cccds(OUTPUT_XLSX)
    except (OSError, ValueError, BadZipFile) as exc:
        print(f"KHONG THE DOC KET QUA EXCEL HIEN CO: {exc}")
        input("Nhan Enter de thoat...")
        return

    INPUT_DIR.mkdir(exist_ok=True)
    FAILED_DIR.mkdir(exist_ok=True)

    all_image_files = sorted(
        p for p in INPUT_DIR.iterdir()
        if p.is_file() and p.suffix.lower() in IMAGE_EXTENSIONS
    )
    image_files = [
        path for path in all_image_files if not is_processed_image(path)
    ]
    skipped_count = len(all_image_files) - len(image_files)

    if skipped_count:
        print(f"Bo qua anh da duoc danh dau: {skipped_count}")

    if not image_files:
        print("=" * 60)
        if skipped_count:
            print("KHONG CO ANH MOI CAN QUET")
        else:
            print("KHONG TIM THAY ANH")
            print(f"Hay bo anh vao thu muc: {INPUT_DIR}")
        print("=" * 60)
        input("Nhan Enter de thoat...")
        return

    rows = []
    failed_files = []
    duplicate_count = 0
    pending_file_moves = []

    print("=" * 60)
    print("        CHUONG TRINH DOC QR CCCD")
    print("=" * 60)
    print(f"So anh can doc: {len(image_files)}")
    print()

    for index, image_path in enumerate(image_files, start=1):
        print(f"[{index}/{len(image_files)}] Dang doc: {image_path.name}")

        image = cv2.imread(str(image_path))

        if image is None:
            print("   -> KHONG MO DUOC ANH")
            destination = move_failed_image(image_path, FAILED_DIR)
            pending_file_moves.append((destination, image_path))
            failed_files.append({
                "Ten anh": destination.name,
                "Ly do": "Khong mo duoc anh",
                "Duong dan": str(destination)
            })
            continue

        data = decode_qr_robust(image)

        if data:
            cccd, ho_ten, ngay_sinh, gioi_tinh, dia_chi, ngay_cap = parse_cccd_qr(data)
            cccd_key = str(cccd or "").strip()

            if cccd_key and cccd_key in existing_cccds:
                print(f"   -> TRUNG CCCD: {cccd_key}; khong them vao Excel")
                try:
                    duplicate_path = rename_duplicate_image(image_path, ho_ten)
                    print(f"   -> Da danh dau anh trung: {duplicate_path.name}")
                except (OSError, ValueError) as exc:
                    print(f"   -> KHONG DANH DAU DUOC ANH TRUNG: {exc}")
                duplicate_count += 1
                continue

            if cccd_key:
                existing_cccds.add(cccd_key)

            dia_chi_moi, ma_xa_moi, trang_thai_chuyen_doi = (
                convert_cccd_address(dia_chi, mapping_entries)
            )
            output_image_path = image_path

            try:
                output_image_path = rename_scanned_image(image_path, ho_ten)
                if output_image_path != image_path:
                    pending_file_moves.append((output_image_path, image_path))
            except (OSError, ValueError) as exc:
                print(f"   -> KHONG DANH DAU DUOC ANH DA QUET: {exc}")
            if not ho_ten:
                print(
                    "   -> QR khong co truong ho ten hop le; "
                    "giu ten goc va them dau da quet"
                )

            rows.append({
                "Ten anh": output_image_path.name,
                "So CCCD": cccd,
                "Ho va ten": ho_ten,
                "Ngay sinh": ngay_sinh,
                "Gioi tinh": gioi_tinh,
                "Dia chi": dia_chi,
                "Dia chi moi": dia_chi_moi,
                "Ma xa moi": ma_xa_moi,
                "Trang thai chuyen doi": trang_thai_chuyen_doi,
                "Ngay cap": ngay_cap,
                "QR_Goc": data,
                "Trang thai": "Doc thanh cong"
            })

            print(
                f"   -> OK: {cccd} | {ho_ten} "
                f"| Anh: {output_image_path.name}"
            )
            print(
                f"   -> DIA CHI MOI: "
                f"{dia_chi_moi or trang_thai_chuyen_doi} "
                f"| MA XA MOI: {ma_xa_moi or 'CHUA XAC DINH'}"
            )
        else:
            print("   -> KHONG TIM THAY / KHONG DOC DUOC QR")

            destination = move_failed_image(image_path, FAILED_DIR)
            pending_file_moves.append((destination, image_path))

            failed_files.append({
                "Ten anh": destination.name,
                "Ly do": "Khong co QR hoac khong giai ma duoc QR",
                "Duong dan": str(destination)
            })

    try:
        create_excel(rows, failed_files)
    except (OSError, ValueError, BadZipFile) as exc:
        print(f"KHONG THE CAP NHAT FILE EXCEL: {exc}")
        print("Hay dong file Excel neu dang mo, sau do chay lai.")
        for moved_path, original_path in reversed(pending_file_moves):
            try:
                if moved_path.exists() and not original_path.exists():
                    moved_path.rename(original_path)
            except OSError as rollback_error:
                print(
                    f"CANH BAO | Khong the khoi phuc {moved_path} "
                    f"ve {original_path}: {rollback_error}"
                )
        input("Nhan Enter de thoat...")
        return

    print()
    print("=" * 60)
    print("HOAN TAT")
    print(f"Them moi       : {len(rows)} anh")
    print(f"Anh trung      : {duplicate_count} anh")
    print(f"Anh loi        : {len(failed_files)} anh")
    print(f"File Excel     : {OUTPUT_XLSX}")
    print(f"Anh loi        : {FAILED_DIR}")
    print("=" * 60)
    input("Nhan Enter de thoat...")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    main()
