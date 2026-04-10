from __future__ import annotations
import os
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Any
import cv2
import numpy as np
import easyocr
from pyzbar.pyzbar import decode as pyzbar_decode

_EASYOCR_READERS: dict[tuple[bool, str, tuple[str, ...]], Any] = {}
KERNEL = np.ones((2, 2), np.uint8)
OUTPUT_IMAGES_DIR = Path(__file__).resolve().parents[1] / "output_images_ich"
allow_list = '0123456789!#%()*+,-./:;<=>?№_ ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzАБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯабвгдеёжзийклмнопрстуфхцчшщъыьэюя'

def get_easyocr_reader(*, use_gpu: bool | None = None) -> Any:
    model_dir = os.getenv("EASYOCR_MODEL_DIR", "./model")
    languages = os.getenv("EASYOCR_LANGS", "ru,en").split(",")
    languages = [x.strip() for x in languages if x.strip()]
    langs_key = tuple(languages)



    if use_gpu is None:
        use_gpu = os.getenv("EASYOCR_GPU", "0").strip() in {"1", "true", "True", "yes", "YES"}

    key = (bool(use_gpu), model_dir, langs_key)
    if key in _EASYOCR_READERS:
        return _EASYOCR_READERS[key]

    reader = easyocr.Reader(languages, gpu=bool(use_gpu), model_storage_directory=model_dir) #user_network_directory="user_network",recog_network='custom_example',
    _EASYOCR_READERS[key] = reader
    return reader


def extract_header_text(img_cv: np.ndarray, doc_info: dict[str, Any], *, reader: Any) -> dict[str, Any]:
    if not doc_info.get("is_new_doc", False):
        return {}

    header_y = int(doc_info.get("header_y", 0) or 0)
    content_y = int(doc_info.get("content_y", 0) or 0)
    if content_y <= header_y:
        return {}

    header_crop = img_cv[header_y:content_y, :]
    if header_crop.size == 0:
        return {}
    gray = cv2.cvtColor(header_crop, cv2.COLOR_BGR2GRAY)
    result = reader.readtext(gray, detail=0)# blocklist='`{|}~ €₽'
    text = " ".join(result) if result else ""
    print(text)
    return {"full_text": text.strip()}


def decode_qr_code(img_cv: np.ndarray, qr_info: dict[str, Any]) -> tuple[str | None, str | None]:
    if pyzbar_decode is None:
        return None, None

    x = int(qr_info.get("x", 0) or 0)
    y = int(qr_info.get("y", 0) or 0)
    w = int(qr_info.get("w", 0) or 0)
    h = int(qr_info.get("h", 0) or 0)
    padding = 0
    x1 = max(0, x - padding)
    y1 = max(0, y - padding)
    x2 = min(img_cv.shape[1], x + w + padding)
    y2 = min(img_cv.shape[0], y + h + padding)
    qr_crop = img_cv[y1:y2, x1:x2]
    if qr_crop.size == 0:
        return None, None

    gray = cv2.cvtColor(qr_crop, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)
    closed = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, KERNEL, iterations=2)
    restored = cv2.bitwise_not(closed)
    decoded = pyzbar_decode(restored)
    if not decoded:
        return None, None
    try:
        data = decoded[0].data.decode("utf-8", errors="ignore")
        return data, getattr(decoded[0], "type", None)
    except Exception:
        return None, None


def extract_cell_text(img_cv: np.ndarray, abs_x: int, abs_y: int, w: int, h: int, *, reader: Any) -> str:
    y1 = max(0, abs_y)
    y2 = min(img_cv.shape[0], abs_y + h)
    x1 = max(0, abs_x)
    x2 = min(img_cv.shape[1], abs_x + w)
    if y1 >= y2 or x1 >= x2:
        return ""
    cell_img = img_cv[y1:y2, x1:x2]
    if cell_img.size == 0:
        return ""
    gray = cv2.cvtColor(cell_img, cv2.COLOR_BGR2GRAY)

    #cv2.imwrite(str(OUTPUT_IMAGES_DIR / f"page_{abs_y}_table_{abs_x}_cells.jpg"),gray)

    result = reader.readtext(gray, detail=0,min_size=1)
    text = " ".join(result) if result else ""
    return text.strip()


def enrich_structures_with_ocr(structures: list[dict[str, Any]], images_folder: str | Path = "output_images",
                               max_workers: int = 4, use_gpu: bool = True, ) -> list[dict[str, Any]]:
    images_folder = Path(images_folder)
    all_results: list[dict[str, Any]] = []
    doc_headers: dict[int, dict[str, Any]] = {}

    reader = get_easyocr_reader(use_gpu=use_gpu)

    for doc_info in structures:
        page_num = int(doc_info.get("page", 0) or 0)
        doc_idx = int(doc_info.get("doc_idx", 1) or 1)
        is_new_doc = bool(doc_info.get("is_new_doc", False))

        img_path = images_folder / f"page_{page_num}_original.jpg"
        if not img_path.exists():
            continue

        img_cv = cv2.imread(str(img_path))
        if img_cv is None:
            continue

        if is_new_doc:
            header_data = extract_header_text(img_cv, doc_info, reader=reader)
            doc_info["header"] = header_data
            doc_headers[doc_idx] = header_data
        else:
            doc_info["header"] = doc_headers.get(doc_idx, {})

        qr_codes = doc_info.get("qr_codes", []) or []
        if isinstance(qr_codes, list) and qr_codes:
            new_qr_codes = []
            for i, qr_info in enumerate(qr_codes):
                if not isinstance(qr_info, dict):
                    continue
                data, qr_type = decode_qr_code(img_cv, qr_info)
                new_qr_codes.append(
                    {
                        "qr_idx": i + 1,
                        "x": int(qr_info.get("x", 0) or 0),
                        "y": int(qr_info.get("y", 0) or 0),
                        "w": int(qr_info.get("w", 0) or 0),
                        "h": int(qr_info.get("h", 0) or 0),
                        "type": qr_type,
                        "data": data,
                    }
                )
            doc_info["qr_codes"] = new_qr_codes

        tables = doc_info.get("tables", []) or []
        if isinstance(tables, list) and tables:
            for t_idx, table in enumerate(tables):
                if not isinstance(table, dict):
                    continue

                rows = table.get("rows", []) or []
                if not isinstance(rows, list) or not rows:
                    continue

                cell_ptrs: list[tuple[int, int]] = []
                cell_args: list[tuple[int, int, int, int]] = []

                for r_idx, row in enumerate(rows):
                    if not isinstance(row, dict):
                        continue
                    row_cells = row.get("cells", []) or []
                    if not isinstance(row_cells, list) or not row_cells:
                        continue
                    for c_idx, cell in enumerate(row_cells):
                        if not isinstance(cell, dict):
                            continue
                        cell_ptrs.append((r_idx, c_idx))
                        cell_args.append(
                            (
                                int(cell.get("abs_x", 0) or 0),
                                int(cell.get("abs_y", 0) or 0),
                                int(cell.get("w", 0) or 0),
                                int(cell.get("h", 0) or 0),
                            )
                        )

                def _do_one(args: tuple[int, int, int, int]) -> str:
                    ax, ay, cw, ch = args
                    return extract_cell_text(img_cv, ax, ay, cw, ch, reader=reader)

                with ThreadPoolExecutor(max_workers=max_workers) as ex:
                    texts = list(ex.map(_do_one, cell_args))

                new_table = dict(table)
                new_rows = list(rows)

                for i, (r_idx, c_idx) in enumerate(cell_ptrs):
                    text = texts[i] if i < len(texts) else ""
                    row = new_rows[r_idx]
                    if not isinstance(row, dict):
                        continue
                    row_cells = row.get("cells", [])
                    if not isinstance(row_cells, list) or c_idx >= len(row_cells):
                        continue
                    cell = row_cells[c_idx]
                    if not isinstance(cell, dict):
                        continue
                    new_cell = dict(cell)
                    new_cell["text"] = text
                    row_cells = list(row_cells)
                    row_cells[c_idx] = new_cell
                    new_row = dict(row)
                    new_row["cells"] = row_cells
                    new_rows[r_idx] = new_row

                new_table["rows"] = new_rows
                doc_info["tables"][t_idx] = new_table

        all_results.append(doc_info)

    return all_results
