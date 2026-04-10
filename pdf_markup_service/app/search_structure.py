import fitz
import cv2
import numpy as np
from pathlib import Path
from pyzbar.pyzbar import decode


OUTPUT_IMAGES_DIR = Path(__file__).resolve().parents[1] / "output_images"
OUTPUT_IMAGES_DIR.mkdir(parents=True, exist_ok=True)

KERNEL = np.ones((2, 2), np.uint8)

def has_document_header(img_cv, page_num, qr_codes=None):
    h, w = img_cv.shape[:2]
    page_area = float(max(1, h * w))
    qr_codes = qr_codes or []

    max_ratio = 0.0
    for qr in qr_codes:
        qw = float(max(0, qr.get("w", 0)))
        qh = float(max(0, qr.get("h", 0)))
        qr_area = qw * qh
        max_ratio = max(max_ratio, qr_area / page_area)

    has_header = max_ratio >= 0.005

    if has_header:
        print(f"   ✅ Новый документ по QR (max_ratio={max_ratio:.4f})")
    else:
        print(f"   ⏭️  Продолжение документа (max_ratio={max_ratio:.4f})")

    return has_header, 0
def find_qr_codes(img_cv, page_num):

    qr_codes = []

    gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)
    closed = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, KERNEL, iterations=2)
    restored = cv2.bitwise_not(closed)
    debug_path = OUTPUT_IMAGES_DIR / f"page_{page_num}_qr_restored.jpg"
    cv2.imwrite(str(debug_path), restored)

    decoded_objects = decode(restored)

    for obj in decoded_objects:
        x, y, w_rect, h_rect = obj.rect
        factor = 1
        qr_codes.append({
            'type': obj.type,
            'x': int(x / factor),
            'y': int(y / factor),
            'w': int(w_rect / factor),
            'h': int(h_rect / factor)
        })
        print(f"   🟩 НАЙДЕН: ")
    return qr_codes


def find_tables(img_cv, page_num):
    """Находит таблицы с адаптивными параметрами"""
    gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

    h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (30, 1))
    h_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, h_kernel, iterations=1)

    v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 30))
    v_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, v_kernel, iterations=1)

    intersections = cv2.bitwise_and(h_lines, v_lines)
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
    intersections_dilated = cv2.dilate(intersections, kernel, iterations=2)

    v_contours, _ = cv2.findContours(v_lines, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    v_lines_clean = np.zeros_like(v_lines)

    for cnt in v_contours:
        x, y, w, h = cv2.boundingRect(cnt)
        top_roi = intersections_dilated[y:y + 10, x:x + w]
        bottom_roi = intersections_dilated[y + h - 10:y + h, x:x + w]
        has_top = cv2.countNonZero(top_roi) > 0
        has_bottom = cv2.countNonZero(bottom_roi) > 0
        if has_top and has_bottom:
            cv2.drawContours(v_lines_clean, [cnt], -1, (255), -1)

    v_lines = v_lines_clean
    table_mask = cv2.add(h_lines, v_lines)

    contours, _ = cv2.findContours(table_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    tables = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        min_width = img_cv.shape[1] * 0.2
        min_height = img_cv.shape[0] * 0.01
        if w > min_width and h > min_height:
            tables.append((x, y, w, h))
            print(f"   ✅ Таблица: {w}x{h}")
        else:
            print(f"   ❌ Пропущено: {w}x{h}")

    tables = sorted(tables, key=lambda k: k[1])
    print(f"   📊 Найдено таблиц: {len(tables)}")
    return tables


def find_cells_in_table(table_img):
    gray = cv2.cvtColor(table_img, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)

    h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (15, 1))
    v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 15))
    h_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, h_kernel, iterations=1)
    v_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, v_kernel, iterations=1)
    grid = cv2.add(h_lines, v_lines)

    _, cells_mask = cv2.threshold(grid, 10, 255, cv2.THRESH_BINARY_INV)
    contours, _ = cv2.findContours(cells_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    cells = []
    table_h, table_w = table_img.shape[:2]

    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        if w < 30 or h < 20:
            continue
        if w > table_w * 0.95 and h > table_h * 0.95:
            continue

        margin = 2
        has_top = cv2.countNonZero(h_lines[y - margin:y + margin, x:x + w]) > 5 if y >= margin else False
        has_bottom = cv2.countNonZero(
            h_lines[y + h - margin:y + h + margin, x:x + w]) > 5 if y + h + margin <= table_h else False
        has_left = cv2.countNonZero(v_lines[y:y + h, x - margin:x + margin]) > 5 if x >= margin else False
        has_right = cv2.countNonZero(
            v_lines[y:y + h, x + w - margin:x + w + margin]) > 5 if x + w + margin <= table_w else False

        if (has_top and has_bottom) or (has_left and has_right):
            cells.append({'x': x, 'y': y, 'w': w, 'h': h})

    cells = group_cells_by_rows(cells, row_threshold=15)
    return cells


def group_cells_by_rows(cells, row_threshold=15):
    if not cells:
        return []
    cells_sorted = sorted(cells, key=lambda k: k['y'])
    rows = []
    current_row = [cells_sorted[0]]
    for i in range(1, len(cells_sorted)):
        prev_cell = current_row[-1]
        curr_cell = cells_sorted[i]
        if abs(curr_cell['y'] - prev_cell['y']) <= row_threshold:
            current_row.append(curr_cell)
        else:
            rows.append(current_row)
            current_row = [curr_cell]
    rows.append(current_row)
    result = []
    cell_idx = 1
    for row_num, row in enumerate(rows, 1):
        row_sorted = sorted(row, key=lambda k: k['x'])
        for col_num, cell in enumerate(row_sorted, 1):
            cell['row_num'] = row_num
            cell['col_num'] = col_num
            cell['cell_idx'] = cell_idx
            result.append(cell)
            cell_idx += 1
    return result


def process_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    zoom = 420 / 72
    mat = fitz.Matrix(zoom, zoom)

    all_results = []
    current_doc_idx = 0
    last_doc_id = None

    for page_num, page in enumerate(doc, 1):
        print(f"\n{'=' * 70}")
        print(f"📄 СТРАНИЦА {page_num}")
        print('=' * 70)

        pix = page.get_pixmap(matrix=mat)
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape((pix.height, pix.width, 3))
        img_cv = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)

        cv2.imwrite(str(OUTPUT_IMAGES_DIR / f"page_{page_num}_original.jpg"), img_cv)

        print("\n🔍 Поиск QR-кодов...")
        qr_codes = find_qr_codes(img_cv, page_num)
        print(f"   🟩 Найдено QR-кодов: {len(qr_codes)}")

        print("\n🔍 Поиск таблиц...")
        tables = find_tables(img_cv, page_num)

        print("\n🔍 Проверка заголовка документа...")
        has_header, _ = has_document_header(img_cv, page_num, qr_codes=qr_codes)

        if has_header:
            current_doc_idx += 1
            print(f"   🆕 Новый документ #{current_doc_idx}")
        else:
            print(f"   ➡️  Продолжение документа #{current_doc_idx}")

        header_y = 0
        content_y = 0

        if has_header and tables:
            content_y = tables[0][1]
            print(f"   📋 Шапка: y=0 до y={content_y} (верх первой таблицы)")
        elif has_header and not tables:
            content_y = int(img_cv.shape[0] * 0.4)  # Нет таблиц - шапка ~40%
            print(f"   📋 Шапка: y=0 до y={content_y} (таблиц не найдено)")
        else:
            content_y = 0  # Продолжение - шапки нет
            print(f"   ➡️  Продолжение документа (без шапки)")

        page_results = {
            'page': page_num,
            'doc_idx': current_doc_idx,
            'is_new_doc': has_header,
            'header_y': header_y,
            'content_y': content_y,
            'header': {},
            'tables': [],
            'qr_codes': qr_codes
        }

        debug_img = img_cv.copy()

        if has_header:
            cv2.line(debug_img, (0, content_y), (img_cv.shape[1], content_y), (0, 0, 255), 2)
            cv2.putText(debug_img, f"DOC #{current_doc_idx} HEADER", (10, content_y - 10),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 0, 255), 2)

        for table_idx, (x, y, w, h) in enumerate(tables, 1):
            print(f"\n{'─' * 60}")
            print(f"📊 ТАБЛИЦА {table_idx}")
            table_img = img_cv[y:y + h, x:x + w]
            cells = find_cells_in_table(table_img)
            print(f"   Найдено ячеек: {len(cells)}")

            table_result = {
                'table_idx': table_idx,
                'x': x, 'y': y, 'w': w, 'h': h,
                'rows': []
            }

            rows_map = {}
            for cell in cells:
                rows_map.setdefault(cell['row_num'], []).append(cell)

            for row_num in sorted(rows_map.keys()):
                row_cells = sorted(rows_map[row_num], key=lambda c: c['col_num'])
                row_result = {'row_num': row_num, 'cells': []}
                for cell in row_cells:
                    abs_x = x + cell['x']
                    abs_y = y + cell['y']
                    row_result['cells'].append({
                        'cell_idx': cell['cell_idx'],
                        'col_num': cell['col_num'],
                        'abs_x': abs_x,
                        'abs_y': abs_y,
                        'w': cell['w'],
                        'h': cell['h']
                    })
                    cv2.rectangle(table_img, (cell['x'], cell['y']),
                                  (cell['x'] + cell['w'], cell['y'] + cell['h']), (255, 0, 0), 1)
                table_result['rows'].append(row_result)

            cv2.imwrite(str(OUTPUT_IMAGES_DIR / f"page_{page_num}_table_{table_idx}_cells.jpg"), table_img)
            cv2.rectangle(debug_img, (x, y), (x + w, y + h), (0, 255, 0), 2)
            page_results['tables'].append(table_result)

        cv2.imwrite(str(OUTPUT_IMAGES_DIR / f"page_{page_num}_tables.jpg"), debug_img)
        all_results.append(page_results)


    doc.close()
    return all_results

