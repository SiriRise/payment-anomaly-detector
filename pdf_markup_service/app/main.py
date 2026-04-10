import shutil
import uuid
import json
import time
from contextlib import asynccontextmanager
from functools import partial
from pathlib import Path
from typing import Any
print("GHBDTN")
import anyio
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from ocr_extract import enrich_structures_with_ocr, get_easyocr_reader
from validate_2 import validate_qr_vs_tables, save_report_to_docx, check_table_formulas


SERVICE_ROOT = Path(__file__).resolve().parents[1]  # pdf_markup_service/
PROJECT_ROOT = SERVICE_ROOT.parents[0]  # testikOSR/

UPLOADS_DIR = SERVICE_ROOT / "uploads"
UPLOADS_DIR.mkdir(parents=True, exist_ok=True)

TEMPLATES_DIR = SERVICE_ROOT / "templates"
templates = Jinja2Templates(directory=str(TEMPLATES_DIR))
REPORTS_DIR = SERVICE_ROOT / "reports"
REPORTS_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_IMAGES_DIR = SERVICE_ROOT / "output_images"
OUTPUT_IMAGES_DIR.mkdir(parents=True, exist_ok=True)


def _debug_log(run_id: str, hypothesis_id: str, location: str, message: str, data: dict[str, Any]) -> None:
    payload = {
        "sessionId": "a34a8d",
        "runId": run_id,
        "hypothesisId": hypothesis_id,
        "location": location,
        "message": message,
        "data": data,
        "timestamp": int(time.time() * 1000),
    }
    with open("debug-a34a8d.log", "a", encoding="utf-8") as f:
        f.write(json.dumps(payload, ensure_ascii=False) + "\n")

def _ensure_project_root() -> None:
    import sys

    if str(PROJECT_ROOT) not in sys.path:
        sys.path.insert(0, str(PROJECT_ROOT))


def _warm_spacy() -> None:
    from ner_5_server import warm_spacy

    warm_spacy()


@asynccontextmanager
async def lifespan(_: FastAPI):
    # Прогрев EasyOCR и spaCy (чтобы первый запрос не ждал загрузки моделей)
    await anyio.to_thread.run_sync(partial(get_easyocr_reader, use_gpu=True))
    await anyio.to_thread.run_sync(_warm_spacy)
    yield


app = FastAPI(title="PDF Markup Service", lifespan=lifespan)

app.mount("/output-images", StaticFiles(directory=str(OUTPUT_IMAGES_DIR)), name="output-images")


def _is_pdf(filename: str, content_type: str | None) -> bool:
    if content_type and content_type.lower() == "application/pdf":
        return True
    return filename.lower().endswith(".pdf")


def _run_search(pdf_path: str) -> Any:
    _ensure_project_root()
    from search_structure import process_pdf  # type: ignore

    return process_pdf(pdf_path)


def _clear_output_images() -> None:
    for p in OUTPUT_IMAGES_DIR.glob("*"):
        if p.is_file():
            p.unlink(missing_ok=True)


def _run_ner(pages: list) -> dict:
    from ner_5_server import ner_from_pages

    return ner_from_pages(pages)


def _run_extract(img_folder: str, structures: Any, *, use_gpu: bool) -> Any:
    if not isinstance(structures, list):
        return structures
    return enrich_structures_with_ocr(
        structures, images_folder=img_folder, max_workers=4, use_gpu=use_gpu
    )


def _build_validation_view(
    ner_data: dict, discrepancies: list, formula_result: dict, image_version: str
) -> dict[str, Any]:
    run_id = f"view-{uuid.uuid4().hex[:8]}"
    pages: dict[int, dict[str, Any]] = {}
    headers = ner_data.get("headers", []) or []
    tables = ner_data.get("tables", []) or []
    qr_codes = ner_data.get("qr_codes", []) or []
    table_by_key = {(int(t.get("doc_idx") or 0), int(t.get("table_idx") or 0), int(t.get("page") or 0)): t for t in tables}
    qr_by_key = {(int(q.get("doc_idx") or 0), int(q.get("qr_idx") or 0)): q for q in qr_codes}
    qr_by_page_idx = {(int(q.get("page") or 0), int(q.get("qr_idx") or 0)): q for q in qr_codes}
    # region agent log
    _debug_log(
        run_id,
        "H1",
        "main.py:_build_validation_view:start",
        "validation view input counts",
        {"headers": len(headers), "tables": len(tables), "qr_codes": len(qr_codes), "discrepancies": len(discrepancies)},
    )
    # endregion

    def ensure_page(page_num: int) -> dict[str, Any]:
        if page_num not in pages:
            pages[page_num] = {
                "page": page_num,
                "image_url": f"/output-images/page_{page_num}_original.jpg?v={image_version}",
                "highlights": [],
                "messages": [],
            }
        return pages[page_num]

    # Показываем все страницы документа, даже если на странице нет ошибок.
    for p in headers:
        pn = int(p.get("page") or 0)
        if pn > 0:
            ensure_page(pn)
    for p in tables:
        pn = int(p.get("page") or 0)
        if pn > 0:
            ensure_page(pn)
    for p in qr_codes:
        pn = int(p.get("page") or 0)
        if pn > 0:
            ensure_page(pn)
    # region agent log
    tables_by_page: dict[int, list[int]] = {}
    for t in tables:
        pn = int(t.get("page") or 0)
        ti = int(t.get("table_idx") or 0)
        if pn > 0 and ti > 0:
            tables_by_page.setdefault(pn, []).append(ti)
    _debug_log(
        run_id,
        "H6",
        "main.py:_build_validation_view:tables_by_page",
        "table indexes available per page",
        {str(k): sorted(v) for k, v in tables_by_page.items()},
    )
    # endregion

    for d in discrepancies:
        page = int(d.get("page") or 0)
        if page <= 0:
            continue
        p = ensure_page(page)
        doc_idx = int(d.get("doc_idx") or 0)
        qr_idx = int(d.get("qr_idx") or 0)
        table_idx = int(d.get("matched_table_idx") or 0)
        qr = qr_by_key.get((doc_idx, qr_idx)) or qr_by_page_idx.get((page, qr_idx))
        # region agent log
        _debug_log(
            run_id,
            "H2",
            "main.py:_build_validation_view:discrepancy_lookup",
            "discrepancy key lookup",
            {
                "doc_idx": doc_idx,
                "page": page,
                "qr_idx": qr_idx,
                "table_idx": table_idx,
                "matched_table_idx_raw": d.get("matched_table_idx"),
                "qr_found": qr is not None,
                "table_found": (doc_idx, table_idx, page) in table_by_key,
            },
        )
        # endregion
        if qr:
            qc = qr.get("coords") or {}
            p["highlights"].append(
                {
                    "kind": "qr",
                    "label": f"QR #{qr_idx or '?'}",
                    "x": qc.get("x", qr.get("x", 0)),
                    "y": qc.get("y", qr.get("y", 0)),
                    "w": qc.get("w", qr.get("w", 0)),
                    "h": qc.get("h", qr.get("h", 0)),
                }
            )
        table = table_by_key.get((doc_idx, table_idx, page))
        if table:
            c = table.get("table_coords", {})
            p["highlights"].append(
                {
                    "kind": "table",
                    "label": f"Таблица #{table_idx or '?'}",
                    "x": c.get("x", 0),
                    "y": c.get("y", 0),
                    "w": c.get("w", 0),
                    "h": c.get("h", 0),
                }
            )
        else:
            # region agent log
            _debug_log(
                run_id,
                "H3",
                "main.py:_build_validation_view:missing_table",
                "table from discrepancy not found in ner_data.tables",
                {"doc_idx": doc_idx, "page": page, "table_idx": table_idx},
            )
            # endregion
        details = [f"{f}: табл={tv} / qr={qv}" for f, tv, qv in d.get("diffs", [])]
        if d.get("missing_from_table"):
            details.append(f"Нет в таблице (есть в QR): {', '.join(d.get('missing_from_table', []))}")
        if d.get("missing_requisites"):
            details.append(f"В таблице неполный набор реквизитов, отсутствуют: {', '.join(d.get('missing_requisites', []))}")
        if d.get("ls_error"):
            ls = d["ls_error"]
            details.append(f"ЛС: шапка={ls.get('header_val')} / qr={ls.get('qr_val')}")
        p["messages"].append(
            {
                "kind": "qr_mismatch",
                "title": f"{d.get('type', 'Несовпадение')} (QR #{qr_idx or '?'} ↔ Таблица #{table_idx or '?'})",
                "text": "; ".join(details) if details else d.get("details", ""),
            }
        )

    for err in formula_result.get("errors", []) or []:
        page = int(err.get("page") or 0)
        if page <= 0:
            continue
        p = ensure_page(page)
        # region agent log
        _debug_log(
            run_id,
            "H7",
            "main.py:_build_validation_view:formula_error_item",
            "formula error source row",
            {"page": page, "table_idx": err.get("table_idx"), "row": err.get("row")},
        )
        # endregion
        table = next((t for t in tables if t.get("page") == page and t.get("table_idx") == err.get("table_idx")), None)
        if table is None:
            # region agent log
            _debug_log(
                run_id,
                "H4",
                "main.py:_build_validation_view:formula_table_lookup",
                "formula error table lookup failed by page+table_idx",
                {"page": page, "table_idx": err.get("table_idx"), "row": err.get("row")},
            )
            # endregion
        if table:
            row_cells = [c for c in table.get("cells_data", []) if c.get("row_num") == err.get("row")]
            if row_cells:
                xs = [c["coords"]["x"] for c in row_cells]
                ys = [c["coords"]["y"] for c in row_cells]
                x2 = [c["coords"]["x"] + c["coords"]["w"] for c in row_cells]
                y2 = [c["coords"]["y"] + c["coords"]["h"] for c in row_cells]
                p["highlights"].append(
                    {"kind": "formula", "x": min(xs), "y": min(ys), "w": max(x2) - min(xs), "h": max(y2) - min(ys)}
                )
        msg = "; ".join(
            f"{x.get('formula')}: ожид. {x.get('expected'):.2f}, факт {x.get('actual'):.2f}" for x in err.get("errors", [])
        )
        p["messages"].append({"kind": "formula_error", "title": f"Таблица {err.get('table_idx')} строка {err.get('row')}", "text": msg})

    result_pages = [pages[k] for k in sorted(pages.keys())]
    # region agent log
    _debug_log(
        run_id,
        "H5",
        "main.py:_build_validation_view:end",
        "validation view output pages",
        {
            "pages_total": len(result_pages),
            "pages_with_messages": sum(1 for p in result_pages if p.get("messages")),
            "pages_with_highlights": sum(1 for p in result_pages if p.get("highlights")),
        },
    )
    # endregion
    return {"pages": result_pages}


def _run_pipeline(pdf_path: str, *, use_gpu: bool, report_filename: str) -> dict[str, Any]:
    _clear_output_images()
    structures = _run_search(pdf_path)
    # nfhod_tablic_3 пишет картинки относительно cwd в output_images/
    enriched = _run_extract(str(OUTPUT_IMAGES_DIR), structures, use_gpu=use_gpu)
    empty_ner: dict[str, Any] = {"headers": [], "tables": [], "qr_codes": []}


    if not isinstance(enriched, list):
        return {"ner": empty_ner}

    ner_data = _run_ner(enriched)

    # Передаём в том же формате, который ожидает валидатор

    discrepancies, stats = validate_qr_vs_tables({"ner": ner_data})
    # save_report_to_docx(discrepancies, stats)
    tables = ner_data.get("tables", [])  # таблицы уже обогащены cells_data
    formula_result = check_table_formulas(tables)

    # 3. Сохраняем единый отчёт
    report_path = save_report_to_docx(
        discrepancies, stats, formula_result, filename=str(REPORTS_DIR / report_filename)
    )
    print(f"Отчёт сохранён: {report_path}")

    return {
        "ner": ner_data,
        "validation": {
                          "qr_mismatches": stats,
                          "formula_errors": formula_result["stats"]["errors_count"],
                          "status": "ok" if (stats["mismatches"] == 0 and formula_result["stats"][
                              "errors_count"] == 0) else "mismatch_found"
                      },
        "validation_view": _build_validation_view(
            ner_data, discrepancies, formula_result, image_version=report_filename
        ),
        "report_url": f"/api/report/{report_filename}",
    }


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse(request, "index.html", {"request": request})


@app.post("/api/parse")
async def parse_pdf(
    file: UploadFile = File(...),
    use_gpu: bool = Form(False),
):
    if not file.filename:
        raise HTTPException(status_code=400, detail="Имя файла пустое")

    if not _is_pdf(file.filename, file.content_type):
        raise HTTPException(status_code=400, detail="Нужен PDF файл")

    job_id = str(uuid.uuid4())
    safe_name = Path(file.filename).name
    saved_path = UPLOADS_DIR / f"{job_id}_{safe_name}"

    try:
        with saved_path.open("wb") as f:
            shutil.copyfileobj(file.file, f)
    finally:
        await file.close()

    try:
        # process_pdf тяжелый и CPU-bound → в отдельный поток
        # В некоторых версиях anyio run_sync не принимает kwargs → используем partial
        report_filename = f"verification_{job_id}.docx"
        payload = await anyio.to_thread.run_sync(
            partial(_run_pipeline, str(saved_path), use_gpu=use_gpu, report_filename=report_filename)
        )
        return JSONResponse(content={"job_id": job_id, "file": safe_name, **payload})
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка обработки: {e}") from e


@app.get("/api/report/{report_name}")
async def download_report(report_name: str):
    safe = Path(report_name).name
    path = REPORTS_DIR / safe
    if not path.exists():
        raise HTTPException(status_code=404, detail="Отчет не найден")
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=safe,
    )

