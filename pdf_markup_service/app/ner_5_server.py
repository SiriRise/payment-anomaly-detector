from __future__ import annotations
import re
from functools import lru_cache
from typing import Any
import spacy
from spacy.matcher import Matcher

try:
    nlp = spacy.load("ru_core_news_sm")
except Exception as e:
    raise RuntimeError(
        "Модель ru_core_news_sm не найдена. Установи: python -m spacy download ru_core_news_sm"
    ) from e


def _digits_only(s: str) -> str:
    return "".join(c for c in s if c.isdigit())

def parse_qr_gost(text: str | None) -> dict:
    if not text or "ST00012" not in text:
        return {}

    result: dict = {}
    parts = text.split("|")

    data_map: dict[str, str] = {}
    for part in parts:
        if "=" in part:
            key, val = part.split("=", 1)
            data_map[key] = val

    if "Name" in data_map:
        result["org_name"] = data_map["Name"]
    if "PayeeINN" in data_map:
        inn = _digits_only(data_map["PayeeINN"])
        result["inn"] = inn
    if "BIC" in data_map:
        bik = _digits_only(data_map["BIC"])
        result["bik"] = bik
    if "KPP" in data_map:
        kpp = _digits_only(data_map["KPP"])
        result["kpp"] = kpp
    if "PersonalAcc" in data_map:
        acc = _digits_only(data_map["PersonalAcc"])
        result["rs_account"] = acc
    if "Sum" in data_map:
        result["sum"] = data_map["Sum"]
    if "PaymPeriod" in data_map:
        result["period"] = data_map["PaymPeriod"]
    if "Purpose" in data_map:
        result["purpose"] = data_map["Purpose"]
    if "LastName" in data_map:
        result["lastName"] = data_map["LastName"]
    if "FirstName" in data_map:
        result["firstName"] = data_map["FirstName"]
    if "MiddleName" in data_map:
        result["middleName"] = data_map["MiddleName"]
    if "PayerAddress" in data_map:
        result["payerAddress"] = data_map["PayerAddress"]
    if "CorrespAcc" in data_map:
        result["correspAcc"] = data_map["CorrespAcc"]
    result["source_type"] = "QR_CODE"
    return result

def create_requisites_matcher() -> Matcher:
    matcher = Matcher(nlp.vocab)

    inn_variants = ["инн", "inn", "ihh", "iнн", "1нн"]
    matcher.add(
        "INN_TAG",
        [
            [
                {"LOWER": {"IN": inn_variants}},
                {"IS_PUNCT": True, "OP": "*"},
                {"IS_SPACE": True, "OP": "*"},
                {"LIKE_NUM": True, "OP": "?"},
            ]
        ],
    )

    kpp_variants = ["кпп", "kpp"]
    matcher.add(
        "KPP_TAG",
        [
            [
                {"LOWER": {"IN": kpp_variants}},
                {"IS_PUNCT": True, "OP": "*"},
                {"IS_SPACE": True, "OP": "*"},
                {"LIKE_NUM": True, "OP": "?"},
            ]
        ],
    )

    bik_variants = ["бик", "bik"]
    matcher.add(
        "BIK_TAG",
        [
            [
                {"LOWER": {"IN": bik_variants}},
                {"IS_PUNCT": True, "OP": "*"},
                {"IS_SPACE": True, "OP": "*"},
                {"LIKE_NUM": True, "OP": "?"},
            ]
        ],
    )

    rs_variants = [
        "р/с","р с","p/c","plc","p c","plc","pc","рc:","рc","c","c:","pic","pic:","pIс:","pic","рic","piс",
        "рiс","р/сч","с:","p с","p с:","с","р/сч","p/cч","р/cч","p/сч","р/сч",

    ]

    matcher.add(
        "RS_TAG",
        [
            [
                {"LOWER": {"IN": rs_variants}},
                {"IS_PUNCT": True, "OP": "*"},
                {"IS_SPACE": True, "OP": "*"},
                {"LIKE_NUM": True, "OP": "?"},
            ],
            [
                {"LOWER": {"IN": ["р","p"]}},
                {"IS_PUNCT": True, "OP": "*"},
                {"LOWER": {"IN": ["сч","c:","c","ic:","ic"]}},
                {"IS_PUNCT": True, "OP": "*"},
                {"IS_SPACE": True, "OP": "*"},
                {"LIKE_NUM": True, "OP": "?"},

            ],
        ],
    )

    _num = {"TEXT": {"REGEX": r"^\d{4,18}[A-Za-zА-Яа-яёЁ]?$"}, "OP": "?"}

    # Разделитель между меткой и номером (двоеточие, пробелы или их отсутствие)
    _mid = [{"IS_PUNCT": True, "OP": "*"}, {"IS_SPACE": True, "OP": "*"}]

    matcher.add("LS_TAG", [
        # 1. Слитные/однословные варианты
        [{"LOWER": {"IN": ["л/с:","л/с", "лс", "l/s", "ls", "лич.сч", "личсч"]}}, *_mid, _num],

        # 2. Разбитые варианты (Л / С, лич . сч)
        [
            {"LOWER": {"IN": ["л", "l", "лич"]}},
            {"IS_PUNCT": True, "OP": "*"},
            {"LOWER": {"IN": ["с","с:", "s", "сч"]}},
            *_mid,
            _num
        ]
    ])

    area_variants = ["пом"]
    matcher.add(
        "AREA_TAG",
        [
            [
                {"LOWER": {"IN": area_variants}},
                {"IS_PUNCT": True, "OP": "*"},
                {"IS_SPACE": True, "OP": "*"},
                {"LIKE_NUM": True, "OP": "?"},
            ]
        ],
    )
    sum_owner_variants_1 = ["кол-во"]
    sum_owner_variants_2 = ["собственников"]
    matcher.add(
        "SUM_OWNER_TAG",
        [
            [
                {"LOWER": {"IN": sum_owner_variants_1}, "OP":"*"},
                {"LOWER": {"IN": sum_owner_variants_2}},
                {"IS_PUNCT": True, "OP": "*"},
                {"IS_SPACE": True, "OP": "*"},
                {"LIKE_NUM": True, "OP": "?"},
            ]
        ],
    )
    address_triggers = [
        "адрес", "адрес помещения", "адрес объекта", "местонахождение", "расположен","помещения"
    ]

    matcher.add(
        "ADDRESS_TAG",
        [[
                {"LOWER": {"IN": address_triggers}},
                {"ORTH": ":", "OP": "?"},
                {"IS_SPACE": True, "OP": "*"},
                # Захватываем ВСЁ до конца строки
                #{"TEXT": {"REGEX": r".*"}, "OP": "+"}]
                {"TEXT": {"NOT_IN": ["Оплатить", "Внимание!", "Площадь", "Кол-во", "Количество",
                "Временно", "Тип", "ФИО", "Плательщик", "Собственник", "Наниматель",
                "собственника/нанимателя", "нанимателя","Плата"]}, "OP": "+"}]
        ]
    )

    owner_variants = [
        "наниматель", "нанимателя", "нанимателю", "нанимателем", "наниматели", "нанимателей","собственника/нанимателя", "нанимателя/собственника"  # варианты со слешем
    ]
    matcher.add(
        "OWNER_TAG",
        [[
                {"LOWER": {"IN": owner_variants}},
                {"ORTH": ":", "OP": "?"},
                {"IS_SPACE": True, "OP": "*"},
                {"TEXT": {"NOT_IN": ["Оплатить", "Внимание", "Площадь","Внимание!", "Кол-во"]}, "OP": "+"}]
        ]


    )


    return matcher

def extract_entities_from_text(text: str, matcher: Matcher) -> dict:
    if not text or len(text.strip()) < 5:
        return {}

    doc = nlp(text)
    matches = matcher(doc)
    result: dict = {}

    for match_id, start, end in matches:
        rule_name = matcher.vocab.strings[match_id]
        span = doc[start:end]
        if rule_name == "LS_TAG":
            num_match = re.search(r'(\d{4,18}\s?[A-Za-zА-Яа-яёЁ]?)', span.text)
            if num_match:
                account = num_match.group(1).replace(" ", "")  # склеиваем, если spaCy разбил "123 R"
                result["ls_account"] = account
            continue

        numbers = "".join(t.text for t in span if t.like_num)
        clean_num = re.sub(r"\D", "", numbers)


        if rule_name == "AREA_TAG":
            result["area"] = numbers
        if rule_name == "SUM_OWNER_TAG":
            result["sum_owner"] = clean_num
        if rule_name == "OWNER_TAG":
            owner = re.sub(r'^[^:]*:\s*', '', span.text, flags=re.IGNORECASE).strip()
            result["owner"] = owner

        if rule_name == "ADDRESS_TAG":
            address = re.sub(r'^[^:]*:\s*', '', span.text, flags=re.IGNORECASE).strip()
            address_clean = address.rstrip(';,. ')
            result["address"] = address_clean

        if rule_name == "INN_TAG" and len(clean_num) in (10, 12):
            if "inn" not in result:
                result["inn"] = clean_num
                name_span = doc[:start]
                raw_name = name_span.text
                clean_name = re.sub(r"[\s\-:,]+$", "", raw_name).strip()
                if len(clean_name) > 3 and not re.match(r"^[\d\s]+$", clean_name):
                    result["org_name"] = clean_name

        elif rule_name == "KPP_TAG" and len(clean_num) == 9:
            if "kpp" not in result:
                result["kpp"] = clean_num

        elif rule_name == "BIK_TAG" and len(clean_num) == 9 and clean_num.startswith("04"):
            if "bik" not in result:
                result["bik"] = clean_num

        elif rule_name == "RS_TAG" and len(clean_num) == 20:
            if "rs_account" not in result:
                result["rs_account"] = clean_num

    return result


@lru_cache(maxsize=1)
def _matcher() -> Matcher:
    return create_requisites_matcher()


def warm_spacy() -> None:
    _matcher()
    nlp("Прогрев модели.")


def iter_cells(table: dict) -> Any:
    if "cells" in table and isinstance(table.get("cells"), list):
        for c in table["cells"]:
            yield c, (c.get("row_num", 0) if isinstance(c, dict) else 0)
        return
    for row in table.get("rows") or []:
        if not isinstance(row, dict):
            continue
        rn = int(row.get("row_num", 0) or 0)
        for c in row.get("cells") or []:
            yield c, rn


def _table_record(table: dict, page: int, doc_idx: int, matcher: Matcher) -> dict:
    texts: list[str] = []
    cells_data: list[dict] = []
    for cell, rrow in iter_cells(table):
        ct = cell.get("text", "") if isinstance(cell, dict) else ""
        if ct:
            texts.append(ct)
        rn = cell.get("row_num") if isinstance(cell, dict) else None
        if rn is None or rn == 0:
            rn = rrow
        cells_data.append(
            {
                "cell_idx": cell.get("cell_idx", 0) if isinstance(cell, dict) else 0,
                "row_num": rn,
                "col_num": cell.get("col_num", 0) if isinstance(cell, dict) else 0,
                "text": ct,
                "coords": {
                    "x": cell.get("abs_x", 0) if isinstance(cell, dict) else 0,
                    "y": cell.get("abs_y", 0) if isinstance(cell, dict) else 0,
                    "w": cell.get("w", 0) if isinstance(cell, dict) else 0,
                    "h": cell.get("h", 0) if isinstance(cell, dict) else 0,
                },
            }
        )
    blob = " ".join(texts)
    table_entities = extract_entities_from_text(blob, matcher) if blob.strip() else {}
    return {
        "page": page,
        "doc_idx": doc_idx,
        "table_idx": table.get("table_idx", 0),
        "cells_data": cells_data,
        "extracted_entities": table_entities,
        "table_coords": {k: table.get(k, 0) for k in ("x", "y", "w", "h")},
    }


def ner_from_pages(pages: list[dict]) -> dict[str, Any]:
    matcher = _matcher()
    hdrs: list[dict] = []
    tbls: list[dict] = []
    qrs: list[dict] = []

    for item in pages:
        p = item["page"]
        di = item.get("doc_idx", 1)
        h = (item.get("header") or {}).get("full_text") or ""
        if str(h).strip():
            ent = extract_entities_from_text(h, matcher)
            if ent:
                hdrs.append(
                    {
                        "page": p,
                        "doc_idx": di,
                        "is_new_doc": item.get("is_new_doc", False),
                        "extracted_entities": {
                            **ent,

                        },
                    }
                )

        for qi, qr in enumerate(item.get("qr_codes") or []):
            q = qr.get("data")
            if not q:
                continue
            parsed = parse_qr_gost(q)
            if not parsed:
                continue
            parsed["page"] = p
            parsed["doc_idx"] = di
            parsed["qr_idx"] = qi + 1
            parsed["coords"] = {
                "x": qr.get("x"),
                "y": qr.get("y"),
                "w": qr.get("w"),
                "h": qr.get("h"),
            }
            qrs.append(parsed)

        for t in item.get("tables") or []:
            tbls.append(_table_record(t, p, di, matcher))

    return {"headers": hdrs, "tables": tbls, "qr_codes": qrs}
