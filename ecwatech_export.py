# ecwatech_export.py
# -*- coding: utf-8 -*-

"""
Экспорт экспонентов EcwaTech:
- читает JSON ({"exhibitors":[...]}) ИЛИ DOCX, где внутри лежит JSON,
- нормализует данные в единую таблицу (без колонки логотип),
- даёт хелперы для сохранения в Excel/CSV/JSON.

Зависимости (для DOCX и экспорта):
    pip install lxml pandas openpyxl
"""

import json, re, zipfile
from pathlib import Path
from typing import List, Dict, Tuple

# lxml нужен только для чтения DOCX; если его нет — JSON всё равно работает
try:
    import lxml.etree as ET
except Exception:
    ET = None

BASE_URL = "https://personal-account.expovr.ru"

# ───────────────────────── helpers ─────────────────────────

def pick_text(node):
    """Вернёт node['ru'] или node['en'] или ''."""
    if isinstance(node, dict):
        return node.get("ru") or node.get("en") or ""
    return node or ""

def build_url(path: str, file_name: str = "") -> str:
    """Собирает абсолютный URL к файлу по паре (path, file)."""
    if not path:
        return ""
    p = str(path).lstrip("./")
    if not p.startswith("/"):
        p = "/" + p
    if file_name:
        if not p.endswith("/"):
            p += "/"
        return f"{BASE_URL}{p}{file_name}"
    return f"{BASE_URL}{p}"

def clean_sites(val: str) -> str:
    """Нормализует строку сайтов: split по пробелам/запятым, +https:// при необходимости, уникализация."""
    if not val:
        return ""
    parts = re.split(r"[,\s]+", val.strip())
    parts = [p for p in parts if p]

    def norm(u):
        if re.match(r"^https?://", u, re.I):
            return u
        if re.match(r"^[\w.-]+\.[a-z]{2,}$", u, re.I):
            return "https://" + u
        return u

    seen, out = set(), []
    for p in (norm(x) for x in parts):
        if p and p not in seen:
            seen.add(p); out.append(p)
    return "; ".join(out)

# ───────────────────────── core ─────────────────────────

def normalize_exhibitor(ex: Dict) -> Dict:
    """Приводит объект экспонента к плоской строке таблицы (без «Логотип»)."""
    contacts = ex.get("contacts", {}) or {}
    city = contacts.get("city", {})
    city_txt = pick_text(city) if isinstance(city, dict) else (city or "")
    country = contacts.get("country", "") or ""

    # соцсети
    links = ex.get("links", {}) or {}
    socials = "; ".join(f"{k.upper()}: {v}" for k, v in links.items() if v)

    # описание: descr → about → brief
    descr = pick_text(ex.get("descr", {})) or pick_text(ex.get("about", {})) or pick_text(ex.get("brief", {}))

    # PDF-материалы
    pdfs = []
    for f in (ex.get("files", []) or []):
        url = build_url(f.get("path", ""), f.get("file", ""))
        nm = pick_text(f.get("name", {}))
        if url:
            pdfs.append(f"{nm}: {url}" if nm else url)

    # сайт(ы)
    site = clean_sites(contacts.get("www", "") or "")

    return {
        "Название компании": pick_text(ex.get("name", {})),
        "Сайт": site,
        "E-mail": contacts.get("email", "") or "",
        "Телефон": contacts.get("phone", "") or "",
        "Стенд": ex.get("stand", "") or "",
        "Город": city_txt,
        "Страна": country,
        "Соцсети": socials,
        "Описание": descr,
        "PDF-материалы": " | ".join(pdfs),
    }

def load_json(path: Path) -> Dict:
    """Читает обычный JSON-файл."""
    return json.loads(path.read_text(encoding="utf-8"))

def _docx_extract_text(path: Path) -> str:
    """Извлекает текст из DOCX (word/document.xml). Нужен lxml."""
    if ET is None:
        raise RuntimeError("Для разбора DOCX требуется lxml. Установите: pip install lxml")
    with zipfile.ZipFile(path, 'r') as z:
        xml = z.read('word/document.xml')
    root = ET.fromstring(xml)
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    paras = ["".join(t.text or "" for t in p.findall('.//w:t', ns)) for p in root.findall('.//w:p', ns)]
    return "\n".join(paras)

def load_from_docx(path: Path) -> Dict:
    """Достаёт JSON из DOCX и возвращает dict с ключом 'exhibitors'."""
    full_text = _docx_extract_text(path)
    s, e = full_text.find("{"), full_text.rfind("}")
    if s == -1 or e == -1 or e <= s:
        raise ValueError("Не нашёл JSON в DOCX.")
    raw = full_text[s:e+1]

    # вырезаем по-объектно массив exhibitors (устойчиво к мусору/переносам)
    m = re.search(r'"exhibitors"\s*:\s*\[', raw)
    if not m:
        raise ValueError('Не найден ключ "exhibitors" в DOCX.')
    pos = m.end()
    depth = 0
    in_str = False
    esc = False
    start = None
    chunks = []
    while pos < len(raw):
        ch = raw[pos]
        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
        else:
            if ch == '"':
                in_str = True
            elif ch == "{":
                if depth == 0:
                    start = pos
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0 and start is not None:
                    chunks.append(raw[start:pos+1])
                    start = None
            elif ch == "]" and depth == 0:
                break
        pos += 1

    items = []
    for chunk in chunks:
        cleaned = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", " ", chunk).replace("\u00a0", " ")
        try:
            items.append(json.loads(cleaned))
        except json.JSONDecodeError:
            cleaned2 = re.sub(r"\\x[0-9a-fA-F]{2}", " ", cleaned)
            try:
                items.append(json.loads(cleaned2))
            except Exception:
                # пропускаем битый объект, но продолжаем
                continue

    return {"exhibitors": items}

def load_exhibitors_any(path_str: str) -> Tuple[List[Dict], int]:
    """
    Читает JSON или DOCX с JSON.
    Возвращает (rows, count), где rows — нормализованные строки таблицы.
    """
    path = Path(path_str)
    if not path.exists():
        raise FileNotFoundError(f"Файл не найден: {path}")

    if path.suffix.lower() == ".docx":
        data = load_from_docx(path)
    else:
        data = load_json(path)

    if isinstance(data, dict) and isinstance(data.get("exhibitors"), list):
        exhibitors = data["exhibitors"]
    elif isinstance(data, list):
        exhibitors = data
    else:
        raise ValueError("Ожидаю {'exhibitors': [...]} или массив объектов.")

    rows = [normalize_exhibitor(ex) for ex in exhibitors]
    rows.sort(key=lambda r: (r.get("Название компании") or "").lower())
    return rows, len(rows)

# ───────────────────────── export helpers ─────────────────────────

def save_excel(rows: List[Dict], out_path: str):
    """Сохраняет таблицу в Excel (без колонки логотипа)."""
    try:
        import pandas as pd
    except Exception:
        raise RuntimeError("Нужен pandas и openpyxl: pip install pandas openpyxl")
    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(rows).to_excel(out_path, index=False)
    return out_path

def save_csv(rows: List[Dict], out_path: str):
    """Сохраняет таблицу в CSV (UTF-8 BOM)."""
    try:
        import pandas as pd
    except Exception:
        raise RuntimeError("Нужен pandas: pip install pandas")
    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(rows).to_csv(out_path, index=False, encoding="utf-8-sig")
    return out_path

def save_json(rows: List[Dict], out_path: str):
    """Сохраняет нормализованный JSON."""
    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    Path(out_path).write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    return out_path
