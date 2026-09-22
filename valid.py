"""
Валидация документов из папки ./documents на соответствие требованиям базы знаний:

    - Структура документов
    - Требования к языку и формулировкам
    - Актуальность документов (по внутренним признакам файла и кросс-документая проверка по порогу)
    - Отсутствие дублей (внутри файла — через LLM,
                            между файлами — через сравнение текстов)

Результат:
    ./Отчет_валидации.xlsx
        лист "Сводка"      — по каждому файлу и каждому критерию: статус + находки
        лист "Дубли между файлами" — пары похожих/дублирующихся файлов
        лист "Версии документов" - пары кандидатов с разными версиями одного и того же
        лист "Ошибки"      — файлы, которые не удалось обработать


python valid.py --docs ./documents
"""

import os
import re
import json
import time
import hashlib
import difflib
import argparse
from pathlib import Path
from collections import defaultdict
from datetime import datetime

import pandas as pd
import requests

# ------------------------------------------------------------
# Извлечение текста из документов: пробуем "богатый" парсер
# (anydoc), затем — стандартные библиотеки под конкретный формат.
# Если ничего не установлено — файл будет пропущен с пояснением
# в листе "Ошибки", а не уронит весь скрипт.
# ------------------------------------------------------------

try:
    import anydoc
    ANYDOC_AVAILABLE = True
except ImportError:
    ANYDOC_AVAILABLE = False

try:
    import docx  # python-docx
    from docx.oxml.ns import qn
    from docx.table import Table as DocxTable
    from docx.text.paragraph import Paragraph as DocxParagraph
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False


# Подпись к рисунку/таблице: "Рис. 1", "Рисунок 2:", "Табл. 3", "Таблица 1.",
# "Figure 1", "Table 2" и т.п. — используется и для docx, и для PDF.
CAPTION_PATTERN = re.compile(
    r"^(рис(\.|унок)?|скриншот|табл(\.|ица)?|figure|fig\.|table)\s*\.?\s*№?\s*\d*",
    re.IGNORECASE,
)


# ============================================================
# КОНФИГУРАЦИЯ
# ============================================================

class Config:

    INPUT_DOCS_DIR = Path("./documents")
    OUTPUT_REPORT = Path("./Отчет_валидации.xlsx")

    TOKEN = os.environ.get("TOKEN", "")
    MODEL = os.environ.get("MODEL", "qwen3.6-35b-a3b")
    API_URL = os.environ.get("API_URL", "")

    TEMPERATURE = 0.0
    SEED = 42
    MAX_TOKENS = 4000
    MAX_RETRIES = 5

    CHUNK_CHARS = 40000

    # Порог схожести текста (0..1) для признания двух файлов дублями
    DUPLICATE_SIMILARITY_THRESHOLD = 0.85

    SUPPORTED_EXTENSIONS = {".docx", ".doc", ".pdf", ".txt", ".md"}

    TEST_MODE = False


# ============================================================
# ИЗВЛЕЧЕНИЕ ТЕКСТА ИЗ ДОКУМЕНТА
# ============================================================

def extract_text_docx(file_path):
    if not DOCX_AVAILABLE:
        return ""
    try:
        d = docx.Document(str(file_path))
        parts = []
        for para in d.paragraphs:
            if para.text.strip():
                parts.append(para.text)
        for table in d.tables:
            for row in table.rows:
                cells = [c.text.strip() for c in row.cells]
                if any(cells):
                    parts.append(" | ".join(cells))
        return "\n".join(parts)
    except Exception as e:
        print(f"   ⚠️ python-docx не справился: {e}")
        return ""


def extract_text_pdf(file_path):
    if not PDFPLUMBER_AVAILABLE:
        return ""
    try:
        parts = []
        with pdfplumber.open(str(file_path)) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    parts.append(text)
        return "\n".join(parts)
    except Exception as e:
        print(f"   ⚠️ pdfplumber не справился: {e}")
        return ""


def get_document_text(file_path):

    file_path = Path(file_path)
    suffix = file_path.suffix.lower()

    # 1. anydoc — если установлен, даёт лучший markdown с разметкой
    if ANYDOC_AVAILABLE:
        try:
            markdown = anydoc.to_markdown(str(file_path))
            if markdown and len(markdown.strip()) > 20:
                return markdown
        except Exception as e:
            print(f"   ⚠️ anydoc не справился с {file_path.name}: {e}")

    # 2. Фолбэки по типу файла
    if suffix in (".docx", ".doc"):
        text = extract_text_docx(file_path)
        if text:
            return text

    if suffix == ".pdf":
        text = extract_text_pdf(file_path)
        if text:
            return text

    if suffix in (".txt", ".md"):
        try:
            return file_path.read_text(encoding="utf-8", errors="ignore")
        except Exception as e:
            print(f"   ⚠️ Не удалось прочитать {file_path.name}: {e}")
            return ""

    return ""


# ============================================================
# РАЗБОР РАЗМЕТКИ (docx/pdf): где физически расположены
# рисунки/таблицы относительно текста и подписей
# ============================================================
#
# Извлечённый "плоский" текст не показывает, вклеена ли картинка
# в строку текста или вынесена отдельным абзацем — эта информация
# есть только в исходной разметке файла. Ниже — программный (не
# LLM) разбор структуры документа, дающий детерминированный ответ.

def iter_docx_block_items(document):
    """
    Идёт по телу docx-документа в порядке следования и отдаёт
    параграфы и таблицы как единый поток (стандартный рецепт для
    python-docx, т.к. document.paragraphs и document.tables отдают
    их РАЗДЕЛЬНО, без сохранения общего порядка и связи "рядом").
    """
    body = document.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield "paragraph", DocxParagraph(child, document)
        elif child.tag == qn("w:tbl"):
            yield "table", DocxTable(child, document)


def paragraph_has_image(paragraph):
    for run in paragraph.runs:
        if run._element.findall(qn("w:drawing")) or run._element.findall(qn("w:pict")):
            return True
    return False


def find_nearby_caption(blocks, idx, max_empty_skip=2):
    """
    Ищет подпись только у БЛИЖАЙШЕГО непустого абзаца до и после
    блока idx (пропуская до max_empty_skip пустых абзацев-разделителей).
    Специально НЕ смотрит дальше первого содержательного абзаца в
    каждую сторону — иначе подпись соседнего, другого рисунка/таблицы
    может ошибочно "приписаться" текущему объекту.
    """
    captioned = False
    context_snippet = None

    for direction in (1, -1):
        steps, j = 0, idx + direction
        while 0 <= j < len(blocks) and steps <= max_empty_skip:
            kind, block = blocks[j]
            if kind != "paragraph":
                break  # соседняя таблица — не переходим через неё в поисках подписи
            text_j = block.text.strip()
            if text_j:
                if context_snippet is None:
                    context_snippet = text_j[:80]
                if CAPTION_PATTERN.match(text_j):
                    captioned = True
                    context_snippet = text_j[:80]
                break
            j += direction
            steps += 1

    return captioned, context_snippet


def analyze_docx_structure(file_path):
    """
    Возвращает {"figures": [...], "tables": [...]} — для каждой
    картинки: isolated (True если абзац с картинкой не содержит
    другого текста) и captioned (найдена ли подпись рядом); для
    каждой таблицы: captioned.
    """
    if not DOCX_AVAILABLE:
        return None
    try:
        document = docx.Document(str(file_path))
    except Exception as e:
        return {"figures": [], "tables": [], "parse_error": str(e)}

    blocks = list(iter_docx_block_items(document))
    figures, tables_info = [], []

    for idx, (kind, block) in enumerate(blocks):

        if kind == "paragraph" and paragraph_has_image(block):
            own_text = block.text.strip()
            captioned, context_snippet = find_nearby_caption(blocks, idx)
            figures.append({
                "paragraph_index": idx,
                "isolated": len(own_text) == 0,
                "captioned": captioned,
                "own_text_sample": own_text[:80] if own_text else None,
                "context_snippet": context_snippet,
            })

        elif kind == "table":
            captioned, context_snippet = find_nearby_caption(blocks, idx)
            tables_info.append({
                "table_index": idx,
                "captioned": captioned,
                "context_snippet": context_snippet,
            })

    return {"figures": figures, "tables": tables_info}


def get_pdf_text_lines(page, y_tolerance=3):
    """Группирует слова PDF-страницы в текстовые строки по вертикали."""
    words = page.extract_words()
    lines_map = defaultdict(list)
    for w in words:
        key = round(w["top"] / y_tolerance) * y_tolerance
        lines_map[key].append(w)

    lines = []
    for _, ws in lines_map.items():
        ws_sorted = sorted(ws, key=lambda w: w["x0"])
        lines.append({
            "text": " ".join(w["text"] for w in ws_sorted),
            "top": min(w["top"] for w in ws_sorted),
            "bottom": max(w["bottom"] for w in ws_sorted),
        })
    return lines


def analyze_pdf_structure(file_path, caption_margin=40):
    """
    Для каждой картинки на странице PDF проверяет:
    - isolated: нет ли текстовых строк, вертикально пересекающихся
      с картинкой (кроме коротких служебных, типа номера страницы)
      — если пересекаются, значит текст обтекает/делит с ней строку;
    - captioned: есть ли строка-подпись в пределах caption_margin
      пунктов сверху/снизу от картинки.
    Аналогично для таблиц, найденных через page.find_tables().
    """
    if not PDFPLUMBER_AVAILABLE:
        return None

    figures, tables_info = [], []

    try:
        with pdfplumber.open(str(file_path)) as pdf:
            for page_num, page in enumerate(pdf.pages, start=1):
                lines = get_pdf_text_lines(page)

                for img_idx, img in enumerate(page.images, start=1):
                    img_top, img_bottom = img.get("top"), img.get("bottom")

                    overlapping = [
                        l for l in lines
                        if not (l["bottom"] < img_top or l["top"] > img_bottom)
                        and len(l["text"].strip()) > 3
                    ]
                    nearby = [
                        l for l in lines
                        if (img_bottom <= l["top"] <= img_bottom + caption_margin)
                        or (img_top - caption_margin <= l["bottom"] <= img_top)
                    ]
                    captioned = any(CAPTION_PATTERN.match(l["text"].strip()) for l in nearby)

                    figures.append({
                        "page": page_num,
                        "index": img_idx,
                        "isolated": len(overlapping) == 0,
                        "captioned": captioned,
                        "overlap_text_sample": overlapping[0]["text"][:80] if overlapping else None,
                    })

                if hasattr(page, "find_tables"):
                    try:
                        found_tables = page.find_tables()
                    except Exception:
                        found_tables = []

                    for t_idx, t in enumerate(found_tables, start=1):
                        bbox = t.bbox  # (x0, top, x1, bottom)
                        nearby = [
                            l for l in lines
                            if (bbox[1] - caption_margin <= l["bottom"] <= bbox[1])
                            or (bbox[3] <= l["top"] <= bbox[3] + caption_margin)
                        ]
                        captioned = any(CAPTION_PATTERN.match(l["text"].strip()) for l in nearby)
                        tables_info.append({
                            "page": page_num, "index": t_idx, "captioned": captioned
                        })

    except Exception as e:
        return {"figures": [], "tables": [], "parse_error": str(e)}

    return {"figures": figures, "tables": tables_info}


def analyze_layout(file_path):
    suffix = Path(file_path).suffix.lower()
    if suffix in (".docx", ".doc"):
        return analyze_docx_structure(file_path)
    if suffix == ".pdf":
        return analyze_pdf_structure(file_path)
    return None  # .txt/.md — рисунков/таблиц как объектов разметки нет


def build_layout_summary_text(layout):
    """Текстовая сводка для передачи в промпт LLM как проверенный факт."""
    if layout is None:
        return ("Автоматический разбор разметки для этого формата не выполнялся "
                "(например, .txt/.md — рисунков и таблиц как объектов разметки в "
                "таком формате нет).")
    if layout.get("parse_error"):
        return (f"Не удалось программно разобрать разметку файла: "
                f"{layout['parse_error']}. Оценивай структуру только по тексту.")

    figures, tables = layout.get("figures", []), layout.get("tables", [])
    if not figures and not tables:
        return "Программный разбор разметки не обнаружил встроенных изображений или таблиц в этом файле."

    lines = [
        "Технические сигналы, полученные ПРОГРАММНЫМ разбором разметки файла "
        "(это точный анализ структуры, а не догадка по тексту — доверяй этим "
        "данным больше, чем собственному впечатлению, и используй их для полей "
        "tables_figures_isolated_in_paragraphs / tables_figures_captioned):"
    ]
    for f in figures:
        loc = f.get("paragraph_index", f.get("page"))
        lines.append(
            f"- Рисунок (позиция {loc}): "
            f"{'выделен отдельным абзацем' if f.get('isolated') else 'НАХОДИТСЯ В ОДНОЙ СТРОКЕ/АБЗАЦЕ С ТЕКСТОМ'}; "
            f"{'подпись найдена рядом' if f.get('captioned') else 'подпись НЕ найдена рядом'}."
        )
    for t in tables:
        loc = t.get("table_index", t.get("page"))
        lines.append(
            f"- Таблица (позиция {loc}): "
            f"{'подпись найдена рядом' if t.get('captioned') else 'подпись НЕ найдена рядом'}."
        )
    return "\n".join(lines)


def merge_layout_into_verdict(verdict, layout):
    """
    Переопределяет/дополняет оценку LLM по структуре точными данными
    разбора разметки — там, где разбор реально что-то нашёл. Если
    рисунков/таблиц в файле нет или разбор не удался, оценка LLM
    остаётся как есть.
    """
    if not layout or layout.get("parse_error"):
        return verdict
    figures, tables = layout.get("figures", []), layout.get("tables", [])
    if not figures and not tables:
        return verdict

    structure = verdict.setdefault("structure", {})
    findings = list(structure.get("findings", []) or [])

    inline_figures = [f for f in figures if not f.get("isolated", True)]
    uncaptioned_figures = [f for f in figures if not f.get("captioned", False)]
    uncaptioned_tables = [t for t in tables if not t.get("captioned", False)]

    for f in inline_figures:
        loc = f.get("paragraph_index", f.get("page"))
        findings.append(
            f"[автоматически] Рисунок (позиция {loc}) вклеен в тот же абзац/строку, "
            f"где есть обычный текст — распознавание при обработке базы знаний может не сработать."
        )
    for f in uncaptioned_figures:
        loc = f.get("paragraph_index", f.get("page"))
        findings.append(f"[автоматически] Не найдена подпись рядом с рисунком (позиция {loc}).")
    for t in uncaptioned_tables:
        loc = t.get("table_index", t.get("page"))
        findings.append(f"[автоматически] Не найдена подпись рядом с таблицей (позиция {loc}).")

    structure["findings"] = findings
    if figures:
        structure["tables_figures_isolated_in_paragraphs"] = (len(inline_figures) == 0)
    if figures or tables:
        structure["tables_figures_captioned"] = (
            len(uncaptioned_figures) == 0 and len(uncaptioned_tables) == 0
        )
    if inline_figures or uncaptioned_figures or uncaptioned_tables:
        structure["status"] = "issues"

    verdict["structure"] = structure
    return verdict


# ============================================================
# ПРОМПТ ДЛЯ LLM
# ============================================================

VALIDATION_SYSTEM_PROMPT = """
Ты — эксперт по документационному менеджменту и методологии ведения
базы знаний компании. Тебе дан текст одного документа. Проверь его
по чек-листу ниже и верни результат СТРОГО в JSON по заданной схеме.

Требования, на соответствие которым нужно проверить документ
(проверяй ТОЛЬКО эти пункты, ничего не выдумывай сверх них):

## Структура документа
Обязательные элементы:
- чёткие заголовки и подзаголовки;
- деление на разделы и смысловые блоки;
- понятные подписи к таблицам, скриншотам и т.д.;
- если в файле описано НЕСКОЛЬКО разных процессов — они должны быть
  логически разделены заголовками/подзаголовками.
Важно: рисунки, таблицы, скриншоты должны быть выделены отдельными
абзацами, а не вклеены прямо в строку текста (иначе распознавание
не сработает). Для этого пункта тебе может быть передан отдельный
блок "Технические сигналы" — это результат ТОЧНОГО программного
разбора разметки файла (не догадка по тексту). Если такой блок
есть — ОБЯЗАТЕЛЬНО опирайся на него для полей
tables_figures_isolated_in_paragraphs и tables_figures_captioned,
а не на собственное впечатление от текста. Если блока нет или он
говорит, что рисунков/таблиц не найдено — оценивай по тексту как
обычно.
Не рекомендуется:
- объединение разных регламентов в одном файле без структуры;
- использование черновых материалов как рабочих инструкций (ищи
  маркеры: "черновик", "draft", "TODO", "не финально", незакрытые
  комментарии, следы правок "удалить этот абзац" и т.п.)

## Язык и формулировки
- простой деловой язык (без излишнего канцелярита и жаргона);
- аббревиатуры расшифрованы при первом употреблении;
- единый словарь терминов в конце документа — это ОПЦИОНАЛЬНО,
  отсутствие словаря — НЕ нарушение, просто отметь has_glossary=false;
- один документ должен быть выдержан на ОДНОМ языке (если два языка
  смешаны в теле текста, а не представляют собой две явно отдельные
  версии/секции перевода — это нарушение).

## Актуальность документа
Ты НЕ видишь, где физически хранится файл и есть ли у него другие
версии в системе — поэтому проверяй ТОЛЬКО внутренние признаки:
- есть ли в тексте указание версии/номера редакции/даты утверждения;
- есть ли явные маркеры устаревания в самом тексте: "устарело",
  "не применяется", "заменено на...", "старая версия",
  перечёркнутые/помеченные как неактуальные разделы.
Если таких явных внутренних признаков в тексте нет — status
должен быть "needs_manual_check" (не "issues"), т.к. полноценная
проверка актуальности требует сверки с хранилищем документов, а не
только с текстом внутри файла.

## Отсутствие дублей (ТОЛЬКО внутри этого одного файла)
Проверь, нет ли внутри ЭТОГО документа повторяющихся разделов/
абзацев с одинаковым или почти одинаковым содержанием (дублирование
информации внутри одного файла). Сравнение с ДРУГИМИ файлами тебе
делать не нужно — это делается отдельно.

ВАЖНО
-----
- Не придумывай нарушений, которых нет в тексте — если критерий
  выполнен, статус "ok" и findings пустой список.
- findings — краткие, конкретные, с указанием на место в тексте
  (можно короткой цитатой до 10-12 слов или пересказом), а не общие
  фразы вроде "есть проблемы со структурой".
- Если текст документа обрезан/является только частью большого
  файла (см. пометку в начале текста, если она есть) — оценивай
  честно то, что видишь, не додумывай остальное.

Ответ — ТОЛЬКО JSON по схеме, без пояснений вне JSON.
"""

VALIDATION_USER_PROMPT_TEMPLATE = """
Файл: {filename}
{chunk_note}

{layout_signals}

Текст документа:

{document_text}
"""


# ============================================================
# JSON PARSING (как в исходном скрипте)
# ============================================================

def extract_balanced_json(text):
    if not text:
        return None
    start = text.find("{")
    if start == -1:
        return None
    depth = 0
    for i in range(start, len(text)):
        char = text[i]
        if char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                candidate = text[start:i + 1]
                try:
                    return json.loads(candidate)
                except json.JSONDecodeError:
                    return None
    return None


def parse_json_response(text, finish_reason=None):
    if not text:
        return None
    text = text.strip()
    text = re.sub(r"^```json\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"^```\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    result = extract_balanced_json(text)
    if result is not None:
        return result
    if finish_reason == "length":
        print("   ❌ Ответ LLM обрезан по MAX_TOKENS")
    else:
        print(f"   ⚠️ JSON не распознан (finish_reason={finish_reason})")
        print(f"   Ответ: {text[:500]!r}")
    return None


# ============================================================
# ВЫЗОВ LLM
# ============================================================

SCHEMA = None


def load_schema(schema_path="schema.json"):
    try:
        with open(schema_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"❌ Файл схемы не найден: {schema_path}")
        return None
    except json.JSONDecodeError as e:
        print(f"❌ Ошибка JSON-схемы: {e}")
        return None


def call_llm_validate(document_text, filename, chunk_note="", layout_signals_text="",
                       max_retries=Config.MAX_RETRIES):

    if Config.TEST_MODE:
        return {
            "detected_languages": ["ru"],
            "single_language_ok": True,
            "structure": {
                "status": "ok", "has_headings": True, "has_section_division": True,
                "tables_figures_captioned": True, "tables_figures_isolated_in_paragraphs": True,
                "multiple_processes_separated": None,
                "looks_like_merged_unstructured_regulations": False,
                "looks_like_draft_material": False, "findings": []
            },
            "language_style": {
                "status": "ok", "plain_business_language": True,
                "abbreviations_undecoded": [], "has_glossary": False, "findings": []
            },
            "currency_signals": {
                "status": "needs_manual_check", "version_or_date_found": None,
                "outdated_markers_found": [], "findings": []
            },
            "internal_duplication": {
                "status": "ok", "duplicated_fragments_found": [], "findings": []
            },
            "overall_summary": "TEST MODE"
        }

    user_prompt = VALIDATION_USER_PROMPT_TEMPLATE.format(
        filename=filename,
        chunk_note=chunk_note,
        layout_signals=layout_signals_text or "",
        document_text=document_text[:Config.CHUNK_CHARS],
    )

    headers = {
        "Authorization": f"Bearer {Config.TOKEN}",
        "Content-Type": "application/json",
    }

    payload = {
        "model": Config.MODEL,
        "messages": [
            {"role": "system", "content": VALIDATION_SYSTEM_PROMPT},
            {"role": "user", "content": user_prompt},
        ],
        "temperature": Config.TEMPERATURE,
        "max_tokens": Config.MAX_TOKENS,
        "seed": Config.SEED,
        "response_format": {
            "type": "json_schema",
            "json_schema": {
                "name": SCHEMA.get("name", "document_validation_result"),
                "schema": SCHEMA.get("schema", SCHEMA),
                "strict": True,
            },
        },
    }

    for attempt in range(1, max_retries + 1):
        try:
            response = requests.post(Config.API_URL, headers=headers, json=payload, timeout=180)

            if response.status_code == 200:
                data = response.json()
                try:
                    message = data["choices"][0]["message"]
                    text = message["content"]
                    finish_reason = data["choices"][0].get("finish_reason")
                except (KeyError, IndexError, TypeError):
                    print("   ❌ Неожиданный формат ответа API")
                    return None

                result = parse_json_response(text, finish_reason)
                return result

            if response.status_code == 429:
                wait_time = 15 * attempt
                print(f"   ⏳ HTTP 429. Ждём {wait_time} сек.")
                time.sleep(wait_time)
                continue

            if response.status_code in (500, 502, 503, 504):
                wait_time = 10 * attempt
                print(f"   ⏳ HTTP {response.status_code}. Повтор через {wait_time} сек.")
                time.sleep(wait_time)
                continue

            print(f"   ❌ HTTP {response.status_code}: {response.text[:500]}")
            return None

        except requests.exceptions.Timeout:
            print(f"   ⏳ Timeout {attempt}/{max_retries}")
            time.sleep(10 * attempt)
        except requests.exceptions.RequestException as e:
            print(f"   ⚠️ Ошибка сети: {e}")
            time.sleep(10 * attempt)
        except Exception as e:
            print(f"   ⚠️ Неожиданная ошибка: {e}")
            time.sleep(5 * attempt)

    return None


def merge_chunk_results(chunk_results):
    """
    Если документ был разбит на несколько кусков (очень длинный
    файл), объединяет вердикты по кускам в один: статус — худший
    из встреченных ("issues" > "partial"/"needs_manual_check" >
    "ok"), находки — конкатенируются.
    """

    if not chunk_results:
        return None
    if len(chunk_results) == 1:
        return chunk_results[0]

    status_priority = {"issues": 2, "partial": 1, "needs_manual_check": 1, "ok": 0}

    def merge_section(key, status_field="status", list_fields=("findings",)):
        best_status = "ok"
        merged_lists = {f: [] for f in list_fields}
        extra = defaultdict(list)

        for r in chunk_results:
            section = r.get(key, {}) or {}
            st = section.get(status_field, "ok")
            if status_priority.get(st, 0) > status_priority.get(best_status, 0):
                best_status = st
            for f in list_fields:
                merged_lists[f].extend(section.get(f, []) or [])
            for k, v in section.items():
                if k not in (status_field,) and k not in list_fields:
                    if isinstance(v, bool):
                        extra[k].append(v)
                    elif isinstance(v, list):
                        extra[k].extend(v)
                    else:
                        extra[k].append(v)

        out = {status_field: best_status}
        for f in list_fields:
            out[f] = merged_lists[f]
        for k, vals in extra.items():
            if all(isinstance(v, bool) for v in vals):
                out[k] = all(vals)  # консервативно: все куски должны подтвердить
            else:
                # берём непустые/непустые уникальные значения
                uniq = [v for v in vals if v not in (None, "")]
                out[k] = uniq[0] if uniq else (vals[0] if vals else None)
        return out

    languages = set()
    for r in chunk_results:
        languages.update(r.get("detected_languages", []) or [])

    return {
        "detected_languages": sorted(languages),
        "single_language_ok": all(r.get("single_language_ok", True) for r in chunk_results),
        "structure": merge_section("structure", list_fields=("findings",)),
        "language_style": merge_section(
            "language_style", list_fields=("findings", "abbreviations_undecoded")
        ),
        "currency_signals": merge_section(
            "currency_signals", list_fields=("findings", "outdated_markers_found")
        ),
        "internal_duplication": merge_section(
            "internal_duplication", list_fields=("findings", "duplicated_fragments_found")
        ),
        "overall_summary": " | ".join(
            r.get("overall_summary", "") for r in chunk_results if r.get("overall_summary")
        ),
    }


def analyze_long_document(text, filename, layout_signals_text=""):
    """
    Режет очень длинный текст на куски по Config.CHUNK_CHARS (со
    небольшим перехлёстом, чтобы не резать абзац точно по стыку) и
    анализирует каждый кусок отдельно, затем объединяет вердикты.
    layout_signals_text (результат разбора разметки) передаётся в
    КАЖДЫЙ кусок — сама разметка (расположение картинок/таблиц)
    не режется на части, это данные по документу в целом.
    """

    chunk_size = Config.CHUNK_CHARS
    overlap = 500

    if len(text) <= chunk_size:
        result = call_llm_validate(text, filename, layout_signals_text=layout_signals_text)
        return [result] if result else []

    chunks = []
    start = 0

    while start < len(text):
        end = min(start + chunk_size, len(text))
        chunks.append(text[start:end])
        if end == len(text):
            break
        start = end - overlap

    results = []
    for i, chunk in enumerate(chunks, start=1):
        note = f"(часть {i} из {len(chunks)} — документ разбит из-за большого размера)"
        print(f"   📄 Анализ части {i}/{len(chunks)} файла {filename}")
        result = call_llm_validate(chunk, filename, chunk_note=note, layout_signals_text=layout_signals_text)
        if result:
            results.append(result)

    return results


# ============================================================
# МЕЖФАЙЛОВЫЕ ДУБЛИ (не через LLM — быстрее и детерминированнее)
# ============================================================

def normalize_for_comparison(text):
    text = text.lower()
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[^\w\s]", "", text, flags=re.UNICODE)
    return text.strip()


def text_similarity(a, b):
    if not a or not b:
        return 0.0
    # Для очень длинных текстов SequenceMatcher может быть медленным —
    # ограничиваем сравнение первыми N символами, этого обычно
    # достаточно, чтобы отличить разные регламенты от дублей.
    a_short = a[:40000]
    b_short = b[:40000]
    return difflib.SequenceMatcher(None, a_short, b_short).ratio()


def compute_similarity_pairs(file_texts):
    """Считает попарную текстовую схожесть один раз — переиспользуется
    и для поиска дублей , и для группировки версий."""
    names = list(file_texts.keys())
    normalized = {n: normalize_for_comparison(file_texts[n]) for n in names}
    pairs = []
    for i in range(len(names)):
        for j in range(i + 1, len(names)):
            a, b = names[i], names[j]
            pairs.append((a, b, text_similarity(normalized[a], normalized[b])))
    return pairs


def find_cross_file_duplicates(file_texts, pairs=None, threshold=None):
    """
    file_texts: dict {filename: raw_text}
    pairs: результат compute_similarity_pairs(file_texts) — если None,
           считается заново.
    Возвращает список найденных пар-дублей/почти-дублей:
        [{"file_a":..., "file_b":..., "similarity": 0.0-1.0, "note": ...}]
    """

    threshold = threshold if threshold is not None else Config.DUPLICATE_SIMILARITY_THRESHOLD
    pairs = pairs if pairs is not None else compute_similarity_pairs(file_texts)

    names = list(file_texts.keys())
    normalized = {name: normalize_for_comparison(file_texts[name]) for name in names}

    # быстрый предфильтр по точному хешу — 100% дубли содержимого
    hashes = defaultdict(list)
    for name in names:
        h = hashlib.sha256(normalized[name].encode("utf-8")).hexdigest()
        hashes[h].append(name)

    duplicates = []
    exact_dupe_pairs = set()

    for h, group in hashes.items():
        if len(group) > 1:
            for i in range(len(group)):
                for j in range(i + 1, len(group)):
                    duplicates.append({
                        "file_a": group[i], "file_b": group[j],
                        "similarity": 1.0, "note": "Полностью идентичное содержимое"
                    })
                    exact_dupe_pairs.add(frozenset((group[i], group[j])))

    # похожесть содержимого (без точного совпадения хеша)
    for name_a, name_b, sim in pairs:
        if frozenset((name_a, name_b)) in exact_dupe_pairs:
            continue
        if sim >= threshold:
            duplicates.append({
                "file_a": name_a, "file_b": name_b,
                "similarity": round(sim, 3),
                "note": "Похожее содержимое (возможный дубль/почти-дубль)"
            })

    # похожие имена файлов (напр. "Регламент_v1.docx" и "Регламент_v2.docx"
    # или "Регламент (копия).docx") — сигнал, даже если
    # содержимое чуть разошлось
    for i in range(len(names)):
        for j in range(i + 1, len(names)):
            name_sim = difflib.SequenceMatcher(
                None, names[i].lower(), names[j].lower()
            ).ratio()
            already_flagged = any(
                {d["file_a"], d["file_b"]} == {names[i], names[j]} for d in duplicates
            )
            if name_sim >= 0.8 and not already_flagged:
                duplicates.append({
                    "file_a": names[i], "file_b": names[j],
                    "similarity": round(name_sim, 3),
                    "note": "Очень похожие ИМЕНА файлов — проверить, не дубль/копия ли это"
                })

    return duplicates


# ------------------------------------------------------------
# ГРУППЫ ВЕРСИЙ ОДНОГО ДОКУМЕНТА
# ------------------------------------------------------------
# Правило: "в базе знаний должна быть одна актуальная версия
# документа, устаревшие — в архиве в другом месте". Раз все файлы,
# которые мы видим, лежат в ОДНОЙ папке ./documents — само наличие
# нескольких файлов, похожих по содержанию ИЛИ по имени (без учёта
# версии/даты в имени), уже сигнал: либо это нарушение (несколько
# версий не разнесены по местам хранения), либо один из файлов
# ошибочно остался не архивированным.

VERSION_FILENAME_STRIP_PATTERN = re.compile(
    r"(_?v\d+(\.\d+)?|_?версия\s*\d+|\(\d+\)|_?copy|_?копия|_?final|_?старая|"
    r"_?устарел\w*|_?draft|_?черновик|_?актуальн\w*|"
    r"_?\d{4}[-_]\d{2}[-_]\d{2}|_?\d{2}[.\-_]\d{2}[.\-_]\d{2,4})",
    re.IGNORECASE,
)


def normalize_filename_stem(filename):
    stem = Path(filename).stem.lower()
    stem = VERSION_FILENAME_STRIP_PATTERN.sub("", stem)
    stem = re.sub(r"[\s_\-]+", " ", stem).strip()
    return stem


def _uf_find(parent, x):
    while parent[x] != x:
        parent[x] = parent[parent[x]]
        x = parent[x]
    return x


def _uf_union(parent, a, b):
    ra, rb = _uf_find(parent, a), _uf_find(parent, b)
    if ra != rb:
        parent[ra] = rb


def find_version_groups(file_paths, pairs, threshold_low=0.5):
    """
    Группирует файлы, которые, вероятно, являются версиями одного и
    того же документа: похожи по содержанию (>= threshold_low —
    более точные совпадения тоже сюда попадают, они же отдельно
    считаются дублями в 2.5) ИЛИ имеют совпадающее "имя без версии".
    Возвращает список групп (списков имён файлов), где len(group) > 1.
    """
    names = list(file_paths.keys())
    parent = {n: n for n in names}

    for a, b, sim in pairs:
        if sim >= threshold_low:
            _uf_union(parent, a, b)

    stems = defaultdict(list)
    for n in names:
        stem = normalize_filename_stem(n)
        if stem:
            stems[stem].append(n)
    for group in stems.values():
        if len(group) > 1:
            for k in range(1, len(group)):
                _uf_union(parent, group[0], group[k])

    groups = defaultdict(list)
    for n in names:
        groups[_uf_find(parent, n)].append(n)

    return [sorted(g) for g in groups.values() if len(g) > 1]


def guess_current_candidate(group, file_paths, verdicts_by_file):
    """
    Эвристика "какой файл в группе, вероятно, актуальный":
    1) нет явных маркеров устаревания в тексте (по вердикту LLM);
    2) выше номер версии, если он есть в имени файла (v1/v2/...);
    3) при прочих равных — более свежая дата изменения файла.
    Это ТОЛЬКО рекомендация для ручной проверки, не окончательное решение.
    """
    def score(name):
        verdict = verdicts_by_file.get(name, {}) or {}
        cs = verdict.get("currency_signals", {}) or {}
        has_outdated_marker = bool(cs.get("outdated_markers_found"))
        m = re.search(r"v(?:ersion)?[\s_\-]?(\d+)", name, re.IGNORECASE)
        version_num = int(m.group(1)) if m else -1
        try:
            mtime = file_paths[name].stat().st_mtime
        except Exception:
            mtime = 0
        return (0 if has_outdated_marker else 1, version_num, mtime)

    return max(group, key=score)


def build_version_group_rows(version_groups, file_paths, verdicts_by_file):
    rows = []
    for group in version_groups:
        candidate = guess_current_candidate(group, file_paths, verdicts_by_file)
        for name in group:
            verdict = verdicts_by_file.get(name, {}) or {}
            cs = verdict.get("currency_signals", {}) or {}
            try:
                mtime_str = datetime.fromtimestamp(
                    file_paths[name].stat().st_mtime
                ).strftime("%Y-%m-%d %H:%M")
            except Exception:
                mtime_str = "—"
            rows.append({
                "Группа файлов": " / ".join(group),
                "Файл": name,
                "Вердикт": (
                    "Похоже на актуальную версию (проверьте вручную)"
                    if name == candidate else
                    "Вероятно устаревшая версия/копия — кандидат на архивацию"
                ),
                "Версия/дата из текста": cs.get("version_or_date_found") or "—",
                "Маркеры устаревания в тексте": ", ".join(
                    cs.get("outdated_markers_found", []) or []
                ) or "—",
                "Дата изменения файла": mtime_str,
            })
    return rows


# ============================================================
# ОБРАБОТКА ПАПКИ И ФОРМИРОВАНИЕ ОТЧЁТА
# ============================================================

STATUS_RU = {
    "ok": "ОК",
    "partial": "Частично",
    "issues": "Есть нарушения",
    "needs_manual_check": "Нужна ручная проверка",
}


def status_to_ru(status):
    return STATUS_RU.get(status, status or "—")


def build_report_rows(filename, verdict):
    rows = []

    languages = ", ".join(verdict.get("detected_languages", []) or [])
    single_lang_ok = verdict.get("single_language_ok", True)

    s = verdict.get("structure", {}) or {}
    rows.append({
        "Файл": filename,
        "Критерий": "2.2 Структура документа",
        "Статус": status_to_ru(s.get("status")),
        "Находки": "; ".join(s.get("findings", []) or []) or "—",
        "Доп. детали": (
            f"заголовки={s.get('has_headings')}, "
            f"деление на разделы={s.get('has_section_division')}, "
            f"подписи таблиц/рисунков={s.get('tables_figures_captioned')}, "
            f"иллюстрации отдельными абзацами={s.get('tables_figures_isolated_in_paragraphs')}, "
            f"разные процессы разделены={s.get('multiple_processes_separated')}, "
            f"похоже на объединённые регламенты={s.get('looks_like_merged_unstructured_regulations')}, "
            f"похоже на черновик={s.get('looks_like_draft_material')}"
        ),
    })

    ls = verdict.get("language_style", {}) or {}
    rows.append({
        "Файл": filename,
        "Критерий": "2.3 Язык и формулировки",
        "Статус": status_to_ru(ls.get("status")),
        "Находки": "; ".join(ls.get("findings", []) or []) or "—",
        "Доп. детали": (
            f"языки в документе={languages or '—'}, один язык на документ={single_lang_ok}, "
            f"простой деловой язык={ls.get('plain_business_language')}, "
            f"нерасшифрованные аббревиатуры={', '.join(ls.get('abbreviations_undecoded', []) or []) or '—'}, "
            f"есть словарь терминов={ls.get('has_glossary')}"
        ),
    })

    cs = verdict.get("currency_signals", {}) or {}
    rows.append({
        "Файл": filename,
        "Критерий": "2.4 Актуальность (только внутренние признаки)",
        "Статус": status_to_ru(cs.get("status")),
        "Находки": "; ".join(cs.get("findings", []) or []) or "—",
        "Доп. детали": (
            f"версия/дата в тексте={cs.get('version_or_date_found') or '—'}, "
            f"маркеры устаревания={', '.join(cs.get('outdated_markers_found', []) or []) or '—'}. "
            f"Полная проверка актуальности требует сверки с местом хранения — вне возможностей LLM по тексту."
        ),
    })

    idup = verdict.get("internal_duplication", {}) or {}
    rows.append({
        "Файл": filename,
        "Критерий": "2.5 Дубли внутри файла",
        "Статус": status_to_ru(idup.get("status")),
        "Находки": "; ".join(idup.get("findings", []) or []) or "—",
        "Доп. детали": (
            "; ".join(idup.get("duplicated_fragments_found", []) or []) or "—"
        ),
    })

    return rows


def process_all_documents(docs_dir):

    docs_dir = Path(docs_dir)

    if not docs_dir.exists():
        print(f"❌ Папка не найдена: {docs_dir}")
        return [], {}, []

    files = [
        f for f in docs_dir.rglob("*")
        if f.is_file() and f.suffix.lower() in Config.SUPPORTED_EXTENSIONS
    ]

    print(f"📁 Найдено документов: {len(files)}")

    report_rows = []
    file_texts = {}
    file_paths = {}
    verdicts_by_file = {}
    error_rows = []

    for file_path in files:
        print(f"\n📄 Обработка: {file_path.name}")

        text = get_document_text(file_path)

        if not text or len(text.strip()) < 20:
            print(f"   ❌ Не удалось извлечь текст (или файл пустой)")
            error_rows.append({
                "Файл": file_path.name,
                "Ошибка": "Не удалось извлечь текст. Проверьте, установлены ли "
                          "python-docx / pdfplumber / anydoc для этого формата, "
                          "или файл действительно пустой/сканирован без OCR.",
            })
            continue

        file_texts[file_path.name] = text
        file_paths[file_path.name] = file_path

        layout = analyze_layout(file_path)
        layout_signals_text = build_layout_summary_text(layout)

        chunk_results = analyze_long_document(text, file_path.name, layout_signals_text)

        if not chunk_results:
            print(f"   ❌ LLM не вернула результат для {file_path.name}")
            error_rows.append({
                "Файл": file_path.name,
                "Ошибка": "LLM не вернула валидный результат после всех попыток.",
            })
            continue

        verdict = merge_chunk_results(chunk_results)
        verdict = merge_layout_into_verdict(verdict, layout)

        verdicts_by_file[file_path.name] = verdict
        report_rows.extend(build_report_rows(file_path.name, verdict))

        print(f"   ✅ Готово: {file_path.name}")

    return report_rows, file_texts, file_paths, verdicts_by_file, error_rows


def save_report(report_rows, duplicate_rows, version_rows, error_rows, output_path):

    output_path = Path(output_path)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

        if report_rows:
            df = pd.DataFrame(report_rows, columns=["Файл", "Критерий", "Статус", "Находки", "Доп. детали"])
        else:
            df = pd.DataFrame(columns=["Файл", "Критерий", "Статус", "Находки", "Доп. детали"])
        df.to_excel(writer, sheet_name="Сводка", index=False)

        if duplicate_rows:
            df_dup = pd.DataFrame(duplicate_rows, columns=["file_a", "file_b", "similarity", "note"])
            df_dup.columns = ["Файл A", "Файл B", "Схожесть (0-1)", "Комментарий"]
        else:
            df_dup = pd.DataFrame(columns=["Файл A", "Файл B", "Схожесть (0-1)", "Комментарий"])
        df_dup.to_excel(writer, sheet_name="Дубли между файлами", index=False)

        version_cols = ["Группа файлов", "Файл", "Вердикт", "Версия/дата из текста",
                        "Маркеры устаревания в тексте", "Дата изменения файла"]
        if version_rows:
            df_ver = pd.DataFrame(version_rows, columns=version_cols)
        else:
            df_ver = pd.DataFrame(columns=version_cols)
        df_ver.to_excel(writer, sheet_name="Версии документов", index=False)

        if error_rows:
            df_err = pd.DataFrame(error_rows, columns=["Файл", "Ошибка"])
        else:
            df_err = pd.DataFrame(columns=["Файл", "Ошибка"])
        df_err.to_excel(writer, sheet_name="Ошибки", index=False)

    print(f"\n✅ Отчёт сохранён: {output_path}")


# ============================================================
# ОСНОВНАЯ ФУНКЦИЯ
# ============================================================

def main():
    parser = argparse.ArgumentParser(description="Валидация документов по требованиям")
    parser.add_argument("--docs", default="./documents", help="Папка с документами")
    parser.add_argument("--output", default="./Отчет_валидации.xlsx", help="Файл отчёта")
    parser.add_argument("--schema", default="schema.json", help="JSON-схема для structured output")
    parser.add_argument("--model", default=None, help="Название модели (или переменная окружения MODEL)")
    parser.add_argument("--api-url", default=None, help="URL LLM API (или переменная окружения API_URL)")
    parser.add_argument("--dup-threshold", type=float, default=None,
                        help="Порог схожести текста для межфайловых дублей, 0-1 (по умолчанию 0.85)")
    parser.add_argument("--version-threshold", type=float, default=0.5,
                        help="Порог схожести текста для группировки версий одного документа, 0-1 (по умолчанию 0.5)")
    parser.add_argument("--skip-cross-file-duplicates", action="store_true",
                        help="Не искать дубли и версии между файлами")
    parser.add_argument("--test-mode", action="store_true", help="Тестовый режим без реальных вызовов LLM")

    args = parser.parse_args()

    Config.INPUT_DOCS_DIR = Path(args.docs)
    Config.OUTPUT_REPORT = Path(args.output)
    Config.TEST_MODE = args.test_mode

    if args.model:
        Config.MODEL = args.model
    if args.api_url:
        Config.API_URL = args.api_url
    if args.dup_threshold is not None:
        Config.DUPLICATE_SIMILARITY_THRESHOLD = args.dup_threshold

    global SCHEMA
    SCHEMA = load_schema(args.schema)

    if not Config.TEST_MODE and SCHEMA is None:
        print("❌ JSON-схема не загружена.")
        return

    if not Config.TEST_MODE and not Config.API_URL:
        print("❌ Не задан API_URL (флаг --api-url или переменная окружения API_URL).")
        return

    if not Config.TEST_MODE and not Config.MODEL:
        print("❌ Не задана модель (флаг --model или переменная окружения MODEL).")
        return

    print("\n" + "=" * 80)
    print("📄 ШАГ 1. АНАЛИЗ ДОКУМЕНТОВ")
    print("=" * 80)

    report_rows, file_texts, file_paths, verdicts_by_file, error_rows = process_all_documents(
        Config.INPUT_DOCS_DIR
    )

    duplicate_rows = []
    version_rows = []

    if not args.skip_cross_file_duplicates:

        print("\n" + "=" * 80)
        print("🔁 ШАГ 2. ПОИСК ДУБЛЕЙ И ВЕРСИЙ МЕЖДУ ФАЙЛАМИ")
        print("=" * 80)

        pairs = compute_similarity_pairs(file_texts)

        duplicate_rows = find_cross_file_duplicates(file_texts, pairs=pairs)
        print(f"   Найдено потенциальных дублей/почти-дублей: {len(duplicate_rows)}")

        version_groups = find_version_groups(
            file_paths, pairs, threshold_low=args.version_threshold
        )
        version_rows = build_version_group_rows(version_groups, file_paths, verdicts_by_file)
        print(f"   Найдено групп возможных версий одного документа: {len(version_groups)}")


        files_in_version_groups = {name: group for group in version_groups for name in group}

        for row in report_rows:
            if row["Критерий"].startswith("2.4") and row["Файл"] in files_in_version_groups:
                group = files_in_version_groups[row["Файл"]]
                others = [f for f in group if f != row["Файл"]]
                note = (
                    f"В папке documents найдены другие файлы, похожие на этот документ "
                    f"(вероятно, другие версии/копии): {', '.join(others)}. По правилам "
                    f"в базе знаний должна храниться только 1 актуальная версия, "
                    f"остальные должны быть архивированы в другом месте (см. лист "
                    f"'Версии документов')."
                )
                row["Статус"] = "Есть нарушения"
                row["Находки"] = (
                    (row["Находки"] + "; " if row["Находки"] and row["Находки"] != "—" else "")
                    + note
                )

    print("\n" + "=" * 80)
    print("💾 ШАГ 3. СОХРАНЕНИЕ ОТЧЁТА")
    print("=" * 80)

    save_report(report_rows, duplicate_rows, version_rows, error_rows, Config.OUTPUT_REPORT)

    print("\n" + "=" * 80)
    print("🏁 ГОТОВО")
    print("=" * 80)
    print(f"📄 Обработано файлов: {len(file_texts)}")
    print(f"❌ Ошибок обработки: {len(error_rows)}")
    print(f"🔁 Найдено дублей/похожих пар: {len(duplicate_rows)}")
    print(f"🗂️ Найдено групп версий одного документа: {len(version_rows) and len(set(r['Группа файлов'] for r in version_rows))}")


if __name__ == "__main__":
    main()