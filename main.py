# -*- coding: utf-8 -*-
import os
import csv
import json
import re
import sqlite3
import sys
import time
from itertools import islice

APP_TITLE = "Prom Generator"
from copy import deepcopy
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from formula_engine import FormulaEngine, FormulaError


def _show_dependency_error(message: str) -> None:
    """Display a blocking error for a missing runtime dependency."""

    root = None
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(APP_TITLE, message)
    except Exception:
        print(message, file=sys.stderr)
    finally:
        if root is not None:
            try:
                root.destroy()
            except Exception:
                pass


DEPENDENCY_WARNINGS = []

CSV_JSON_FALLBACK_NOTE = "\nЕкспорт у CSV (.csv) та JSON (.json) залишається доступним."
PANDAS_INSTALL_HINT = (
    "Встановіть бібліотеку командою 'pip install pandas' і перезапустіть застосунок, щоб увімкнути цей формат."
)
OPENPYXL_INSTALL_HINT = (
    "Встановіть бібліотеку командою 'pip install openpyxl' і перезапустіть застосунок, щоб увімкнути цей формат."
)

try:
    import customtkinter as ctk
except ModuleNotFoundError:
    _show_dependency_error(
        "Бібліотека CustomTkinter не знайдена.\n"
        "Встановіть її командою 'pip install customtkinter' і перезапустіть застосунок."
    )
    sys.exit(1)

PANDAS_IMPORT_ERROR_DETAIL = ""
PANDAS_EXPORT_BLOCKED_MESSAGE = ""
EXCEL_ENGINE_IMPORT_ERROR_DETAIL = ""
EXCEL_EXPORT_BLOCKED_MESSAGE = ""

try:
    import pandas as pd
except ModuleNotFoundError as exc:
    pd = None
    PANDAS_IMPORT_ERROR_DETAIL = str(exc)
    PANDAS_EXPORT_BLOCKED_MESSAGE = (
        "Експорт у формат Excel (.xlsx) недоступний: бібліотека pandas не встановлена."
    )
except ImportError as exc:  # e.g. missing binary dependencies
    pd = None
    PANDAS_IMPORT_ERROR_DETAIL = str(exc)
    PANDAS_EXPORT_BLOCKED_MESSAGE = (
        "Експорт у формат Excel (.xlsx) недоступний: не вдалося завантажити бібліотеку pandas."
    )
else:
    PANDAS_IMPORT_ERROR_DETAIL = ""
    PANDAS_EXPORT_BLOCKED_MESSAGE = ""

if PANDAS_EXPORT_BLOCKED_MESSAGE:
    detail_suffix = f"\nДеталі: {PANDAS_IMPORT_ERROR_DETAIL}" if PANDAS_IMPORT_ERROR_DETAIL else ""
    DEPENDENCY_WARNINGS.append(
        PANDAS_EXPORT_BLOCKED_MESSAGE
        + detail_suffix
        + "\n"
        + PANDAS_INSTALL_HINT
        + CSV_JSON_FALLBACK_NOTE
    )
else:
    PANDAS_IMPORT_ERROR_DETAIL = ""

if pd is None:
    EXCEL_EXPORT_BLOCKED_MESSAGE = PANDAS_EXPORT_BLOCKED_MESSAGE
else:
    try:
        import openpyxl  # noqa: F401  # type: ignore[import-not-found]
    except ModuleNotFoundError as exc:
        EXCEL_ENGINE_IMPORT_ERROR_DETAIL = str(exc)
        EXCEL_EXPORT_BLOCKED_MESSAGE = (
            "Експорт у формат Excel (.xlsx) недоступний: бібліотека openpyxl не встановлена."
        )
    except ImportError as exc:
        EXCEL_ENGINE_IMPORT_ERROR_DETAIL = str(exc)
        EXCEL_EXPORT_BLOCKED_MESSAGE = (
            "Експорт у формат Excel (.xlsx) недоступний: не вдалося завантажити бібліотеку openpyxl."
        )

if EXCEL_EXPORT_BLOCKED_MESSAGE and pd is not None:
    detail_suffix = (
        f"\nДеталі: {EXCEL_ENGINE_IMPORT_ERROR_DETAIL}"
        if EXCEL_ENGINE_IMPORT_ERROR_DETAIL
        else ""
    )
    DEPENDENCY_WARNINGS.append(
        EXCEL_EXPORT_BLOCKED_MESSAGE
        + detail_suffix
        + "\n"
        + OPENPYXL_INSTALL_HINT
        + CSV_JSON_FALLBACK_NOTE
    )
else:
    EXCEL_ENGINE_IMPORT_ERROR_DETAIL = ""

try:
    from jinja2 import Template, TemplateError
except ModuleNotFoundError:
    _show_dependency_error(
        "Бібліотека Jinja2 не знайдена.\n"
        "Встановіть її командою 'pip install jinja2' і перезапустіть застосунок."
    )
    sys.exit(1)

DB_FILE = "catalog.db"
TEMPLATES_FILE = "templates.json"
EXPORT_FIELDS_FILE = "export_fields.json"
TITLE_TAGS_FILE = "title_tags_templates.json"
FILM_TYPE_DEFAULT_LABEL = "Універсальний шаблон"

# ============================ ШАБЛОНИ / НАЛАШТУВАННЯ ============================

DEFAULT_TEMPLATES = {
    "title_template": "Гідрогелева плівка {{ film_type }} {{ brand }} {{ model }}",
    "tags_template": "{{ brand }} {{ model }}, плівка {{ brand }} {{ model }}, hydrogel film {{ brand }} {{ model }}, {{ film_type }} {{ brand }} {{ model }}",
    # Окремі описи для категорій і типів плівок (fallback -> "default")
    "descriptions": {
        "Смартфони": {
            "прозора": "Прозора плівка для {{ brand }} {{ model }} — базовий прозорий захист, висока чутливість та легка поклейка.",
            "матова": "Матова плівка для {{ brand }} {{ model }} — мінімум відблисків, комфорт на сонці, приємний тактильний ефект.",
            "anti-blue": "Anti-Blue плівка для {{ brand }} {{ model }} — фільтрація синього світла для зниження втоми очей.",
            "privacy clear": "Privacy Clear для {{ brand }} {{ model }} — захист від підглядання під прямим кутом, прозора фронтальна видимість.",
            "privacy mate": "Privacy Mate для {{ brand }} {{ model }} — матова з приватністю, менше відблисків і захист від бічних кутів огляду.",
            "default": "Універсальна плівка для {{ brand }} {{ model }} — захист від подряпин та відбитків."
        },
        "Планшети": {
            "прозора": "Прозора плівка для планшета {{ brand }} {{ model }} — чиста картинка на великому екрані, проста поклейка.",
            "матова": "Матова плівка для планшета {{ brand }} {{ model }} — мінімум відблисків, комфорт для роботи/навчання.",
            "anti-blue": "Anti-Blue для планшета {{ brand }} {{ model }} — зниження синього спектру, довша робота без втоми очей.",
            "default": "Універсальна плівка для планшета {{ brand }} {{ model }} — збалансований захист поверхні."
        }
    },
    # Типи плівок (можна вмикати/вимикати в коді при потребі)
    "film_types": [
        {"name": "прозора", "enabled": True},
        {"name": "матова", "enabled": True},
        {"name": "privacy clear", "enabled": True},
        {"name": "privacy mate", "enabled": True},
        {"name": "anti-blue", "enabled": True}
    ]
}

DEFAULT_EXPORT_FIELDS = [
    {"field": "Категорія", "enabled": True, "template": "{{ category }}"},
    {"field": "Бренд", "enabled": True, "template": "{{ brand }}"},
    {"field": "Модель", "enabled": True, "template": "{{ model }}"},
    {"field": "Тип_плівки", "enabled": True, "template": "{{ film_type }}"},
    {"field": "Назва_позиції", "enabled": True, "template": "{{ title }}"},
    {"field": "Назва_позиції_укр", "enabled": False, "template": "{{ title }}"},
    {"field": "Пошукові_запити", "enabled": True, "template": "{{ tags }}"},
    {"field": "Пошукові_запити_укр", "enabled": False, "template": "{{ tags }}"},
    {"field": "Опис", "enabled": True, "template": "{{ description }}"},
    {"field": "Опис_укр", "enabled": False, "template": "{{ description }}"},
    {"field": "Код_товару", "enabled": False, "template": "{{ spec('Код_товару') }}"},
    {"field": "Тип_товару", "enabled": False, "template": "{{ film_type }}"},
    {"field": "Ціна", "enabled": False, "template": ""},
    {"field": "Валюта", "enabled": False, "template": ""},
    {"field": "Одиниця_виміру", "enabled": False, "template": ""},
    {"field": "Мінімальний_обсяг_замовлення", "enabled": False, "template": ""},
    {"field": "Оптова_ціна", "enabled": False, "template": ""},
    {"field": "Мінімальне_замовлення_опт", "enabled": False, "template": ""},
    {"field": "Посилання_зображення", "enabled": False, "template": ""},
    {"field": "Наявність", "enabled": False, "template": ""},
    {"field": "Кількість", "enabled": False, "template": ""},
    {"field": "Номер_групи", "enabled": False, "template": "{{ category_id }}"},
    {"field": "Назва_групи", "enabled": False, "template": "{{ category }}"},
    {"field": "Посилання_підрозділу", "enabled": False, "template": ""},
    {"field": "Можливість_поставки", "enabled": False, "template": ""},
    {"field": "Термін_поставки", "enabled": False, "template": ""},
    {"field": "Спосіб_пакування", "enabled": False, "template": ""},
    {"field": "Спосіб_пакування_укр", "enabled": False, "template": ""},
    {"field": "Унікальний_ідентифікатор", "enabled": False, "template": ""},
    {"field": "Ідентифікатор_товару", "enabled": False, "template": "{{ model_id }}"},
    {"field": "Ідентифікатор_підрозділу", "enabled": False, "template": ""},
    {"field": "Ідентифікатор_групи", "enabled": False, "template": "{{ category_id }}"},
    {"field": "Виробник", "enabled": False, "template": "{{ brand }}"},
    {"field": "Країна_виробник", "enabled": False, "template": ""},
    {"field": "Знижка", "enabled": False, "template": ""},
    {"field": "ID_групи_різновидів", "enabled": False, "template": ""},
    {"field": "Особисті_нотатки", "enabled": False, "template": ""},
    {"field": "Продукт_на_сайті", "enabled": False, "template": ""},
    {"field": "Термін_дії_знижки_від", "enabled": False, "template": ""},
    {"field": "Термін_дії_знижки_до", "enabled": False, "template": ""},
    {"field": "Ціна_від", "enabled": False, "template": ""},
    {"field": "Ярлик", "enabled": False, "template": ""},
    {"field": "HTML_заголовок", "enabled": False, "template": "{{ title }}"},
    {"field": "HTML_заголовок_укр", "enabled": False, "template": "{{ title }}"},
    {"field": "HTML_опис", "enabled": False, "template": "{{ description }}"},
    {"field": "HTML_опис_укр", "enabled": False, "template": "{{ description }}"},
    {"field": "Код_маркування_(GTIN)", "enabled": False, "template": ""},
    {"field": "Номер_пристрою_(MPN)", "enabled": False, "template": ""},
    {"field": "Вага,кг", "enabled": False, "template": "{{ spec('Вага, кг') }}"},
    {"field": "Ширина,см", "enabled": False, "template": "{{ spec('Ширина, см') }}"},
    {"field": "Висота,см", "enabled": False, "template": "{{ spec('Висота, см') }}"},
    {"field": "Довжина,см", "enabled": False, "template": "{{ spec('Довжина, см') }}"},
    {"field": "Де_знаходиться_товар", "enabled": False, "template": ""},
    {"field": "Назва_Характеристики", "enabled": False, "template": "{{ spec_items | map(attribute=0) | join('; ') }}"},
    {"field": "Одиниця_виміру,_Характеристики", "enabled": False, "template": ""},
    {"field": "Значення_Характеристики", "enabled": False, "template": "{{ spec_items | map(attribute=1) | join('; ') }}"},
]

EXCEL_FORMAT_LABEL = "Excel (.xlsx)"
CSV_FORMAT_LABEL = "CSV (.csv)"
JSON_FORMAT_LABEL = "JSON (.json)"
EXPORT_FORMAT_OPTIONS = (EXCEL_FORMAT_LABEL, CSV_FORMAT_LABEL, JSON_FORMAT_LABEL)


def get_available_export_formats():
    if pd is None or EXCEL_EXPORT_BLOCKED_MESSAGE:
        return [fmt for fmt in EXPORT_FORMAT_OPTIONS if fmt != EXCEL_FORMAT_LABEL]
    return list(EXPORT_FORMAT_OPTIONS)

def _title_tags_block(title: str, tags: str) -> dict:
    return {
        "title_template": title,
        "tags_template": tags,
    }

def _build_title_tags_defaults(film_type_names, base_title, base_tags):
    base_block = _title_tags_block(base_title, base_tags)
    data = {"default": base_block.copy()}
    for name in film_type_names:
        data[name] = base_block.copy()
    return data

def load_templates():
    if not os.path.exists(TEMPLATES_FILE):
        defaults = deepcopy(DEFAULT_TEMPLATES)
        save_templates(defaults)
        return defaults

    try:
        with open(TEMPLATES_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, json.JSONDecodeError):
        defaults = deepcopy(DEFAULT_TEMPLATES)
        save_templates(defaults)
        return defaults

    if not isinstance(data, dict):
        defaults = deepcopy(DEFAULT_TEMPLATES)
        save_templates(defaults)
        return defaults

    # гарантуємо всі ключі
    for k, v in DEFAULT_TEMPLATES.items():
        if k not in data:
            data[k] = deepcopy(v)
    return data

def save_templates(dct):
    with open(TEMPLATES_FILE, "w", encoding="utf-8") as f:
        json.dump(dct, f, ensure_ascii=False, indent=2)

def load_title_tags_templates(templates: dict):
    film_type_names = [item.get("name") for item in templates.get("film_types", []) if item.get("name")]
    base_title = templates.get("title_template", DEFAULT_TEMPLATES["title_template"])
    base_tags = templates.get("tags_template", DEFAULT_TEMPLATES["tags_template"])
    defaults = _build_title_tags_defaults(film_type_names, base_title, base_tags)

    if not os.path.exists(TITLE_TAGS_FILE):
        save_title_tags_templates(defaults)
        return defaults

    with open(TITLE_TAGS_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)

    changed = False

    if "default" not in data or not isinstance(data["default"], dict):
        data["default"] = defaults["default"].copy()
        changed = True

    for key in ("title_template", "tags_template"):
        if key not in data["default"]:
            data["default"][key] = defaults["default"][key]
            changed = True

    for name in film_type_names:
        block = data.get(name)
        if not isinstance(block, dict):
            data[name] = defaults["default"].copy()
            changed = True
            continue
        for key in ("title_template", "tags_template"):
            if key not in block:
                block[key] = data["default"].get(key, defaults["default"][key])
                changed = True

    if changed:
        save_title_tags_templates(data)

    return data

def save_title_tags_templates(dct):
    with open(TITLE_TAGS_FILE, "w", encoding="utf-8") as f:
        json.dump(dct, f, ensure_ascii=False, indent=2)

def resolve_title_tags(title_tags_templates: dict, templates: dict, film_type: str):
    fallback_title = templates.get("title_template", DEFAULT_TEMPLATES["title_template"])
    fallback_tags = templates.get("tags_template", DEFAULT_TEMPLATES["tags_template"])
    default_block = title_tags_templates.get("default", {})
    film_block = title_tags_templates.get(film_type, {})
    title_template = (
        film_block.get("title_template")
        or default_block.get("title_template")
        or fallback_title
    )
    tags_template = (
        film_block.get("tags_template")
        or default_block.get("tags_template")
        or fallback_tags
    )
    return title_template, tags_template


def load_export_fields():
    if not os.path.exists(EXPORT_FIELDS_FILE):
        fields = [field.copy() for field in DEFAULT_EXPORT_FIELDS]
        save_export_fields(fields)
        return fields

    try:
        with open(EXPORT_FIELDS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, json.JSONDecodeError):
        fields = [field.copy() for field in DEFAULT_EXPORT_FIELDS]
        save_export_fields(fields)
        return fields

    if not isinstance(data, list):
        fields = [field.copy() for field in DEFAULT_EXPORT_FIELDS]
        save_export_fields(fields)
        return fields

    normalized = []
    changed = False
    for item in data:
        if not isinstance(item, dict):
            changed = True
            continue
        field_name = item.get("field") or item.get("name") or item.get("key")
        if field_name is None:
            changed = True
            continue
        field_name = str(field_name).strip()
        if not field_name:
            changed = True
            continue
        template = item.get("template", "")
        if template is None:
            template = ""
        template = str(template)
        enabled = bool(item.get("enabled", False))
        normalized.append({"field": field_name, "template": template, "enabled": enabled})
        if (
            field_name != item.get("field")
            or template != (item.get("template", "") or "")
            or enabled != bool(item.get("enabled", False))
        ):
            changed = True

    if not normalized:
        normalized = [field.copy() for field in DEFAULT_EXPORT_FIELDS]
        save_export_fields(normalized)
        return normalized

    if changed:
        save_export_fields(normalized)

    return normalized


def save_export_fields(fields: list):
    sanitized = []
    for item in fields:
        if not isinstance(item, dict):
            continue
        field_name = item.get("field") or item.get("name") or item.get("key")
        if field_name is None:
            continue
        field_name = str(field_name).strip()
        if not field_name:
            continue
        template = item.get("template", "")
        if template is None:
            template = ""
        template = str(template)
        enabled = bool(item.get("enabled", False))
        sanitized.append({"field": field_name, "template": template, "enabled": enabled})

    with open(EXPORT_FIELDS_FILE, "w", encoding="utf-8") as f:
        json.dump(sanitized, f, ensure_ascii=False, indent=2)

    return sanitized

# ============================ УТИЛІТИ ============================

def ensure_folder(path: str):
    if path and not os.path.exists(path):
        os.makedirs(path, exist_ok=True)


_INPUT_SPLIT_RE = re.compile(r"[\n\r,;\u201a\u201e\uFF0C\u3001]+")


def split_catalog_input(raw: str):
    if not raw:
        return []
    parts = [part.strip() for part in _INPUT_SPLIT_RE.split(raw) if part.strip()]
    unique = []
    seen = set()
    for part in parts:
        if part not in seen:
            unique.append(part)
            seen.add(part)
    return unique

  
def create_inline_entry(parent, text: str):
    entry = tk.Entry(parent)
    font = ctk.CTkFont()
    entry.configure(font=font)
    entry._ctk_font = font  # keep reference to avoid garbage collection
    mode = (ctk.get_appearance_mode() or "light").lower()
    if mode == "dark":
        bg = "#2b2b2b"
        fg = "#f2f2f2"
        border = "#565b5e"
    else:
        bg = "#ffffff"
        fg = "#1f1f1f"
        border = "#a5a5a5"
    entry.configure(
        background=bg,
        foreground=fg,
        insertbackground=fg,
        selectbackground="#1f6aa5",
        selectforeground=fg,
        highlightthickness=1,
        highlightbackground=border,
        highlightcolor="#1f6aa5",
        borderwidth=0,
        relief="flat",
    )
    entry.insert(0, text)
    entry.select_range(0, tk.END)
    entry.focus_set()
    return entry

def show_error(msg: str):
    messagebox.showerror("Помилка", msg)

def show_info(msg: str):
    messagebox.showinfo("Інформація", msg)

# ============================ БАЗА ДАНИХ ============================

def db_connect():
    conn = sqlite3.connect(DB_FILE)
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

def init_db():
    conn = db_connect()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS categories(
            id   INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS brands(
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            category_id INTEGER NOT NULL,
            name        TEXT NOT NULL,
            UNIQUE(category_id, name),
            FOREIGN KEY(category_id) REFERENCES categories(id) ON DELETE CASCADE
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS models(
            id       INTEGER PRIMARY KEY AUTOINCREMENT,
            brand_id INTEGER NOT NULL,
            name     TEXT NOT NULL,
            UNIQUE(brand_id, name),
            FOREIGN KEY(brand_id) REFERENCES brands(id) ON DELETE CASCADE
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS model_specs(
            id       INTEGER PRIMARY KEY AUTOINCREMENT,
            model_id INTEGER NOT NULL,
            key      TEXT NOT NULL,
            value    TEXT,
            FOREIGN KEY(model_id) REFERENCES models(id) ON DELETE CASCADE
        )
    """)
    conn.commit()

    # Стартові категорії, якщо порожньо
    cur.execute("SELECT COUNT(*) FROM categories")
    if cur.fetchone()[0] == 0:
        cur.executemany("INSERT INTO categories(name) VALUES(?)",
                        [("Смартфони",), ("Планшети",)])
        conn.commit()
    conn.close()

# ---- CRUD helpers

def _trimmed_rows(rows):
    """Повертає список записів зі зрізаними пробілами на початку/в кінці назв."""
    trimmed = []
    for row in rows:
        if len(row) == 2:
            idx, name = row
            if isinstance(name, str):
                name = name.strip()
            trimmed.append((idx, name))
        else:
            trimmed.append(row)
    return trimmed


def get_categories():
    conn = db_connect(); cur = conn.cursor()
    cur.execute("SELECT id, name FROM categories ORDER BY name")
    rows = cur.fetchall(); conn.close()
    return _trimmed_rows(rows)

def add_category(name: str):
    name = name.strip()
    if not name: return
    conn = db_connect(); cur = conn.cursor()
    cur.execute("INSERT OR IGNORE INTO categories(name) VALUES(?)", (name,))
    conn.commit(); conn.close()

def rename_category(cat_id: int, new_name: str):
    new_name = new_name.strip()
    if not new_name:
        return False
    conn = db_connect(); cur = conn.cursor()
    try:
        cur.execute("UPDATE categories SET name=? WHERE id=?", (new_name, cat_id))
        conn.commit()
        return True
    except sqlite3.IntegrityError as exc:
        conn.rollback()
        return exc
    finally:
        conn.close()

def delete_category(cat_id: int):
    conn = db_connect(); cur = conn.cursor()
    cur.execute("DELETE FROM categories WHERE id=?", (cat_id,))
    conn.commit(); conn.close()

def get_brands(category_id: int):
    conn = db_connect(); cur = conn.cursor()
    cur.execute("SELECT id, name FROM brands WHERE category_id=? ORDER BY name", (category_id,))
    rows = cur.fetchall(); conn.close()
    return _trimmed_rows(rows)

def add_brand(category_id: int, name: str):
    name = name.strip()
    if not name: return
    conn = db_connect(); cur = conn.cursor()
    cur.execute("INSERT OR IGNORE INTO brands(category_id, name) VALUES(?,?)", (category_id, name))
    conn.commit(); conn.close()

def rename_brand(brand_id: int, new_name: str):
    new_name = new_name.strip()
    if not new_name:
        return False
    conn = db_connect(); cur = conn.cursor()
    try:
        cur.execute("UPDATE brands SET name=? WHERE id=?", (new_name, brand_id))
        conn.commit()
        return True
    except sqlite3.IntegrityError as exc:
        conn.rollback()
        return exc
    finally:
        conn.close()

def delete_brand(brand_id: int):
    conn = db_connect(); cur = conn.cursor()
    cur.execute("DELETE FROM brands WHERE id=?", (brand_id,))
    conn.commit(); conn.close()

def get_models(brand_id: int):
    conn = db_connect(); cur = conn.cursor()
    cur.execute("SELECT id, name FROM models WHERE brand_id=? ORDER BY name", (brand_id,))
    rows = cur.fetchall(); conn.close()
    return _trimmed_rows(rows)

def add_model(brand_id: int, name: str):
    name = name.strip()
    if not name: return
    conn = db_connect(); cur = conn.cursor()
    cur.execute("INSERT OR IGNORE INTO models(brand_id, name) VALUES(?,?)", (brand_id, name))
    conn.commit(); conn.close()

def rename_model(model_id: int, new_name: str):
    new_name = new_name.strip()
    if not new_name:
        return False
    conn = db_connect(); cur = conn.cursor()
    try:
        cur.execute("UPDATE models SET name=? WHERE id=?", (new_name, model_id))
        conn.commit()
        return True
    except sqlite3.IntegrityError as exc:
        conn.rollback()
        return exc
    finally:
        conn.close()

def delete_model(model_id: int):
    conn = db_connect(); cur = conn.cursor()
    cur.execute("DELETE FROM models WHERE id=?", (model_id,))
    conn.commit(); conn.close()

# ---- Specs (key-value)

def get_specs(model_id: int):
    conn = db_connect(); cur = conn.cursor()
    cur.execute("SELECT id, key, value FROM model_specs WHERE model_id=? ORDER BY id", (model_id,))
    rows = cur.fetchall(); conn.close()
    return rows

def insert_spec(model_id: int, key: str, value: str):
    key = key.strip()
    if not key: return
    conn = db_connect(); cur = conn.cursor()
    cur.execute("INSERT INTO model_specs(model_id, key, value) VALUES(?,?,?)",
                (model_id, key, value))
    conn.commit(); conn.close()

def update_spec(spec_id: int, key: str, value: str):
    key = key.strip()
    if not key: return
    conn = db_connect(); cur = conn.cursor()
    cur.execute("UPDATE model_specs SET key=?, value=? WHERE id=?",
                (key, value, spec_id))
    conn.commit(); conn.close()

def delete_spec(spec_id: int):
    conn = db_connect(); cur = conn.cursor()
    cur.execute("DELETE FROM model_specs WHERE id=?", (spec_id,))
    conn.commit(); conn.close()

def load_specs_map(model_ids):
    if not model_ids:
        return {}
    unique_ids = []
    seen = set()
    for mid in model_ids:
        try:
            ivalue = int(mid)
        except (TypeError, ValueError):
            continue
        if ivalue in seen:
            continue
        seen.add(ivalue)
        unique_ids.append(ivalue)
    if not unique_ids:
        return {}

    placeholders = ",".join(["?"] * len(unique_ids))
    query = f"""
        SELECT model_id, key, value
        FROM model_specs
        WHERE model_id IN ({placeholders})
        ORDER BY model_id, id
    """
    conn = db_connect(); cur = conn.cursor()
    cur.execute(query, tuple(unique_ids))
    rows = cur.fetchall(); conn.close()

    specs_map = {}
    for model_id, key, value in rows:
        if isinstance(key, str):
            key = key.strip()
        if isinstance(value, str):
            value = value.strip()
        if not key:
            continue
        specs_map.setdefault(model_id, {})[key] = value
    return specs_map

# ============================ ФОРМУЛИ ============================

_FORMULA_PREFIX_RE = re.compile(r"^\s*=")
_IDENTIFIER_SANITIZE_RE = re.compile(r"[^\w]+", re.UNICODE)
_ASCII_SANITIZE_RE = re.compile(r"[^0-9a-z]+")

_TRANSLIT_TABLE = {
    "а": "a",
    "б": "b",
    "в": "v",
    "г": "h",
    "ґ": "g",
    "д": "d",
    "е": "e",
    "є": "ie",
    "ж": "zh",
    "з": "z",
    "и": "y",
    "і": "i",
    "ї": "yi",
    "й": "y",
    "к": "k",
    "л": "l",
    "м": "m",
    "н": "n",
    "о": "o",
    "п": "p",
    "р": "r",
    "с": "s",
    "т": "t",
    "у": "u",
    "ф": "f",
    "х": "kh",
    "ц": "ts",
    "ч": "ch",
    "ш": "sh",
    "щ": "shch",
    "ь": "",
    "ю": "yu",
    "я": "ya",
    "ъ": "",
    "ы": "y",
    "э": "e",
    "ё": "yo",
}

_COMMON_SPEC_ALIAS = {
    "color": "color",
    "colour": "color",
    "kolir": "color",
    "колір": "color",
    "цвет": "color",
    "brand": "brand",
    "бренд": "brand",
    "weight": "weight",
    "вага": "weight",
    "вес": "weight",
    "material": "material",
    "матеріал": "material",
    "материал": "material",
    "thickness": "thickness",
    "товщина": "thickness",
    "толщина": "thickness",
    "sku": "sku",
    "код": "sku",
    "код_товару": "sku",
}


def _looks_like_formula(text: str) -> bool:
    return bool(text) and bool(_FORMULA_PREFIX_RE.match(text))


def _normalize_identifier(value: str) -> str:
    lowered = value.strip().lower()
    cleaned = _IDENTIFIER_SANITIZE_RE.sub("_", lowered)
    cleaned = re.sub(r"_+", "_", cleaned).strip("_")
    return cleaned


def _transliterate_ascii(value: str) -> str:
    result = []
    for ch in value.lower():
        if ch.isdigit():
            result.append(ch)
            continue
        if "a" <= ch <= "z":
            result.append(ch)
            continue
        if ch in _TRANSLIT_TABLE:
            result.append(_TRANSLIT_TABLE[ch])
            continue
        if ch == "_" or ch.isspace():
            result.append("_")
        else:
            result.append("")
    ascii_candidate = "".join(result)
    ascii_candidate = _ASCII_SANITIZE_RE.sub("_", ascii_candidate)
    ascii_candidate = re.sub(r"_+", "_", ascii_candidate).strip("_")
    return ascii_candidate


def _build_formula_context(base_context):
    formula_context = {}
    for key, value in base_context.items():
        if callable(value):
            continue
        formula_context[key] = value

    if "brand" in base_context and "attr_brand" not in formula_context:
        formula_context["attr_brand"] = base_context.get("brand")
    if "model" in base_context and "attr_model" not in formula_context:
        formula_context["attr_model"] = base_context.get("model")

    specs = base_context.get("specs")
    if isinstance(specs, dict):
        for spec_key, spec_value in specs.items():
            if spec_value is None:
                continue
            key_str = str(spec_key)
            normalized = _normalize_identifier(key_str)
            ascii_name = _transliterate_ascii(key_str)
            for candidate in (normalized, ascii_name):
                if candidate:
                    formula_context.setdefault(f"attr_{candidate}", spec_value)
            alias_source = None
            if ascii_name and ascii_name in _COMMON_SPEC_ALIAS:
                alias_source = _COMMON_SPEC_ALIAS[ascii_name]
            elif normalized and normalized in _COMMON_SPEC_ALIAS:
                alias_source = _COMMON_SPEC_ALIAS[normalized]
            if alias_source:
                formula_context.setdefault(f"attr_{alias_source}", spec_value)

    category_name = base_context.get("category")
    if isinstance(category_name, str):
        cat_slug = _transliterate_ascii(category_name) or _normalize_identifier(category_name)
        if cat_slug:
            formula_context.setdefault(f"category_{cat_slug}", category_name)

    film_type = base_context.get("film_type")
    if isinstance(film_type, str):
        ft_slug = _transliterate_ascii(film_type) or _normalize_identifier(film_type)
        if ft_slug:
            formula_context.setdefault(f"film_type_{ft_slug}", film_type)

    return formula_context

# ============================ ГЕНЕРАЦІЯ ============================

def _normalize_id_list(ids):
    if not ids:
        return []
    if isinstance(ids, (list, tuple, set)):
        result = []
        for value in ids:
            try:
                ivalue = int(value)
            except (TypeError, ValueError):
                continue
            if ivalue:
                result.append(ivalue)
        return result
    try:
        ivalue = int(ids)
    except (TypeError, ValueError):
        return []
    return [ivalue] if ivalue else []


def collect_models(category_ids=None, brand_ids=None, model_ids=None):
    """Повертає список (brand_name, model_name, category_name, model_id) з урахуванням множинних фільтрів."""
    category_ids = _normalize_id_list(category_ids)
    brand_ids = _normalize_id_list(brand_ids)
    model_ids = _normalize_id_list(model_ids)

    
    conn = db_connect(); cur = conn.cursor()
    query = """
        SELECT b.name, m.name, c.name, m.id, b.id, c.id
        FROM models m
        JOIN brands b ON m.brand_id = b.id
        JOIN categories c ON b.category_id = c.id
    """
    conditions = []
    params = []
    if category_ids:
        placeholders = ",".join(["?"] * len(category_ids))
        conditions.append(f"c.id IN ({placeholders})")
        params.extend(category_ids)
    if brand_ids:
        placeholders = ",".join(["?"] * len(brand_ids))
        conditions.append(f"b.id IN ({placeholders})")
        params.extend(brand_ids)
    if model_ids:
        placeholders = ",".join(["?"] * len(model_ids))
        conditions.append(f"m.id IN ({placeholders})")
        params.extend(model_ids)
    if conditions:
        query += " WHERE " + " AND ".join(conditions)
    query += " ORDER BY c.name, b.name, m.name"
    cur.execute(query, tuple(params))
    rows = cur.fetchall(); conn.close()
    cleaned = []
    for brand, model, cat, mid, brand_id, cat_id in rows:
        if isinstance(brand, str):
            brand = brand.strip()
        if isinstance(model, str):
            model = model.strip()
        if isinstance(cat, str):
            cat = cat.strip()
        cleaned.append((brand, model, cat, mid, brand_id, cat_id))
    return cleaned
def generate_export_rows(
    film_types: list,
    templates: dict,
    title_tags_templates: dict,
    export_fields: list,
    category_ids=None,
    brand_ids=None,
    model_ids=None,
    progress_callback=None,
):
    film_types = list(film_types)
    pairs = collect_models(category_ids=category_ids, brand_ids=brand_ids, model_ids=model_ids)
    if not pairs:
        return [], []

    enabled_fields = []
    for field in export_fields:
        if not isinstance(field, dict):
            continue
        name = field.get("field") or field.get("name") or field.get("key")
        if name is None:
            continue
        name = str(name).strip()
        if not name or not field.get("enabled"):
            continue
        template_str = field.get("template", "")
        if template_str is None:
            template_str = ""
        enabled_fields.append({"field": name, "template": str(template_str)})

    if not enabled_fields:
        raise ValueError("Увімкніть хоча б одне поле експорту.")

    specs_map = load_specs_map([mid for _brand, _model, _cat, mid, _bid, _cid in pairs])
    title_tags_cache = {}
    desc_template_cache = {}
    field_template_cache = {}
    rows = []
    now_value = datetime.now()

    column_order = [field["field"] for field in enabled_fields]

    descriptions = templates.get("descriptions", {})

    total_steps = len(pairs) * len(film_types)
    progress_count = 0
    if progress_callback is not None:
        progress_callback(progress_count, total_steps)

    for brand, model, cat, mid, brand_id, cat_id in pairs:
        specs = specs_map.get(mid, {})
        if not isinstance(specs, dict):
            specs = {}
        spec_items = list(specs.items())

        def spec_lookup(key, default=""):
            return specs.get(key, default)

        cat_desc_block = descriptions.get(cat, {}) if isinstance(descriptions, dict) else {}
        for f in film_types:
            film_type = f if isinstance(f, str) else str(f)
            if film_type not in title_tags_cache:
                title_tpl_str, tags_tpl_str = resolve_title_tags(title_tags_templates, templates, film_type)
                try:
                    title_tpl = Template(title_tpl_str)
                    tags_tpl = Template(tags_tpl_str)
                except TemplateError as exc:
                    raise ValueError(
                        f"Помилка в шаблонах заголовку/тегів для типу \"{film_type}\": {exc}"
                    ) from exc
                title_tags_cache[film_type] = (title_tpl, tags_tpl)
            title_t, tags_t = title_tags_cache[film_type]
            try:
                title_value = title_t.render(film_type=film_type, brand=brand, model=model, category=cat)
                tags_value = tags_t.render(film_type=film_type, brand=brand, model=model, category=cat)
            except TemplateError as exc:
                raise ValueError(
                    f"Не вдалося згенерувати заголовок або теги для типу \"{film_type}\": {exc}"
                ) from exc

            desc_key = (cat, film_type)
            desc_tpl = desc_template_cache.get(desc_key)
            if desc_tpl is None:
                desc_template_str = cat_desc_block.get(film_type)
                if desc_template_str is None:
                    desc_template_str = cat_desc_block.get("default")
                if desc_template_str is None:
                    desc_template_str = "Плівка для {{ brand }} {{ model }}"
                try:
                    desc_tpl = Template(desc_template_str)
                except TemplateError as exc:
                    raise ValueError(
                        f"Помилка в шаблоні опису для категорії \"{cat}\" і типу \"{film_type}\": {exc}"
                    ) from exc
                desc_template_cache[desc_key] = desc_tpl
            try:
                desc_value = desc_tpl.render(film_type=film_type, brand=brand, model=model, category=cat)
            except TemplateError as exc:
                raise ValueError(
                    f"Не вдалося сформувати опис для категорії \"{cat}\" і типу \"{film_type}\": {exc}"
                ) from exc

            context = {
                "brand": brand,
                "brand_id": brand_id,
                "model": model,
                "model_id": mid,
                "category": cat,
                "category_id": cat_id,
                "film_type": film_type,
                "title": title_value,
                "description": desc_value,
                "tags": tags_value,
                "specs": specs,
                "spec_items": spec_items,
                "spec": spec_lookup,
                "row_number": len(rows) + 1,
                "now": now_value,
            }

            row_values = []
            formula_context = None
            for field in enabled_fields:
                field_name = field["field"]
                tpl_str = field.get("template", "")
                if tpl_str:
                    if _looks_like_formula(tpl_str):
                        if formula_context is None:
                            formula_context = _build_formula_context(context)
                        try:
                            value = FormulaEngine.evaluate(tpl_str, formula_context)
                        except FormulaError as exc:
                            raise ValueError(
                                f"Помилка у формулі поля \"{field_name}\": {exc}"
                            ) from exc
                    else:
                        tpl = field_template_cache.get(tpl_str)
                        if tpl is None:
                            try:
                                tpl = Template(tpl_str)
                            except TemplateError as exc:
                                raise ValueError(
                                    f"Помилка в шаблоні поля \"{field_name}\": {exc}"
                                ) from exc
                            field_template_cache[tpl_str] = tpl
                        try:
                            value = tpl.render(**context)
                        except TemplateError as exc:
                            raise ValueError(
                                f"Не вдалося згенерувати значення поля \"{field_name}\": {exc}"
                            ) from exc
                else:
                    value = context.get(field_name, "")
                if value is None:
                    value = ""
                elif not isinstance(value, str):
                    value = str(value)
                row_values.append(value)
            rows.append(row_values)
            progress_count += 1
            if progress_callback is not None:
                progress_callback(progress_count, total_steps)

    return rows, column_order


def _row_to_values(record, columns):
    if isinstance(record, dict):
        return [record.get(name, "") for name in columns]
    if isinstance(record, (list, tuple)):
        values = list(record)
        if len(values) < len(columns):
            values.extend([""] * (len(columns) - len(values)))
        elif len(values) > len(columns):
            values = values[: len(columns)]
        return values
    return ["" for _ in columns]


def _make_unique_column_keys(columns):
    counts = {}
    unique_keys = []
    for name in columns:
        count = counts.get(name, 0) + 1
        counts[name] = count
        if count == 1:
            unique_keys.append(name)
        else:
            unique_keys.append(f"{name}__{count}")
    return unique_keys


def export_products(records: list, columns: list, fmt: str, folder: str):
    ensure_folder(folder)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = os.path.join(folder, f"products_{ts}")

    if fmt == EXCEL_FORMAT_LABEL:
        if pd is None or EXCEL_EXPORT_BLOCKED_MESSAGE:
            if pd is None:
                message = PANDAS_EXPORT_BLOCKED_MESSAGE or "Експорт у Excel недоступний."
                detail = PANDAS_IMPORT_ERROR_DETAIL
                install_hint = PANDAS_INSTALL_HINT
            else:
                message = EXCEL_EXPORT_BLOCKED_MESSAGE or "Експорт у Excel недоступний."
                detail = EXCEL_ENGINE_IMPORT_ERROR_DETAIL
                install_hint = OPENPYXL_INSTALL_HINT
            if detail and detail not in message:
                message = f"{message} (деталі: {detail})"
            message = f"{message}\n{install_hint}{CSV_JSON_FALLBACK_NOTE}"
            raise RuntimeError(message)
        out_products = base + ".xlsx"
        df = pd.DataFrame.from_records(records, columns=columns)
        try:
            df.to_excel(out_products, index=False)
        except (ImportError, ModuleNotFoundError) as exc:
            exc_text = str(exc)
            if "openpyxl" in exc_text.lower():
                message = (
                    "Не вдалося зберегти Excel-файл: бібліотека openpyxl не встановлена або пошкоджена.\n"
                    "Встановіть її командою 'pip install openpyxl' і перезапустіть застосунок, або оберіть інший формат експорту."
                )
                raise RuntimeError(message) from exc
            raise
    elif fmt == CSV_FORMAT_LABEL:
        out_products = base + ".csv"
        if pd is not None:
            pd.DataFrame.from_records(records, columns=columns).to_csv(
                out_products, index=False, encoding="utf-8-sig"
            )
        else:
            with open(out_products, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                if columns:
                    writer.writerow(columns)
                for record in records:
                    row = _row_to_values(record, columns)
                    writer.writerow(row)
    elif fmt == JSON_FORMAT_LABEL:
        out_products = base + ".json"
        json_records = []
        json_columns = _make_unique_column_keys(columns)
        for record in records:
            values = _row_to_values(record, columns)
            json_records.append({key: value for key, value in zip(json_columns, values)})
        with open(out_products, "w", encoding="utf-8") as f:
            json.dump(json_records, f, ensure_ascii=False, indent=2)
    else:
        raise ValueError(f"Невідомий формат експорту: {fmt}")

    return out_products

# ============================ GUI: ВІКНО ХАРАКТЕРИСТИК ============================

class SpecsWindow(ctk.CTkToplevel):
    def __init__(self, master, model_id: int, model_name: str):
        super().__init__(master)
        self.model_id = model_id
        self.title(f"Характеристики: {model_name}")
        self.geometry("700x480")
        self.resizable(True, True)

        # Таблиця
        self.tree = ttk.Treeview(self, columns=("key", "value"), show="headings", height=16)
        self.tree.heading("key", text="Назва параметра")
        self.tree.heading("value", text="Значення")
        self.tree.column("key", width=260, anchor="w")
        self.tree.column("value", width=360, anchor="w")
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)

        scroll = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        scroll.place(relx=1.0, rely=0.0, relheight=1.0, anchor="ne")
        self.tree.bind("<Delete>", self._on_delete_key)
        self.tree.bind("<Button-1>", self._on_tree_click, add="+")

        self._rename_click = (None, None, 0.0)
        self._rename_entry = None
        self._rename_meta = None
        self._rename_delay_min = 0.35
        self._rename_delay_max = 4.0

        # Панель керування
        ctrl = ctk.CTkFrame(self)
        ctrl.pack(fill="x", padx=10, pady=(0,10))

        self.key_entry = ctk.CTkEntry(ctrl, placeholder_text="Напр.: Діагональ екрану")
        self.key_entry.pack(side="left", fill="x", expand=True, padx=5)
        self.val_entry = ctk.CTkEntry(ctrl, placeholder_text="Напр.: 6.7''")
        self.val_entry.pack(side="left", fill="x", expand=True, padx=5)

        ctk.CTkButton(ctrl, text="Додати", command=self._add).pack(side="left", padx=5)
        ctk.CTkButton(ctrl, text="Оновити", command=self._edit).pack(side="left", padx=5)
        ctk.CTkButton(ctrl, text="Видалити", fg_color="#8b0000", hover_color="#a40000", command=self._delete).pack(side="left", padx=5)

        self._refresh()

    def _on_delete_key(self, _event):
        self._delete()
        return "break"

    def _refresh(self):
        if self._rename_entry is not None:
            self._finish_inline_edit(save=False)
        self.tree.delete(*self.tree.get_children())
        for sid, k, v in get_specs(self.model_id):
            self.tree.insert("", "end", iid=f"spec_{sid}", values=(k, v))

    def _on_tree_click(self, event):
        row = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        region = self.tree.identify_region(event.x, event.y)
        now = time.time()
        if region not in {"cell", "tree"}:
            self._rename_click = (None, None, now)
            return
        if not row or column not in {"#1", "#2"}:
            self._rename_click = (row, column, now)
            return
        last_row, last_col, last_time = self._rename_click
        self._rename_click = (row, column, now)
        delay = now - last_time
        if row == last_row and column == last_col and self._rename_delay_min <= delay <= self._rename_delay_max:
            self.after(0, lambda: self._start_inline_edit(row, column))

    def _start_inline_edit(self, iid: str, column: str):
        if not self.tree.exists(iid):
            return
        if column not in {"#1", "#2"}:
            return
        values = self.tree.item(iid, "values")
        if not values:
            return
        index = 0 if column == "#1" else 1
        original = values[index]
        bbox = self.tree.bbox(iid, column)
        if not bbox:
            return
        if self._rename_entry is not None:
            self._finish_inline_edit(save=False)
        entry = create_inline_entry(self.tree, original)
        x, y, width, height = bbox
        entry.place(x=x, y=y, width=width, height=height)
        field = "key" if column == "#1" else "value"
        self._rename_entry = entry
        self._rename_meta = (iid, field)
        entry.bind("<Return>", lambda _e: self._finish_inline_edit(save=True))
        entry.bind("<KP_Enter>", lambda _e: self._finish_inline_edit(save=True))
        entry.bind("<Escape>", lambda _e: self._finish_inline_edit(save=False))
        entry.bind("<FocusOut>", lambda _e: self._finish_inline_edit(save=True))

    def _finish_inline_edit(self, save: bool):
        if not self._rename_entry or not self._rename_meta:
            return
        entry = self._rename_entry
        iid, field = self._rename_meta
        self._rename_entry = None
        self._rename_meta = None
        new_value = entry.get().strip()
        entry.destroy()
        if not save:
            self._restore_selection(iid)
            return
        if field == "key" and not new_value:
            show_error("Назва параметра не може бути порожньою.")
            self._restore_selection(iid)
            return
        sid = int(iid.split("_")[1])
        values = self.tree.item(iid, "values")
        if not values or len(values) < 2:
            self._restore_selection(iid)
            return
        key, value = values[0], values[1]
        if field == "key":
            if new_value == key:
                self._restore_selection(iid)
                return
            key = new_value
        else:
            if new_value == value:
                self._restore_selection(iid)
                return
            value = new_value
        update_spec(sid, key, value)
        self._refresh()
        self.after(10, lambda: self._restore_selection(f"spec_{sid}"))

    def _restore_selection(self, iid: str):
        def _select():
            if not self.tree.exists(iid):
                return
            self.tree.selection_set(iid)
            self.tree.focus(iid)
            self.tree.see(iid)
        self.after(10, _select)

    def _add(self):
        k = self.key_entry.get().strip()
        v = self.val_entry.get().strip()
        if not k:
            show_error("Введіть назву параметра.")
            return
        insert_spec(self.model_id, k, v)
        self.key_entry.delete(0, tk.END); self.val_entry.delete(0, tk.END)
        self._refresh()

    def _edit(self):
        sel = self.tree.selection()
        if not sel:
            show_error("Оберіть рядок у таблиці.")
            return
        sid = int(sel[0].split("_")[1])
        k = self.key_entry.get().strip()
        v = self.val_entry.get().strip()
        if not k:
            show_error("Введіть назву параметра.")
            return
        update_spec(sid, k, v)
        self._refresh()

    def _delete(self):
        selection = list(self.tree.selection())
        if not selection:
            show_error("Оберіть рядок у таблиці.")
            return
        prompt = (
            "Видалити вибрану характеристику?"
            if len(selection) == 1
            else "Видалити вибрані характеристики?"
        )
        if not messagebox.askyesno("Підтвердження", prompt):
            return
        ids = [int(iid.split("_")[1]) for iid in selection]
        for sid in ids:
            delete_spec(sid)
        self._refresh()

# ============================ GUI: ОСНОВНИЙ ДОДАТОК ============================

class _TkAppCompatProxy:
    """Proxy that allows attaching arbitrary Python attributes to a tkapp."""

    __slots__ = ("_tkapp", "_extras")

    def __init__(self, tkapp):
        object.__setattr__(self, "_tkapp", tkapp)
        object.__setattr__(self, "_extras", {})

    def __getattr__(self, name):
        extras = object.__getattribute__(self, "_extras")
        if name in extras:
            return extras[name]
        return getattr(object.__getattribute__(self, "_tkapp"), name)

    def __setattr__(self, name, value):
        if name in {"_tkapp", "_extras"}:
            object.__setattr__(self, name, value)
            return
        extras = object.__getattribute__(self, "_extras")
        extras[name] = value

    def __delattr__(self, name):
        if name in {"_tkapp", "_extras"}:
            raise AttributeError(name)
        extras = object.__getattribute__(self, "_extras")
        if name in extras:
            del extras[name]
            return
        delattr(object.__getattribute__(self, "_tkapp"), name)

    def __dir__(self):
        extras = object.__getattribute__(self, "_extras")
        base_dir = dir(object.__getattribute__(self, "_tkapp"))
        return sorted(set(base_dir).union(extras.keys()))


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        raw_tkapp = object.__getattribute__(self, "tk")
        if not isinstance(raw_tkapp, _TkAppCompatProxy):
            object.__setattr__(self, "tk", _TkAppCompatProxy(raw_tkapp))
        tkapp = object.__getattribute__(self, "tk")
        self.title(APP_TITLE)
        self.geometry("1100x680")
        self.minsize(980, 640)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.templates = load_templates()
        self.title_tags_templates = load_title_tags_templates(self.templates)
        self.export_fields = load_export_fields()
        self.current_category_id = None
        self.current_brand_id = None
        self._current_film_type_key = None
        self._current_desc_category = None
        self._template_scope_display_to_pair = {}
        self._template_scope_pair_to_display = {}
        self._gen_tree = None
        self._gen_tree_states = {}
        self._gen_tree_meta = {}
        self._gen_tree_labels = {}
        self._rename_clicks = {
            "cat": (None, 0.0),
            "brand": (None, 0.0),
            "model": (None, 0.0),
        }
        self._rename_entry = None
        self._rename_entry_meta = None
        self._rename_delay_min = 0.35
        self._rename_delay_max = 4.0
        self._export_selected_index = None
        self._export_tree_updating = False
        self.progress_bar = None
        self.progress_label = None
        self._preview_window = None
        # Compatibility: some flows expect the filmtype name variable to exist during tab
        # construction even if the dedicated film type tab is hidden. Older widgets access
        # the variable through the low-level Tk interpreter (self.tk), so expose it there
        # via a proxy that can store Python attributes.
        self.filmtype_name_var = tk.StringVar(master=self, value="")
        tkapp.filmtype_name_var = self.filmtype_name_var

        self.filmtype_enabled_var = tk.BooleanVar(master=self, value=True)
        tkapp.filmtype_enabled_var = self.filmtype_enabled_var

        self._build_header()
        self._build_tabs()

        # Початкові дані
        self._refresh_categories()
        self._refresh_filmtype_checkboxes()

        if DEPENDENCY_WARNINGS:
            self.after(200, self._show_dependency_warnings)

    # -------- верхній бар
    def _build_header(self):
        top = ctk.CTkFrame(self)
        top.pack(fill="x", padx=10, pady=(10, 0))
        ctk.CTkLabel(top, text=APP_TITLE, font=ctk.CTkFont(size=18, weight="bold")).pack(side="left")

        self.theme_var = tk.StringVar(value="dark")
        theme = ctk.CTkOptionMenu(top, values=["dark", "light", "system"], variable=self.theme_var, width=120,
                                  command=lambda v: ctk.set_appearance_mode(v))
        theme.set("dark")
        theme.pack(side="right")

    # -------- вкладки
    def _build_tabs(self):
        tabs = ctk.CTkTabview(self, width=1040, height=600)
        tabs.pack(fill="both", expand=True, padx=10, pady=10)

        self.tab_catalog   = tabs.add("Каталог")
        self.tab_templates = tabs.add("Шаблони")
        self.tab_export    = tabs.add("Експорт")
        self.tab_generate  = tabs.add("Генерація")

        self._build_tab_catalog()
        self._build_tab_templates()
        self._build_tab_export()
        self._build_tab_generate()

    def _show_dependency_warnings(self):
        for warning in DEPENDENCY_WARNINGS:
            messagebox.showwarning(APP_TITLE, warning)
        DEPENDENCY_WARNINGS.clear()

    def _film_type_names(self):
        return [item.get("name") for item in self.templates.get("film_types", []) if item.get("name")]

    def _film_type_menu_items(self):
        items = [(FILM_TYPE_DEFAULT_LABEL, "default")]
        for name in self._film_type_names():
            items.append((name, name))
        return items

    def _template_scope_items(self):
        categories = list(self.templates.get("descriptions", {}).keys())
        film_items = self._film_type_menu_items()
        if not categories:
            categories = [""]
        items = []
        for cat in categories:
            display_cat = cat if cat else "—"
            for film_label, film_key in film_items:
                label = f"{display_cat} — {film_label}"
                items.append((label, (cat, film_key)))
        return items

    def _selected_film_type_key(self) -> str:
        key = getattr(self, "_current_film_type_key", None)
        if key in (None, ""):
            self._current_film_type_key = "default"
            return "default"
        if key == "default" or key in set(self._film_type_names()):
            return key
        self._current_film_type_key = "default"
        return "default"

    # -------- Каталог (перший дизайн)
    def _build_tab_catalog(self):
        # Ліва колона: Категорії + Бренди
        left = ctk.CTkFrame(self.tab_catalog)
        left.pack(side="left", fill="both", expand=True, padx=(0,10), pady=10)

        # Категорії
        ctk.CTkLabel(left, text="Категорія").pack(anchor="w", padx=10, pady=(8,0))
        cat_frame = ctk.CTkFrame(left); cat_frame.pack(fill="x", padx=10, pady=5)
        self.cat_tree = ttk.Treeview(cat_frame, columns=("name",), show="headings", height=6)
        self.cat_tree.heading("name", text="Назва")
        self.cat_tree.column("name", width=260, anchor="w")
        self.cat_tree.pack(side="left", fill="x", expand=True)
        cat_scroll = ttk.Scrollbar(cat_frame, orient="vertical", command=self.cat_tree.yview)
        cat_scroll.pack(side="right", fill="y"); self.cat_tree.configure(yscrollcommand=cat_scroll.set)
        self.cat_tree.bind("<<TreeviewSelect>>", self._on_category_select)
        self.cat_tree.bind("<Button-1>", lambda e: self._handle_tree_click(e, "cat", self.cat_tree), add="+")
        self.cat_tree.bind("<Delete>", lambda e: self._handle_tree_delete("cat"))

        cat_ctrl = ctk.CTkFrame(left); cat_ctrl.pack(fill="x", padx=10, pady=(0,10))
        self.cat_entry = ctk.CTkEntry(cat_ctrl, placeholder_text="Назва категорії")
        self.cat_entry.pack(side="left", fill="x", expand=True, padx=(0,5))
        ctk.CTkButton(cat_ctrl, text="Додати", command=self._cat_add, width=90).pack(side="left", padx=3)
        ctk.CTkButton(cat_ctrl, text="Перейменувати", command=self._cat_rename, width=120).pack(side="left", padx=3)
        ctk.CTkButton(cat_ctrl, text="Видалити", command=self._cat_delete, width=90,
                      fg_color="#8b0000", hover_color="#a40000").pack(side="left", padx=3)

        # Бренди
        ctk.CTkLabel(left, text="Бренд").pack(anchor="w", padx=10, pady=(8,0))
        brand_frame = ctk.CTkFrame(left); brand_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.brand_tree = ttk.Treeview(brand_frame, columns=("name",), show="headings", height=11)
        self.brand_tree.heading("name", text="Назва")
        self.brand_tree.column("name", width=260, anchor="w")
        self.brand_tree.pack(side="left", fill="both", expand=True)
        brand_scroll = ttk.Scrollbar(brand_frame, orient="vertical", command=self.brand_tree.yview)
        brand_scroll.pack(side="right", fill="y"); self.brand_tree.configure(yscrollcommand=brand_scroll.set)
        self.brand_tree.bind("<<TreeviewSelect>>", self._on_brand_select)
        self.brand_tree.bind("<Button-1>", lambda e: self._handle_tree_click(e, "brand", self.brand_tree), add="+")
        self.brand_tree.bind("<Delete>", lambda e: self._handle_tree_delete("brand"))

        brand_ctrl = ctk.CTkFrame(left); brand_ctrl.pack(fill="x", padx=10, pady=(0,10))
        self.brand_entry = ctk.CTkEntry(brand_ctrl, placeholder_text="Назва бренду")
        self.brand_entry.pack(side="left", fill="x", expand=True, padx=(0,5))
        ctk.CTkButton(brand_ctrl, text="Додати", command=self._brand_add, width=90).pack(side="left", padx=3)
        ctk.CTkButton(brand_ctrl, text="Перейменувати", command=self._brand_rename, width=120).pack(side="left", padx=3)
        ctk.CTkButton(brand_ctrl, text="Видалити", command=self._brand_delete, width=90,
                      fg_color="#8b0000", hover_color="#a40000").pack(side="left", padx=3)

        # Права колона: Моделі
        right = ctk.CTkFrame(self.tab_catalog)
        right.pack(side="left", fill="both", expand=True, padx=(10,0), pady=10)

        ctk.CTkLabel(right, text="Модель").pack(anchor="w", padx=10, pady=(8,0))
        model_frame = ctk.CTkFrame(right); model_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.model_tree = ttk.Treeview(model_frame, columns=("name",), show="headings", height=22)
        self.model_tree.heading("name", text="Назва")
        self.model_tree.column("name", width=360, anchor="w")
        self.model_tree.pack(side="left", fill="both", expand=True)
        model_scroll = ttk.Scrollbar(model_frame, orient="vertical", command=self.model_tree.yview)
        model_scroll.pack(side="right", fill="y"); self.model_tree.configure(yscrollcommand=model_scroll.set)
        self.model_tree.bind("<Button-1>", lambda e: self._handle_tree_click(e, "model", self.model_tree), add="+")
        self.model_tree.bind("<Double-1>", self._on_model_double_click, add="+")
        self.model_tree.bind("<Delete>", lambda e: self._handle_tree_delete("model"))

        model_ctrl = ctk.CTkFrame(right); model_ctrl.pack(fill="x", padx=10, pady=(0,10))
        self.model_entry = ctk.CTkEntry(model_ctrl, placeholder_text="Назва моделі")
        self.model_entry.pack(side="left", fill="x", expand=True, padx=(0,5))
        ctk.CTkButton(model_ctrl, text="Додати", command=self._model_add, width=90).pack(side="left", padx=3)
        ctk.CTkButton(model_ctrl, text="Перейменувати", command=self._model_rename, width=120).pack(side="left", padx=3)
        ctk.CTkButton(model_ctrl, text="Видалити", command=self._model_delete, width=90,
                      fg_color="#8b0000", hover_color="#a40000").pack(side="left", padx=3)
        ctk.CTkButton(model_ctrl, text="Характеристики", command=self._open_specs, width=140).pack(side="left", padx=6)

    def _on_model_double_click(self, event):
        row = self.model_tree.identify_row(event.y)
        if not row:
            return
        self.model_tree.selection_set(row)
        self._open_specs()

    def _handle_tree_click(self, event, kind, tree):
        row = tree.identify_row(event.y)
        now = time.time()
        if not row:
            self._rename_clicks[kind] = (None, now)
            return
        last_row, last_time = self._rename_clicks.get(kind, (None, 0.0))
        self._rename_clicks[kind] = (row, now)
        delay = now - last_time
        if row == last_row and self._rename_delay_min <= delay <= self._rename_delay_max:
            self.after(0, lambda: self._start_tree_rename(kind, tree, row))

    def _handle_tree_delete(self, kind):
        if kind == "cat":
            self._cat_delete()
        elif kind == "brand":
            self._brand_delete()
        else:
            self._model_delete()
        return "break"

    def _start_tree_rename(self, kind, tree, iid):
        if not tree.exists(iid):
            return
        values = tree.item(iid, "values")
        if not values:
            return
        original = values[0]
        bbox = tree.bbox(iid, column="#1")
        if not bbox:
            return
        if self._rename_entry is not None:
            self._finish_inline_rename(save=False)
        entry = create_inline_entry(tree, original)
        x, y, width, height = bbox
        entry.place(x=x, y=y, width=width, height=height)
        self._rename_entry = entry
        self._rename_entry_meta = (kind, tree, iid, original)
        entry.bind("<Return>", lambda _e: self._finish_inline_rename(save=True))
        entry.bind("<KP_Enter>", lambda _e: self._finish_inline_rename(save=True))
        entry.bind("<Escape>", lambda _e: self._finish_inline_rename(save=False))
        entry.bind("<FocusOut>", lambda _e: self._finish_inline_rename(save=True))

    def _finish_inline_rename(self, save: bool):
        if not self._rename_entry or not self._rename_entry_meta:
            return
        entry = self._rename_entry
        kind, tree, iid, original = self._rename_entry_meta
        self._rename_entry = None
        self._rename_entry_meta = None
        new_value = entry.get().strip()
        entry.destroy()
        if not save:
            self._restore_tree_selection(kind, iid)
            return
        if not new_value:
            if kind == "cat":
                show_error("Назва категорії не може бути порожньою.")
            elif kind == "brand":
                show_error("Назва бренду не може бути порожньою.")
            else:
                show_error("Назва моделі не може бути порожньою.")
            self._restore_tree_selection(kind, iid)
            return
        if new_value == original:
            self._restore_tree_selection(kind, iid)
            return
        self._apply_tree_rename(kind, iid, new_value)

    def _apply_tree_rename(self, kind, iid, new_value):
        if kind == "cat":
            cat_id = int(iid.split("_")[1])
            result = rename_category(cat_id, new_value)
            if result is not True:
                if isinstance(result, sqlite3.IntegrityError):
                    show_error("Категорія з такою назвою вже існує.")
                else:
                    show_error("Не вдалося перейменувати категорію.")
            self._refresh_categories()
            self.after(10, lambda: self._restore_tree_selection("cat", f"cat_{cat_id}"))
        elif kind == "brand":
            brand_id = int(iid.split("_")[1])
            result = rename_brand(brand_id, new_value)
            if result is not True:
                if isinstance(result, sqlite3.IntegrityError):
                    show_error("Бренд з такою назвою вже існує.")
                else:
                    show_error("Не вдалося перейменувати бренд.")
            if self.current_category_id:
                self._refresh_brands(self.current_category_id)
            else:
                self._refresh_brands(None)
            self._reload_gen_tree()
            self.after(10, lambda: self._restore_tree_selection("brand", f"brand_{brand_id}"))
        else:
            model_id = int(iid.split("_")[1])
            result = rename_model(model_id, new_value)
            if result is not True:
                if isinstance(result, sqlite3.IntegrityError):
                    show_error("Модель з такою назвою вже існує.")
                else:
                    show_error("Не вдалося перейменувати модель.")
            if self.current_brand_id:
                self._refresh_models(self.current_brand_id)
            else:
                self._refresh_models(None)
            self._reload_gen_tree()
            self.after(10, lambda: self._restore_tree_selection("model", f"model_{model_id}"))

    def _restore_tree_selection(self, kind, iid):
        tree = {
            "cat": getattr(self, "cat_tree", None),
            "brand": getattr(self, "brand_tree", None),
            "model": getattr(self, "model_tree", None),
        }.get(kind)
        if tree is None:
            return
        def _select():
            if not tree.exists(iid):
                return
            tree.selection_set(iid)
            tree.focus(iid)
            tree.see(iid)
            if kind == "cat":
                self._on_category_select()
            elif kind == "brand":
                self._on_brand_select()
        self.after(10, _select)

    # ---- catalog actions
    def _refresh_categories(self):
        self.cat_tree.delete(*self.cat_tree.get_children())
        for cid, name in get_categories():
            self.cat_tree.insert("", "end", iid=f"cat_{cid}", values=(name,))
        self.current_category_id = None
        self._refresh_brands(None)
        self._refresh_models(None)
        self._reload_gen_tree()

    def _refresh_brands(self, category_id):
        self.brand_tree.delete(*self.brand_tree.get_children())
        if category_id:
            for bid, name in get_brands(category_id):
                self.brand_tree.insert("", "end", iid=f"brand_{bid}", values=(name,))
        self.current_brand_id = None

    def _refresh_models(self, brand_id):
        self.model_tree.delete(*self.model_tree.get_children())
        if brand_id:
            for mid, name in get_models(brand_id):
                self.model_tree.insert("", "end", iid=f"model_{mid}", values=(name,))

    def _on_category_select(self, _evt=None):
        sel = self.cat_tree.selection()
        if not sel:
            self.current_category_id = None
            self._refresh_brands(None); self._refresh_models(None)
            return
        self.current_category_id = int(sel[0].split("_")[1])
        self._refresh_brands(self.current_category_id)
        self._refresh_models(None)

    def _on_brand_select(self, _evt=None):
        sel = self.brand_tree.selection()
        if not sel:
            self.current_brand_id = None
            self._refresh_models(None)
            return
        self.current_brand_id = int(sel[0].split("_")[1])
        self._refresh_models(self.current_brand_id)

    def _cat_add(self):
        name = self.cat_entry.get().strip()
        if not name: return show_error("Введіть назву категорії.")
        add_category(name); self.cat_entry.delete(0, tk.END); self._refresh_categories()

    def _cat_rename(self):
        if not self.current_category_id: return show_error("Виберіть категорію.")
        name = self.cat_entry.get().strip()
        if not name: return show_error("Введіть нову назву категорії.")
        cat_id = self.current_category_id
        result = rename_category(cat_id, name)
        if result is not True:
            if isinstance(result, sqlite3.IntegrityError):
                show_error("Категорія з такою назвою вже існує.")
            else:
                show_error("Не вдалося перейменувати категорію.")
        self._refresh_categories()
        if cat_id:
            self.after(10, lambda: self._restore_tree_selection("cat", f"cat_{cat_id}"))

    def _cat_delete(self):
        selection = list(self.cat_tree.selection())
        if not selection:
            return show_error("Виберіть категорію.")
        prompt = (
            "Видалити вибрану категорію та всі її бренди/моделі?"
            if len(selection) == 1
            else "Видалити вибрані категорії та всі їх бренди/моделі?"
        )
        if not messagebox.askyesno("Підтвердження", prompt):
            return
        ids = [int(iid.split("_")[1]) for iid in selection]
        for cat_id in ids:
            delete_category(cat_id)
        self._refresh_categories()

    def _brand_add(self):
        if not self.current_category_id: return show_error("Спочатку виберіть категорію.")
        raw = self.brand_entry.get()
        names = split_catalog_input(raw)
        if not names: return show_error("Введіть назву бренду (через кому для декількох).")
        for name in names:
            add_brand(self.current_category_id, name)
        self.brand_entry.delete(0, tk.END)
        self._refresh_brands(self.current_category_id)
        self._reload_gen_tree()

    def _brand_rename(self):
        if not self.current_brand_id: return show_error("Виберіть бренд.")
        name = self.brand_entry.get().strip()
        if not name: return show_error("Введіть нову назву бренду.")
        brand_id = self.current_brand_id
        result = rename_brand(brand_id, name)
        if result is not True:
            if isinstance(result, sqlite3.IntegrityError):
                show_error("Бренд з такою назвою вже існує.")
            else:
                show_error("Не вдалося перейменувати бренд.")
        self._refresh_brands(self.current_category_id)
        if brand_id:
            self.after(10, lambda: self._restore_tree_selection("brand", f"brand_{brand_id}"))
        self._reload_gen_tree()

    def _brand_delete(self):
        selection = list(self.brand_tree.selection())
        if not selection:
            return show_error("Виберіть бренд.")
        prompt = (
            "Видалити вибраний бренд та всі його моделі?"
            if len(selection) == 1
            else "Видалити вибрані бренди та всі їх моделі?"
        )
        if not messagebox.askyesno("Підтвердження", prompt):
            return
        ids = [int(iid.split("_")[1]) for iid in selection]
        for brand_id in ids:
            delete_brand(brand_id)
        self._refresh_brands(self.current_category_id)
        self._refresh_models(None)
        self._reload_gen_tree()

    def _model_add(self):
        if not self.current_brand_id: return show_error("Спочатку виберіть бренд.")
        raw = self.model_entry.get()
        names = split_catalog_input(raw)
        if not names: return show_error("Введіть назву моделі (через кому для декількох).")
        for name in names:
            add_model(self.current_brand_id, name)
        self.model_entry.delete(0, tk.END)
        self._refresh_models(self.current_brand_id)
        self._reload_gen_tree()

    def _model_rename(self):
        sel = self.model_tree.selection()
        if not sel: return show_error("Виберіть модель.")
        model_id = int(sel[0].split("_")[1])
        name = self.model_entry.get().strip()
        if not name: return show_error("Введіть нову назву моделі.")
        result = rename_model(model_id, name)
        if result is not True:
            if isinstance(result, sqlite3.IntegrityError):
                show_error("Модель з такою назвою вже існує.")
            else:
                show_error("Не вдалося перейменувати модель.")
        self._refresh_models(self.current_brand_id)
        self.after(10, lambda: self._restore_tree_selection("model", f"model_{model_id}"))
        self._reload_gen_tree()

    def _model_delete(self):
        selection = list(self.model_tree.selection())
        if not selection:
            return show_error("Виберіть модель.")
        prompt = (
            "Видалити вибрану модель?"
            if len(selection) == 1
            else "Видалити вибрані моделі?"
        )
        if not messagebox.askyesno("Підтвердження", prompt):
            return
        ids = [int(iid.split("_")[1]) for iid in selection]
        brand_id = self.current_brand_id
        for model_id in ids:
            delete_model(model_id)
        self._refresh_models(brand_id)
        self._reload_gen_tree()

    def _open_specs(self):
        sel = self.model_tree.selection()
        if not sel: return show_error("Виберіть модель.")
        model_id = int(sel[0].split("_")[1])
        model_name = self.model_tree.item(sel[0], "values")[0]
        SpecsWindow(self, model_id, model_name)

    # -------- Шаблони
    def _build_tab_templates(self):
        wrap = ctk.CTkFrame(self.tab_templates)
        wrap.pack(fill="both", expand=True, padx=10, pady=10)

        selector = ctk.CTkFrame(wrap)
        selector.pack(fill="x", padx=10, pady=(6,2))
        ctk.CTkLabel(selector, text="Категорія та тип плівки:").pack(side="left", padx=(0,6))
        self.template_scope_var = tk.StringVar(value="")
        self.template_scope_menu = ctk.CTkOptionMenu(
            selector,
            values=["—"],
            variable=self.template_scope_var,
            width=320,
            command=self._on_template_scope_change
        )
        self.template_scope_menu.pack(side="left")

        # Ліва колонка: Заголовок і Теги
        left = ctk.CTkFrame(wrap)
        left.pack(side="left", fill="both", expand=True, padx=(0,10), pady=5)

        ctk.CTkLabel(left, text="Шаблон заголовка ({{ brand }}, {{ model }}, {{ film_type }})").pack(anchor="w", padx=10, pady=(10,0))
        self.title_box = ctk.CTkTextbox(left, height=80)
        self.title_box.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(left, text="Шаблон тегів ({{ brand }}, {{ model }}, {{ film_type }})").pack(anchor="w", padx=10, pady=(10,0))
        self.tags_box = ctk.CTkTextbox(left, height=110)
        self.tags_box.pack(fill="x", padx=10, pady=5)

        ctk.CTkButton(left, text="Зберегти заголовок/теги", command=self._save_title_tags).pack(anchor="e", padx=10, pady=10)

        # Права колонка: Опис для Категорії + Типу плівки
        right = ctk.CTkFrame(wrap)
        right.pack(side="left", fill="both", expand=True, padx=(10,0), pady=5)

        categories = list(self.templates.get("descriptions", {}).keys())
        initial_category = categories[0] if categories else ""
        self.desc_cat_var = tk.StringVar(value=initial_category)
        if self._current_desc_category is None:
            self._current_desc_category = initial_category

        ctk.CTkLabel(right, text="Шаблон опису (доступні {{ brand }}, {{ model }}, {{ film_type }})").pack(anchor="w", padx=10, pady=(10,0))
        self.desc_box = ctk.CTkTextbox(right)
        self.desc_box.pack(fill="both", expand=True, padx=10, pady=5)

        btn_row = ctk.CTkFrame(right); btn_row.pack(fill="x", padx=10, pady=6)
        ctk.CTkButton(btn_row, text="Зберегти опис", command=self._save_desc_template).pack(side="right")

        self._refresh_template_scope_menu()

    def _save_title_tags(self, show_message: bool = True):
        film = self._selected_film_type_key()
        title_value = self.title_box.get("1.0", "end").strip()
        tags_value = self.tags_box.get("1.0", "end").strip()

        if film not in self.title_tags_templates:
            self.title_tags_templates[film] = {}

        self.title_tags_templates[film]["title_template"] = title_value
        self.title_tags_templates[film]["tags_template"] = tags_value
        save_title_tags_templates(self.title_tags_templates)

        if film == "default":
            self.templates["title_template"] = title_value
            self.templates["tags_template"] = tags_value
            save_templates(self.templates)

        if show_message:
            show_info("Шаблони заголовку та тегів збережено.")

    def _load_title_tags_template(self):
        if not hasattr(self, "title_box") or not hasattr(self, "tags_box"):
            return
        film = self._selected_film_type_key()
        title_template, tags_template = resolve_title_tags(self.title_tags_templates, self.templates, film)
        self.title_box.delete("1.0", "end")
        self.title_box.insert("1.0", title_template)
        self.tags_box.delete("1.0", "end")
        self.tags_box.insert("1.0", tags_template)

    def _refresh_template_scope_menu(self):
        if not hasattr(self, "template_scope_menu"):
            return
        items = self._template_scope_items()
        if not items:
            return
        display_values = [label for label, _ in items]
        self._template_scope_display_to_pair = {label: pair for label, pair in items}
        self._template_scope_pair_to_display = {pair: label for label, pair in items}
        self.template_scope_menu.configure(values=display_values)

        current_cat = getattr(self, "_current_desc_category", None)
        current_film = self._selected_film_type_key()
        if current_cat is None:
            current_cat = items[0][1][0]
        pair = (current_cat, current_film)
        if pair not in self._template_scope_pair_to_display:
            pair = items[0][1]
            current_cat, current_film = pair

        self._current_desc_category = current_cat
        if hasattr(self, "desc_cat_var"):
            self.desc_cat_var.set(current_cat or "")
        self._current_film_type_key = current_film
        current_label = self._template_scope_pair_to_display[pair]
        self.template_scope_var.set(current_label)
        self.template_scope_menu.set(current_label)
        self._on_template_scope_change(current_label)

    def _on_template_scope_change(self, selected_label=None):
        if not hasattr(self, "template_scope_var"):
            return
        if selected_label is None:
            selected_label = self.template_scope_var.get()
        pair = self._template_scope_display_to_pair.get(selected_label)
        if pair is None and self._template_scope_display_to_pair:
            pair = next(iter(self._template_scope_display_to_pair.values()))
        if pair is None:
            return
        cat, film = pair
        if cat is None:
            cat = ""
        self._current_desc_category = cat
        if hasattr(self, "desc_cat_var"):
            self.desc_cat_var.set(cat)
        self._current_film_type_key = film or "default"
        self._load_title_tags_template()
        self._load_desc_template()

    def _load_desc_template(self):
        if not hasattr(self, "desc_box") or not hasattr(self, "desc_cat_var"):
            return
        cat = self.desc_cat_var.get()
        if not cat:
            categories = list(self.templates.get("descriptions", {}).keys())
            if categories:
                cat = categories[0]
                self.desc_cat_var.set(cat)
        self._current_desc_category = cat
        film = self._selected_film_type_key()
        descs = self.templates["descriptions"].get(cat, {})
        if film == "default":
            txt = descs.get("default", "")
        else:
            txt = descs.get(film)
            if txt is None:
                txt = descs.get("default", "")
        if txt is None:
            txt = ""
        self.desc_box.delete("1.0", "end")
        self.desc_box.insert("1.0", txt)

    def _save_desc_template(self):
        cat = self.desc_cat_var.get()
        film = self._selected_film_type_key()
        txt = self.desc_box.get("1.0", "end").strip()
        self.templates["descriptions"].setdefault(cat, {})[film] = txt
        save_templates(self.templates)
        show_info("Шаблон опису збережено.")

    # -------- Експортні поля
    def _build_tab_export(self):
        wrap = ctk.CTkFrame(self.tab_export)
        wrap.pack(fill="both", expand=True, padx=10, pady=10)
        wrap.grid_columnconfigure(0, weight=0)
        wrap.grid_columnconfigure(1, weight=1)
        wrap.grid_rowconfigure(0, weight=1)

        list_frame = ctk.CTkFrame(wrap)
        list_frame.grid(row=0, column=0, sticky="ns", padx=(0, 10))
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        self.export_fields_tree = ttk.Treeview(
            list_frame,
            columns=("field", "enabled"),
            show="headings",
            selectmode="browse",
            height=18,
        )
        self.export_fields_tree.heading("field", text="Поле")
        self.export_fields_tree.heading("enabled", text="Увімкнено")
        self.export_fields_tree.column("field", width=220, anchor="w")
        self.export_fields_tree.column("enabled", width=90, anchor="center")
        self.export_fields_tree.grid(row=0, column=0, sticky="nsew")

        y_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.export_fields_tree.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        self.export_fields_tree.configure(yscrollcommand=y_scroll.set)
        self.export_fields_tree.bind("<<TreeviewSelect>>", self._on_export_field_select)

        btn_frame = ctk.CTkFrame(list_frame)
        btn_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ctk.CTkButton(btn_frame, text="Додати поле", command=self._export_add_field).pack(side="left", padx=4)
        ctk.CTkButton(btn_frame, text="Видалити", command=self._export_delete_field).pack(side="left", padx=4)
        ctk.CTkButton(btn_frame, text="Вгору", command=lambda: self._export_move_field(-1)).pack(side="right", padx=4)
        ctk.CTkButton(btn_frame, text="Вниз", command=lambda: self._export_move_field(1)).pack(side="right", padx=4)

        detail = ctk.CTkFrame(wrap)
        detail.grid(row=0, column=1, sticky="nsew")
        detail.grid_rowconfigure(4, weight=1)

        self.export_field_name_var = tk.StringVar(value="")
        self.export_field_enabled_var = tk.BooleanVar(value=False)

        ctk.CTkLabel(detail, text="Назва поля").pack(anchor="w", padx=10, pady=(8, 0))
        self.export_field_name_entry = ctk.CTkEntry(detail, textvariable=self.export_field_name_var)
        self.export_field_name_entry.pack(fill="x", padx=10, pady=(0, 8))

        self.export_field_enabled_check = ctk.CTkCheckBox(
            detail,
            text="Увімкнути поле",
            variable=self.export_field_enabled_var,
        )
        self.export_field_enabled_check.pack(anchor="w", padx=10, pady=(0, 8))

        ctk.CTkLabel(detail, text="Шаблон (формула =IF(...) або Jinja2)").pack(anchor="w", padx=10, pady=(0, 4))
        self.export_field_template = ctk.CTkTextbox(detail, height=220)
        self.export_field_template.pack(fill="both", expand=True, padx=10, pady=(0, 6))
        self.export_field_template.bind("<<Modified>>", self._on_export_template_modified)

        hint_text = (
            "Почніть рядок зі знаком '=' щоб використати формулу у стилі Google Sheets (аргументи через ';').\n"
            "Без '=' шаблон рендериться через Jinja2. Доступні змінні: {{ brand }}, {{ model }}, {{ category }}, {{ film_type }}, "
            "{{ title }}, {{ description }}, {{ tags }}, {{ spec('Назва') }}, {{ specs['Ключ'] }}, {{ row_number }}, {{ now }}."
        )
        ctk.CTkLabel(detail, text=hint_text, justify="left", anchor="w", wraplength=360).pack(
            fill="x", padx=10, pady=(0, 4)
        )

        self.export_template_status = ctk.CTkLabel(
            detail,
            text="",
            justify="left",
            anchor="w",
            wraplength=360,
            text_color="#888888",
        )
        self.export_template_status.pack(fill="x", padx=10, pady=(0, 8))
        self._update_template_status("")

        action_row = ctk.CTkFrame(detail)
        action_row.pack(fill="x", padx=10, pady=(0, 10))
        self.export_apply_button = ctk.CTkButton(action_row, text="Застосувати", command=lambda: self._export_apply_detail(False))
        self.export_apply_button.pack(side="right", padx=4)
        ctk.CTkButton(action_row, text="Зберегти", command=self._export_save_all).pack(side="right", padx=4)
        ctk.CTkButton(action_row, text="Відновити стандартні", command=self._export_reset_defaults).pack(side="left", padx=4)

        self._refresh_export_fields_tree(select_index=0 if self.export_fields else None)
        if self.export_fields:
            self._load_export_field_detail(0)
        else:
            self._load_export_field_detail(None)

    def _refresh_export_fields_tree(self, select_index=None):
        tree = getattr(self, "export_fields_tree", None)
        if tree is None:
            return
        self._export_tree_updating = True
        tree.delete(*tree.get_children())
        for idx, field in enumerate(self.export_fields):
            name = str(field.get("field", "")).strip()
            status = "Так" if field.get("enabled") else "Ні"
            tree.insert("", "end", iid=f"exp_{idx}", values=(name, status))
        if select_index is not None and 0 <= select_index < len(self.export_fields):
            iid = f"exp_{select_index}"
            tree.selection_set(iid)
            tree.focus(iid)
        self._export_tree_updating = False

    def _set_export_detail_state(self, enabled: bool):
        state = tk.NORMAL if enabled else tk.DISABLED
        if hasattr(self, "export_field_name_entry"):
            self.export_field_name_entry.configure(state=state)
        if hasattr(self, "export_field_enabled_check"):
            self.export_field_enabled_check.configure(state=state)
        if hasattr(self, "export_apply_button"):
            self.export_apply_button.configure(state=state)
        if hasattr(self, "export_field_template"):
            self.export_field_template.configure(state="normal" if enabled else "disabled")
        self._update_template_status()

    def _load_export_field_detail(self, index):
        if index is None or index < 0 or index >= len(self.export_fields):
            self._export_selected_index = None
            self._set_export_detail_state(False)
            if hasattr(self, "export_field_name_var"):
                self.export_field_name_var.set("")
            if hasattr(self, "export_field_enabled_var"):
                self.export_field_enabled_var.set(False)
            if hasattr(self, "export_field_template"):
                self.export_field_template.configure(state="normal")
                self.export_field_template.delete("1.0", "end")
                self.export_field_template.configure(state="disabled")
            self._update_template_status("")
            return

        self._export_selected_index = index
        field = self.export_fields[index]
        name = str(field.get("field", ""))
        template = field.get("template", "")
        if template is None:
            template = ""
        enabled = bool(field.get("enabled"))

        self._set_export_detail_state(True)
        self.export_field_name_var.set(name)
        self.export_field_enabled_var.set(enabled)
        self.export_field_template.configure(state="normal")
        self.export_field_template.delete("1.0", "end")
        self.export_field_template.insert("1.0", str(template))
        self.export_field_template.edit_modified(False)
        self._update_template_status(template)

    def _on_export_field_select(self, _event):
        if getattr(self, "_export_tree_updating", False):
            return
        selection = self.export_fields_tree.selection() if hasattr(self, "export_fields_tree") else []
        if not selection:
            self._load_export_field_detail(None)
            return
        iid = selection[0]
        try:
            idx = int(iid.split("_", 1)[1])
        except (IndexError, ValueError):
            idx = None
        self._export_apply_detail(False)
        if idx is None:
            self._load_export_field_detail(None)
        else:
            self._load_export_field_detail(idx)

    def _on_export_template_modified(self, _event):
        widget = getattr(self, "export_field_template", None)
        if widget is None:
            return
        try:
            modified = bool(widget.edit_modified())
        except Exception:
            modified = True
        if not modified:
            return
        try:
            widget.edit_modified(False)
        except Exception:
            pass
        text = widget.get("1.0", "end").rstrip()
        self._update_template_status(text)

    def _update_template_status(self, template_text: str | None = None):
        label = getattr(self, "export_template_status", None)
        if label is None:
            return
        if template_text is None:
            widget = getattr(self, "export_field_template", None)
            if widget is None:
                template_text = ""
            else:
                template_text = widget.get("1.0", "end").rstrip()
        message, color = self._analyze_template_text(template_text)
        label.configure(text=message, text_color=color)

    def _analyze_template_text(self, template_text: str):
        trimmed = template_text.strip()
        if not trimmed:
            return (
                "Введіть формулу (=IF(...)) або шаблон Jinja2. Аргументи формули розділяйте ';'.",
                "#888888",
            )
        if _looks_like_formula(trimmed):
            try:
                info = FormulaEngine.describe(trimmed)
            except FormulaError as exc:
                return (f"❌ Помилка формули: {exc}", "#c94a4a")
            variables = sorted(info.get("variables", []))
            if variables:
                vars_text = ", ".join(variables)
                hint = f"Змінні: {vars_text}"
            else:
                hint = "Змінні не використовуються."
            return (f"✅ Формула валідна. {hint}", "#4c9a2a")
        try:
            Template(trimmed)
        except TemplateError as exc:
            return (f"❌ Помилка шаблону Jinja2: {exc}", "#c94a4a")
        return (
            "ℹ️ Використовується шаблон Jinja2. Доступні змінні: {{ brand }}, {{ model }}, {{ category }}, {{ film_type }}, {{ title }}, {{ description }}, {{ tags }}, {{ row_number }}, {{ now }}.",
            "#888888",
        )

    def _export_apply_detail(self, save_to_file: bool):
        idx = getattr(self, "_export_selected_index", None)
        if idx is None or idx < 0 or idx >= len(self.export_fields):
            return False
        field = self.export_fields[idx]
        name = self.export_field_name_var.get().strip()
        if not name:
            name = field.get("field") or f"Поле_{idx + 1}"
            self.export_field_name_var.set(name)
        template = self.export_field_template.get("1.0", "end").rstrip()
        enabled = bool(self.export_field_enabled_var.get())

        trimmed_template = template.strip()
        if trimmed_template:
            if _looks_like_formula(trimmed_template):
                try:
                    FormulaEngine.describe(trimmed_template)
                except FormulaError as exc:
                    show_error(f"Помилка у формулі: {exc}")
                    self._update_template_status(template)
                    return False
            else:
                try:
                    Template(trimmed_template)
                except TemplateError as exc:
                    show_error(f"Помилка у шаблоні Jinja2: {exc}")
                    self._update_template_status(template)
                    return False

        changed = False
        if field.get("field") != name:
            field["field"] = name
            changed = True
        if field.get("template", "") != template:
            field["template"] = template
            changed = True
        if bool(field.get("enabled")) != enabled:
            field["enabled"] = enabled
            changed = True

        self._update_template_status(template)

        if changed:
            self._refresh_export_fields_tree(select_index=idx)

        if save_to_file:
            save_export_fields(self.export_fields)
            show_info("Налаштування експорту збережено.")

        return True

    def _export_save_all(self):
        self._export_apply_detail(False)
        save_export_fields(self.export_fields)
        show_info("Налаштування експорту збережено.")

    def _export_add_field(self):
        self._export_apply_detail(False)
        new_field = {"field": "Нове_поле", "template": "", "enabled": True}
        self.export_fields.append(new_field)
        idx = len(self.export_fields) - 1
        self._refresh_export_fields_tree(select_index=idx)
        self._load_export_field_detail(idx)

    def _export_delete_field(self):
        idx = getattr(self, "_export_selected_index", None)
        if idx is None or idx < 0 or idx >= len(self.export_fields):
            show_error("Оберіть поле для видалення.")
            return
        if not messagebox.askyesno("Підтвердження", "Видалити вибране поле?"):
            return
        self.export_fields.pop(idx)
        if self.export_fields:
            new_idx = min(idx, len(self.export_fields) - 1)
            self._refresh_export_fields_tree(select_index=new_idx)
            self._load_export_field_detail(new_idx)
        else:
            self._refresh_export_fields_tree(select_index=None)
            self._load_export_field_detail(None)

    def _export_move_field(self, direction: int):
        idx = getattr(self, "_export_selected_index", None)
        if idx is None:
            return
        new_idx = idx + direction
        if new_idx < 0 or new_idx >= len(self.export_fields):
            return
        self._export_apply_detail(False)
        self.export_fields[idx], self.export_fields[new_idx] = self.export_fields[new_idx], self.export_fields[idx]
        self._refresh_export_fields_tree(select_index=new_idx)
        self._load_export_field_detail(new_idx)

    def _export_reset_defaults(self):
        if not messagebox.askyesno("Підтвердження", "Відновити стандартний список полів?"):
            return
        self.export_fields = [field.copy() for field in DEFAULT_EXPORT_FIELDS]
        save_export_fields(self.export_fields)
        if self.export_fields:
            self._refresh_export_fields_tree(select_index=0)
            self._load_export_field_detail(0)
        else:
            self._refresh_export_fields_tree(select_index=None)
            self._load_export_field_detail(None)
        show_info("Стандартні поля відновлено.")

    # -------- Генерація
    def _build_tab_generate(self):
        wrap = ctk.CTkFrame(self.tab_generate)
        wrap.pack(fill="both", expand=True, padx=10, pady=10)
        
        left = ctk.CTkFrame(wrap)
        left.pack(side="left", fill="both", expand=True, padx=(0, 10), pady=10)

        ctk.CTkLabel(left, text="Категорії / бренди / моделі").pack(anchor="w", padx=10, pady=(6, 4))

        tree_container = ctk.CTkFrame(left)
        tree_container.pack(fill="both", expand=True, padx=10, pady=(0, 6))
        tree_container.grid_columnconfigure(0, weight=1)
        tree_container.grid_rowconfigure(0, weight=1)

        self._gen_tree = ttk.Treeview(tree_container, show="tree", selectmode="extended")
        self._gen_tree.column("#0", anchor="w", width=320)
        self._gen_tree.grid(row=0, column=0, sticky="nsew")

        y_scroll = ttk.Scrollbar(tree_container, orient="vertical", command=self._gen_tree.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(tree_container, orient="horizontal", command=self._gen_tree.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")

        self._gen_tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self._gen_tree.bind("<Button-1>", self._on_gen_tree_click, add="+")
        self._gen_tree.bind("<space>", self._toggle_selected_gen_node)
        self._gen_tree.bind("<Return>", self._toggle_selected_gen_node)

        controls = ctk.CTkFrame(left)
        controls.pack(fill="x", padx=10, pady=(0, 8))
        ctk.CTkButton(controls, text="Вибрати все", command=self._select_all_gen_tree, width=120).pack(side="left", padx=4)
        ctk.CTkButton(controls, text="Очистити", command=self._clear_all_gen_tree, width=120).pack(side="left", padx=4)
        ctk.CTkButton(controls, text="Розгорнути все", command=self._expand_all_gen_tree, width=140).pack(side="right", padx=4)
        ctk.CTkButton(controls, text="Згорнути все", command=self._collapse_all_gen_tree, width=140).pack(side="right", padx=4)

        tip_text = (
            "Порада: клацніть або натисніть пробіл, щоб поставити/зняти галочку. "
            "Без вибору буде згенеровано увесь каталог."
        )
        tip = ctk.CTkLabel(left, text=tip_text, anchor="w", wraplength=320)
        tip.pack(fill="x", padx=10, pady=(0, 6))

        right = ctk.CTkFrame(wrap)
        right.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)

        fmt_frame = ctk.CTkFrame(right)
        fmt_frame.pack(fill="x", padx=10, pady=(6, 6))
        ctk.CTkLabel(fmt_frame, text="Формат експорту:").pack(anchor="w", padx=6, pady=(4, 4))
        export_formats = get_available_export_formats()
        if not export_formats:
            export_formats = ["JSON (.json)"]
        default_format = export_formats[0]
        self.export_fmt_var = tk.StringVar(value=default_format)
        self.export_fmt_menu = ctk.CTkOptionMenu(
            fmt_frame,
            values=export_formats,
            variable=self.export_fmt_var,
            width=200,
        )
        self.export_fmt_menu.pack(anchor="w", padx=6, pady=(0, 4))

        path_frame = ctk.CTkFrame(right)
        path_frame.pack(fill="x", padx=10, pady=(4, 6))
        ctk.CTkLabel(path_frame, text="Папка збереження:").pack(anchor="w", padx=6, pady=(4, 4))
        self.out_folder_var = tk.StringVar(value=os.getcwd())
        path_entry = ctk.CTkEntry(path_frame, textvariable=self.out_folder_var)
        path_entry.pack(fill="x", padx=6, pady=(0, 4))
        ctk.CTkButton(path_frame, text="Обрати...", command=self._choose_folder, width=110).pack(anchor="e", padx=6, pady=(0, 4))

        types_frame = ctk.CTkFrame(right)
        types_frame.pack(fill="both", expand=True, padx=10, pady=(4, 6))
        ctk.CTkLabel(types_frame, text="Типи плівок:").pack(anchor="w", padx=6, pady=(4, 2))
        self.filmtype_frame = ctk.CTkFrame(types_frame, fg_color="transparent")
        self.filmtype_frame.pack(fill="x", padx=6, pady=(2, 6))
        self.ft_vars = []

        action_row = ctk.CTkFrame(right)
        action_row.pack(fill="x", padx=10, pady=(6, 0))
        ctk.CTkButton(
            action_row,
            text="Попередній перегляд",
            command=self._preview_generation,
            height=36,
        ).pack(side="right", padx=6)
        ctk.CTkButton(action_row, text="Згенерувати", command=self._generate, height=36).pack(side="right", padx=6)

        progress_frame = ctk.CTkFrame(right)
        progress_frame.pack(fill="x", padx=10, pady=(10, 0))
        self.progress_bar = ctk.CTkProgressBar(progress_frame)
        self.progress_bar.pack(fill="x", padx=6, pady=(6, 4))
        self.progress_bar.set(0)
        self.progress_label = ctk.CTkLabel(progress_frame, text="Очікування", anchor="w")
        self.progress_label.pack(fill="x", padx=6, pady=(0, 6))

        self._reload_gen_tree()

    def _progress_reset(self, message: str = "Очікування"):
        bar = getattr(self, "progress_bar", None)
        if bar is None:
            return
        if hasattr(bar, "stop"):
            bar.stop()
        bar.configure(mode="determinate")
        bar.set(0)
        label = getattr(self, "progress_label", None)
        if label is not None and message is not None:
            label.configure(text=message)
        self.update_idletasks()

    def _progress_update(self, current: int, total: int, stage: str = "Генерація"):
        bar = getattr(self, "progress_bar", None)
        if bar is None:
            return
        total = total or 0
        if total > 0:
            fraction = max(0.0, min(float(current) / float(total), 1.0))
        else:
            fraction = 0.0
        bar.configure(mode="determinate")
        bar.set(fraction)
        label = getattr(self, "progress_label", None)
        if label is not None:
            if total > 0:
                text = f"{stage}: {current} з {total}"
            else:
                text = f"{stage}: {current}"
            label.configure(text=text)
        self.update_idletasks()

    def _progress_message(self, message: str):
        label = getattr(self, "progress_label", None)
        if label is not None and message is not None:
            label.configure(text=message)
        self.update_idletasks()

    def _progress_finish(self, message: str = "Готово"):
        bar = getattr(self, "progress_bar", None)
        if bar is not None:
            if hasattr(bar, "stop"):
                bar.stop()
            bar.configure(mode="determinate")
            bar.set(1)
        label = getattr(self, "progress_label", None)
        if label is not None and message is not None:
            label.configure(text=message)
        self.update_idletasks()

    def _reload_gen_tree(self):
        tree = getattr(self, "_gen_tree", None)
        if tree is None:
            return

        prev_checked = self._collect_checked_model_ids()
        prev_open = set()

        def _gather_open(iid):
            if tree.item(iid, "open"):
                prev_open.add(iid)
            for child in tree.get_children(iid):
                _gather_open(child)

        for root_iid in tree.get_children(""):
            _gather_open(root_iid)

        tree.delete(*tree.get_children(""))
        self._gen_tree_states.clear()
        self._gen_tree_meta.clear()
        self._gen_tree_labels.clear()

        def _clean_label(value):
            if isinstance(value, str):
                return value.strip()
            if value is None:
                return ""
            return str(value)

        for cat_id, cat_name in get_categories():
            label = _clean_label(cat_name)
            cat_iid = f"cat_{cat_id}"
            self._gen_tree_labels[cat_iid] = label
            self._gen_tree_states[cat_iid] = 0
            self._gen_tree_meta[cat_iid] = {"type": "category", "id": cat_id}
            tree.insert("", "end", iid=cat_iid, text=f"{self._state_symbol(0)} {label}")

            for brand_id, brand_name in get_brands(cat_id):
                b_label = _clean_label(brand_name)
                brand_iid = f"brand_{brand_id}"
                self._gen_tree_labels[brand_iid] = b_label
                self._gen_tree_states[brand_iid] = 0
                self._gen_tree_meta[brand_iid] = {"type": "brand", "id": brand_id, "category_id": cat_id}
                tree.insert(cat_iid, "end", iid=brand_iid, text=f"{self._state_symbol(0)} {b_label}")

                for model_id, model_name in get_models(brand_id):
                    m_label = _clean_label(model_name)
                    model_iid = f"model_{model_id}"
                    self._gen_tree_labels[model_iid] = m_label
                    self._gen_tree_states[model_iid] = 0
                    self._gen_tree_meta[model_iid] = {"type": "model", "id": model_id, "brand_id": brand_id, "category_id": cat_id}
                    tree.insert(brand_iid, "end", iid=model_iid, text=f"{self._state_symbol(0)} {m_label}")

        if not prev_open:
            roots = tree.get_children("")
            if roots:
                tree.item(roots[0], open=True)
        else:
            for iid in prev_open:
                if tree.exists(iid):
                    tree.item(iid, open=True)

        for mid in sorted(prev_checked):
            iid = f"model_{mid}"
            if tree.exists(iid):
                self._set_gen_tree_state(iid, 2, propagate=False)
                self._update_parent_states(iid)

    def _state_symbol(self, state: int) -> str:
        if state == 2:
            return "☑"
        if state == 1:
            return "◪"
        return "☐"

    def _set_gen_tree_state(self, iid: str, state: int, propagate: bool = False):
        tree = getattr(self, "_gen_tree", None)
        if tree is None or not tree.exists(iid):
            return
        self._gen_tree_states[iid] = state
        label = self._gen_tree_labels.get(iid, tree.item(iid, "text"))
        display = f"{self._state_symbol(state)} {label}"
        tree.item(iid, text=display)
        if propagate:
            for child in tree.get_children(iid):
                self._set_gen_tree_state(child, state, propagate=True)

    def _update_parent_states(self, iid: str):
        tree = getattr(self, "_gen_tree", None)
        if tree is None:
            return
        parent = tree.parent(iid)
        if not parent:
            return
        child_states = [self._gen_tree_states.get(child, 0) for child in tree.get_children(parent)]
        if all(state == 2 for state in child_states):
            new_state = 2
        elif all(state == 0 for state in child_states):
            new_state = 0
        else:
            new_state = 1
        self._set_gen_tree_state(parent, new_state, propagate=False)
        self._update_parent_states(parent)

    def _collect_checked_model_ids(self):
        ids = set()
        for iid, state in self._gen_tree_states.items():
            if state != 2:
                continue
            meta = self._gen_tree_meta.get(iid)
            if meta and meta.get("type") == "model":
                mid = meta.get("id")
                if mid:
                    try:
                        ids.add(int(mid))
                    except (TypeError, ValueError):
                        continue
        return ids

    def _on_gen_tree_click(self, event):
        tree = getattr(self, "_gen_tree", None)
        if tree is None:
            return
        iid = tree.identify_row(event.y)
        if not iid:
            return
        element = tree.identify_element(event.x, event.y)
        if element != "text":
            return
        self._toggle_gen_tree_node(iid)

    def _toggle_selected_gen_node(self, event=None):
        tree = getattr(self, "_gen_tree", None)
        if tree is None:
            return "break"
        iid = tree.focus()
        if not iid:
            selection = tree.selection()
            if selection:
                iid = selection[0]
        if not iid:
            return "break"
        self._toggle_gen_tree_node(iid)
        return "break"

    def _toggle_gen_tree_node(self, iid: str):
        if iid not in self._gen_tree_states:
            return
        current = self._gen_tree_states.get(iid, 0)
        new_state = 0 if current == 2 else 2
        self._set_gen_tree_state(iid, new_state, propagate=True)
        self._update_parent_states(iid)

    def _set_gen_tree_open_recursive(self, iid: str, value: bool):
        tree = getattr(self, "_gen_tree", None)
        if tree is None or not tree.exists(iid):
            return
        tree.item(iid, open=value)
        for child in tree.get_children(iid):
            self._set_gen_tree_open_recursive(child, value)

    def _expand_all_gen_tree(self):
        tree = getattr(self, "_gen_tree", None)
        if tree is None:
            return
        for iid in tree.get_children(""):
            self._set_gen_tree_open_recursive(iid, True)

    def _collapse_all_gen_tree(self):
        tree = getattr(self, "_gen_tree", None)
        if tree is None:
            return
        for iid in tree.get_children(""):
            self._set_gen_tree_open_recursive(iid, False)

    def _select_all_gen_tree(self):
        tree = getattr(self, "_gen_tree", None)
        if tree is None:
            return
        for iid in tree.get_children(""):
            self._set_gen_tree_state(iid, 2, propagate=True)

    def _clear_all_gen_tree(self):
        tree = getattr(self, "_gen_tree", None)
        if tree is None:
            return
        for iid in tree.get_children(""):
            self._set_gen_tree_state(iid, 0, propagate=True)

    def _refresh_filmtype_checkboxes(self):
        for w in getattr(self, "filmtype_frame", []).winfo_children():
            w.destroy()
        self.ft_vars.clear()
        for item in self.templates["film_types"]:
            var = tk.BooleanVar(value=bool(item.get("enabled", True)))
            ctk.CTkCheckBox(self.filmtype_frame, text=item["name"], variable=var).pack(side="left", padx=6, pady=2)
            self.ft_vars.append((item["name"], var))
        self._refresh_template_scope_menu()

    def _choose_folder(self):
        folder = filedialog.askdirectory(title="Виберіть папку для файлів")
        if folder: self.out_folder_var.set(folder)

    def _preview_generation(self):
        # оновити шаблони/поля перед переглядом
        self._save_title_tags(show_message=False)
        self._export_apply_detail(save_to_file=False)

        selected_types = [name for name, var in self.ft_vars if var.get()]
        if not selected_types:
            return show_error("Оберіть хоча б один тип плівки.")

        for name, var in self.ft_vars:
            for item in self.templates["film_types"]:
                if item["name"] == name:
                    item["enabled"] = bool(var.get())
                    break
        save_templates(self.templates)

        selected_models = sorted(self._collect_checked_model_ids())

        try:
            if selected_models:
                records, columns = generate_export_rows(
                    selected_types,
                    self.templates,
                    self.title_tags_templates,
                    self.export_fields,
                    model_ids=selected_models,
                )
            else:
                records, columns = generate_export_rows(
                    selected_types,
                    self.templates,
                    self.title_tags_templates,
                    self.export_fields,
                )
        except ValueError as err:
            return show_error(str(err))

        if not records:
            return show_error("Немає даних для генерації (перевірте моделі).")

        preview_limit = 20
        preview_records = list(islice(records, preview_limit))

        if getattr(self, "_preview_window", None) is not None:
            try:
                if self._preview_window.winfo_exists():
                    self._preview_window.destroy()
            except Exception:
                pass

        preview_window = ctk.CTkToplevel(self)
        preview_window.title("Попередній перегляд генерації")
        preview_window.geometry("960x480")
        self._preview_window = preview_window

        info_text = f"Показано перші {len(preview_records)} з {len(records)} рядків."
        ctk.CTkLabel(preview_window, text=info_text, anchor="w").pack(fill="x", padx=14, pady=(12, 4))

        table_frame = ctk.CTkFrame(preview_window)
        table_frame.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)

        column_ids = [f"preview_col_{idx}" for idx in range(len(columns))]
        tree = ttk.Treeview(table_frame, columns=column_ids, show="headings")
        tree.grid(row=0, column=0, sticky="nsew")

        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        for col_id, header in zip(column_ids, columns):
            tree.heading(col_id, text=header)
            tree.column(col_id, anchor="w", stretch=True, width=160)

        for record in preview_records:
            values = _row_to_values(record, columns)
            tree.insert("", "end", values=values)

        if len(records) > preview_limit:
            note = f"(Доступно більше рядків: всього {len(records)}.)"
            ctk.CTkLabel(preview_window, text=note, anchor="w").pack(fill="x", padx=14, pady=(0, 10))

        ctk.CTkButton(preview_window, text="Закрити", command=preview_window.destroy).pack(pady=(0, 12))

    def _generate(self):
        self._progress_reset("Підготовка...")
        # зберегти (на випадок якщо змінювали шаблони перед тим)
        self._save_title_tags(show_message=False)
        self._export_apply_detail(save_to_file=False)

        selected_types = [name for name, var in self.ft_vars if var.get()]
        if not selected_types:
            self._progress_reset("Очікування")
            return show_error("Оберіть хоча б один тип плівки.")

        # оновимо enabled у файлі шаблонів
        for name, var in self.ft_vars:
            for item in self.templates["film_types"]:
                if item["name"] == name:
                    item["enabled"] = bool(var.get())
                    break
        save_templates(self.templates)

        # вибір моделей через дерево
        selected_models = sorted(self._collect_checked_model_ids())
        self._progress_message("Генерація даних...")

        def progress_callback(current, total):
            self._progress_update(current, total, stage="Генерація")

        try:
            if selected_models:
                records, columns = generate_export_rows(
                    selected_types,
                    self.templates,
                    self.title_tags_templates,
                    self.export_fields,
                    model_ids=selected_models,
                    progress_callback=progress_callback,
                )
            else:
                records, columns = generate_export_rows(
                    selected_types,
                    self.templates,
                    self.title_tags_templates,
                    self.export_fields,
                    progress_callback=progress_callback,
                )
        except ValueError as err:
            self._progress_reset("Помилка генерації")
            return show_error(str(err))

        if not records:
            self._progress_reset("Очікування")
            return show_error("Немає даних для генерації (перевірте моделі).")

        # експорт
        self._progress_message("Експорт файлів...")
        try:
            products_file = export_products(
                records,
                columns,
                self.export_fmt_var.get(),
                self.out_folder_var.get().strip(),
            )
        except Exception as e:
            self._progress_reset("Помилка експорту")
            return show_error(f"Не вдалося зберегти файли: {e}")

        self._progress_finish(f"Готово: {len(records)} рядків")
        msg = f"✅ Згенеровано {len(records)} рядків.\nФайл експорту: {products_file}"
        show_info(msg)

# ============================ ENTRY ============================

def main():
    init_db()
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()