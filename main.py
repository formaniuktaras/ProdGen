# -*- coding: utf-8 -*-
import os
import json
import re
import sqlite3
import time
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
import pandas as pd
from jinja2 import Template

APP_TITLE = "Prom Generator"
DB_FILE = "catalog.db"
TEMPLATES_FILE = "templates.json"
SPECS_FILE = "specs.json"
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
        with open(TEMPLATES_FILE, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_TEMPLATES, f, ensure_ascii=False, indent=2)
        return DEFAULT_TEMPLATES.copy()
    with open(TEMPLATES_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
    # гарантуємо всі ключі
    for k, v in DEFAULT_TEMPLATES.items():
        if k not in data:
            data[k] = v
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
    export_specs_json()

def update_spec(spec_id: int, key: str, value: str):
    key = key.strip()
    if not key: return
    conn = db_connect(); cur = conn.cursor()
    cur.execute("UPDATE model_specs SET key=?, value=? WHERE id=?",
                (key, value, spec_id))
    conn.commit(); conn.close()
    export_specs_json()

def delete_spec(spec_id: int):
    conn = db_connect(); cur = conn.cursor()
    cur.execute("DELETE FROM model_specs WHERE id=?", (spec_id,))
    conn.commit(); conn.close()
    export_specs_json()

def export_specs_json():
    """Експорт усіх характеристик у переносимий SPECS_FILE."""
    conn = db_connect(); cur = conn.cursor()
    cur.execute("""
        SELECT m.name, s.key, s.value
        FROM model_specs s
        JOIN models m ON s.model_id = m.id
        ORDER BY m.name, s.id
    """)
    rows = cur.fetchall(); conn.close()
    data = {}
    for model, key, value in rows:
        if isinstance(model, str):
            model = model.strip()
        data.setdefault(model, {})[key] = value
    with open(SPECS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

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
        SELECT b.name, m.name, c.name, m.id
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
    for brand, model, cat, mid in rows:
        if isinstance(brand, str):
            brand = brand.strip()
        if isinstance(model, str):
            model = model.strip()
        if isinstance(cat, str):
            cat = cat.strip()
        cleaned.append((brand, model, cat, mid))
    return cleaned
def generate_dataset(film_types: list, templates: dict, title_tags_templates: dict,
                     category_ids=None, brand_ids=None, model_ids=None):
    pairs = collect_models(category_ids=category_ids, brand_ids=brand_ids, model_ids=model_ids)

    records = []
    title_tags_cache = {}
    for brand, model, cat, _mid in pairs:
        # блок описів для категорії
        cat_desc_block = templates.get("descriptions", {}).get(cat, {})
        for f in film_types:
            if f not in title_tags_cache:
                title_tpl_str, tags_tpl_str = resolve_title_tags(title_tags_templates, templates, f)
                title_tags_cache[f] = (Template(title_tpl_str), Template(tags_tpl_str))
            title_t, tags_t = title_tags_cache[f]
            desc_template_str = cat_desc_block.get(f, cat_desc_block.get("default", "Плівка для {{ brand }} {{ model }}"))
            desc_t = Template(desc_template_str)

            records.append({
                "Категорія": cat,
                "Бренд": brand,
                "Модель": model,
                "Тип плівки": f,
                "Назва": title_t.render(film_type=f, brand=brand, model=model),
                "Опис": desc_t.render(film_type=f, brand=brand, model=model),
                "Теги": tags_t.render(film_type=f, brand=brand, model=model)
            })
    return records

def export_products_and_specs(records: list, fmt: str, folder: str):
    ensure_folder(folder)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = os.path.join(folder, f"products_{ts}")

    # ТОВАРИ
    if fmt == "Excel (.xlsx)":
        out_products = base + ".xlsx"
        pd.DataFrame.from_records(records).to_excel(out_products, index=False)
    elif fmt == "CSV (.csv)":
        out_products = base + ".csv"
        pd.DataFrame.from_records(records).to_csv(out_products, index=False, encoding="utf-8-sig")
    else:
        out_products = base + ".json"
        with open(out_products, "w", encoding="utf-8") as f:
            json.dump(records, f, ensure_ascii=False, indent=2)

    # ХАРАКТЕРИСТИКИ → окремий Excel/CSV/JSON з трьома колонками
    conn = db_connect(); cur = conn.cursor()
    cur.execute("""
        SELECT m.name AS model, s.key, s.value
        FROM model_specs s
        JOIN models m ON s.model_id = m.id
        ORDER BY m.name, s.id
    """)
    specs_rows = cur.fetchall(); conn.close()

    out_specs = None
    if specs_rows:
        df_specs = pd.DataFrame(specs_rows, columns=["Модель", "Характеристика", "Значення"])
        if fmt == "Excel (.xlsx)":
            out_specs = base + "_specs.xlsx"
            df_specs.to_excel(out_specs, index=False)
        elif fmt == "CSV (.csv)":
            out_specs = base + "_specs.csv"
            df_specs.to_csv(out_specs, index=False, encoding="utf-8-sig")
        else:
            out_specs = base + "_specs.json"
            with open(out_specs, "w", encoding="utf-8") as f:
                json.dump(df_specs.to_dict(orient="records"), f, ensure_ascii=False, indent=2)

    return out_products, out_specs

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

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1100x680")
        self.minsize(980, 640)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.templates = load_templates()
        self.title_tags_templates = load_title_tags_templates(self.templates)
        self.current_category_id = None
        self.current_brand_id = None
        self._current_film_type_key = None
        self._film_type_display_to_key = {}
        self._film_type_key_to_display = {}
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

        self._build_header()
        self._build_tabs()

        # Початкові дані
        self._refresh_categories()
        self._refresh_filmtype_checkboxes()

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
        self.tab_generate  = tabs.add("Генерація")

        self._build_tab_catalog()
        self._build_tab_templates()
        self._build_tab_generate()

    def _film_type_names(self):
        return [item.get("name") for item in self.templates.get("film_types", []) if item.get("name")]

    def _film_type_menu_items(self):
        items = [(FILM_TYPE_DEFAULT_LABEL, "default")]
        for name in self._film_type_names():
            items.append((name, name))
        return items

    def _film_type_key_from_display(self, label: str) -> str:
        if not label:
            return "default"
        return self._film_type_display_to_key.get(label, "default")

    def _selected_film_type_key(self) -> str:
        key = getattr(self, "_current_film_type_key", None)
        if key == "default":
            return "default"
        names = set(self._film_type_names())
        if key in names:
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
        ctk.CTkLabel(selector, text="Тип плівки:").pack(side="left", padx=(0,6))
        self.film_type_var = tk.StringVar(value="")
        self.film_type_menu = ctk.CTkOptionMenu(
            selector,
            values=["—"],
            variable=self.film_type_var,
            width=220,
            command=self._on_film_type_change
        )
        self.film_type_menu.pack(side="left")

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

        top = ctk.CTkFrame(right); top.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(top, text="Категорія:").pack(side="left", padx=(0,6))
        self.desc_cat_var = tk.StringVar(value=list(self.templates["descriptions"].keys())[0])
        self.desc_cat_menu = ctk.CTkOptionMenu(top, values=list(self.templates["descriptions"].keys()),
                                               variable=self.desc_cat_var, width=200,
                                               command=lambda _v: self._load_desc_template())
        self.desc_cat_menu.pack(side="left")

        ctk.CTkLabel(right, text="Шаблон опису (доступні {{ brand }}, {{ model }}, {{ film_type }})").pack(anchor="w", padx=10, pady=(10,0))
        self.desc_box = ctk.CTkTextbox(right)
        self.desc_box.pack(fill="both", expand=True, padx=10, pady=5)

        btn_row = ctk.CTkFrame(right); btn_row.pack(fill="x", padx=10, pady=6)
        ctk.CTkButton(btn_row, text="Зберегти опис", command=self._save_desc_template).pack(side="right")

        self._refresh_film_type_menu()

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

    def _refresh_film_type_menu(self):
        if not hasattr(self, "film_type_menu"):
            return
        items = self._film_type_menu_items()
        display_values = [label for label, _ in items]
        self._film_type_display_to_key = {label: key for label, key in items}
        self._film_type_key_to_display = {key: label for label, key in items}
        self.film_type_menu.configure(values=display_values)

        valid_keys = set(self._film_type_key_to_display.keys())
        current_key = getattr(self, "_current_film_type_key", None)
        if current_key not in valid_keys:
            if len(items) > 1:
                current_key = items[1][1]
            else:
                current_key = items[0][1]
        self._current_film_type_key = current_key
        current_display = self._film_type_key_to_display[current_key]
        self.film_type_var.set(current_display)
        self.film_type_menu.set(current_display)
        self._on_film_type_change()

    def _on_film_type_change(self, selected_label=None):
        if selected_label is not None:
            self._current_film_type_key = self._film_type_key_from_display(selected_label)
        elif not getattr(self, "_current_film_type_key", None):
            current_label = self.film_type_var.get() if hasattr(self, "film_type_var") else None
            self._current_film_type_key = self._film_type_key_from_display(current_label)
        self._load_title_tags_template()
        self._load_desc_template()

    def _load_desc_template(self):
        if not hasattr(self, "desc_box") or not hasattr(self, "desc_cat_var"):
            return
        cat = self.desc_cat_var.get()
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
        self.export_fmt_var = tk.StringVar(value="Excel (.xlsx)")
        self.export_fmt_menu = ctk.CTkOptionMenu(
            fmt_frame,
            values=["Excel (.xlsx)", "CSV (.csv)", "JSON (.json)"],
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
        ctk.CTkButton(action_row, text="Згенерувати", command=self._generate, height=36).pack(side="right", padx=6)

        self._reload_gen_tree()

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
        self._refresh_film_type_menu()

    def _choose_folder(self):
        folder = filedialog.askdirectory(title="Виберіть папку для файлів")
        if folder: self.out_folder_var.set(folder)

    def _generate(self):
        # зберегти (на випадок якщо змінювали шаблони перед тим)
        self._save_title_tags(show_message=False)

        selected_types = [name for name, var in self.ft_vars if var.get()]
        if not selected_types:
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
        if selected_models:
            records = generate_dataset(
                selected_types,
                self.templates,
                self.title_tags_templates,
                model_ids=selected_models,
            )
        else:
            records = generate_dataset(
                selected_types,
                self.templates,
                self.title_tags_templates,
            )
        if not records:
            return show_error("Немає даних для генерації (перевірте моделі).")

        # експорт
        try:
            products_file, specs_file = export_products_and_specs(
                records, self.export_fmt_var.get(), self.out_folder_var.get().strip()
            )
        except Exception as e:
            return show_error(f"Не вдалося зберегти файли: {e}")

        msg = f"✅ Згенеровано {len(records)} товарів.\nФайл товарів: {products_file}"
        if specs_file: msg += f"\nФайл характеристик: {specs_file}"
        show_info(msg)

# ============================ ENTRY ============================

def main():
    init_db()
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()