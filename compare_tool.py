#!/usr/bin/env python3
"""
Инструмент сравнения таблицы уставок (Word .docx) и файла конфигурации (XML).

Использование:
  GUI-режим:  запустить без аргументов (двойной щелчок по exe)
  CLI-режим:  compare_tool.exe <таблица.docx> <конфиг.xml> [--group N] [-o отчёт.txt]
"""

import re
import sys
import argparse
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime

try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
    HAS_TK = True
except ImportError:
    HAS_TK = False

# ─────────────────────────────────────────────────────────────────────────────
# Константы и утилиты
# ─────────────────────────────────────────────────────────────────────────────

ID_PATTERN = re.compile(r'\[(\d{6})\]')


def normalize_value(val: str) -> str:
    """
    Нормализация значения для сравнения:
      - "1 - предусмотрено"  →  "1"   (перечислимые настройки: берём только код)
      - "1000.000"           →  "1000"
      - "0.30"               →  "0.3"
      - "0,30"               →  "0.3"  (запятая вместо точки)
    """
    if val is None:
        return ''
    val = str(val).strip()
    val = re.sub(r'[\n\r\t]+', ' ', val)
    val = re.sub(r'\s+', ' ', val).strip()
    if not val:
        return ''

    # "N - описание" или "N – описание" → берём только N
    m = re.match(r'^(-?\d+(?:[.,]\d+)?)\s*[-–—]\s*\S', val)
    if m:
        val = m.group(1)

    # Попытка числового сравнения
    try:
        f = float(val.replace(',', '.').replace(' ', ''))
        if f == int(f):
            return str(int(f))
        # Убрать незначимые нули (через %g)
        return f'{f:.10g}'
    except (ValueError, OverflowError):
        return val.lower().strip()


# ─────────────────────────────────────────────────────────────────────────────
# Парсинг XML-конфига
# ─────────────────────────────────────────────────────────────────────────────

def parse_xml_config(xml_path: str, group: int = 1) -> dict:
    """
    Читает XML-файл конфигурации.
    Возвращает dict: '6-значный ID' → {id, name, value, unit, source, path, all_values}.
    Обрабатывает секции: Ratios, Initials, Settings (рекурсивно).
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()
    params: dict = {}

    # ── Ratios (коэффициенты ТТ/ТН) ──────────────────────────────────────────
    ratios_node = root.find('Ratios')
    if ratios_node is not None:
        for item in ratios_node:
            m = re.match(r'\d+:(\d+)', item.get('ID', ''))
            if not m:
                continue
            nid = m.group(1)
            val = item.get('Value', '')
            params[nid] = {
                'id': nid,
                'full_id': item.get('ID', ''),
                'name': f'[{nid}]',
                'value': val,
                'unit': item.get('Unit', ''),
                'source': 'Ratios',
                'path': 'Коэффициенты трансформации',
                'all_values': {1: val},
            }

    # ── Initials (общие сведения) ─────────────────────────────────────────────
    initials_node = root.find('Initials')
    if initials_node is not None:
        for item in initials_node:
            m = re.match(r'\d+:(\d+)', item.get('ID', ''))
            if not m:
                continue
            nid = m.group(1)
            val = item.get('Value', '')
            params[nid] = {
                'id': nid,
                'full_id': item.get('ID', ''),
                'name': f'[{nid}]',
                'value': val,
                'unit': '',
                'source': 'Initials',
                'path': 'Общие сведения',
                'all_values': {1: val},
            }

    # ── Settings (уставки защит, рекурсивно) ─────────────────────────────────
    settings_root = root.find('Settings')
    if settings_root is not None:
        def _collect(node, path: str = ''):
            node_name = node.get('Name', '')
            cur_path = (path + ' / ' + node_name).strip(' /')
            for child in node:
                if child.tag == 'Setting':
                    m = re.match(r'\d+:(\d+)', child.get('ID', ''))
                    if not m:
                        continue
                    nid = m.group(1)
                    # Собрать все группы значений
                    all_vals = {}
                    for g in range(1, 5):
                        v = child.get(f'Value{g}')
                        if v is not None:
                            all_vals[g] = v
                    # Значение для запрошенной группы
                    value = child.get(f'Value{group}') or child.get('Value1') or ''
                    params[nid] = {
                        'id': nid,
                        'full_id': child.get('ID', ''),
                        'name': child.get('Name', f'[{nid}]'),
                        'value': value,
                        'unit': child.get('Unit', ''),
                        'source': 'Settings',
                        'path': cur_path,
                        'all_values': all_vals,
                    }
                elif child.tag == 'Node':
                    _collect(child, cur_path)

        _collect(settings_root)

    return params


# ─────────────────────────────────────────────────────────────────────────────
# Парсинг Word-таблицы
# ─────────────────────────────────────────────────────────────────────────────

def _unique_cells(row):
    """
    Возвращает список уникальных ячеек строки таблицы Word.
    В объединённых ячейках (merged) один и тот же XML-элемент (_tc) встречается
    несколько раз — убираем дубли.
    """
    seen = set()
    result = []
    for cell in row.cells:
        tc_id = id(cell._tc)
        if tc_id not in seen:
            seen.add(tc_id)
            result.append(cell)
    return result


def _clean_text(cell) -> str:
    """Извлечь текст ячейки, нормализовать пробелы и переносы строк."""
    t = cell.text.strip()
    t = re.sub(r'[\n\r]+', ' ', t)
    t = re.sub(r'\s+', ' ', t)
    return t.strip()


def parse_word_table(docx_path: str) -> dict:
    """
    Читает .docx и извлекает параметры из всех таблиц.
    Ищет строки с шаблоном [XXXXXX] (6-значный ID) в любом столбце.

    Возвращает dict: '6-значный ID' → {id, name, value, unit, primary_values}.
    """
    doc = Document(docx_path)
    params: dict = {}

    for table in doc.tables:
        # Определить, есть ли в заголовке таблицы пометка "(первичные величины)"
        is_primary = False
        for row in table.rows[:2]:
            for cell in _unique_cells(row):
                if re.search(r'первичн', _clean_text(cell), re.IGNORECASE):
                    is_primary = True

        for row in table.rows:
            cells = [_clean_text(c) for c in _unique_cells(row)]

            # Найти ячейку с ID [XXXXXX]
            id_idx = None
            for i, txt in enumerate(cells):
                if ID_PATTERN.search(txt):
                    id_idx = i
                    break
            if id_idx is None:
                continue

            name_text = cells[id_idx]
            m = ID_PATTERN.search(name_text)
            if not m:
                continue
            nid = m.group(1)

            # Значение = следующая ячейка; единица измерения = ещё следующая
            value = cells[id_idx + 1] if id_idx + 1 < len(cells) else ''
            unit  = cells[id_idx + 2] if id_idx + 2 < len(cells) else ''

            # Пропустить строки-заголовки таблицы
            skip_re = r'значени|наименован|диапазон'
            if re.search(skip_re, value, re.IGNORECASE):
                continue

            # Первое вхождение — актуальное (не перезаписывать)
            if nid not in params:
                params[nid] = {
                    'id': nid,
                    'name': name_text,
                    'value': value,
                    'unit': unit,
                    'primary_values': is_primary,
                }

    return params


# ─────────────────────────────────────────────────────────────────────────────
# Сравнение
# ─────────────────────────────────────────────────────────────────────────────

def compare(word_params: dict, xml_params: dict) -> dict:
    """
    Возвращает dict с тремя списками:
      differences  — есть в обоих, но значения различаются
      only_word    — есть в таблице, нет в конфиге
      only_xml     — есть в конфиге (со значением), нет в таблице
    """
    all_ids = sorted(set(word_params) | set(xml_params))
    differences = []
    only_word = []
    only_xml = []

    for nid in all_ids:
        in_word = nid in word_params
        in_xml  = nid in xml_params

        if in_word and not in_xml:
            only_word.append(word_params[nid])

        elif in_xml and not in_word:
            xp = xml_params[nid]
            # Добавляем в "только в конфиге" только если значение не пустое
            if xp.get('value', '').strip():
                only_xml.append(xp)

        else:  # в обоих
            wp = word_params[nid]
            xp = xml_params[nid]
            w_norm = normalize_value(wp['value'])
            x_norm = normalize_value(xp['value'])

            # Пропустить, если оба пустые
            if not w_norm and not x_norm:
                continue

            if w_norm != x_norm:
                differences.append({
                    'id': nid,
                    'name': xp['name'],
                    'path': xp.get('path', ''),
                    'value_word': wp['value'],
                    'unit_word':  wp.get('unit', ''),
                    'value_xml':  xp['value'],
                    'unit_xml':   xp.get('unit', ''),
                    'primary_values': wp.get('primary_values', False),
                    'all_xml_values': xp.get('all_values', {}),
                })

    return {
        'differences': differences,
        'only_word':   only_word,
        'only_xml':    only_xml,
    }


# ─────────────────────────────────────────────────────────────────────────────
# Форматирование отчёта
# ─────────────────────────────────────────────────────────────────────────────

def format_report(results: dict,
                  word_path: str, xml_path: str,
                  word_params: dict, xml_params: dict,
                  group: int = 1) -> str:
    SEP = '=' * 72
    lines = []

    def add(text=''):
        lines.append(text)

    now = datetime.now().strftime('%d.%m.%Y %H:%M:%S')
    add(SEP)
    add('ОТЧЁТ СРАВНЕНИЯ ТАБЛИЦЫ УСТАВОК И ФАЙЛА КОНФИГУРАЦИИ')
    add(SEP)
    add(f'Сформирован:          {now}')
    add(f'Таблица уставок:      {word_path}')
    add(f'Файл конфигурации:    {xml_path}')
    add(f'Группа уставок XML:   {group}')
    add()
    add(f'Параметров в таблице:         {len(word_params)}')
    add(f'Параметров в конфиге (всего): {len(xml_params)}')
    add()

    # ── 1. Различия значений ─────────────────────────────────────────────────
    diffs = results['differences']
    add(SEP)
    add(f'1. РАЗЛИЧИЯ ЗНАЧЕНИЙ — {len(diffs)} позиций')
    add(SEP)
    if diffs:
        for d in diffs:
            add(f"  {d['name']}")
            if d.get('path'):
                add(f"    Раздел: {d['path']}")
            w_str = d['value_word'] + (f" {d['unit_word']}" if d['unit_word'] else '')
            x_str = d['value_xml']  + (f" {d['unit_xml']}"  if d['unit_xml']  else '')
            add(f"    Таблица:  {w_str or '(пусто)'}")
            add(f"    Конфиг:   {x_str or '(пусто)'}")
            if d.get('primary_values'):
                add(f"    * Значение в таблице указано в первичных единицах")
            # Показать все группы конфига, если их несколько
            av = d.get('all_xml_values', {})
            if len(av) > 1:
                groups_str = '  '.join(f'Гр{g}={v}' for g, v in sorted(av.items()))
                add(f"    Конфиг (все группы): {groups_str}")
            add()
    else:
        add('  Различий нет.')
        add()

    # ── 2. Есть в таблице, нет в конфиге ────────────────────────────────────
    ow = results['only_word']
    add(SEP)
    add(f'2. ЕСТЬ В ТАБЛИЦЕ, НЕТ В КОНФИГЕ — {len(ow)} позиций')
    add(SEP)
    if ow:
        for p in ow:
            add(f"  {p['name']}")
            v_str = p['value'] + (f" {p['unit']}" if p.get('unit') else '')
            if v_str.strip():
                add(f"    Значение в таблице: {v_str}")
            add()
    else:
        add('  Нет таких параметров.')
        add()

    # ── 3. Есть в конфиге, нет в таблице ────────────────────────────────────
    ox = results['only_xml']
    add(SEP)
    add(f'3. ЕСТЬ В КОНФИГЕ (со значением), НЕТ В ТАБЛИЦЕ — {len(ox)} позиций')
    add(SEP)
    if ox:
        for p in ox:
            add(f"  {p['name']}")
            if p.get('path'):
                add(f"    Раздел: {p['path']}")
            v_str = p['value'] + (f" {p['unit']}" if p.get('unit') else '')
            add(f"    Значение в конфиге: {v_str}")
            add()
    else:
        add('  Нет таких параметров.')
        add()

    add(SEP)
    add('КОНЕЦ ОТЧЁТА')
    add(SEP)
    return '\n'.join(lines)


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

def run_cli():
    if not HAS_DOCX:
        print('ОШИБКА: не установлен пакет python-docx. '
              'Выполните: pip install python-docx')
        sys.exit(1)

    ap = argparse.ArgumentParser(
        description='Сравнение таблицы уставок (Word) с файлом конфигурации (XML)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument('word', help='Таблица уставок (.docx)')
    ap.add_argument('xml',  help='Файл конфигурации (.xml)')
    ap.add_argument('--group', type=int, default=1, choices=[1, 2, 3, 4],
                    metavar='N', help='Группа уставок XML для сравнения (1-4, по умолч. 1)')
    ap.add_argument('-o', '--output', help='Файл для сохранения отчёта')
    args = ap.parse_args()

    print(f'Чтение таблицы уставок: {args.word} ...')
    word_params = parse_word_table(args.word)
    print(f'  Найдено параметров: {len(word_params)}')

    print(f'Чтение конфигурации:   {args.xml} ...')
    xml_params = parse_xml_config(args.xml, group=args.group)
    print(f'  Найдено параметров: {len(xml_params)}')

    print('Сравнение ...')
    results = compare(word_params, xml_params)
    report  = format_report(results, args.word, args.xml,
                            word_params, xml_params, group=args.group)

    if args.output:
        Path(args.output).write_text(report, encoding='utf-8-sig')
        print(f'Отчёт сохранён: {args.output}')
    else:
        print()
        print(report)


# ─────────────────────────────────────────────────────────────────────────────
# GUI
# ─────────────────────────────────────────────────────────────────────────────

def run_gui():
    if not HAS_TK:
        print('tkinter недоступен — используйте CLI-режим.')
        sys.exit(1)
    if not HAS_DOCX:
        messagebox.showerror(
            'Ошибка зависимостей',
            'Не найден пакет python-docx.\n'
            'Выполните: pip install python-docx'
        )
        sys.exit(1)

    root = tk.Tk()
    root.title('Сравнение таблицы уставок и конфигурации')
    root.geometry('950x720')
    root.minsize(700, 500)

    # Переменные формы
    word_var  = tk.StringVar()
    xml_var   = tk.StringVar()
    group_var = tk.IntVar(value=1)
    status_var = tk.StringVar(value='Готов к работе.')

    # ── Фрейм: выбор файлов ───────────────────────────────────────────────────
    frm_files = ttk.LabelFrame(root, text=' Файлы ', padding=8)
    frm_files.pack(fill='x', padx=12, pady=(10, 4))
    frm_files.columnconfigure(1, weight=1)

    ttk.Label(frm_files, text='Таблица уставок (.docx):').grid(
        row=0, column=0, sticky='w', padx=4, pady=3)
    ttk.Entry(frm_files, textvariable=word_var).grid(
        row=0, column=1, sticky='ew', padx=4)
    ttk.Button(frm_files, text='Обзор…',
               command=lambda: word_var.set(
                   filedialog.askopenfilename(
                       title='Таблица уставок',
                       filetypes=[('Word документы', '*.docx'), ('Все файлы', '*.*')]
                   ) or word_var.get()
               )).grid(row=0, column=2, padx=4)

    ttk.Label(frm_files, text='Файл конфигурации (.xml):').grid(
        row=1, column=0, sticky='w', padx=4, pady=3)
    ttk.Entry(frm_files, textvariable=xml_var).grid(
        row=1, column=1, sticky='ew', padx=4)
    ttk.Button(frm_files, text='Обзор…',
               command=lambda: xml_var.set(
                   filedialog.askopenfilename(
                       title='Файл конфигурации',
                       filetypes=[('XML файлы', '*.xml'), ('Все файлы', '*.*')]
                   ) or xml_var.get()
               )).grid(row=1, column=2, padx=4)

    # ── Фрейм: группа уставок ─────────────────────────────────────────────────
    frm_opts = ttk.Frame(root, padding=(12, 0))
    frm_opts.pack(fill='x')
    ttk.Label(frm_opts, text='Группа уставок XML:').pack(side='left')
    for i in range(1, 5):
        ttk.Radiobutton(frm_opts, text=f' {i} ', variable=group_var,
                        value=i).pack(side='left', padx=2)

    # ── Кнопки действий ──────────────────────────────────────────────────────
    frm_btns = ttk.Frame(root, padding=(12, 4))
    frm_btns.pack()

    _report_holder = {'text': None}

    def do_compare():
        word_path = word_var.get().strip()
        xml_path  = xml_var.get().strip()

        if not word_path:
            messagebox.showerror('Ошибка', 'Укажите файл таблицы уставок.')
            return
        if not xml_path:
            messagebox.showerror('Ошибка', 'Укажите файл конфигурации.')
            return
        if not Path(word_path).is_file():
            messagebox.showerror('Ошибка', f'Файл не найден:\n{word_path}')
            return
        if not Path(xml_path).is_file():
            messagebox.showerror('Ошибка', f'Файл не найден:\n{xml_path}')
            return

        status_var.set('Обработка…')
        root.update_idletasks()

        try:
            wp = parse_word_table(word_path)
            xp = parse_xml_config(xml_path, group=group_var.get())
            res = compare(wp, xp)
            report = format_report(res, word_path, xml_path, wp, xp,
                                   group=group_var.get())

            txt.config(state='normal')
            txt.delete('1.0', tk.END)
            txt.insert(tk.END, report)
            txt.config(state='disabled')

            _report_holder['text'] = report
            d  = len(res['differences'])
            ow = len(res['only_word'])
            ox = len(res['only_xml'])
            status_var.set(
                f'Готово.  Различий: {d}  |  '
                f'Только в таблице: {ow}  |  '
                f'Только в конфиге: {ox}'
            )
        except Exception as exc:
            messagebox.showerror('Ошибка при обработке', str(exc))
            status_var.set('Ошибка.')

    def do_save():
        report = _report_holder.get('text')
        if not report:
            messagebox.showinfo('Информация', 'Сначала выполните сравнение.')
            return
        out = filedialog.asksaveasfilename(
            title='Сохранить отчёт',
            defaultextension='.txt',
            initialfile='отчет_сравнения.txt',
            filetypes=[('Текстовый файл', '*.txt'), ('Все файлы', '*.*')],
        )
        if out:
            Path(out).write_text(report, encoding='utf-8-sig')
            messagebox.showinfo('Сохранено', f'Отчёт сохранён:\n{out}')

    ttk.Button(frm_btns, text='  Сравнить  ', command=do_compare).pack(
        side='left', padx=8)
    ttk.Button(frm_btns, text='  Сохранить отчёт  ', command=do_save).pack(
        side='left', padx=8)

    # ── Область результатов ───────────────────────────────────────────────────
    frm_res = ttk.LabelFrame(root, text=' Результат ', padding=4)
    frm_res.pack(fill='both', expand=True, padx=12, pady=4)

    txt = scrolledtext.ScrolledText(
        frm_res, font=('Courier New', 9), state='disabled', wrap='none')
    txt.pack(fill='both', expand=True)

    # ── Строка статуса ────────────────────────────────────────────────────────
    ttk.Label(root, textvariable=status_var, relief='sunken',
              anchor='w', padding=(8, 2)).pack(fill='x', side='bottom')

    root.mainloop()


# ─────────────────────────────────────────────────────────────────────────────
# Точка входа
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    # Если запущен с аргументами — CLI, иначе GUI
    if len(sys.argv) > 1 and not sys.argv[1].startswith('--help'):
        run_cli()
    else:
        if HAS_TK:
            run_gui()
        else:
            print('Укажите аргументы: compare_tool <таблица.docx> <конфиг.xml>')
            sys.exit(1)
