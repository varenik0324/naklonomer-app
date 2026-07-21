import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import io
import re
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import openpyxl  # <-- добавлено для чтения через openpyxl

# ------------------------------------------------------------
# Настройки страницы
# ------------------------------------------------------------
st.set_page_config(
    page_title="Анализ наклономера + осадки",
    page_icon="📐",
    layout="wide"
)
st.title("📐 Анализ данных накладного инклинометра и осадок фундамента")
st.markdown("Загрузите Excel-файл с данными измерений, укажите параметры, и приложение построит графики, профили и сформирует отчёты.")

# ------------------------------------------------------------
# ПАРСИНГ ДАННЫХ НАКЛОНОМЕРА (через openpyxl)
# ------------------------------------------------------------
def parse_inclinometer_data(file_bytes, manual_floor_rows=None, search_start=None, search_end=None):
    """
    Парсит лист 'Наклономер' через openpyxl (data_only=True).
    Если manual_floor_rows задан (словарь {этаж: номер_строки}), использует его.
    Иначе ищет автоматически в диапазоне [search_start, search_end) или до появления 'Таблица'.
    Возвращает DataFrame с колонками: Цикл, Этаж, αx, αy.
    """
    # Загружаем книгу через openpyxl (data_only=True – читаем вычисленные значения)
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    
    # Находим лист с наклономером
    sheet = None
    for name in wb.sheetnames:
        if 'наклономер' in name.lower() or 'крен' in name.lower():
            sheet = wb[name]
            break
    if sheet is None:
        st.error("Не найден лист с данными наклономера. Проверьте файл.")
        return None

    total_rows = sheet.max_row
    total_cols = sheet.max_column

    # 1. Находим строку с заголовками циклов
    cycle_header_row = None
    for row_idx in range(1, total_rows + 1):  # openpyxl строки с 1
        row_values = [sheet.cell(row=row_idx, column=col).value for col in range(1, total_cols + 1)]
        row_str = ' '.join(str(cell) for cell in row_values if cell is not None)
        if 'Цикл' in row_str and re.search(r'\d{2}\.\d{2}\.\d{4}', row_str):
            cycle_header_row = row_idx
            break

    if cycle_header_row is None:
        st.warning("Не найдена строка с заголовками циклов. Будем использовать порядковые номера.")
        cycle_labels = None
    else:
        cycle_labels = []
        for col in range(2, total_cols + 1):  # с колонки B (индекс 2)
            cell = sheet.cell(row=cycle_header_row, column=col).value
            if cell is not None:
                cell_str = str(cell)
                match = re.search(r'(\d{2}\.\d{2}\.\d{4})', cell_str)
                if match:
                    try:
                        date_obj = pd.to_datetime(match.group(1), dayfirst=True)
                        cycle_labels.append(date_obj.strftime('%Y-%m-%d'))
                    except:
                        cycle_labels.append(cell_str)
                else:
                    cycle_labels.append(cell_str)
            else:
                cycle_labels.append("")
        cycle_labels = [c for c in cycle_labels if c]

    # 2. Определяем строки с этажами
    if manual_floor_rows is not None:
        floor_rows = {}
        for floor, row_idx in manual_floor_rows.items():
            if 1 <= row_idx <= total_rows:
                row_values = [sheet.cell(row=row_idx, column=col).value for col in range(1, total_cols + 1)]
                num_count = sum(1 for v in row_values if v is not None and isinstance(v, (int, float)))
                if num_count >= 6:
                    floor_rows[floor] = row_idx
                else:
                    st.warning(f"Строка {row_idx} содержит мало числовых данных, возможно, это не данные этажа {floor}.")
            else:
                st.warning(f"Строка {row_idx} выходит за пределы листа (всего строк: {total_rows}).")
        if len(floor_rows) < 2:
            st.error("Недостаточно валидных строк с этажами. Проверьте введённые индексы.")
            return None
    else:
        # Автоматический поиск
        if search_start is None:
            search_start = cycle_header_row + 1 if cycle_header_row is not None else 1
        if search_end is None:
            # Ищем до появления "Таблица" или до конца
            end_search = total_rows + 1
            for row_idx in range(search_start, total_rows + 1):
                row_values = [sheet.cell(row=row_idx, column=col).value for col in range(1, total_cols + 1)]
                row_str = ' '.join(str(cell) for cell in row_values if cell is not None)
                if 'Таблица' in row_str:
                    end_search = row_idx
                    break
            search_end = end_search
        else:
            search_end = min(search_end, total_rows)

        search_start = max(1, min(search_start, total_rows))
        search_end = max(search_start, search_end)

        floor_rows = {}
        for row_idx in range(search_start, search_end + 1):
            row_values = [sheet.cell(row=row_idx, column=col).value for col in range(1, total_cols + 1)]
            found_floor = None
            for cell in row_values:
                if cell is None:
                    continue
                cell_str = str(cell).strip()
                match = re.search(r'\b(5|15|27)\b', cell_str)
                if match:
                    found_floor = int(match.group(1))
                    break
            if found_floor is not None:
                num_count = sum(1 for v in row_values if v is not None and isinstance(v, (int, float)))
                if num_count >= 6:
                    floor_rows[found_floor] = row_idx
            if len(floor_rows) == 3:
                break

        if len(floor_rows) < 2:
            st.warning("Не удалось автоматически найти строки с этажами (5, 15).")
            # Покажем первые 30 строк для диагностики
            preview = []
            for r in range(1, min(31, total_rows + 1)):
                row_vals = [sheet.cell(row=r, column=c).value for c in range(1, min(10, total_cols + 1))]
                preview.append(row_vals)
            st.write("Первые 30 строк листа (первые 10 колонок):")
            st.dataframe(pd.DataFrame(preview))
            return None

    # 3. Сортируем и собираем данные
    floor_rows_sorted = [floor_rows[f] for f in sorted(floor_rows.keys())]

    max_pairs = 0
    for floor_idx in floor_rows_sorted:
        values = []
        for col in range(2, total_cols + 1):  # с колонки B
            cell = sheet.cell(row=floor_idx, column=col).value
            if cell is not None and isinstance(cell, (int, float)):
                values.append(float(cell))
        pairs = []
        if len(values) % 2 == 0:
            pairs = [(values[i], values[i+1]) for i in range(0, len(values), 2)]
        else:
            pairs = [(values[i], values[i+1]) for i in range(0, len(values)-1, 2)]
        if len(pairs) > max_pairs:
            max_pairs = len(pairs)

    if cycle_labels is None:
        cycle_labels = [f"Цикл {i+1}" for i in range(max_pairs)]
    elif len(cycle_labels) < max_pairs:
        for i in range(len(cycle_labels), max_pairs):
            cycle_labels.append(f"Цикл {i+1}")

    data = []
    for floor_idx in floor_rows_sorted:
        # Определяем этаж из ячейки с номером
        floor_val = None
        for col in range(1, total_cols + 1):
            cell = sheet.cell(row=floor_idx, column=col).value
            if cell is None:
                continue
            cell_str = str(cell).strip()
            match = re.search(r'\b(5|15|27)\b', cell_str)
            if match:
                floor_val = int(match.group(1))
                break
        if floor_val is None:
            continue

        values = []
        for col in range(2, total_cols + 1):
            cell = sheet.cell(row=floor_idx, column=col).value
            if cell is not None and isinstance(cell, (int, float)):
                values.append(float(cell))
        pairs = []
        if len(values) % 2 == 0:
            pairs = [(values[i], values[i+1]) for i in range(0, len(values), 2)]
        else:
            pairs = [(values[i], values[i+1]) for i in range(0, len(values)-1, 2)]

        for i, (ax, ay) in enumerate(pairs):
            if i < len(cycle_labels):
                data.append({
                    'Цикл': cycle_labels[i],
                    'Этаж': floor_val,
                    'αx': ax,
                    'αy': ay
                })

    if not data:
        st.error("Не удалось извлечь данные наклономера.")
        return None

    df = pd.DataFrame(data)
    df['Цикл'] = df['Цикл'].astype(str)
    df['Этаж'] = df['Этаж'].astype(int)
    df['αx'] = pd.to_numeric(df['αx'], errors='coerce')
    df['αy'] = pd.to_numeric(df['αy'], errors='coerce')
    df = df.dropna(subset=['αx', 'αy'])
    return df

# ------------------------------------------------------------
# ПАРСИНГ ДАННЫХ ОСАДОК (без изменений – через pandas)
# ------------------------------------------------------------
def parse_settlement_data(file_bytes, sheet_name, corner_marks, L, B):
    df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)
    ... # (остаётся как было)

# ------------------------------------------------------------
# ГЕНЕРАЦИЯ ОТЧЁТОВ (без изменений)
# ------------------------------------------------------------
def generate_excel_report(...):
    ...

def generate_pdf_report(...):
    ...

def generate_word_report(...):
    ...

# ------------------------------------------------------------
# ОСНОВНАЯ ЛОГИКА ПРИЛОЖЕНИЯ (без изменений, использует новые функции)
# ------------------------------------------------------------
uploaded_file = st.file_uploader(...)
if uploaded_file is not None:
    try:
        file_bytes = uploaded_file.read()
        xl = pd.ExcelFile(io.BytesIO(file_bytes))
        all_sheets = xl.sheet_names

        # --- Боковая панель ---
        # ... всё как было
        df_incl = parse_inclinometer_data(file_bytes, search_start=search_start, search_end=search_end)
        if df_incl is None:
            # Ручной ввод строк
            ...
        else:
            st.success(...)
        # Остальной код приложения полностью сохраняется

    except Exception as e:
        st.error(f"Ошибка обработки: {e}")
        st.code(traceback.format_exc())
