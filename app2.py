import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import io
import re
from datetime import datetime
from typing import Optional, Tuple, List, Dict, Any

# ------------------------------------------------------------
# КОНСТАНТЫ
# ------------------------------------------------------------
FLOORS_NEEDED = [5, 15, 27]
TABLE_INCLINOMETER = "Таблица 8"
TABLE_SETTLEMENT_END = "Таблица 9"
DEFAULT_BUILDING_LENGTH = 70.46
DEFAULT_BUILDING_WIDTH = 18.69
DEFAULT_VERTICAL_SCALE = 1.0

# ------------------------------------------------------------
# НАСТРОЙКИ СТРАНИЦЫ И СОСТОЯНИЯ
# ------------------------------------------------------------
st.set_page_config(
    page_title="Анализ наклономера + 3D-модель здания",
    page_icon="📐",
    layout="wide"
)
st.title("📐 3D-модель здания по данным накладного инклинометра и осадок")
st.markdown(
    "Загрузите Excel-файл с данными измерений, укажите параметры, "
    "и приложение построит 3D-модель деформаций здания."
)

# Инициализация состояния сессии
if 'building_length' not in st.session_state:
    st.session_state.building_length = DEFAULT_BUILDING_LENGTH
if 'building_width' not in st.session_state:
    st.session_state.building_width = DEFAULT_BUILDING_WIDTH
if 'vertical_scale' not in st.session_state:
    st.session_state.vertical_scale = DEFAULT_VERTICAL_SCALE
if 'current_index' not in st.session_state:
    st.session_state.current_index = 0
if 'selected_floors' not in st.session_state:
    st.session_state.selected_floors = []

# ------------------------------------------------------------
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ (кэшируемые)
# ------------------------------------------------------------
@st.cache_resource
def load_excel_file(file_bytes: bytes) -> pd.ExcelFile:
    """Загружает Excel-файл из байтов и возвращает объект ExcelFile."""
    return pd.ExcelFile(io.BytesIO(file_bytes))

@st.cache_data
def extract_cycle_headers(df_raw: pd.DataFrame) -> Tuple[Optional[int], List[str]]:
    """
    Находит строку с заголовками циклов и извлекает полные названия.
    Возвращает (индекс строки, список названий циклов).
    """
    total_cols = len(df_raw.columns)
    for idx, row in df_raw.iterrows():
        row_str = ' '.join(str(cell) for cell in row if pd.notna(cell))
        if 'Цикл' in row_str and re.search(r'\d{2}\.\d{2}\.\d{4}', row_str):
            cycle_names = []
            for col in range(1, total_cols):
                cell = df_raw.iloc[idx, col]
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    if 'Цикл' in cell_str:
                        cycle_names.append(' '.join(cell_str.split()))
            return idx, cycle_names
    return None, []

def parse_cycle_labels(cycle_names: List[str]) -> List[str]:
    """Преобразует названия циклов в ключи-даты (YYYY-MM-DD)."""
    labels = []
    for name in cycle_names:
        match = re.search(r'(\d{2}\.\d{2}\.\d{4})', name)
        if match:
            try:
                date_obj = pd.to_datetime(match.group(1), dayfirst=True)
                labels.append(date_obj.strftime('%Y-%m-%d'))
            except:
                labels.append(name)
        else:
            labels.append(name)
    if len(labels) != len(cycle_names):
        return cycle_names.copy()
    return labels

@st.cache_data
def find_floor_rows(df_raw: pd.DataFrame, start_row: int, end_row: int) -> Dict[int, int]:
    """
    Ищет строки с номерами этажей в первых двух столбцах.
    Возвращает словарь {этаж: индекс_строки}.
    """
    floor_rows = {}
    for idx in range(start_row, end_row):
        row = df_raw.iloc[idx]
        found_floor = None
        for col in [0, 1]:
            if col < len(row):
                cell = row.iloc[col] if hasattr(row, 'iloc') else row[col]
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    match = re.search(r'\b(\d+)\b', cell_str)
                    if match:
                        found_floor = int(match.group(1))
                        break
        if found_floor is not None:
            num_count = sum(1 for v in row if pd.notna(v) and isinstance(v, (int, float)))
            if num_count >= 6:
                floor_rows[found_floor] = idx
    return floor_rows

@st.cache_data
def extract_angle_pairs(df_raw: pd.DataFrame, floor_idx: int, total_cols: int) -> List[Tuple[float, float]]:
    """Извлекает пары (αx, αy) из строки, начиная с колонки 1."""
    values = []
    for col in range(1, total_cols):
        cell = df_raw.iloc[floor_idx, col]
        if pd.notna(cell) and isinstance(cell, (int, float)):
            values.append(float(cell))
    pairs = []
    if len(values) % 2 == 0:
        pairs = [(values[i], values[i+1]) for i in range(0, len(values), 2)]
    else:
        pairs = [(values[i], values[i+1]) for i in range(0, len(values)-1, 2)]
    return pairs

# ------------------------------------------------------------
# ПАРСИНГ ДАННЫХ НАКЛОНОМЕРА
# ------------------------------------------------------------
@st.cache_data
def parse_inclinometer_data(
    file_bytes: bytes,
    sheet_name: str,
    manual_floor_rows: Optional[Dict[int, int]] = None,
    search_start: Optional[int] = None,
    search_end: Optional[int] = None
) -> Optional[pd.DataFrame]:
    """
    Парсит лист Excel с данными наклономера.
    Возвращает DataFrame с колонками: Цикл, Цикл_полное, Этаж, αx, αy.
    """
    try:
        df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)
    except Exception as e:
        st.error(f"Не удалось прочитать лист '{sheet_name}': {e}")
        return None

    total_rows = len(df_raw)
    total_cols = len(df_raw.columns)

    # --- Извлечение заголовков циклов ---
    cycle_header_row, cycle_names = extract_cycle_headers(df_raw)
    if cycle_header_row is None:
        st.warning("Не найдена строка с заголовками циклов. Будем использовать порядковые номера.")
        cycle_labels = None
        cycle_full_names = None
    else:
        cycle_labels = parse_cycle_labels(cycle_names)
        cycle_full_names = cycle_names
        if len(cycle_labels) != len(cycle_full_names):
            cycle_labels = cycle_full_names.copy()

    # --- Поиск строк с этажами ---
    if manual_floor_rows is not None:
        floor_rows = {}
        for floor, row_idx in manual_floor_rows.items():
            if 0 <= row_idx < total_rows:
                row = df_raw.iloc[row_idx]
                num_count = sum(1 for v in row if pd.notna(v) and isinstance(v, (int, float)))
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
        start_search = None
        end_search = None
        for idx, row in df_raw.iterrows():
            row_str = ' '.join(str(cell) for cell in row if pd.notna(cell))
            if TABLE_INCLINOMETER in row_str:
                start_search = idx + 1
                break
        if start_search is not None:
            for idx in range(start_search, total_rows):
                row_str = ' '.join(str(cell) for cell in df_raw.iloc[idx] if pd.notna(cell))
                if TABLE_SETTLEMENT_END in row_str or (TABLE_INCLINOMETER in row_str and idx > start_search):
                    end_search = idx
                    break
        if start_search is None:
            start_search = search_start if search_start is not None else (cycle_header_row + 1 if cycle_header_row is not None else 0)
            if search_end is None:
                end_search = total_rows
                for idx in range(start_search, total_rows):
                    row_str = ' '.join(str(cell) for cell in df_raw.iloc[idx] if pd.notna(cell))
                    if 'Таблица' in row_str:
                        end_search = idx
                        break
            else:
                end_search = search_end
        start_search = max(0, min(start_search, total_rows - 1))
        end_search = max(start_search, min(end_search, total_rows))

        floor_rows = find_floor_rows(df_raw, start_search, end_search)
        if len(floor_rows) < 2:
            st.warning("Не удалось автоматически найти строки с этажами. Попробуйте ручной ввод.")
            st.write("Первые 30 строк листа (первые 10 колонок):")
            st.dataframe(df_raw.iloc[:30, :10])
            return None

    floor_rows_sorted = [floor_rows[f] for f in sorted(floor_rows.keys())]

    # --- Определяем максимальное количество пар углов ---
    max_pairs = 0
    for floor_idx in floor_rows_sorted:
        pairs = extract_angle_pairs(df_raw, floor_idx, total_cols)
        if len(pairs) > max_pairs:
            max_pairs = len(pairs)

    # --- Дополняем метки циклов ---
    if cycle_labels is None:
        cycle_labels = [f"Цикл {i+1}" for i in range(max_pairs)]
        cycle_full_names = cycle_labels.copy()
    elif len(cycle_labels) < max_pairs:
        for i in range(len(cycle_labels), max_pairs):
            cycle_labels.append(f"Цикл {i+1}")
            cycle_full_names.append(f"Цикл {i+1}")

    # --- Сбор данных ---
    data = []
    for floor_idx in floor_rows_sorted:
        floor_val = None
        for col in [0, 1]:
            if col < len(df_raw.columns):
                cell = df_raw.iloc[floor_idx, col]
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    match = re.search(r'\b(\d+)\b', cell_str)
                    if match:
                        floor_val = int(match.group(1))
                        break
        if floor_val is None:
            continue

        pairs = extract_angle_pairs(df_raw, floor_idx, total_cols)
        for i, (ax, ay) in enumerate(pairs):
            if i < len(cycle_labels):
                data.append({
                    'Цикл': cycle_labels[i],
                    'Цикл_полное': cycle_full_names[i] if cycle_full_names and i < len(cycle_full_names) else cycle_labels[i],
                    'Этаж': floor_val,
                    'αx': ax,
                    'αy': ay
                })

    if not data:
        st.error("Не удалось извлечь данные наклономера.")
        return None

    df = pd.DataFrame(data)
    df['Цикл'] = df['Цикл'].astype(str)
    df['Цикл_полное'] = df['Цикл_полное'].astype(str)
    df['Этаж'] = df['Этаж'].astype(int)
    df['αx'] = pd.to_numeric(df['αx'], errors='coerce')
    df['αy'] = pd.to_numeric(df['αy'], errors='coerce')
    df = df.dropna(subset=['αx', 'αy'])
    return df

# ------------------------------------------------------------
# ПАРСИНГ ОСАДОК (исправленная версия с надёжным поиском колонок)
# ------------------------------------------------------------
@st.cache_data
def parse_settlement_data(
    file_bytes: bytes,
    sheet_name: str,
    corner_marks: List[str],
    L: float,
    B: float,
    mark_col: int = 0,
    zero_cycle_sett: Optional[str] = None,
    manual_osad_col: Optional[int] = None
):
    """
    Парсит лист Excel с данными осадок, вычисляет углы крена по осадкам.
    Возвращает (DataFrame с углами, словарь с абсолютными осадками, список всех циклов).
    """
    df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)

    # ========== ОТЛАДКА: покажем структуру листа ==========
    st.write("=== ОТЛАДКА ПАРСИНГА ОСАДОК ===")
    st.write(f"Лист: {sheet_name}")
    st.write("Первые 20 строк (первые 10 колонок):")
    st.dataframe(df_raw.iloc[:20, :10])

    # --- Извлечение заголовков циклов ---
    cycle_header_row, cycle_names = extract_cycle_headers(df_raw)
    st.write(f"Строка с заголовками циклов: {cycle_header_row}")
    st.write(f"Найденные названия циклов: {cycle_names}")

    if cycle_header_row is None:
        st.error("Не найдена строка с заголовками циклов в листе осадок.")
        return None

    # --- Определение колонок с осадками (улучшенная логика) ---
    cycle_cols = {}
    all_cycles = []

    # Если задан ручной номер колонки, используем его для всех циклов
    if manual_osad_col is not None and manual_osad_col >= 0:
        st.write(f"Ручной выбор колонки осадок: {manual_osad_col}")
        # Найдём все колонки с заголовками циклов
        for col_idx, cell in df_raw.iloc[cycle_header_row, :].items():
            if pd.notna(cell):
                cell_str = str(cell).strip()
                if 'Цикл' in cell_str:
                    date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', cell_str)
                    if date_match:
                        try:
                            date_obj = pd.to_datetime(date_match.group(1), dayfirst=True)
                            cycle_label = date_obj.strftime('%Y-%m-%d')
                        except:
                            cycle_label = cell_str
                    else:
                        cycle_label = cell_str
                    all_cycles.append(cycle_label)
                    if manual_osad_col < len(df_raw.columns):
                        cycle_cols[cycle_label] = manual_osad_col
                    else:
                        st.error(f"Столбец {manual_osad_col} выходит за пределы листа.")
                        return None
    else:
        st.write("Автоматический поиск колонок с осадками...")
        # Сначала найдём все колонки, где в заголовке есть "осадк" (регистронезависимо)
        osad_cols = {}
        for col_idx, cell in df_raw.iloc[cycle_header_row, :].items():
            if pd.notna(cell):
                cell_str = str(cell).strip().lower()
                if 'осадк' in cell_str:
                    osad_cols[col_idx] = cell_str
                    st.write(f"  Найдена колонка с осадками: {col_idx} ('{cell_str}')")

        # Теперь проходим по колонкам с циклами
        for col_idx, cell in df_raw.iloc[cycle_header_row, :].items():
            if pd.notna(cell):
                cell_str = str(cell).strip()
                if 'Цикл' in cell_str:
                    date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', cell_str)
                    if date_match:
                        try:
                            date_obj = pd.to_datetime(date_match.group(1), dayfirst=True)
                            cycle_label = date_obj.strftime('%Y-%m-%d')
                        except:
                            cycle_label = cell_str
                    else:
                        cycle_label = cell_str
                    all_cycles.append(cycle_label)
                    st.write(f"  Цикл '{cycle_label}' найден в колонке {col_idx}")

                    # Ищем колонку с осадками для этого цикла
                    # Сначала проверим, нет ли в osad_cols колонки, которая находится рядом (в пределах 3)
                    found = False
                    for offset in range(1, 4):
                        check_col = col_idx + offset
                        if check_col in osad_cols:
                            cycle_cols[cycle_label] = check_col
                            st.write(f"    -> Колонка с осадками для этого цикла: {check_col} (найдена по слову 'осадк')")
                            found = True
                            break
                    if not found:
                        # Если не нашли по слову, пробуем взять колонку +2 (типичная структура)
                        default_col = col_idx + 2
                        if default_col < len(df_raw.columns):
                            cycle_cols[cycle_label] = default_col
                            st.write(f"    -> Колонка с осадками не найдена по слову 'осадк', берём по умолчанию: {default_col}")
                        else:
                            # fallback
                            cycle_cols[cycle_label] = col_idx + 1
                            st.write(f"    -> Колонка с осадками не найдена, берём по умолчанию: {col_idx + 1}")

    st.write(f"Итоговые колонки осадок: {cycle_cols}")
    st.write(f"Все найденные циклы: {all_cycles}")

    if not cycle_cols:
        st.error("Не найдены колонки с осадками для циклов.")
        return None

    # --- Определяем нулевой цикл ---
    if zero_cycle_sett is None:
        def try_parse_date(s):
            for fmt in ('%Y-%m-%d', '%d.%m.%Y', '%Y-%m-%d'):
                try:
                    return pd.to_datetime(s, format=fmt)
                except:
                    continue
            return None
        try:
            sorted_cycles = sorted(all_cycles, key=lambda x: try_parse_date(x) or x)
        except:
            sorted_cycles = sorted(all_cycles)
        zero_cycle_sett = sorted_cycles[0] if sorted_cycles else None
        st.write(f"Нулевой цикл определён автоматически: {zero_cycle_sett}")
    else:
        st.write(f"Нулевой цикл выбран пользователем: {zero_cycle_sett}")

    # --- Поиск строк с марками ---
    mark_rows = []
    for idx in range(cycle_header_row + 1, len(df_raw)):
        cell_val = df_raw.iloc[idx, mark_col]
        if pd.notna(cell_val):
            try:
                float(cell_val)
                mark_rows.append(idx)
            except (ValueError, TypeError):
                if isinstance(cell_val, str) and cell_val.strip():
                    text_lower = cell_val.strip().lower()
                    if not any(word in text_lower for word in ['нет', 'доступ', 'нов', 'уничтож', 'примечание', 'таблица']):
                        mark_rows.append(idx)

    st.write(f"Найдены строки с марками (индексы): {mark_rows}")

    if not mark_rows:
        st.warning(f"Не найдены строки с марками в столбце {mark_col}.")
        return None

    # --- Извлечение абсолютных осадок ---
    marks_abs_data = {}
    for cycle_label, col_idx in cycle_cols.items():
        marks_abs_data[cycle_label] = {}
        for mark_idx in mark_rows:
            mark_num = df_raw.iloc[mark_idx, mark_col]
            if isinstance(mark_num, (int, float)):
                mark_str = str(int(mark_num)) if mark_num == int(mark_num) else str(mark_num)
            else:
                mark_str = str(mark_num).strip()
            settlement = df_raw.iloc[mark_idx, col_idx]
            if pd.notna(settlement) and isinstance(settlement, (int, float)):
                marks_abs_data[cycle_label][mark_str] = settlement
            else:
                marks_abs_data[cycle_label][mark_str] = np.nan

    st.write("Извлечённые данные осадок по циклам и маркам (только выбранные):")
    for cycle, marks in marks_abs_data.items():
        # покажем только для выбранных марок
        filtered = {m: marks.get(m, np.nan) for m in corner_marks}
        st.write(f"  Цикл {cycle}: {filtered}")

    # --- Определяем нулевые осадки ---
    zero_marks = {}
    for mark in corner_marks:
        mark_str = str(mark)
        if zero_cycle_sett in marks_abs_data and mark_str in marks_abs_data[zero_cycle_sett]:
            zero_marks[mark_str] = marks_abs_data[zero_cycle_sett][mark_str]
        else:
            zero_marks[mark_str] = np.nan

    st.write(f"Нулевые осадки для выбранных марок: {zero_marks}")

    if all(np.isnan(list(zero_marks.values()))):
        st.warning(f"Нулевой цикл {zero_cycle_sett} не содержит данных для выбранных марок. Используем первый доступный цикл.")
        first_cycle = list(marks_abs_data.keys())[0]
        for mark in corner_marks:
            mark_str = str(mark)
            if mark_str in marks_abs_data[first_cycle]:
                zero_marks[mark_str] = marks_abs_data[first_cycle][mark_str]
            else:
                zero_marks[mark_str] = np.nan
        zero_cycle_sett = first_cycle

    st.write(f"После корректировки нулевой цикл: {zero_cycle_sett}, нулевые осадки: {zero_marks}")

    if all(np.isnan(list(zero_marks.values()))):
        st.error("Не удалось найти данные для выбранных угловых марок ни в одном цикле.")
        return None

    # --- Вычисление приростов осадок ---
    data = []
    for cycle_label, cols in marks_abs_data.items():
        if cycle_label == zero_cycle_sett:
            continue
        for mark_str, sett_val in cols.items():
            if mark_str in zero_marks and not np.isnan(zero_marks[mark_str]) and not np.isnan(sett_val):
                delta = sett_val - zero_marks[mark_str]
                data.append({
                    'Цикл': cycle_label,
                    'Марка': mark_str,
                    'Осадка_мм': delta
                })

    if not data:
        st.error("Не удалось извлечь приросты осадок для выбранных марок.")
        return None

    df_sett = pd.DataFrame(data)

    # --- Расчёт углов крена по осадкам ---
    marks_order = [str(m) for m in corner_marks]
    results = []
    for cycle in df_sett['Цикл'].unique():
        cycle_data = df_sett[df_sett['Цикл'] == cycle]
        s = {}
        for mark in marks_order:
            val = cycle_data[cycle_data['Марка'] == mark]['Осадка_мм']
            if not val.empty:
                s[mark] = val.iloc[0]
            else:
                s[mark] = np.nan
        if any(np.isnan(list(s.values()))):
            continue
        s1, s2, s3, s4 = s[marks_order[0]], s[marks_order[1]], s[marks_order[2]], s[marks_order[3]]
        a = ((s2 - s1) + (s4 - s3)) / (2 * L) if L != 0 else 0
        b = ((s3 - s1) + (s4 - s2)) / (2 * B) if B != 0 else 0
        results.append({
            'Цикл': cycle,
            'a_мм_м': a,
            'b_мм_м': b,
            'a_град': np.degrees(np.arctan(a / 1000)),
            'b_град': np.degrees(np.arctan(b / 1000))
        })

    if not results:
        st.error("Не удалось рассчитать углы. Возможно, для выбранных марок нет данных в одном из циклов.")
        return None

    df_angles = pd.DataFrame(results)
    return df_angles, marks_abs_data, list(marks_abs_data.keys())

# ------------------------------------------------------------
# РАСЧЁТ ДЕФОРМИРОВАННОЙ ОСИ
# ------------------------------------------------------------
def calculate_displacement_points(
    df_incl: pd.DataFrame,
    selected_cycle: str,
    L: float,
    vertical_scale: float = 1.0,
    floors: List[int] = None
) -> Tuple[List[Tuple[float, float, float]], float, float, float, float]:
    """
    Вычисляет точки деформированной оси по данным наклономеров для указанных этажей.
    Возвращает: список точек (x, y, z), координаты верхней точки и максимальную высоту.
    """
    if floors is None:
        floors = sorted(df_incl['Этаж'].unique())
    df_cycle = df_incl[df_incl['Цикл'] == selected_cycle]
    df_floors = df_cycle[df_cycle['Этаж'].isin(floors)].sort_values('Этаж')

    points = [(0, 0, 0)]
    cum_x, cum_y = 0.0, 0.0
    prev_floor = 0

    for _, row in df_floors.iterrows():
        floor = row['Этаж']
        alpha_x = row['αx_abs']
        alpha_y = row['αy_abs']
        delta_h = (floor - prev_floor) * L
        dx = delta_h * np.sin(np.radians(alpha_x))
        dy = delta_h * np.sin(np.radians(alpha_y))
        cum_x += dx
        cum_y += dy
        points.append((cum_x, cum_y, floor * L * vertical_scale))
        prev_floor = floor

    top_x, top_y, top_z = points[-1]
    max_z = max([p[2] for p in points]) if points else 0
    return points, top_x, top_y, top_z, max_z

# ------------------------------------------------------------
# ПОСТРОЕНИЕ 3D-МОДЕЛИ
# ------------------------------------------------------------
def plot_building_3d(
    df_incl: pd.DataFrame,
    selected_cycle: str,
    L: float,
    building_length: float,
    building_width: float,
    vertical_scale: float = 1.0,
    df_sett_angles: Optional[pd.DataFrame] = None,
    floors: List[int] = None
) -> Optional[go.Figure]:
    """
    Строит 3D-модель здания с деформациями, наклономерами и креном.
    Возвращает объект Figure или None, если данных недостаточно.
    """
    if floors is None:
        floors = sorted(df_incl['Этаж'].unique())
    points, top_x, top_y, top_z, max_z = calculate_displacement_points(
        df_incl, selected_cycle, L, vertical_scale, floors
    )

    if len(points) < 2:
        return None

    df_cycle = df_incl[df_incl['Цикл'] == selected_cycle]
    df_floors = df_cycle[df_cycle['Этаж'].isin(floors)].sort_values('Этаж')

    fig = go.Figure()

    # Исходная вертикаль
    fig.add_trace(go.Scatter3d(
        x=[0, 0], y=[0, 0], z=[0, max_z],
        mode='lines',
        line=dict(color='gray', width=2, dash='dash'),
        name='Исходная вертикаль'
    ))

    # Деформированная ось
    xs = [p[0] for p in points]
    ys = [p[1] for p in points]
    zs = [p[2] for p in points]
    fig.add_trace(go.Scatter3d(
        x=xs, y=ys, z=zs,
        mode='lines+markers',
        line=dict(color='red', width=5),
        marker=dict(size=8, color='red'),
        name='Деформированная ось'
    ))

    # Векторы смещений и метки
    for i, (x, y, z) in enumerate(points[1:], start=1):
        floor = df_floors.iloc[i-1]['Этаж']
        fig.add_trace(go.Scatter3d(
            x=[0, x], y=[0, y], z=[z, z],
            mode='lines+markers',
            line=dict(color='green', width=3, dash='dot'),
            marker=dict(size=6, color='green', symbol='circle'),
            name=f'Смещение {floor} эт.',
            showlegend=False
        ))
        dist = np.sqrt(x**2 + y**2)
        fig.add_trace(go.Scatter3d(
            x=[x], y=[y], z=[z],
            mode='text',
            text=[f"{floor}эт: {dist:.3f} м"],
            textposition='top center',
            textfont=dict(color='green', size=10),
            showlegend=False
        ))

    # Наклономеры
    for i, (x, y, z) in enumerate(points[1:], start=1):
        floor = df_floors.iloc[i-1]['Этаж']
        alpha_x = df_floors.iloc[i-1]['αx_abs']
        alpha_y = df_floors.iloc[i-1]['αy_abs']
        fig.add_trace(go.Scatter3d(
            x=[x], y=[y], z=[z],
            mode='markers+text',
            marker=dict(size=14, color='blue', symbol='square'),
            text=[f"Этаж {floor}<br>αx={alpha_x:.3f}°<br>αy={alpha_y:.3f}°"],
            textposition='top center',
            name=f'Наклономер {floor}'
        ))

    # Общий крен
    fig.add_trace(go.Scatter3d(
        x=[0, top_x], y=[0, top_y], z=[0, top_z],
        mode='lines+markers',
        line=dict(color='orange', width=6),
        marker=dict(size=8, color='orange', symbol='diamond'),
        name='Общий крен здания'
    ))
    kren_angle = np.degrees(np.arctan2(np.sqrt(top_x**2 + top_y**2), top_z))
    fig.add_trace(go.Scatter3d(
        x=[top_x], y=[top_y], z=[top_z],
        mode='text',
        text=[f"Крен: {kren_angle:.2f}°"],
        textposition='top center',
        textfont=dict(color='orange', size=14),
        showlegend=False
    ))

    # Крен по осадкам
    if df_sett_angles is not None and not df_sett_angles.empty:
        sett_row = df_sett_angles[df_sett_angles['Цикл'] == selected_cycle]
        if not sett_row.empty:
            a = sett_row['a_мм_м'].values[0]
            b = sett_row['b_мм_м'].values[0]
            scale = 10.0
            dx_os = a * scale
            dy_os = b * scale
            fig.add_trace(go.Scatter3d(
                x=[0, dx_os], y=[0, dy_os], z=[0, 0],
                mode='lines+markers',
                line=dict(color='purple', width=5, dash='dash'),
                marker=dict(size=10, color='purple', symbol='diamond'),
                name=f'Крен по осадкам (a={a:.2f}, b={b:.2f})'
            ))

    # Каркас здания
    half_len = building_length / 2
    half_wid = building_width / 2
    corners = [
        (-half_len, -half_wid),
        ( half_len, -half_wid),
        ( half_len,  half_wid),
        (-half_len,  half_wid)
    ]
    for cx, cy in corners:
        fig.add_trace(go.Scatter3d(
            x=[cx, cx + top_x],
            y=[cy, cy + top_y],
            z=[0, top_z],
            mode='lines',
            line=dict(color='black', width=2),
            showlegend=False
        ))
    for z_level, (x_shift, y_shift) in [(0, (0, 0)), (top_z, (top_x, top_y))]:
        shifted_corners = [(cx + x_shift, cy + y_shift) for cx, cy in corners]
        for i in range(4):
            x1, y1 = shifted_corners[i]
            x2, y2 = shifted_corners[(i + 1) % 4]
            fig.add_trace(go.Scatter3d(
                x=[x1, x2], y=[y1, y2], z=[z_level, z_level],
                mode='lines',
                line=dict(color='black', width=2),
                showlegend=False
            ))

    fig.update_layout(
        title=f"3D-модель здания – цикл {selected_cycle} (верт. масштаб {vertical_scale:.1f})",
        scene=dict(
            xaxis_title="Смещение X, м",
            yaxis_title="Смещение Y, м",
            zaxis_title="Высота, м",
            aspectmode='data',
            camera=dict(eye=dict(x=1.8, y=1.8, z=1.2))
        ),
        width=900,
        height=750,
        template="plotly_white",
        legend=dict(
            orientation="v",
            yanchor="top",
            y=1,
            xanchor="left",
            x=1.02
        )
    )
    return fig

# ------------------------------------------------------------
# ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ ДЛЯ ОТОБРАЖЕНИЯ ЦИКЛОВ
# ------------------------------------------------------------
def get_cycle_display_options(df_incl: pd.DataFrame) -> Dict[str, str]:
    """Возвращает словарь {ключ: полное_название} для циклов."""
    unique = df_incl[['Цикл', 'Цикл_полное']].drop_duplicates()
    unique_sorted = unique.sort_values('Цикл')
    return dict(zip(unique_sorted['Цикл'], unique_sorted['Цикл_полное']))

# ------------------------------------------------------------
# ОСНОВНАЯ ЛОГИКА ПРИЛОЖЕНИЯ
# ------------------------------------------------------------
uploaded_file = st.file_uploader(
    "Загрузите Excel-файл с данными наклономера и осадок",
    type=["xlsx", "xls"],
    help="Файл должен содержать лист с наклономером (строки с этажами 5,15,27) и, опционально, листы с осадками."
)

if uploaded_file is not None:
    try:
        file_bytes = uploaded_file.read()
        xl = load_excel_file(file_bytes)
        all_sheets = xl.sheet_names

        # --- Боковая панель ---
        st.sidebar.header("Выбор листов")
        incl_sheet_name = st.sidebar.selectbox(
            "Лист с наклономером",
            all_sheets,
            index=all_sheets.index('Наклономер') if 'Наклономер' in all_sheets else 0
        )

        st.sidebar.header("Параметры поиска строк")
        use_manual_range = st.sidebar.checkbox("Ручной диапазон для поиска этажей", value=False)
        if use_manual_range:
            search_start = st.sidebar.number_input("Начальная строка (индекс)", min_value=0, step=1, value=7)
            search_end = st.sidebar.number_input("Конечная строка (индекс)", min_value=0, step=1, value=13)
        else:
            search_start = None
            search_end = None

        df_incl = parse_inclinometer_data(
            file_bytes, incl_sheet_name,
            search_start=search_start, search_end=search_end
        )

        if df_incl is None:
            st.subheader("🔧 Ручной ввод строк с этажами")
            st.write("Введите номера строк (индексы, начиная с 0), в которых расположены данные для этажей.")
            col1, col2, col3 = st.columns(3)
            with col1:
                row_5 = st.number_input("Строка для этажа 5", min_value=0, step=1, value=0)
            with col2:
                row_15 = st.number_input("Строка для этажа 15", min_value=0, step=1, value=1)
            with col3:
                row_27 = st.number_input("Строка для этажа 27", min_value=0, step=1, value=2)

            if st.button("Повторить парсинг с указанными строками"):
                manual_rows = {}
                if row_5 >= 0:
                    manual_rows[5] = row_5
                if row_15 >= 0:
                    manual_rows[15] = row_15
                if row_27 >= 0:
                    manual_rows[27] = row_27
                df_incl = parse_inclinometer_data(
                    file_bytes, incl_sheet_name,
                    manual_floor_rows=manual_rows
                )
                if df_incl is None:
                    st.error("Не удалось извлечь данные даже с ручным указанием строк. Проверьте структуру листа.")
                    st.stop()
                else:
                    st.success("✅ Данные наклономера успешно загружены!")
            else:
                st.stop()
        else:
            st.success(f"✅ Данные наклономера загружены. Найдено {len(df_incl)} записей.")

        if df_incl is None:
            st.stop()

        # --- Боковая панель: выбор этажей ---
        all_available_floors = sorted(df_incl['Этаж'].unique())
        default_floors = all_available_floors.copy()
        selected_floors = st.sidebar.multiselect(
            "Этажи с наклономерами (выберите для построения модели)",
            options=all_available_floors,
            default=default_floors,
            help="Выберите этажи, по которым будут строиться деформации. Если данных для этажа нет, он будет пропущен."
        )
        st.session_state.selected_floors = selected_floors

        # --- Боковая панель: размеры здания ---
        st.sidebar.subheader("Размеры здания для 3D-модели")
        st.session_state.building_length = st.sidebar.number_input(
            "Длина здания в плане (X), м",
            value=st.session_state.get("building_length", DEFAULT_BUILDING_LENGTH),
            step=0.1,
            key="building_length_input",
            help="Горизонтальный размер вдоль оси X"
        )
        st.session_state.building_width = st.sidebar.number_input(
            "Ширина здания в плане (Y), м",
            value=st.session_state.get("building_width", DEFAULT_BUILDING_WIDTH),
            step=0.1,
            key="building_width_input",
            help="Горизонтальный размер вдоль оси Y"
        )
        st.session_state.vertical_scale = st.sidebar.slider(
            "Вертикальный масштаб (высота)",
            min_value=0.5,
            max_value=2.0,
            value=st.session_state.get("vertical_scale", DEFAULT_VERTICAL_SCALE),
            step=0.1,
            help="Коэффициент визуального увеличения высоты здания (1.0 = реальная высота)"
        )

        # --- Основные вкладки ---
        tab1, tab2, tab3 = st.tabs([
            "📊 Данные и параметры",
            "🏢 3D-модель здания",
            "📘 Наклономер"
        ])

        cycle_display_map = get_cycle_display_options(df_incl)
        cycle_keys = list(cycle_display_map.keys())

        with tab1:
            if len(cycle_keys) == 0:
                st.error("Нет циклов в данных наклономера.")
                st.stop()

            st.subheader("⚙️ Параметры расчёта")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                zero_cycle_display = st.selectbox(
                    "Нулевой цикл",
                    options=[cycle_display_map[k] for k in cycle_keys],
                    index=0
                )
                zero_cycle = [k for k, v in cycle_display_map.items() if v == zero_cycle_display][0]
            with col2:
                alpha0_x = st.number_input("Начальный угол X (αx0), °", value=0.0, step=0.001, format="%.3f")
            with col3:
                alpha0_y = st.number_input("Начальный угол Y (αy0), °", value=0.0, step=0.001, format="%.3f")
            with col4:
                L = st.number_input("Высота этажа (L), м", value=3.0, step=0.1, format="%.1f")

            zero_data = df_incl[df_incl['Цикл'] == zero_cycle][['Этаж', 'αx', 'αy']].rename(
                columns={'αx': 'αx0_inc', 'αy': 'αy0_inc'}
            )
            df_incl = df_incl.merge(zero_data, on='Этаж', how='left')
            df_incl['αx_abs'] = df_incl['αx'] - df_incl['αx0_inc'] + alpha0_x
            df_incl['αy_abs'] = df_incl['αy'] - df_incl['αy0_inc'] + alpha0_y
            df_incl['Смещение X'] = L * np.sin(np.radians(df_incl['αx_abs']))
            df_incl['Смещение Y'] = L * np.sin(np.radians(df_incl['αy_abs']))

            st.session_state.zero_cycle = zero_cycle
            st.session_state.L = L

            # ---------- Блок осадок ----------
            sett_sheets = [s for s in all_sheets if
                           'стилобат' in s.lower() or 'высотн' in s.lower() or 'осадк' in s.lower()]
            if sett_sheets:
                st.subheader("📐 Данные осадок (для расчёта крена)")
                with st.expander("Настройки осадок", expanded=False):
                    selected_sett_sheet = st.selectbox("Выберите лист с осадками", sett_sheets, key="sett_sheet")
                    corner_marks_str = st.text_input(
                        "Номера угловых марок (через запятую, в порядке: нижний левый, нижний правый, верхний левый, верхний правый)",
                        "1,4,11,14"
                    )
                    mark_col = st.number_input(
                        "Номер столбца с марками (0-индекс, обычно 0 или 1)",
                        min_value=0, step=1, value=0, key="mark_col"
                    )
                    manual_osad_col = st.number_input(
                        "Номер столбца с осадками (0-индекс, если не уверены, оставьте -1 для автоопределения)",
                        min_value=-1, step=1, value=-1, key="manual_osad_col"
                    )
                    try:
                        corner_marks = [x.strip() for x in corner_marks_str.split(',') if x.strip()]
                        corner_marks_parsed = []
                        for m in corner_marks:
                            try:
                                corner_marks_parsed.append(int(m))
                            except ValueError:
                                corner_marks_parsed.append(m)
                        corner_marks = corner_marks_parsed
                        if len(corner_marks) != 4:
                            st.warning("Введите ровно 4 номера марок.")
                    except:
                        corner_marks = []
                    L_sett = st.number_input("Длина фундамента L, м", value=70.46, step=0.1, key="L_sett")
                    B_sett = st.number_input("Ширина фундамента B, м", value=18.69, step=0.1, key="B_sett")

                    # --- Получаем список циклов осадок для выбора нулевого ---
                    try:
                        df_raw_test = pd.read_excel(io.BytesIO(file_bytes), sheet_name=selected_sett_sheet, header=None)
                        cycle_header_row_test = None
                        for idx, row in df_raw_test.iterrows():
                            row_str = ' '.join(str(cell) for cell in row if pd.notna(cell))
                            if 'Цикл' in row_str:
                                cycle_header_row_test = idx
                                break
                        all_cycles_temp = []
                        if cycle_header_row_test is not None:
                            for col_idx, cell in df_raw_test.iloc[cycle_header_row_test, :].items():
                                if pd.notna(cell):
                                    cell_str = str(cell).strip()
                                    if 'Цикл' in cell_str:
                                        date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', cell_str)
                                        if date_match:
                                            try:
                                                date_obj = pd.to_datetime(date_match.group(1), dayfirst=True)
                                                cycle_label = date_obj.strftime('%Y-%m-%d')
                                                all_cycles_temp.append(cycle_label)
                                            except:
                                                pass
                        def try_parse_date(s):
                            for fmt in ('%Y-%m-%d', '%d.%m.%Y', '%Y-%m-%d'):
                                try:
                                    return pd.to_datetime(s, format=fmt)
                                except:
                                    continue
                            return None
                        sorted_cycles_temp = sorted(all_cycles_temp, key=lambda x: try_parse_date(x) or x)
                        zero_dt = pd.to_datetime(zero_cycle)
                        best_cycle = sorted_cycles_temp[0] if sorted_cycles_temp else None
                        if sorted_cycles_temp:
                            for cyc in sorted_cycles_temp:
                                cyc_dt = pd.to_datetime(cyc)
                                if cyc_dt >= zero_dt:
                                    best_cycle = cyc
                                    break
                        if best_cycle is None and sorted_cycles_temp:
                            best_cycle = sorted_cycles_temp[0]
                    except:
                        best_cycle = None
                        sorted_cycles_temp = []

                    if sorted_cycles_temp:
                        zero_cycle_sett = st.selectbox(
                            "Нулевой цикл осадок (от которого считать прирост)",
                            options=sorted_cycles_temp,
                            index=sorted_cycles_temp.index(best_cycle) if best_cycle in sorted_cycles_temp else 0,
                            key="zero_cycle_sett"
                        )
                    else:
                        zero_cycle_sett = None

                    if st.button("Рассчитать углы по осадкам"):
                        if len(corner_marks) == 4 and zero_cycle_sett is not None:
                            result = parse_settlement_data(
                                file_bytes, selected_sett_sheet, corner_marks, L_sett, B_sett,
                                mark_col=mark_col, zero_cycle_sett=zero_cycle_sett,
                                manual_osad_col=manual_osad_col if manual_osad_col >= 0 else None
                            )
                            if result is not None:
                                df_sett_angles, marks_data, all_cycles_sett = result
                                st.success(f"✅ Углы по осадкам рассчитаны для {len(df_sett_angles)} циклов.")
                                st.session_state.res_df_sett_angles = df_sett_angles
                            else:
                                st.error("Не удалось рассчитать углы. Проверьте правильность введённых данных.")
                        else:
                            st.error("Укажите 4 угловые марки и выберите нулевой цикл.")

                if 'res_df_sett_angles' in st.session_state and st.session_state.res_df_sett_angles is not None:
                    df_angles = st.session_state.res_df_sett_angles.copy()
                    st.subheader("Таблица углов крена по осадкам")
                    st.caption("**Примечание:** показаны углы для всех циклов (кроме нулевого, где значения равны 0).")
                    st.dataframe(
                        df_angles,
                        column_config={
                            "Цикл": "Цикл",
                            "a_мм_м": st.column_config.NumberColumn("a, мм/м", format="%.3f"),
                            "b_мм_м": st.column_config.NumberColumn("b, мм/м", format="%.3f"),
                            "a_град": st.column_config.NumberColumn("a, град", format="%.4f"),
                            "b_град": st.column_config.NumberColumn("b, град", format="%.4f"),
                        },
                        use_container_width=True
                    )

        with tab2:
            st.subheader("🏢 3D-модель здания с креном и наклономерами")
            st.markdown("""
            **ℹ️ Модель показывает реальную деформацию здания**:
            - **Серая пунктирная линия** – исходная вертикаль.
            - **Красная линия** – деформированная ось (накопленные смещения).
            - **Зелёные векторы** – смещения на выбранных этажах.
            - **Синие квадраты** – места установки наклономеров с углами.
            - **Оранжевая стрелка** – общий крен здания (от фундамента до верха).
            - **Фиолетовая стрелка** (если есть) – крен из данных осадок.
            - **Чёрный каркас** – контур здания с наклоном.
            """)

            total_cycles = len(cycle_keys)
            if total_cycles == 0:
                st.warning("Нет доступных циклов для отображения.")
            else:
                if 'current_index' not in st.session_state or st.session_state.current_index >= total_cycles:
                    st.session_state.current_index = total_cycles - 1

                selected_index = st.slider(
                    "Выбор цикла",
                    min_value=0,
                    max_value=total_cycles-1,
                    value=st.session_state.current_index,
                    step=1,
                    format="%d",
                    help="Перетащите для выбора цикла"
                )
                st.session_state.current_index = selected_index
                selected_cycle_key = cycle_keys[selected_index]
                st.caption(f"**Текущий цикл:** {cycle_display_map[selected_cycle_key]}")

                selected_cycle_building = selected_cycle_key
                building_length = st.session_state.get("building_length", DEFAULT_BUILDING_LENGTH)
                building_width = st.session_state.get("building_width", DEFAULT_BUILDING_WIDTH)
                vertical_scale = st.session_state.get("vertical_scale", DEFAULT_VERTICAL_SCALE)
                df_sett_angles = st.session_state.get("res_df_sett_angles", None)
                floors = st.session_state.get("selected_floors", [])

                fig_building = plot_building_3d(
                    df_incl,
                    selected_cycle_building,
                    st.session_state.L,
                    building_length,
                    building_width,
                    vertical_scale,
                    df_sett_angles,
                    floors
                )
                if fig_building:
                    st.plotly_chart(fig_building, use_container_width=True)
                else:
                    st.warning("Для выбранного цикла нет данных на выбранных этажах. Попробуйте изменить выбор этажей или цикл.")

            # ---------- Раздел с формулами ----------
            with st.expander("📐 Как строится модель (формулы и пояснения)", expanded=False):
                st.markdown(r"""
                **Построение 3D-модели деформаций здания** основано на данных накладного инклинометра, установленного на выбранных этажах (вы задаёте их в боковой панели).

                ---

                ### 1. Исходные данные

                Для каждого цикла измерений и для каждого выбранного этажа \(i\) известны:

                * \(\alpha_{x}^{(i)}\) – угол наклона по оси **X** (градусы);
                * \(\alpha_{y}^{(i)}\) – угол наклона по оси **Y** (градусы).

                Эти углы — это **прирост** относительно нулевого цикла, скорректированный на начальный угол (\( \alpha_{0x}, \alpha_{0y} \)), который задаёт пользователь.  
                Таким образом, **абсолютные** углы наклона для расчётов:

                \[
                \alpha_{x,\text{abs}}^{(i)} = \alpha_{x}^{(i)} - \alpha_{x,0}^{(i)} + \alpha_{0x}
                \]
                \[
                \alpha_{y,\text{abs}}^{(i)} = \alpha_{y}^{(i)} - \alpha_{y,0}^{(i)} + \alpha_{0y}
                \]

                где \(\alpha_{x,0}^{(i)}\), \(\alpha_{y,0}^{(i)}\) – значения углов в нулевом цикле для этажа \(i\).

                ---

                ### 2. Накопление смещений

                Смещение на каждом этаже вычисляется как **сумма приращений** по высоте от фундамента (0-й этаж) до текущего этажа.

                Для участка между двумя соседними выбранными этажами \(k-1\) и \(k\) высота участка:

                \[
                L_k = (\text{этаж}_k - \text{этаж}_{k-1}) \cdot L
                \]

                где \(L\) – высота одного этажа (задаётся пользователем).  
                На этом участке угол наклона считается постоянным (линейная интерполяция).

                Тогда приращения смещения на участке:

                \[
                \Delta M_x^{(k)} = L_k \cdot \sin\left(\alpha_{x,\text{abs}}^{(k)}\right)
                \]
                \[
                \Delta M_y^{(k)} = L_k \cdot \sin\left(\alpha_{y,\text{abs}}^{(k)}\right)
                \]

                Накопленное смещение на этаже \(n\):

                \[
                M_x^{(n)} = \sum_{k=1}^{n} \Delta M_x^{(k)}
                \]
                \[
                M_y^{(n)} = \sum_{k=1}^{n} \Delta M_y^{(k)}
                \]

                Здесь суммирование ведётся по всем выбранным этажам от фундамента до \(n\).

                > **Важно:** углы \(\alpha_{x,\text{abs}}\) и \(\alpha_{y,\text{abs}}\) перед вычислением синуса переводятся из градусов в радианы:  
                > \(\alpha_{\text{рад}} = \alpha_{\text{град}} \cdot \dfrac{\pi}{180}\).

                ---

                ### 3. Построение деформированной оси

                Для каждого выбранного этажа \(n\) вычисляется точка в трёхмерном пространстве:

                \[
                \text{Точка}_n = \bigl( M_x^{(n)},\; M_y^{(n)},\; H_n \bigr)
                \]

                где  
                \(H_n = \text{этаж}_n \cdot L\) – высота этажа над фундаментом.

                Дополнительно добавляется точка фундамента: \((0,\;0,\;0)\).

                **Красная линия**, соединяющая эти точки в порядке возрастания этажей, называется **деформированной осью** здания.

                ---

                ### 4. Общий крен здания

                Вектор от фундамента до самой верхней выбранной точки (максимальный этаж \(N\)) определяет **общий крен** здания.

                Горизонтальное смещение верхней точки:

                \[
                M_{\text{top}} = \sqrt{ \left(M_x^{(N)}\right)^2 + \left(M_y^{(N)}\right)^2 }
                \]

                Высота верхней точки:

                \[
                H_{\text{top}} = \text{этаж}_N \cdot L
                \]

                Угол крена (в градусах):

                \[
                \theta = \arctan\left( \frac{M_{\text{top}}}{H_{\text{top}}} \right) \cdot \frac{180}{\pi}
                \]

                **Оранжевая стрелка** на графике показывает направление и длину этого вектора.

                ---

                ### 5. Векторы смещений на этажах

                Для каждого выбранного этажа отображается **зелёный вектор** – он проведён от вертикальной оси \((0,0,H_n)\) до точки деформированной оси \((M_x^{(n)}, M_y^{(n)}, H_n)\).  
                Это наглядно показывает горизонтальное смещение каждого этажа относительно исходного положения.

                ---

                ### 6. Каркас здания

                **Чёрный каркас** строится на основе размеров здания в плане (длина \(X\) и ширина \(Y\), задаются пользователем).  
                Вертикальные рёбра каркаса наклоняются так, чтобы верхняя их часть совпадала со смещением верхней точки деформированной оси. Так создаётся иллюзия наклона всего объёма здания.

                ---

                ### 7. Крен по данным осадок (опционально)

                Если загружены и рассчитаны данные осадок фундамента, на уровне земли (\(z=0\)) отображается **фиолетовый вектор**.  
                Его направление и длина соответствуют углам крена, вычисленным по осадкам четырёх угловых марок:

                \[
                a = \frac{(s_2 - s_1) + (s_4 - s_3)}{2 L_f}
                \]
                \[
                b = \frac{(s_3 - s_1) + (s_4 - s_2)}{2 B_f}
                \]

                где:
                * \(s_1, s_2, s_3, s_4\) – осадки четырёх угловых марок в порядке обхода;
                * \(L_f\) – длина фундамента;
                * \(B_f\) – ширина фундамента.

                Полученные значения \(a\) и \(b\) – это наклон фундамента в мм/м по осям X и Y соответственно.

                ---

                **Все расчёты выполняются в метрах** (кроме специально оговоренных случаев). Углы для тригонометрических функций всегда переводятся в радианы с помощью `np.radians()`.  
                Результат – интерактивная 3D-модель, которую можно вращать и масштабировать.
                """)

        with tab3:
            st.header("📘 Накладной инклинометр УСМ-ИСН-П")
            st.markdown("""
            **Инклинометр накладной портативный** предназначен для ручных измерений угла наклона (крена) различных конструкций зданий и сооружений в точке монтажа измерительной пластины.  
            Устройство позволяет производить замеры по двум осям (**X** и **Y**) одновременно, что в разы сокращает процесс измерений.
            """)

            st.subheader("📋 Технические характеристики")
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**Общие параметры**")
                st.table(pd.DataFrame({
                    'Параметр': ['Габариты (в сборе)', 'Масса', 'Время автономной работы', 'Время полной зарядки', 'Аккумулятор', 'Разъём зарядки'],
                    'Значение': ['174×134×160 мм', '≤ 4500 г', '≈ 18 ч', '≈ 12 ч', 'Li‑ion 18650, 3400 мА·ч, 3,7 В', 'USB Type‑B']
                }).set_index('Параметр'))
            with col2:
                st.markdown("**Измерительный модуль**")
                st.table(pd.DataFrame({
                    'Параметр': ['Диапазон измерений', 'Разрешающая способность', 'Входное напряжение', 'Потребляемый ток'],
                    'Значение': ['±10°, ±15° или ±30°', '0.001°', '7.5–15 В DC', '≤ 30 мА']
                }).set_index('Параметр'))

            st.subheader("🔄 Измерительная пластина")
            st.table(pd.DataFrame({
                'Параметр': ['Внешний диаметр', 'Внутренний диаметр', 'Диаметр опорного элемента', 'Диаметр монтажного отверстия', 'Толщина пластины', 'Общая толщина (выступ)', 'Масса'],
                'Значение': ['120 мм', '60 мм', '16 мм', '8 мм', '6 мм', '21 мм', '≤ 600 г']
            }).set_index('Параметр'))

            st.subheader("📐 Формулы обработки данных")
            st.latex(r"\alpha_n = \frac{\alpha_0 - \alpha_{180^\circ}}{2}")
            st.markdown("где α₀ – угол при первом измерении, α₁₈₀° – угол после разворота на 180°.")
            st.latex(r"\Delta\alpha = \alpha_n - \alpha_0")
            st.markdown("где Δα – деформация относительно нулевого цикла, α₀ – значение в нулевом цикле.")

            with st.expander("📌 Рекомендации по эксплуатации"):
                st.markdown("""
                - Перед выездом на объект убедитесь в полном заряде аккумулятора.
                - Избегайте падений, ударов и загрязнений прибора.
                - Монтаж пластин выполняйте строго по уровню (используйте пузырьковый или электронный уровень).
                - «Нулевые» циклы проводите после полного закрепления пластин (для химических анкеров – после схватывания состава).
                - Измерения проводите при неподвижном приборе, строго следуя двухшаговому алгоритму (0° и 180°).
                - Результаты записывайте в журнал измерений.
                """)
            st.caption("Источник: Руководство по эксплуатации УСМ-ИСН-П (ООО «СПС», 2023)")

    except Exception as e:
        st.error(f"Ошибка обработки: {e}")
        import traceback
        st.code(traceback.format_exc())

else:
    st.info("👆 Загрузите Excel-файл для начала работы.")
