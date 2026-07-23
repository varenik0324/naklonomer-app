import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import io
import re
import logging
from datetime import datetime
from typing import Optional, Tuple, List, Dict, Any

# ------------------------------------------------------------
# НАСТРОЙКА ЛОГИРОВАНИЯ
# ------------------------------------------------------------
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# ------------------------------------------------------------
# КОНСТАНТЫ
# ------------------------------------------------------------
DEFAULT_BUILDING_LENGTH = 70.46
DEFAULT_BUILDING_WIDTH = 18.69
DEFAULT_VERTICAL_SCALE = 1.0
MAX_ANGLE_DEG = 30
DEFAULT_FILTER_WINDOW = 3
FILTER_TYPES = ["Нет", "Скользящее среднее", "Медианный фильтр"]
DEFAULT_L = 3.0
DEFAULT_ALPHA0_X = 0.0
DEFAULT_ALPHA0_Y = 0.0

# ------------------------------------------------------------
# ИНИЦИАЛИЗАЦИЯ СОСТОЯНИЯ СЕССИИ
# ------------------------------------------------------------
def init_session_state():
    defaults = {
        'building_length': DEFAULT_BUILDING_LENGTH,
        'building_width': DEFAULT_BUILDING_WIDTH,
        'vertical_scale': DEFAULT_VERTICAL_SCALE,
        'selected_floors': [],
        'zero_cycle': None,
        'L': DEFAULT_L,
        'alpha0_x': DEFAULT_ALPHA0_X,
        'alpha0_y': DEFAULT_ALPHA0_Y,
        'selected_cycle_key': None,
        'filter_type': 'Нет',
        'filter_window': DEFAULT_FILTER_WINDOW,
        'df_incl_raw': None,
        'df_incl_filtered': None,
        'df_sett_angles': None,
        'cycles': [],
        'cycle_display_map': {},
        'file_loaded': False,
        'correct_by_sett': False,
    }
    for key, default in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = default

# ------------------------------------------------------------
# ПАРСЕР EXCEL (ИСПРАВЛЕННАЯ ВЕРСИЯ ДЛЯ ОСАДОК)
# ------------------------------------------------------------
class ExcelParser:
    @staticmethod
    @st.cache_resource
    def load_excel(file_bytes: bytes) -> pd.ExcelFile:
        return pd.ExcelFile(io.BytesIO(file_bytes))

    @staticmethod
    def _extract_cycle_headers(df_raw: pd.DataFrame) -> Tuple[Optional[int], List[str]]:
        for idx, row in df_raw.iterrows():
            row_str = ' '.join(str(cell) for cell in row if pd.notna(cell))
            if 'Цикл' in row_str and re.search(r'\d{2}\.\d{2}\.\d{4}', row_str):
                cycle_names = []
                for col in range(1, len(row)):
                    cell = row.iloc[col]
                    if pd.notna(cell) and 'Цикл' in str(cell):
                        cycle_names.append(str(cell).strip())
                return idx, cycle_names
        return None, []

    @staticmethod
    def _parse_cycle_labels(cycle_names: List[str]) -> List[str]:
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
        return labels

    @staticmethod
    def _find_floor_rows(df_raw: pd.DataFrame, cycle_count: int) -> Dict[int, int]:
        floor_rows = {}
        for idx, row in df_raw.iterrows():
            first_val = row.iloc[0] if len(row) > 0 else None
            if pd.isna(first_val):
                continue
            try:
                floor = int(float(first_val))
            except (ValueError, TypeError):
                continue
            numbers = [v for v in row.iloc[1:] if pd.notna(v) and isinstance(v, (int, float))]
            if len(numbers) >= 2 * cycle_count:
                floor_rows[floor] = idx
        return floor_rows

    @staticmethod
    def _extract_angle_pairs(row: pd.Series) -> List[Tuple[float, float]]:
        values = [float(v) for v in row.iloc[1:] if pd.notna(v) and isinstance(v, (int, float))]
        pairs = [(values[i], values[i+1]) for i in range(0, len(values)-1, 2)]
        return pairs

    @classmethod
    def parse_inclinometer(cls, file_bytes: bytes, sheet_name: str) -> Optional[pd.DataFrame]:
        try:
            df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)
        except Exception as e:
            st.error(f"Не удалось прочитать лист '{sheet_name}': {e}")
            return None

        header_row, cycle_names = cls._extract_cycle_headers(df_raw)
        if header_row is None:
            st.error("Не найдена строка с заголовками циклов.")
            return None
        cycle_labels = cls._parse_cycle_labels(cycle_names)
        if not cycle_labels:
            st.error("Не удалось распознать даты циклов.")
            return None

        floor_rows = cls._find_floor_rows(df_raw, len(cycle_labels))
        if len(floor_rows) < 2:
            st.warning("Найдено менее двух этажей. Проверьте структуру файла.")
            return None

        data = []
        for floor, row_idx in floor_rows.items():
            row = df_raw.iloc[row_idx]
            pairs = cls._extract_angle_pairs(row)
            pairs = pairs[:len(cycle_labels)]
            for i, (ax, ay) in enumerate(pairs):
                data.append({
                    'Цикл': cycle_labels[i] if i < len(cycle_labels) else f"Цикл_{i+1}",
                    'Цикл_полное': cycle_names[i] if i < len(cycle_names) else f"Цикл {i+1}",
                    'Этаж': floor,
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
        if df.empty:
            st.error("Все данные наклономера оказались пустыми.")
            return None
        return df

    @classmethod
    def parse_settlement(cls, file_bytes: bytes, sheet_name: str,
                         corner_marks: List[str], L_fund: float, B_fund: float,
                         mark_col: int = 0) -> Optional[pd.DataFrame]:
        """
        Парсит лист с осадками.
        Теперь ищет столбцы с осадками следующим образом:
        1. Сначала пытается найти столбцы с подписью, содержащей 'осадк' (регистронезависимо).
        2. Если не найдены, для каждого заголовка цикла берёт столбец, следующий сразу за ним (col_idx+1).
        3. Если и так не находит числа, выдаёт ошибку.
        """
        try:
            df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)
        except Exception as e:
            st.error(f"Не удалось прочитать лист '{sheet_name}': {e}")
            return None

        header_row, cycle_names = cls._extract_cycle_headers(df_raw)
        if header_row is None:
            st.error("Не найдена строка с заголовками циклов в листе осадок.")
            return None
        cycle_labels = cls._parse_cycle_labels(cycle_names)
        if not cycle_labels:
            st.error("Не удалось распознать даты циклов в осадках.")
            return None

        # Находим столбцы, которые могут содержать осадки (по наличию слова "осадк")
        osad_cols = {}
        for col_idx, cell in df_raw.iloc[header_row, :].items():
            if pd.notna(cell) and 'осадк' in str(cell).lower():
                osad_cols[col_idx] = str(cell).strip()

        cycle_to_col = {}
        for idx, label in enumerate(cycle_labels):
            col_found = None
            # Ищем столбец, в котором есть заголовок с этим циклом
            for col_idx, cell in df_raw.iloc[header_row, :].items():
                if pd.notna(cell) and 'Цикл' in str(cell) and label in str(cell):
                    # Сначала проверяем столбцы с явной подписью "осадк"
                    for offset in [1, 2, 3]:
                        check_col = col_idx + offset
                        if check_col in osad_cols:
                            col_found = check_col
                            break
                    # Если не нашли, берём col_idx+1 (соседний столбец)
                    if col_found is None:
                        col_found = col_idx + 1
                    break
            if col_found is not None and col_found < len(df_raw.columns):
                cycle_to_col[label] = col_found
            else:
                st.warning(f"Не найден столбец с осадками для цикла {label}")

        # Если всё ещё не нашли ни одного столбца, пробуем последний вариант:
        # берём для каждого цикла столбец, следующий за заголовком (без поиска по подписи)
        if not cycle_to_col:
            for idx, label in enumerate(cycle_labels):
                for col_idx, cell in df_raw.iloc[header_row, :].items():
                    if pd.notna(cell) and 'Цикл' in str(cell) and label in str(cell):
                        col_found = col_idx + 1
                        if col_found < len(df_raw.columns):
                            cycle_to_col[label] = col_found
                        break

        if not cycle_to_col:
            st.error("Не найдены столбцы с осадками. Убедитесь, что в листе 'Стилобат' после каждого заголовка цикла идут числовые значения осадок.")
            return None

        # Ищем строки с марками (в столбце mark_col числа)
        mark_rows = []
        for idx in range(header_row + 1, len(df_raw)):
            val = df_raw.iloc[idx, mark_col]
            if pd.notna(val) and isinstance(val, (int, float)):
                mark_rows.append(idx)

        if not mark_rows:
            st.error(f"Не найдены строки с марками в столбце {mark_col}.")
            return None

        marks_data = {}
        for cycle, col_idx in cycle_to_col.items():
            marks_data[cycle] = {}
            for row_idx in mark_rows:
                mark_num = df_raw.iloc[row_idx, mark_col]
                mark_str = str(int(mark_num)) if mark_num == int(mark_num) else str(mark_num)
                sett_val = df_raw.iloc[row_idx, col_idx]
                if pd.notna(sett_val) and isinstance(sett_val, (int, float)):
                    marks_data[cycle][mark_str] = sett_val
                else:
                    marks_data[cycle][mark_str] = np.nan

        sorted_cycles = sorted(cycle_labels)
        zero_cycle = sorted_cycles[0] if sorted_cycles else None
        if zero_cycle is None:
            st.error("Нет циклов для определения нулевого.")
            return None

        corner_marks_str = [str(m) for m in corner_marks]
        for mark in corner_marks_str:
            if mark not in marks_data[zero_cycle] or np.isnan(marks_data[zero_cycle][mark]):
                st.warning(f"Марка {mark} отсутствует в нулевом цикле {zero_cycle}. Попробуйте другие марки.")
                return None

        results = []
        for cycle in sorted_cycles:
            if cycle == zero_cycle:
                continue
            s = {}
            for mark in corner_marks_str:
                if mark in marks_data[cycle] and mark in marks_data[zero_cycle]:
                    s[mark] = marks_data[cycle][mark] - marks_data[zero_cycle][mark]
                else:
                    s[mark] = np.nan
            if any(np.isnan(list(s.values()))):
                continue
            s1, s2, s3, s4 = s[corner_marks_str[0]], s[corner_marks_str[1]], s[corner_marks_str[2]], s[corner_marks_str[3]]
            a = ((s2 - s1) + (s4 - s3)) / (2 * L_fund) if L_fund != 0 else 0
            b = ((s3 - s1) + (s4 - s2)) / (2 * B_fund) if B_fund != 0 else 0
            results.append({
                'Цикл': cycle,
                'a_мм_м': a,
                'b_мм_м': b,
                'a_град': np.degrees(np.arctan(a / 1000)),
                'b_град': np.degrees(np.arctan(b / 1000))
            })

        if not results:
            st.error("Не удалось рассчитать углы по осадкам.")
            return None

        return pd.DataFrame(results)

# ------------------------------------------------------------
# ОБРАБОТЧИК ДАННЫХ И ВИЗУАЛИЗАТОР (БЕЗ ИЗМЕНЕНИЙ)
# ------------------------------------------------------------
# ... (весь остальной код полностью совпадает с предыдущей версией, 
# но для полноты я приведу его сокращённо, чтобы не дублировать 400 строк)
# В реальности вы можете просто заменить метод parse_settlement на новый.
