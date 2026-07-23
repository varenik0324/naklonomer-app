import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import io
import re
import logging
from datetime import datetime
from typing import Optional, Tuple, List, Dict, Any
from functools import lru_cache

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
REQUIRED_SHEETS = ["Наклономер", "Стилобат"]  # фактически обязателен только "Наклономер"
MAX_ANGLE_DEG = 30
DEFAULT_FILTER_WINDOW = 3
FILTER_TYPES = ["Нет", "Скользящее среднее", "Медианный фильтр"]
DEFAULT_L = 3.0
DEFAULT_ALPHA0_X = 0.0
DEFAULT_ALPHA0_Y = 0.0

# ------------------------------------------------------------
# КЛАСС ДЛЯ УПРАВЛЕНИЯ СОСТОЯНИЕМ
# ------------------------------------------------------------
class AppState:
    """Централизованное хранение состояния приложения."""
    def __init__(self):
        self.building_length = DEFAULT_BUILDING_LENGTH
        self.building_width = DEFAULT_BUILDING_WIDTH
        self.vertical_scale = DEFAULT_VERTICAL_SCALE
        self.selected_floors = []
        self.zero_cycle = None
        self.L = DEFAULT_L
        self.alpha0_x = DEFAULT_ALPHA0_X
        self.alpha0_y = DEFAULT_ALPHA0_Y
        self.selected_cycle_key = None
        self.filter_type = "Нет"
        self.filter_window = DEFAULT_FILTER_WINDOW

        # Данные (заполняются после загрузки)
        self.df_incl_raw = None          # исходные данные после apply_parameters (без фильтра)
        self.df_incl_filtered = None     # после фильтрации
        self.df_sett_angles = None       # углы по осадкам
        self.cycles = []                 # список уникальных циклов
        self.cycle_display_map = {}      # отображение ключ -> полное имя

    def update_from_session(self):
        """Загружает значения из st.session_state (для совместимости)."""
        for key, value in st.session_state.items():
            if hasattr(self, key):
                setattr(self, key, value)

    def save_to_session(self):
        """Сохраняет текущие значения в st.session_state."""
        for key, value in self.__dict__.items():
            st.session_state[key] = value

# ------------------------------------------------------------
# ПАРСЕР EXCEL
# ------------------------------------------------------------
class ExcelParser:
    @staticmethod
    @st.cache_data
    def load_excel(file_bytes: bytes) -> pd.ExcelFile:
        return pd.ExcelFile(io.BytesIO(file_bytes))

    @staticmethod
    def _extract_cycle_headers(df_raw: pd.DataFrame) -> Tuple[Optional[int], List[str]]:
        """Находит строку с заголовками циклов и извлекает названия."""
        for idx, row in df_raw.iterrows():
            row_str = ' '.join(str(cell) for cell in row if pd.notna(cell))
            # Ищем "Цикл" и дату в формате ДД.ММ.ГГГГ
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
        """Преобразует названия в ключи YYYY-MM-DD."""
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
        """
        Ищет строки с данными этажей.
        Критерий: в первом столбце – число (номер этажа), в строке достаточно чисел (не менее 2 * cycle_count).
        """
        floor_rows = {}
        for idx, row in df_raw.iterrows():
            # Проверяем первый столбец
            first_val = row.iloc[0] if len(row) > 0 else None
            if pd.isna(first_val):
                continue
            try:
                floor = int(float(first_val))
            except (ValueError, TypeError):
                continue
            # Считаем числовые значения в строке (начиная со второго столбца)
            numbers = [v for v in row.iloc[1:] if pd.notna(v) and isinstance(v, (int, float))]
            if len(numbers) >= 2 * cycle_count:
                floor_rows[floor] = idx
        return floor_rows

    @staticmethod
    def _extract_angle_pairs(row: pd.Series) -> List[Tuple[float, float]]:
        """Извлекает пары (αx, αy) из строки (начиная со второго столбца)."""
        values = [float(v) for v in row.iloc[1:] if pd.notna(v) and isinstance(v, (int, float))]
        pairs = [(values[i], values[i+1]) for i in range(0, len(values)-1, 2)]
        return pairs

    @classmethod
    def parse_inclinometer(cls, file_bytes: bytes, sheet_name: str) -> Optional[pd.DataFrame]:
        """
        Парсит лист с наклономером.
        Возвращает DataFrame с колонками: Цикл, Цикл_полное, Этаж, αx, αy.
        """
        try:
            df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)
        except Exception as e:
            st.error(f"Не удалось прочитать лист '{sheet_name}': {e}")
            return None

        # Находим строку с заголовками циклов
        header_row, cycle_names = cls._extract_cycle_headers(df_raw)
        if header_row is None:
            st.error("Не найдена строка с заголовками циклов.")
            return None
        cycle_labels = cls._parse_cycle_labels(cycle_names)
        if not cycle_labels:
            st.error("Не удалось распознать даты циклов.")
            return None

        # Ищем строки с этажами
        floor_rows = cls._find_floor_rows(df_raw, len(cycle_labels))
        if len(floor_rows) < 2:
            st.warning("Найдено менее двух этажей. Проверьте структуру файла.")
            return None

        # Собираем данные
        data = []
        for floor, row_idx in floor_rows.items():
            row = df_raw.iloc[row_idx]
            pairs = cls._extract_angle_pairs(row)
            # Обрезаем до количества циклов
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
        Парсит лист с осадками, возвращает DataFrame с углами крена по осадкам.
        """
        try:
            df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)
        except Exception as e:
            st.error(f"Не удалось прочитать лист '{sheet_name}': {e}")
            return None

        # Находим заголовки циклов
        header_row, cycle_names = cls._extract_cycle_headers(df_raw)
        if header_row is None:
            st.error("Не найдена строка с заголовками циклов в листе осадок.")
            return None
        cycle_labels = cls._parse_cycle_labels(cycle_names)
        if not cycle_labels:
            st.error("Не удалось распознать даты циклов в осадках.")
            return None

        # Ищем столбцы с осадками (после заголовка "Осадка" или сразу после цикла)
        # Упрощённо: предполагаем, что осадки находятся в столбце, следующем за заголовком цикла
        # Но для надёжности ищем столбцы, содержащие слово "осадк" (регистронезависимо)
        osad_cols = {}
        for col_idx, cell in df_raw.iloc[header_row, :].items():
            if pd.notna(cell) and 'осадк' in str(cell).lower():
                osad_cols[col_idx] = str(cell).strip()

        # Для каждого цикла определяем столбец с осадками
        cycle_to_col = {}
        for idx, label in enumerate(cycle_labels):
            # Ищем столбец, где в заголовке есть "Цикл" и дата
            col_found = None
            for col_idx, cell in df_raw.iloc[header_row, :].items():
                if pd.notna(cell) and 'Цикл' in str(cell) and label in str(cell):
                    # Ищем осадки через 1-2 столбца
                    for offset in [1, 2, 3]:
                        check_col = col_idx + offset
                        if check_col in osad_cols or (check_col < len(df_raw.columns) and
                                                       any('осадк' in str(df_raw.iloc[header_row, check_col]).lower() for _ in [0])):
                            col_found = check_col
                            break
                    if col_found is None:
                        # если не нашли, берём col_idx+1
                        col_found = col_idx + 1
                    break
            if col_found is not None and col_found < len(df_raw.columns):
                cycle_to_col[label] = col_found
            else:
                st.warning(f"Не найден столбец с осадками для цикла {label}")

        if not cycle_to_col:
            st.error("Не найдены столбцы с осадками.")
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

        # Собираем данные по осадкам для каждого цикла и марки
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

        # Выбираем нулевой цикл – первый по дате
        sorted_cycles = sorted(cycle_labels)
        zero_cycle = sorted_cycles[0] if sorted_cycles else None
        if zero_cycle is None:
            st.error("Нет циклов для определения нулевого.")
            return None

        # Проверяем наличие выбранных марок
        corner_marks_str = [str(m) for m in corner_marks]
        # Проверяем, что все марки есть в нулевом цикле
        for mark in corner_marks_str:
            if mark not in marks_data[zero_cycle] or np.isnan(marks_data[zero_cycle][mark]):
                st.warning(f"Марка {mark} отсутствует в нулевом цикле {zero_cycle}. Попробуйте другие марки.")
                return None

        # Вычисляем приросты для каждого цикла
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
            # Расчёт крена по осадкам
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
# ОБРАБОТЧИК ДАННЫХ (ПАРАМЕТРЫ + ФИЛЬТРАЦИЯ)
# ------------------------------------------------------------
class DataProcessor:
    @staticmethod
    def apply_parameters(df_incl: pd.DataFrame, zero_cycle: str,
                         alpha0_x: float, alpha0_y: float, L: float) -> pd.DataFrame:
        """
        Вычисляет абсолютные углы и смещения.
        """
        df = df_incl.copy()
        zero_data = df[df['Цикл'] == zero_cycle][['Этаж', 'αx', 'αy']].rename(
            columns={'αx': 'αx0_inc', 'αy': 'αy0_inc'}
        )
        df = df.merge(zero_data, on='Этаж', how='left')
        # Если для какого-то этажа нет данных в нулевом цикле – предупреждение и замена на 0
        if df['αx0_inc'].isna().any() or df['αy0_inc'].isna().any():
            st.warning("Некоторые этажи отсутствуют в нулевом цикле. Для них начальные углы приняты за 0.")
            df['αx0_inc'] = df['αx0_inc'].fillna(0)
            df['αy0_inc'] = df['αy0_inc'].fillna(0)

        df['αx_abs'] = df['αx'] - df['αx0_inc'] + alpha0_x
        df['αy_abs'] = df['αy'] - df['αy0_inc'] + alpha0_y
        df['Смещение X'] = L * np.sin(np.radians(df['αx_abs']))
        df['Смещение Y'] = L * np.sin(np.radians(df['αy_abs']))
        return df

    @staticmethod
    def filter_data(df: pd.DataFrame, filter_type: str, window: int, L: float) -> pd.DataFrame:
        """
        Применяет фильтр к абсолютным углам (αx_abs, αy_abs) и пересчитывает смещения.
        """
        if filter_type == "Нет" or window < 2:
            return df.copy()

        df_filtered = df.copy()
        for floor in df['Этаж'].unique():
            mask = df['Этаж'] == floor
            floor_data = df[mask].sort_values('Цикл')
            # Применяем rolling
            if filter_type == "Скользящее среднее":
                filtered_x = floor_data['αx_abs'].rolling(window=window, center=True, min_periods=1).mean()
                filtered_y = floor_data['αy_abs'].rolling(window=window, center=True, min_periods=1).mean()
            elif filter_type == "Медианный фильтр":
                filtered_x = floor_data['αx_abs'].rolling(window=window, center=True, min_periods=1).median()
                filtered_y = floor_data['αy_abs'].rolling(window=window, center=True, min_periods=1).median()
            else:
                continue

            df_filtered.loc[floor_data.index, 'αx_abs'] = filtered_x.values
            df_filtered.loc[floor_data.index, 'αy_abs'] = filtered_y.values
            # Пересчёт смещений
            df_filtered.loc[floor_data.index, 'Смещение X'] = L * np.sin(np.radians(filtered_x.values))
            df_filtered.loc[floor_data.index, 'Смещение Y'] = L * np.sin(np.radians(filtered_y.values))

        return df_filtered

    @staticmethod
    @st.cache_data
    def calculate_displacement(df_incl_filtered: pd.DataFrame, cycle: str, floors: List[int],
                               L: float, vertical_scale: float = 1.0):
        """
        Возвращает точки деформированной оси и параметры крена.
        """
        df_cycle = df_incl_filtered[df_incl_filtered['Цикл'] == cycle]
        if floors:
            df_cycle = df_cycle[df_cycle['Этаж'].isin(floors)]
        df_cycle = df_cycle.sort_values('Этаж')

        points = [(0, 0, 0)]
        cum_x, cum_y = 0.0, 0.0
        prev_floor = 0

        for _, row in df_cycle.iterrows():
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

        if len(points) < 2:
            return None, None, None, None, None

        top_x, top_y, top_z = points[-1]
        max_z = max(p[2] for p in points)
        return points, top_x, top_y, top_z, max_z

# ------------------------------------------------------------
# ВИЗУАЛИЗАТОР
# ------------------------------------------------------------
class Visualizer:
    @staticmethod
    def plot_3d(points, top_x, top_y, top_z, max_z,
                df_floors, cycle, building_length, building_width,
                vertical_scale, df_sett_angles=None):
        """
        Строит 3D-модель.
        """
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

        # Векторы смещений и подписи этажей
        for i, (x, y, z) in enumerate(points[1:], start=1):
            floor = df_floors.iloc[i-1]['Этаж']
            fig.add_trace(go.Scatter3d(
                x=[0, x], y=[0, y], z=[z, z],
                mode='lines+markers',
                line=dict(color='green', width=3, dash='dot'),
                marker=dict(size=6, color='green', symbol='circle'),
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
            sett_row = df_sett_angles[df_sett_angles['Цикл'] == cycle]
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
            title=f"3D-модель здания – цикл {cycle} (верт. масштаб {vertical_scale:.1f})",
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
            legend=dict(orientation="v", yanchor="top", y=1, xanchor="left", x=1.02)
        )
        return fig

# ------------------------------------------------------------
# ГЛАВНОЕ ПРИЛОЖЕНИЕ
# ------------------------------------------------------------
def main():
    st.set_page_config(page_title="Анализ наклономера + 3D-модель", page_icon="📐", layout="wide")
    st.title("📐 3D-модель здания по данным накладного инклинометра и осадок")
    st.markdown("Загрузите Excel-файл и настройте параметры – модель построится автоматически.")

    # Инициализация состояния
    state = AppState()
    state.update_from_session()

    # Боковая панель
    st.sidebar.header("Загрузка файла")
    uploaded_file = st.file_uploader("Выберите Excel-файл", type=["xlsx", "xls"])

    if uploaded_file is not None:
        file_bytes = uploaded_file.read()
        try:
            xl = ExcelParser.load_excel(file_bytes)
            all_sheets = xl.sheet_names
        except Exception as e:
            st.error(f"Ошибка чтения файла: {e}")
            st.stop()

        if "Наклономер" not in all_sheets:
            st.error("Лист 'Наклономер' не найден. Проверьте файл.")
            st.stop()

        # Парсинг наклономера (с кэшированием)
        df_incl = ExcelParser.parse_inclinometer(file_bytes, "Наклономер")
        if df_incl is None:
            st.stop()

        # Получаем список циклов
        cycles = sorted(df_incl['Цикл'].unique())
        if not cycles:
            st.error("Нет данных по циклам.")
            st.stop()

        # Определяем отображение циклов
        cycle_display = {c: df_incl[df_incl['Цикл'] == c]['Цикл_полное'].iloc[0] for c in cycles}
        state.cycles = cycles
        state.cycle_display_map = cycle_display

        # Боковая панель: параметры
        st.sidebar.header("Параметры модели")
        # Нулевой цикл
        zero_cycle_idx = cycles.index(state.zero_cycle) if state.zero_cycle in cycles else 0
        zero_cycle = st.sidebar.selectbox(
            "Нулевой цикл",
            options=cycles,
            index=zero_cycle_idx,
            format_func=lambda x: cycle_display[x],
            key="zero_cycle_select"
        )
        # L
        L = st.sidebar.number_input("Высота этажа L, м", value=state.L, step=0.1, format="%.1f", key="L_input")
        # α0
        alpha0_x = st.sidebar.number_input("αx0, °", value=state.alpha0_x, step=0.001, format="%.3f", key="alpha0_x")
        alpha0_y = st.sidebar.number_input("αy0, °", value=state.alpha0_y, step=0.001, format="%.3f", key="alpha0_y")

        # Фильтр
        st.sidebar.subheader("Фильтрация данных")
        filter_type = st.sidebar.selectbox("Тип фильтра", FILTER_TYPES, index=FILTER_TYPES.index(state.filter_type), key="filter_type")
        filter_window = st.sidebar.number_input("Размер окна (циклы)", min_value=2, max_value=15,
                                                value=state.filter_window, step=1, key="filter_window")

        # Этажи
        all_floors = sorted(df_incl['Этаж'].unique())
        selected_floors = st.sidebar.multiselect("Выберите этажи", all_floors, default=all_floors, key="selected_floors")

        # Размеры здания
        building_length = st.sidebar.number_input("Длина здания, м", value=state.building_length, step=0.1, key="building_length")
        building_width = st.sidebar.number_input("Ширина здания, м", value=state.building_width, step=0.1, key="building_width")
        vertical_scale = st.sidebar.slider("Вертикальный масштаб", 0.5, 2.0, state.vertical_scale, 0.1, key="vertical_scale")

        # Обновляем состояние
        state.zero_cycle = zero_cycle
        state.L = L
        state.alpha0_x = alpha0_x
        state.alpha0_y = alpha0_y
        state.filter_type = filter_type
        state.filter_window = filter_window
        state.selected_floors = selected_floors
        state.building_length = building_length
        state.building_width = building_width
        state.vertical_scale = vertical_scale
        state.save_to_session()

        # ---------- Обработка данных ----------
        # Применяем параметры к сырым данным (с кэшированием)
        @st.cache_data
        def process_raw(df, zero_cycle, a0x, a0y, L):
            return DataProcessor.apply_parameters(df, zero_cycle, a0x, a0y, L)

        df_raw = process_raw(df_incl, zero_cycle, alpha0_x, alpha0_y, L)
        state.df_incl_raw = df_raw

        # Фильтрация (с кэшированием)
        @st.cache_data
        def process_filtered(df_raw, filter_type, filter_window, L):
            return DataProcessor.filter_data(df_raw, filter_type, filter_window, L)

        df_filtered = process_filtered(df_raw, filter_type, filter_window, L)
        state.df_incl_filtered = df_filtered

        # Данные осадок (если лист есть)
        if "Стилобат" in all_sheets:
            st.sidebar.subheader("Осадки")
            if st.sidebar.checkbox("Рассчитать крен по осадкам", value=False):
                corner_marks = st.sidebar.text_input("Марки (через запятую)", "1,4,11,14")
                L_fund = st.sidebar.number_input("Длина фундамента, м", 70.46, key="L_fund")
                B_fund = st.sidebar.number_input("Ширина фундамента, м", 18.69, key="B_fund")
                if st.sidebar.button("Рассчитать"):
                    marks = [int(x.strip()) for x in corner_marks.split(',') if x.strip()]
                    if len(marks) != 4:
                        st.error("Введите ровно 4 марки.")
                    else:
                        df_sett = ExcelParser.parse_settlement(file_bytes, "Стилобат", marks, L_fund, B_fund)
                        if df_sett is not None:
                            state.df_sett_angles = df_sett
                            st.session_state.df_sett_angles = df_sett
                            st.success("Углы по осадкам рассчитаны.")
        else:
            state.df_sett_angles = None

        # ---------- Вкладки ----------
        tab1, tab2, tab3 = st.tabs(["📊 Данные и параметры", "🏢 3D-модель", "📘 О приборе"])

        with tab1:
            st.subheader("Данные наклономера")
            # Отображение таблицы с возможностью выбора этажей
            show_floors = st.multiselect("Показать этажи", all_floors, default=selected_floors, key="show_floors_tab1")
            if show_floors:
                df_show = df_filtered[df_filtered['Этаж'].isin(show_floors)].copy()
                # Преобразуем в удобный формат: сводная таблица
                df_pivot = df_show.pivot_table(index=['Этаж', 'Цикл'],
                                               values=['αx_abs', 'αy_abs', 'Смещение X', 'Смещение Y'],
                                               aggfunc='first').reset_index()
                df_pivot['Цикл'] = df_pivot['Цикл'].map(cycle_display)
                st.dataframe(df_pivot, use_container_width=True, height=400)

                # Сравнение с сырыми данными (если фильтр включён)
                if filter_type != "Нет":
                    with st.expander("📊 Сравнение с сырыми данными (без фильтра)"):
                        df_raw_show = df_raw[df_raw['Этаж'].isin(show_floors)].copy()
                        df_raw_pivot = df_raw_show.pivot_table(index=['Этаж', 'Цикл'],
                                                                values=['αx_abs', 'αy_abs'],
                                                                aggfunc='first').reset_index()
                        df_raw_pivot['Цикл'] = df_raw_pivot['Цикл'].map(cycle_display)
                        st.dataframe(df_raw_pivot, use_container_width=True, height=300)

            # Отображение информации о фильтре
            st.caption(f"Фильтр: {filter_type}, окно = {filter_window}")

            if state.df_sett_angles is not None:
                st.subheader("Крен по осадкам")
                st.dataframe(state.df_sett_angles, use_container_width=True)

        with tab2:
            st.subheader("3D-модель здания")
            if len(cycles) == 0:
                st.warning("Нет циклов для отображения.")
            else:
                # Выбор цикла
                default_cycle = state.selected_cycle_key if state.selected_cycle_key in cycles else cycles[-1]
                selected_cycle = st.selectbox("Выберите цикл", cycles, index=cycles.index(default_cycle),
                                              format_func=lambda x: cycle_display[x], key="cycle_3d")
                state.selected_cycle_key = selected_cycle
                state.save_to_session()

                # Расчёт смещений
                points, top_x, top_y, top_z, max_z = DataProcessor.calculate_displacement(
                    df_filtered, selected_cycle, selected_floors, L, vertical_scale
                )
                if points is None:
                    st.warning("Для выбранного цикла и этажей нет данных.")
                else:
                    df_cycle = df_filtered[df_filtered['Цикл'] == selected_cycle]
                    df_floors = df_cycle[df_cycle['Этаж'].isin(selected_floors)].sort_values('Этаж')
                    fig = Visualizer.plot_3d(
                        points, top_x, top_y, top_z, max_z,
                        df_floors, selected_cycle, building_length, building_width,
                        vertical_scale, state.df_sett_angles
                    )
                    st.plotly_chart(fig, use_container_width=True)

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

    else:
        st.info("👆 Загрузите Excel-файл для начала работы.")

if __name__ == "__main__":
    main()
