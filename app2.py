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
    }
    for key, default in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = default

# ------------------------------------------------------------
# ПАРСЕР EXCEL
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

        osad_cols = {}
        for col_idx, cell in df_raw.iloc[header_row, :].items():
            if pd.notna(cell) and 'осадк' in str(cell).lower():
                osad_cols[col_idx] = str(cell).strip()

        cycle_to_col = {}
        for idx, label in enumerate(cycle_labels):
            col_found = None
            for col_idx, cell in df_raw.iloc[header_row, :].items():
                if pd.notna(cell) and 'Цикл' in str(cell) and label in str(cell):
                    for offset in [1, 2, 3]:
                        check_col = col_idx + offset
                        if check_col in osad_cols or (check_col < len(df_raw.columns) and
                                                       any('осадк' in str(df_raw.iloc[header_row, check_col]).lower() for _ in [0])):
                            col_found = check_col
                            break
                    if col_found is None:
                        col_found = col_idx + 1
                    break
            if col_found is not None and col_found < len(df_raw.columns):
                cycle_to_col[label] = col_found
            else:
                st.warning(f"Не найден столбец с осадками для цикла {label}")

        if not cycle_to_col:
            st.error("Не найдены столбцы с осадками.")
            return None

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
# ОБРАБОТЧИК ДАННЫХ
# ------------------------------------------------------------
class DataProcessor:
    @staticmethod
    def apply_parameters(df_incl: pd.DataFrame, zero_cycle: str,
                         alpha0_x: float, alpha0_y: float, L: float) -> pd.DataFrame:
        df = df_incl.copy()
        zero_data = df[df['Цикл'] == zero_cycle][['Этаж', 'αx', 'αy']].rename(
            columns={'αx': 'αx0_inc', 'αy': 'αy0_inc'}
        )
        df = df.merge(zero_data, on='Этаж', how='left')
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
        if filter_type == "Нет" or window < 2:
            return df.copy()

        df_filtered = df.copy()
        for floor in df['Этаж'].unique():
            mask = df['Этаж'] == floor
            floor_data = df[mask].sort_values('Цикл')
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
            df_filtered.loc[floor_data.index, 'Смещение X'] = L * np.sin(np.radians(filtered_x.values))
            df_filtered.loc[floor_data.index, 'Смещение Y'] = L * np.sin(np.radians(filtered_y.values))

        return df_filtered

    @staticmethod
    @st.cache_data
    def calculate_displacement(df_incl_filtered: pd.DataFrame, cycle: str, floors: List[int],
                               L: float, vertical_scale: float = 1.0):
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
            return None, None, None, None, None, None

        top_x, top_y, top_z = points[-1]
        max_z = max(p[2] for p in points)
        return points, top_x, top_y, top_z, max_z, df_cycle

# ------------------------------------------------------------
# ВИЗУАЛИЗАТОР (УЛУЧШЕННАЯ 3D-МОДЕЛЬ)
# ------------------------------------------------------------
class Visualizer:
    @staticmethod
    def plot_3d(points, top_x, top_y, top_z, max_z,
                df_cycle, cycle, building_length, building_width,
                vertical_scale, floors, df_sett_angles=None):
        # Фильтруем df_floors
        if floors:
            df_floors = df_cycle[df_cycle['Этаж'].isin(floors)].sort_values('Этаж')
        else:
            df_floors = df_cycle.sort_values('Этаж')

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
            line=dict(color='red', width=6),
            marker=dict(size=10, color='red', symbol='circle'),
            name='Деформированная ось'
        ))

        # Векторы смещений
        for i, (x, y, z) in enumerate(points[1:], start=1):
            if i-1 >= len(df_floors):
                continue
            floor = df_floors.iloc[i-1]['Этаж']
            fig.add_trace(go.Scatter3d(
                x=[0, x], y=[0, y], z=[z, z],
                mode='lines+markers',
                line=dict(color='lime', width=4, dash='dot'),
                marker=dict(size=8, color='lime', symbol='circle'),
                showlegend=False
            ))
            dist = np.sqrt(x**2 + y**2)
            fig.add_trace(go.Scatter3d(
                x=[x], y=[y], z=[z],
                mode='text',
                text=[f"<b>{floor}эт</b><br>{dist:.3f} м"],
                textposition='top center',
                textfont=dict(color='darkgreen', size=11),
                showlegend=False
            ))

        # Наклономеры
        for i, (x, y, z) in enumerate(points[1:], start=1):
            if i-1 >= len(df_floors):
                continue
            floor = df_floors.iloc[i-1]['Этаж']
            alpha_x = df_floors.iloc[i-1]['αx_abs']
            alpha_y = df_floors.iloc[i-1]['αy_abs']
            fig.add_trace(go.Scatter3d(
                x=[x], y=[y], z=[z],
                mode='markers+text',
                marker=dict(size=16, color='blue', symbol='square', line=dict(color='darkblue', width=2)),
                text=[f"<b>Этаж {floor}</b><br>αx={alpha_x:.3f}°<br>αy={alpha_y:.3f}°"],
                textposition='top center',
                textfont=dict(color='blue', size=11),
                name=f'Наклономер {floor}'
            ))

        # Общий крен
        fig.add_trace(go.Scatter3d(
            x=[0, top_x], y=[0, top_y], z=[0, top_z],
            mode='lines+markers',
            line=dict(color='darkorange', width=8),
            marker=dict(size=12, color='darkorange', symbol='diamond'),
            name='Общий крен здания'
        ))
        kren_angle = np.degrees(np.arctan2(np.sqrt(top_x**2 + top_y**2), top_z))
        fig.add_trace(go.Scatter3d(
            x=[top_x], y=[top_y], z=[top_z],
            mode='text',
            text=[f"<b>Крен: {kren_angle:.2f}°</b>"],
            textposition='top center',
            textfont=dict(color='darkorange', size=14),
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
                    line=dict(color='purple', width=6, dash='dash'),
                    marker=dict(size=12, color='purple', symbol='diamond'),
                    name=f'Крен по осадкам<br>a={a:.2f} мм/м, b={b:.2f} мм/м'
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
                line=dict(color='rgba(50,50,50,0.6)', width=2),
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
                    line=dict(color='rgba(50,50,50,0.6)', width=2),
                    showlegend=False
                ))

        fig.update_layout(
            title=f"<b>3D-модель здания – цикл {cycle}</b> (верт. масштаб {vertical_scale:.1f})",
            scene=dict(
                xaxis_title="Смещение X, м",
                yaxis_title="Смещение Y, м",
                zaxis_title="Высота, м",
                aspectmode='data',
                camera=dict(eye=dict(x=1.5, y=1.5, z=1.2)),
                bgcolor='rgba(240,240,240,0.9)'
            ),
            width=950,
            height=800,
            template="plotly_white",
            legend=dict(orientation="v", yanchor="top", y=1, xanchor="left", x=1.02, font=dict(size=12)),
            margin=dict(l=0, r=0, b=0, t=50)
        )
        return fig

# ------------------------------------------------------------
# СТРАНИЦА "О ФОРМАТЕ ДАННЫХ" (без генератора примера)
# ------------------------------------------------------------
def show_data_format():
    st.markdown("""
    ## 📋 Ожидаемый формат исходных данных

    Приложение принимает файл Excel (`.xlsx` или `.xls`) с двумя листами:

    ### 1. Лист **«Наклономер»**
    Содержит результаты измерений наклономером на разных этажах.

    **Структура:**
    - Первая строка – заголовки циклов. В ячейках, начиная со столбца **B**, должны быть написаны **Цикл ДД.ММ.ГГГГ** (например, `Цикл 01.01.2025`). Между датами могут быть пустые столбцы (они игнорируются).
    - Начиная со второй строки – данные по этажам. В **столбце A** – номер этажа (число). В столбцах B, C, D, E, … – **попарно** значения углов αx и αy для каждого цикла. То есть для одного цикла используются **две соседние ячейки**: сначала αx, затем αy.

    **Пример:**
    |       | A   | B                 | C   | D                 | E   | ... |
    |-------|-----|-------------------|-----|-------------------|-----|-----|
    | **1** |     | **Цикл 01.01.2025**|     | **Цикл 10.01.2025**|     |     |
    | **2** | **5**| 0.12              | 0.05| 0.15              | 0.07| ... |
    | **3** | **15**| 0.08             | 0.03| 0.09              | 0.04| ... |
    | **4** | **27**| 0.05             | 0.02| 0.06              | 0.02| ... |

    ### 2. Лист **«Стилобат»** (опционально)
    Содержит данные осадок угловых марок фундамента.

    **Структура:**
    - Первая строка – заголовки циклов (аналогично листу «Наклономер»). В столбцах, где написано «Цикл ...», через один-два столбца должны быть значения осадок (можно подписать «Осадка, мм»).
    - В **столбце A** – номера марок (числа). В строках – значения осадок (в мм) для каждого цикла (по одному числу на цикл).

    **Пример:**
    |       | A   | B                 | C          | D                 | E          | ... |
    |-------|-----|-------------------|------------|-------------------|------------|-----|
    | **1** |     | **Цикл 01.01.2025**| **Осадка** | **Цикл 10.01.2025**| **Осадка** |     |
    | **2** | **1**|                   | 0.0        |                   | 0.5        | ... |
    | **3** | **4**|                   | 0.2        |                   | 0.7        | ... |
    | **4** | **11**|                  | 0.1        |                   | 0.6        | ... |
    | **5** | **14**|                  | 0.3        |                   | 0.8        | ... |

    > **Важно:** Нумерация марок должна соответствовать углам здания: сначала нижний левый, затем нижний правый, затем верхний левый, затем верхний правый.
    """)

# ------------------------------------------------------------
# ГЛАВНОЕ ПРИЛОЖЕНИЕ
# ------------------------------------------------------------
def main():
    st.set_page_config(page_title="Анализ наклономера + 3D-модель", page_icon="📐", layout="wide")
    st.title("📐 3D-модель здания по данным накладного инклинометра и осадок")
    st.markdown("Загрузите Excel-файл и настройте параметры – модель построится автоматически.")

    init_session_state()

    # Вкладки
    tab_data, tab_model, tab_info, tab_format = st.tabs([
        "📊 Данные и параметры",
        "🏢 3D-модель",
        "📘 О приборе",
        "📋 Формат данных"
    ])

    # --------------------- ВКЛАДКА "ФОРМАТ ДАННЫХ" ---------------------
    with tab_format:
        show_data_format()

    # --------------------- ЗАГРУЗКА ФАЙЛА ---------------------
    uploaded_file = st.sidebar.file_uploader("Выберите Excel-файл", type=["xlsx", "xls"])

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

        # Парсинг наклономера
        df_incl = ExcelParser.parse_inclinometer(file_bytes, "Наклономер")
        if df_incl is None:
            st.stop()

        cycles = sorted(df_incl['Цикл'].unique())
        if not cycles:
            st.error("Нет данных по циклам.")
            st.stop()

        cycle_display = {c: df_incl[df_incl['Цикл'] == c]['Цикл_полное'].iloc[0] for c in cycles}
        st.session_state["cycles"] = cycles
        st.session_state["cycle_display_map"] = cycle_display
        st.session_state["file_loaded"] = True

        # --------------------- ПАРАМЕТРЫ В БОКОВОЙ ПАНЕЛИ ---------------------
        st.sidebar.header("Параметры модели")
        zero_cycle = st.sidebar.selectbox(
            "Нулевой цикл",
            options=cycles,
            index=cycles.index(st.session_state["zero_cycle"]) if st.session_state["zero_cycle"] in cycles else 0,
            format_func=lambda x: cycle_display[x],
            key="zero_cycle_select"
        )
        L = st.sidebar.number_input("Высота этажа L, м", value=st.session_state["L"], step=0.1, format="%.1f", key="L_input")
        alpha0_x = st.sidebar.number_input("αx0, °", value=st.session_state["alpha0_x"], step=0.001, format="%.3f", key="alpha0_x")
        alpha0_y = st.sidebar.number_input("αy0, °", value=st.session_state["alpha0_y"], step=0.001, format="%.3f", key="alpha0_y")

        st.sidebar.subheader("Фильтрация данных")
        filter_type = st.sidebar.selectbox("Тип фильтра", FILTER_TYPES, index=FILTER_TYPES.index(st.session_state["filter_type"]), key="filter_type")
        filter_window = st.sidebar.number_input("Размер окна (циклы)", min_value=2, max_value=15,
                                                value=st.session_state["filter_window"], step=1, key="filter_window")

        all_floors = sorted(df_incl['Этаж'].unique())
        selected_floors = st.sidebar.multiselect("Выберите этажи", all_floors, default=st.session_state["selected_floors"] or all_floors, key="selected_floors")

        building_length = st.sidebar.number_input("Длина здания, м", value=st.session_state["building_length"], step=0.1, key="building_length")
        building_width = st.sidebar.number_input("Ширина здания, м", value=st.session_state["building_width"], step=0.1, key="building_width")
        vertical_scale = st.sidebar.slider("Вертикальный масштаб", 0.5, 2.0, st.session_state["vertical_scale"], 0.1, key="vertical_scale")

        # Чтение значений из состояния
        zero_cycle = st.session_state["zero_cycle_select"]
        L = st.session_state["L_input"]
        alpha0_x = st.session_state["alpha0_x"]
        alpha0_y = st.session_state["alpha0_y"]
        filter_type = st.session_state["filter_type"]
        filter_window = st.session_state["filter_window"]
        selected_floors = st.session_state["selected_floors"]
        building_length = st.session_state["building_length"]
        building_width = st.session_state["building_width"]
        vertical_scale = st.session_state["vertical_scale"]

        # --------------------- ОБРАБОТКА ДАННЫХ (АВТОМАТИЧЕСКИ ПРИ ИЗМЕНЕНИИ ПАРАМЕТРОВ) ---------------------
        @st.cache_data
        def process_raw(df, zero_cycle, a0x, a0y, L):
            return DataProcessor.apply_parameters(df, zero_cycle, a0x, a0y, L)

        df_raw = process_raw(df_incl, zero_cycle, alpha0_x, alpha0_y, L)
        st.session_state["df_incl_raw"] = df_raw

        @st.cache_data
        def process_filtered(df_raw, filter_type, filter_window, L):
            return DataProcessor.filter_data(df_raw, filter_type, filter_window, L)

        df_filtered = process_filtered(df_raw, filter_type, filter_window, L)
        st.session_state["df_incl_filtered"] = df_filtered

        # Осадки (с явной кнопкой, так как требует парсинга)
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
                            st.session_state["df_sett_angles"] = df_sett
                            st.success("Углы по осадкам рассчитаны.")
        else:
            st.session_state["df_sett_angles"] = None

        # --------------------- ВКЛАДКА "ДАННЫЕ И ПАРАМЕТРЫ" ---------------------
        with tab_data:
            st.subheader("Данные наклономера")
            show_floors = st.multiselect("Показать этажи", all_floors, default=selected_floors, key="show_floors_tab1")
            if show_floors:
                df_show = df_filtered[df_filtered['Этаж'].isin(show_floors)].copy()
                df_pivot = df_show.pivot_table(index=['Этаж', 'Цикл'],
                                               values=['αx_abs', 'αy_abs', 'Смещение X', 'Смещение Y'],
                                               aggfunc='first').reset_index()
                df_pivot['Цикл'] = df_pivot['Цикл'].map(cycle_display)
                st.dataframe(df_pivot, use_container_width=True, height=400)

                if filter_type != "Нет":
                    with st.expander("📊 Сравнение с сырыми данными (без фильтра)"):
                        df_raw_show = df_raw[df_raw['Этаж'].isin(show_floors)].copy()
                        df_raw_pivot = df_raw_show.pivot_table(index=['Этаж', 'Цикл'],
                                                                values=['αx_abs', 'αy_abs'],
                                                                aggfunc='first').reset_index()
                        df_raw_pivot['Цикл'] = df_raw_pivot['Цикл'].map(cycle_display)
                        st.dataframe(df_raw_pivot, use_container_width=True, height=300)

            st.caption(f"Фильтр: {filter_type}, окно = {filter_window}")

            if st.session_state["df_sett_angles"] is not None:
                st.subheader("Крен по осадкам")
                st.dataframe(st.session_state["df_sett_angles"], use_container_width=True)

        # --------------------- ВКЛАДКА "3D-МОДЕЛЬ" ---------------------
        with tab_model:
            st.subheader("3D-модель здания")
            if len(cycles) == 0:
                st.warning("Нет циклов для отображения.")
            else:
                default_cycle = st.session_state["selected_cycle_key"] if st.session_state["selected_cycle_key"] in cycles else cycles[-1]
                selected_cycle = st.selectbox("Выберите цикл", cycles, index=cycles.index(default_cycle),
                                              format_func=lambda x: cycle_display[x], key="cycle_3d")
                st.session_state["selected_cycle_key"] = selected_cycle

                result = DataProcessor.calculate_displacement(
                    df_filtered, selected_cycle, selected_floors, L, vertical_scale
                )
                if result[0] is None:
                    st.warning("Для выбранного цикла и этажей нет данных.")
                else:
                    points, top_x, top_y, top_z, max_z, df_cycle = result
                    fig = Visualizer.plot_3d(
                        points, top_x, top_y, top_z, max_z,
                        df_cycle, selected_cycle, building_length, building_width,
                        vertical_scale, selected_floors, st.session_state["df_sett_angles"]
                    )
                    st.plotly_chart(fig, use_container_width=True)

                    # Дополнительная информация о крене
                    st.markdown("### 📐 Результаты для выбранного цикла")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Общий крен здания", f"{np.degrees(np.arctan2(np.sqrt(top_x**2 + top_y**2), top_z)):.3f}°")
                    with col2:
                        st.metric("Смещение верха (X)", f"{top_x:.3f} м")
                    with col3:
                        st.metric("Смещение верха (Y)", f"{top_y:.3f} м")

        # --------------------- ВКЛАДКА "О ПРИБОРЕ" ---------------------
        with tab_info:
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
        # Показываем информацию о формате даже без загрузки
        with tab_format:
            show_data_format()

if __name__ == "__main__":
    main()
