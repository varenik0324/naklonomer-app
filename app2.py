import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import io
import re
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ------------------------------------------------------------
# Настройки страницы
# ------------------------------------------------------------
st.set_page_config(
    page_title="Анализ наклономера + осадки",
    page_icon="📐",
    layout="wide"
)
st.title("📐 Анализ данных накладного инклинометра и осадок фундамента")
st.markdown("Загрузите Excel-файл с данными измерений, укажите параметры, и приложение построит графики и сформирует отчёты.")

# ------------------------------------------------------------
# ПАРСИНГ ДАННЫХ НАКЛОНОМЕРА
# ------------------------------------------------------------
def parse_inclinometer_data(file_bytes, sheet_name, manual_floor_rows=None, search_start=None, search_end=None):
    try:
        df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)
    except Exception as e:
        st.error(f"Не удалось прочитать лист '{sheet_name}': {e}")
        return None

    total_rows = len(df_raw)
    total_cols = len(df_raw.columns)

    cycle_header_row = None
    for idx, row in df_raw.iterrows():
        row_str = ' '.join(str(cell) for cell in row if pd.notna(cell))
        if 'Цикл' in row_str and re.search(r'\d{2}\.\d{2}\.\d{4}', row_str):
            cycle_header_row = idx
            break

    if cycle_header_row is None:
        st.warning("Не найдена строка с заголовками циклов. Будем использовать порядковые номера.")
        cycle_labels = None
    else:
        cycle_labels = []
        for col in range(1, total_cols):
            cell = df_raw.iloc[cycle_header_row, col]
            if pd.notna(cell):
                cell_str = str(cell).strip()
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
        if search_start is None:
            search_start = cycle_header_row + 1 if cycle_header_row is not None else 0
        if search_end is None:
            end_search = total_rows
            for idx in range(search_start, total_rows):
                row_str = ' '.join(str(cell) for cell in df_raw.iloc[idx] if pd.notna(cell))
                if 'Таблица' in row_str:
                    end_search = idx
                    break
            search_end = end_search
        else:
            search_end = min(search_end, total_rows)

        search_start = max(0, min(search_start, total_rows-1))
        search_end = max(search_start, search_end)

        floor_rows = {}
        for idx in range(search_start, search_end):
            row = df_raw.iloc[idx]
            found_floor = None
            for cell in row:
                if pd.isna(cell):
                    continue
                cell_str = str(cell).strip()
                match = re.search(r'\b(5|15|27)\b', cell_str)
                if match:
                    found_floor = int(match.group(1))
                    break
            if found_floor is not None:
                num_count = sum(1 for v in row if pd.notna(v) and isinstance(v, (int, float)))
                if num_count >= 6:
                    floor_rows[found_floor] = idx
            if len(floor_rows) == 3:
                break

        if len(floor_rows) < 2:
            st.warning("Не удалось автоматически найти строки с этажами (5, 15).")
            st.write("Первые 30 строк листа (первые 10 колонок):")
            st.dataframe(df_raw.iloc[:30, :10])
            return None

    floor_rows_sorted = [floor_rows[f] for f in sorted(floor_rows.keys())]

    max_pairs = 0
    for floor_idx in floor_rows_sorted:
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
        if len(pairs) > max_pairs:
            max_pairs = len(pairs)

    if cycle_labels is None:
        cycle_labels = [f"Цикл {i+1}" for i in range(max_pairs)]
    elif len(cycle_labels) < max_pairs:
        for i in range(len(cycle_labels), max_pairs):
            cycle_labels.append(f"Цикл {i+1}")

    data = []
    for floor_idx in floor_rows_sorted:
        floor_val = None
        for col in range(0, total_cols):
            cell = df_raw.iloc[floor_idx, col]
            if pd.isna(cell):
                continue
            cell_str = str(cell).strip()
            match = re.search(r'\b(5|15|27)\b', cell_str)
            if match:
                floor_val = int(match.group(1))
                break
        if floor_val is None:
            continue

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
# ПАРСИНГ ДАННЫХ ОСАДОК
# ------------------------------------------------------------
def parse_settlement_data(file_bytes, sheet_name, corner_marks, L, B, mark_col=0):
    df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)

    cycle_header_row = None
    for idx, row in df_raw.iterrows():
        row_str = ' '.join(str(cell) for cell in row if pd.notna(cell))
        if 'Цикл' in row_str:
            cycle_header_row = idx
            break
    if cycle_header_row is None:
        st.error("Не найдена строка с заголовками циклов в листе осадок.")
        return None

    cycle_cols = {}
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
                for offset in [1, 2, 3]:
                    if col_idx + offset < len(df_raw.columns):
                        next_cell = df_raw.iloc[cycle_header_row, col_idx + offset]
                        if pd.notna(next_cell):
                            next_str = str(next_cell).strip()
                            if 'осадк' in next_str.lower():
                                cycle_cols[cycle_label] = col_idx + offset
                                break
                if cycle_label not in cycle_cols:
                    if col_idx + 2 < len(df_raw.columns):
                        cycle_cols[cycle_label] = col_idx + 2
                    else:
                        cycle_cols[cycle_label] = col_idx + 1

    if not cycle_cols:
        st.error("Не найдены колонки с осадками для циклов.")
        return None

    mark_rows = []
    for idx in range(cycle_header_row + 1, len(df_raw)):
        cell_val = df_raw.iloc[idx, mark_col]
        if pd.notna(cell_val):
            if isinstance(cell_val, (int, float)):
                mark_rows.append(idx)
            elif isinstance(cell_val, str) and cell_val.strip():
                mark_rows.append(idx)

    if not mark_rows:
        st.warning(f"Не найдены строки с марками в столбце {mark_col}. Показываем превью листа.")
        st.dataframe(df_raw.head(20))
        return None

    data = []
    for cycle_label, col_idx in cycle_cols.items():
        for mark_idx in mark_rows:
            mark_num = df_raw.iloc[mark_idx, mark_col]
            if isinstance(mark_num, (int, float)):
                mark_str = str(int(mark_num)) if mark_num == int(mark_num) else str(mark_num)
            else:
                mark_str = str(mark_num).strip()
            if mark_str not in corner_marks:
                continue
            settlement = df_raw.iloc[mark_idx, col_idx]
            if pd.notna(settlement) and isinstance(settlement, (int, float)):
                data.append({
                    'Цикл': cycle_label,
                    'Марка': mark_str,
                    'Осадка_мм': settlement
                })

    if not data:
        st.error("Не удалось извлечь осадки для выбранных марок. Проверьте правильность номеров марок и столбца.")
        st.write("Найденные марки:", [str(df_raw.iloc[idx, mark_col]) for idx in mark_rows[:10]])
        return None

    df_sett = pd.DataFrame(data)

    if len(corner_marks) != 4:
        st.error("Должно быть ровно 4 угловые марки.")
        return None

    marks_order = list(corner_marks)
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

    df_angles = pd.DataFrame(results)
    return df_angles

# ------------------------------------------------------------
# ГЕНЕРАЦИЯ ОТЧЁТОВ
# ------------------------------------------------------------
def generate_excel_report(df_incl, df_sett_angles, cycles, alpha0_x, alpha0_y, L):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if df_incl is not None:
            df_incl.to_excel(writer, index=False, sheet_name='Наклономер')
        if df_sett_angles is not None:
            df_sett_angles.to_excel(writer, index=False, sheet_name='Осадки_углы')
        params = pd.DataFrame({
            'Параметр': ['Начальный угол X, °', 'Начальный угол Y, °', 'Высота этажа, м', 'Нулевой цикл'],
            'Значение': [alpha0_x, alpha0_y, L, cycles[0] if cycles else '']
        })
        params.to_excel(writer, sheet_name='Параметры', index=False)
    return output.getvalue()

def generate_pdf_report(df_incl, df_sett_angles, cycles, alpha0_x, alpha0_y, L):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, height - 50, "Отчёт по наклономеру и осадкам")
    c.setFont("Helvetica", 12)
    c.drawString(50, height - 80, f"Дата: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
    c.drawString(50, height - 100, f"Параметры: αx0 = {alpha0_x:.3f}°, αy0 = {alpha0_y:.3f}°, L = {L} м")
    c.drawString(50, height - 120, f"Нулевой цикл: {cycles[0] if cycles else ''}")
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, height - 150, "Углы наклона всего здания (последний цикл):")
    y = height - 170
    if df_incl is not None and not df_incl.empty:
        last_cycle = df_incl['Цикл'].iloc[-1]
        # берём минимальный этаж
        min_floor = df_incl['Этаж'].min()
        row = df_incl[(df_incl['Цикл'] == last_cycle) & (df_incl['Этаж'] == min_floor)]
        if not row.empty:
            c.setFont("Helvetica", 10)
            c.drawString(60, y, f"αx = {row['αx_abs'].iloc[0]:.3f}°, αy = {row['αy_abs'].iloc[0]:.3f}°")
    c.save()
    buffer.seek(0)
    return buffer

def generate_word_report(df_incl, df_sett_angles, cycles, alpha0_x, alpha0_y, L):
    doc = Document()
    doc.add_heading('Отчёт по наклономеру и осадкам', level=1)
    doc.add_paragraph(f"Дата: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
    doc.add_paragraph(f"Параметры: αx0 = {alpha0_x:.3f}°, αy0 = {alpha0_y:.3f}°, L = {L} м")
    doc.add_paragraph(f"Нулевой цикл: {cycles[0] if cycles else ''}")
    doc.add_heading('Углы наклона всего здания (последний цикл)', level=2)
    if df_incl is not None and not df_incl.empty:
        last_cycle = df_incl['Цикл'].iloc[-1]
        min_floor = df_incl['Этаж'].min()
        row = df_incl[(df_incl['Цикл'] == last_cycle) & (df_incl['Этаж'] == min_floor)]
        if not row.empty:
            doc.add_paragraph(f"αx = {row['αx_abs'].iloc[0]:.3f}°, αy = {row['αy_abs'].iloc[0]:.3f}°")
    doc.add_paragraph("© Геофундамент, 2026").alignment = WD_ALIGN_PARAGRAPH.CENTER
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ------------------------------------------------------------
# ФОРМАТИРОВАНИЕ МЕТОК ЦИКЛОВ
# ------------------------------------------------------------
def format_cycle_labels(df_incl, zero_cycle):
    unique_cycles = sorted(df_incl['Цикл'].unique())
    def try_parse_date(s):
        for fmt in ('%Y-%m-%d', '%d.%m.%Y', '%d.%m.%Y'):
            try:
                return pd.to_datetime(s, format=fmt)
            except:
                continue
        return None
    try:
        sorted_cycles = sorted(unique_cycles, key=lambda x: try_parse_date(x) or x)
    except:
        sorted_cycles = sorted(unique_cycles)
    
    labels = {}
    for i, cyc in enumerate(sorted_cycles):
        if cyc == zero_cycle:
            try:
                dt = pd.to_datetime(cyc)
                display = f"Нулевой ({dt.strftime('%d.%m.%Y')})"
            except:
                display = f"Нулевой ({cyc})"
        else:
            try:
                dt = pd.to_datetime(cyc)
                display = f"Цикл №{i} ({dt.strftime('%d.%m.%Y')})"
            except:
                display = f"Цикл №{i} ({cyc})"
        labels[cyc] = display
    return labels

# ------------------------------------------------------------
# ОСНОВНАЯ ЛОГИКА ПРИЛОЖЕНИЯ
# ------------------------------------------------------------
uploaded_file = st.file_uploader(
    "Загрузите Excel-файл с данными наклономера и/или осадок",
    type=["xlsx", "xls"],
    help="Файл должен содержать лист с наклономером (строки с этажами 5,15,27) и, опционально, листы с осадками."
)

if uploaded_file is not None:
    try:
        file_bytes = uploaded_file.read()
        xl = pd.ExcelFile(io.BytesIO(file_bytes))
        all_sheets = xl.sheet_names

        st.sidebar.header("Выбор листов")
        incl_sheet_name = st.sidebar.selectbox(
            "Лист с наклономером",
            all_sheets,
            index=all_sheets.index('Наклономер') if 'Наклономер' in all_sheets else 0
        )
        st.sidebar.write(f"Выбран: **{incl_sheet_name}**")

        st.sidebar.header("Параметры поиска строк")
        use_manual_range = st.sidebar.checkbox("Ручной диапазон для поиска этажей", value=False)
        if use_manual_range:
            search_start = st.sidebar.number_input("Начальная строка (индекс)", min_value=0, step=1, value=7)
            search_end = st.sidebar.number_input("Конечная строка (индекс)", min_value=0, step=1, value=13)
        else:
            search_start = None
            search_end = None

        # Парсинг наклономера
        df_incl = parse_inclinometer_data(file_bytes, incl_sheet_name, search_start=search_start, search_end=search_end)

        if df_incl is None:
            st.subheader("🔧 Ручной ввод строк с этажами")
            st.write("Введите номера строк (индексы, начиная с 0), в которых расположены данные для этажей 5, 15, 27.")
            st.info("Номера строк можно увидеть в таблице выше (левая колонка — это индекс).")
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
                df_incl = parse_inclinometer_data(file_bytes, incl_sheet_name, manual_floor_rows=manual_rows)
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

        tab1, tab2, tab3, tab4 = st.tabs(["📊 Данные и параметры", "📈 Графики", "📥 Отчёт", "📘 Наклономер"])

        with tab1:
            st.subheader("Данные наклономера")
            st.dataframe(df_incl, use_container_width=True)

            cycles = sorted(df_incl['Цикл'].unique())
            if len(cycles) == 0:
                st.error("Нет циклов в данных наклономера.")
                st.stop()

            st.subheader("⚙️ Параметры расчёта")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                zero_cycle = st.selectbox("Нулевой цикл", cycles, index=0)
            with col2:
                alpha0_x = st.number_input("Начальный угол X (αx0), °", value=0.0, step=0.001, format="%.3f")
            with col3:
                alpha0_y = st.number_input("Начальный угол Y (αy0), °", value=0.0, step=0.001, format="%.3f")
            with col4:
                L = st.number_input("Высота этажа (L), м", value=3.0, step=0.1, format="%.1f")

            zero_data = df_incl[df_incl['Цикл'] == zero_cycle][['Этаж', 'αx', 'αy']].rename(columns={'αx': 'αx0_inc', 'αy': 'αy0_inc'})
            df_incl = df_incl.merge(zero_data, on='Этаж', how='left')
            df_incl['αx_abs'] = df_incl['αx'] - df_incl['αx0_inc'] + alpha0_x
            df_incl['αy_abs'] = df_incl['αy'] - df_incl['αy0_inc'] + alpha0_y

            cycle_labels = format_cycle_labels(df_incl, zero_cycle)
            st.session_state.cycle_labels = cycle_labels
            st.session_state.zero_cycle = zero_cycle

            st.subheader("🎯 Выбор этажей для отображения на графиках")
            available_floors = sorted(df_incl['Этаж'].unique())
            selected_floors = st.multiselect(
                "Выберите этажи",
                options=available_floors,
                default=available_floors,
                key="floor_selector"
            )
            st.session_state.selected_floors = selected_floors

            # Блок осадок
            sett_sheets = [s for s in all_sheets if 'стилобат' in s.lower() or 'высотн' in s.lower() or 'осадк' in s.lower()]
            if sett_sheets:
                st.subheader("📐 Данные осадок")
                with st.expander("Настройки осадок", expanded=False):
                    selected_sett_sheet = st.selectbox("Выберите лист с осадками", sett_sheets, key="sett_sheet")
                    corner_marks_str = st.text_input(
                        "Номера угловых марок (через запятую, в порядке: нижний левый, нижний правый, верхний левый, верхний правый)",
                        "1,5,9,13"
                    )
                    mark_col = st.number_input(
                        "Номер столбца с марками (0-индекс, обычно 0 или 1)",
                        min_value=0, step=1, value=0
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
                    L_sett = st.number_input("Длина фундамента L, м", value=70.0, step=1.0)
                    B_sett = st.number_input("Ширина фундамента B, м", value=18.0, step=1.0)

                    if st.button("Рассчитать углы по осадкам"):
                        if len(corner_marks) == 4:
                            df_sett_angles = parse_settlement_data(
                                file_bytes, selected_sett_sheet, corner_marks, L_sett, B_sett, mark_col=mark_col
                            )
                            if df_sett_angles is not None:
                                st.success(f"✅ Углы по осадкам рассчитаны для {len(df_sett_angles)} циклов.")
                                st.dataframe(df_sett_angles, use_container_width=True)
                                st.session_state.df_sett_angles = df_sett_angles
                            else:
                                st.error("Не удалось рассчитать углы. Проверьте правильность введённых данных.")
                        else:
                            st.error("Укажите 4 угловые марки.")
                if 'df_sett_angles' in st.session_state:
                    st.dataframe(st.session_state.df_sett_angles, use_container_width=True)
            else:
                st.info("Листов с осадками не найдено.")

        with tab2:
            st.subheader("📈 Изменение абсолютных углов наклона по этажам")
            st.markdown("""
            **ℹ️ Пояснение:** На графике показаны абсолютные углы наклона (в градусах) для выбранных этажей по осям X и Y.  
            - **Сплошная линия** – угол по оси X (αx).  
            - **Пунктирная линия** – угол по оси Y (αy).  
            - Каждый цвет соответствует отдельному этажу.  
            - Аннотации показывают максимальные значения.
            """)

            selected_floors = st.session_state.get('selected_floors', sorted(df_incl['Этаж'].unique()))
            if not selected_floors:
                st.warning("Не выбрано ни одного этажа. Вернитесь на вкладку 'Данные и параметры' и выберите этажи.")
            else:
                df_filtered = df_incl[df_incl['Этаж'].isin(selected_floors)]

                if df_filtered.empty:
                    st.warning("Нет данных для выбранных этажей.")
                else:
                    fig1 = go.Figure()
                    colors = px.colors.qualitative.Plotly
                    cycle_labels = st.session_state.get('cycle_labels', {})

                    for i, floor in enumerate(sorted(selected_floors)):
                        floor_df = df_filtered[df_filtered['Этаж'] == floor].sort_values('Цикл')
                        color = colors[i % len(colors)]
                        x_labels = [cycle_labels.get(c, c) for c in floor_df['Цикл']]

                        fig1.add_trace(go.Scatter(
                            x=x_labels,
                            y=floor_df['αx_abs'],
                            mode='lines+markers',
                            name=f'Этаж {floor} αx',
                            line=dict(color=color, width=2),
                            marker=dict(size=8, symbol='circle'),
                            legendgroup=f'floor_{floor}',
                            legendgrouptitle_text=f'Этаж {floor}'
                        ))
                        fig1.add_trace(go.Scatter(
                            x=x_labels,
                            y=floor_df['αy_abs'],
                            mode='lines+markers',
                            name=f'Этаж {floor} αy',
                            line=dict(color=color, width=2, dash='dot'),
                            marker=dict(size=8, symbol='square'),
                            legendgroup=f'floor_{floor}',
                            showlegend=False
                        ))

                    if not df_filtered.empty:
                        max_x = df_filtered.loc[df_filtered['αx_abs'].idxmax()]
                        max_y = df_filtered.loc[df_filtered['αy_abs'].idxmax()]
                        fig1.add_annotation(
                            x=cycle_labels.get(max_x['Цикл'], max_x['Цикл']),
                            y=max_x['αx_abs'],
                            text=f"max αx = {max_x['αx_abs']:.3f}° (эт.{max_x['Этаж']})",
                            showarrow=True, arrowhead=2, ax=0, ay=-40,
                            font=dict(color='darkblue', size=11)
                        )
                        fig1.add_annotation(
                            x=cycle_labels.get(max_y['Цикл'], max_y['Цикл']),
                            y=max_y['αy_abs'],
                            text=f"max αy = {max_y['αy_abs']:.3f}° (эт.{max_y['Этаж']})",
                            showarrow=True, arrowhead=2, ax=0, ay=40,
                            font=dict(color='darkred', size=11)
                        )

                    fig1.update_layout(
                        title="Абсолютные углы наклона по осям X и Y",
                        xaxis_title="Цикл",
                        yaxis_title="Угол, °",
                        template="plotly_white",
                        hovermode="x unified",
                        legend=dict(
                            orientation="h",
                            yanchor="bottom",
                            y=1.02,
                            xanchor="right",
                            x=1,
                            title="",
                            font=dict(size=10)
                        ),
                        margin=dict(l=40, r=40, t=80, b=40)
                    )
                    st.plotly_chart(fig1, use_container_width=True)

                    # ---- НОВЫЙ ГРАФИК: Абсолютный угол наклона всего здания ----
                    st.subheader("📊 Абсолютный угол наклона всего здания (по выбранному этажу)")
                    st.markdown("""
                    **ℹ️ Пояснение:** Этот график показывает изменение угла наклона всего здания по датам.  
                    Выберите этаж, который лучше всего отражает крен здания (обычно самый нижний).  
                    - **Синяя линия** – угол по оси X.  
                    - **Красная линия** – угол по оси Y.
                    """)
                    
                    available_floors = sorted(df_incl['Этаж'].unique())
                    default_floor = available_floors[0]  # минимальный этаж
                    selected_floor_for_kren = st.selectbox(
                        "Выберите этаж, представляющий крен здания",
                        available_floors,
                        index=0
                    )
                    
                    kren_df = df_incl[df_incl['Этаж'] == selected_floor_for_kren].sort_values('Цикл')
                    if not kren_df.empty:
                        fig4 = go.Figure()
                        x_labels = [cycle_labels.get(c, c) for c in kren_df['Цикл']]
                        fig4.add_trace(go.Scatter(
                            x=x_labels,
                            y=kren_df['αx_abs'],
                            mode='lines+markers',
                            name=f'αx (эт.{selected_floor_for_kren})',
                            line=dict(color='blue', width=2),
                            marker=dict(size=8)
                        ))
                        fig4.add_trace(go.Scatter(
                            x=x_labels,
                            y=kren_df['αy_abs'],
                            mode='lines+markers',
                            name=f'αy (эт.{selected_floor_for_kren})',
                            line=dict(color='red', width=2, dash='dot'),
                            marker=dict(size=8, symbol='square')
                        ))
                        fig4.update_layout(
                            title=f"Изменение угла наклона здания по этажу {selected_floor_for_kren}",
                            xaxis_title="Цикл",
                            yaxis_title="Угол, °",
                            template="plotly_white",
                            hovermode="x unified",
                            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                        )
                        st.plotly_chart(fig4, use_container_width=True)
                        # Таблица
                        st.dataframe(
                            kren_df[['Цикл', 'αx_abs', 'αy_abs']]
                            .rename(columns={'αx_abs': 'αx, °', 'αy_abs': 'αy, °'})
                            .assign(Цикл=lambda d: d['Цикл'].map(cycle_labels))
                        )
                    else:
                        st.warning("Нет данных для выбранного этажа.")

                    # ---- Сравнение с осадками (если есть) ----
                    df_sett_angles = st.session_state.get('df_sett_angles', None)
                    if df_sett_angles is not None and not df_sett_angles.empty:
                        st.subheader("📊 Сравнение углов по осадкам и наклономеру (этаж 5)")
                        st.markdown("""
                        **ℹ️ Пояснение:** Сравнение углов наклона, полученных по данным наклономера (этаж 5) и по расчёту из осадок (углы a и b).  
                        - **Сплошные линии** – наклономер.  
                        - **Пунктирные линии** – осадки.  
                        - Совпадение линий говорит о хорошей сходимости методов.
                        """)
                        floor_for_compare = 5
                        if floor_for_compare not in df_incl['Этаж'].unique():
                            floor_for_compare = df_incl['Этаж'].min()
                        incl_compare = df_incl[df_incl['Этаж'] == floor_for_compare].copy()

                        merged = pd.merge(incl_compare, df_sett_angles, on='Цикл', how='inner')
                        if merged.empty:
                            st.warning(f"Нет общих циклов для сравнения с осадками (этаж {floor_for_compare}). Показываем последние значения.")
                            last_incl = incl_compare.iloc[-1] if not incl_compare.empty else None
                            last_sett = df_sett_angles.iloc[-1] if not df_sett_angles.empty else None
                            if last_incl is not None and last_sett is not None:
                                st.write("**Последний цикл наклономера (этаж {})**:".format(floor_for_compare))
                                st.dataframe(last_incl[['Цикл', 'αx_abs', 'αy_abs']].to_frame().T)
                                st.write("**Последний цикл осадок**:")
                                st.dataframe(last_sett[['Цикл', 'a_град', 'b_град']].to_frame().T)
                        else:
                            merged['Цикл_метка'] = merged['Цикл'].map(cycle_labels)
                            fig2 = go.Figure()
                            fig2.add_trace(go.Scatter(
                                x=merged['Цикл_метка'],
                                y=merged['αx_abs'],
                                mode='lines+markers',
                                name='αx (наклономер)',
                                line=dict(color='#1f77b4', width=2),
                                marker=dict(size=8)
                            ))
                            fig2.add_trace(go.Scatter(
                                x=merged['Цикл_метка'],
                                y=merged['a_град'],
                                mode='lines+markers',
                                name='a (осадки)',
                                line=dict(color='#ff7f0e', width=2, dash='dash'),
                                marker=dict(size=8, symbol='diamond')
                            ))
                            fig2.add_trace(go.Scatter(
                                x=merged['Цикл_метка'],
                                y=merged['αy_abs'],
                                mode='lines+markers',
                                name='αy (наклономер)',
                                line=dict(color='#2ca02c', width=2),
                                marker=dict(size=8)
                            ))
                            fig2.add_trace(go.Scatter(
                                x=merged['Цикл_метка'],
                                y=merged['b_град'],
                                mode='lines+markers',
                                name='b (осадки)',
                                line=dict(color='#d62728', width=2, dash='dash'),
                                marker=dict(size=8, symbol='diamond')
                            ))

                            fig2.update_layout(
                                title=f"Сравнение углов (этаж {floor_for_compare})",
                                xaxis_title="Цикл",
                                yaxis_title="Угол, °",
                                template="plotly_white",
                                hovermode="x unified",
                                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                            )
                            st.plotly_chart(fig2, use_container_width=True)

        with tab3:
            st.subheader("📥 Скачать отчёт")
            col1, col2, col3 = st.columns(3)
            with col1:
                excel_data = generate_excel_report(df_incl, st.session_state.get('df_sett_angles', None), cycles, alpha0_x, alpha0_y, L)
                st.download_button(
                    label="📊 Excel",
                    data=excel_data,
                    file_name=f"inclinometer_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with col2:
                pdf_data = generate_pdf_report(df_incl, st.session_state.get('df_sett_angles', None), cycles, alpha0_x, alpha0_y, L)
                st.download_button(
                    label="📄 PDF",
                    data=pdf_data.getvalue(),
                    file_name=f"inclinometer_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                    mime="application/pdf"
                )
            with col3:
                word_data = generate_word_report(df_incl, st.session_state.get('df_sett_angles', None), cycles, alpha0_x, alpha0_y, L)
                st.download_button(
                    label="📝 Word",
                    data=word_data.getvalue(),
                    file_name=f"inclinometer_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        with tab4:
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
