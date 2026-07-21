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
# ПАРСИНГ ДАННЫХ НАКЛОНОМЕРА (через pandas, с ручным выбором листа)
# ------------------------------------------------------------
def parse_inclinometer_data(file_bytes, sheet_name, manual_floor_rows=None, search_start=None, search_end=None):
    """
    Читает указанный лист через pandas (header=None) и извлекает данные наклономера.
    Если manual_floor_rows задан (словарь {этаж: номер_строки}), использует его.
    Иначе ищет автоматически в диапазоне [search_start, search_end) или до появления 'Таблица'.
    Возвращает DataFrame с колонками: Цикл, Этаж, αx, αy.
    """
    try:
        df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)
    except Exception as e:
        st.error(f"Не удалось прочитать лист '{sheet_name}': {e}")
        return None

    total_rows = len(df_raw)
    total_cols = len(df_raw.columns)

    # 1. Находим строку с заголовками циклов
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
        for col in range(1, total_cols):  # с колонки 1 (второй столбец)
            cell = df_raw.iloc[cycle_header_row, col]
            if pd.notna(cell):
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
            # Ищем до появления "Таблица" или до конца
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
        # Определяем этаж из ячейки с номером
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
def parse_settlement_data(file_bytes, sheet_name, corner_marks, L, B):
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
    for idx, row in df_raw.iterrows():
        if idx <= cycle_header_row:
            continue
        if pd.notna(row[0]) and isinstance(row[0], (int, float)):
            mark_rows.append(idx)

    if not mark_rows:
        st.error("Не найдены строки с марками.")
        return None

    data = []
    for cycle_label, col_idx in cycle_cols.items():
        for mark_idx in mark_rows:
            mark_num = df_raw.iloc[mark_idx, 0]
            if isinstance(mark_num, (int, float)):
                mark_num = int(mark_num)
            else:
                mark_num = str(mark_num).strip()
            if mark_num not in corner_marks:
                continue
            settlement = df_raw.iloc[mark_idx, col_idx]
            if pd.notna(settlement) and isinstance(settlement, (int, float)):
                data.append({
                    'Цикл': cycle_label,
                    'Марка': mark_num,
                    'Осадка_мм': settlement
                })

    if not data:
        st.error("Не удалось извлечь осадки для выбранных марок.")
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
    c.drawString(50, height - 150, "Профиль смещений (последний цикл наклономера):")
    y = height - 170
    if df_incl is not None and not df_incl.empty:
        last_cycle = df_incl['Цикл'].iloc[-1]
        profile = df_incl[df_incl['Цикл'] == last_cycle]
        for _, row in profile.iterrows():
            c.setFont("Helvetica", 10)
            c.drawString(60, y, f"Этаж {row['Этаж']}: смещ. X = {row['Смещение X']:.3f} м, Y = {row['Смещение Y']:.3f} м")
            y -= 20
            if y < 50:
                c.showPage()
                y = height - 50
    if df_sett_angles is not None and not df_sett_angles.empty:
        c.showPage()
        c.setFont("Helvetica-Bold", 12)
        c.drawString(50, height - 50, "Углы наклона по осадкам (последний цикл):")
        y = height - 70
        last_sett = df_sett_angles['Цикл'].iloc[-1]
        row = df_sett_angles[df_sett_angles['Цикл'] == last_sett].iloc[0]
        c.setFont("Helvetica", 10)
        c.drawString(60, y, f"Цикл {last_sett}: a = {row['a_град']:.3f}°, b = {row['b_град']:.3f}°")
    c.save()
    buffer.seek(0)
    return buffer

def generate_word_report(df_incl, df_sett_angles, cycles, alpha0_x, alpha0_y, L):
    doc = Document()
    doc.add_heading('Отчёт по наклономеру и осадкам', level=1)
    doc.add_paragraph(f"Дата: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
    doc.add_paragraph(f"Параметры: αx0 = {alpha0_x:.3f}°, αy0 = {alpha0_y:.3f}°, L = {L} м")
    doc.add_paragraph(f"Нулевой цикл: {cycles[0] if cycles else ''}")
    doc.add_heading('Профиль смещений (последний цикл наклономера)', level=2)
    if df_incl is not None and not df_incl.empty:
        last_cycle = df_incl['Цикл'].iloc[-1]
        profile = df_incl[df_incl['Цикл'] == last_cycle]
        for _, row in profile.iterrows():
            doc.add_paragraph(f"Этаж {row['Этаж']}: смещ. X = {row['Смещение X']:.3f} м, Y = {row['Смещение Y']:.3f} м")
    if df_sett_angles is not None and not df_sett_angles.empty:
        doc.add_heading('Углы наклона по осадкам', level=2)
        for _, row in df_sett_angles.iterrows():
            doc.add_paragraph(f"Цикл {row['Цикл']}: a = {row['a_град']:.3f}°, b = {row['b_град']:.3f}°")
    doc.add_paragraph("© Геофундамент, 2026").alignment = WD_ALIGN_PARAGRAPH.CENTER
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ------------------------------------------------------------
# ОСНОВНАЯ ЛОГИКА ПРИЛОЖЕНИЯ (с вкладками и новой вкладкой "Наклономер")
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

        # --- Боковая панель: выбор листов и поиск ---
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

        # ------------------------------------------------------------
        # 1. Парсинг наклономера
        # ------------------------------------------------------------
        df_incl = parse_inclinometer_data(file_bytes, incl_sheet_name, search_start=search_start, search_end=search_end)

        if df_incl is None:
            # Ручной ввод строк
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

        # --- Основная область с вкладками (теперь их 4) ---
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

            # Расчёт абсолютных углов и смещений
            zero_data = df_incl[df_incl['Цикл'] == zero_cycle][['Этаж', 'αx', 'αy']].rename(columns={'αx': 'αx0_inc', 'αy': 'αy0_inc'})
            df_incl = df_incl.merge(zero_data, on='Этаж', how='left')
            df_incl['αx_abs'] = df_incl['αx'] - df_incl['αx0_inc'] + alpha0_x
            df_incl['αy_abs'] = df_incl['αy'] - df_incl['αy0_inc'] + alpha0_y
            df_incl['Смещение X'] = L * np.sin(np.radians(df_incl['αx_abs']))
            df_incl['Смещение Y'] = L * np.sin(np.radians(df_incl['αy_abs']))

            # Блок осадок (если есть)
            sett_sheets = [s for s in all_sheets if 'стилобат' in s.lower() or 'высотн' in s.lower() or 'осадк' in s.lower()]
            if sett_sheets:
                st.subheader("📐 Данные осадок")
                with st.expander("Настройки осадок", expanded=False):
                    selected_sett_sheet = st.selectbox("Выберите лист с осадками", sett_sheets, key="sett_sheet")
                    corner_marks_str = st.text_input(
                        "Номера угловых марок (через запятую, в порядке: нижний левый, нижний правый, верхний левый, верхний правый)",
                        "1,5,9,13"
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
                            df_sett_angles = parse_settlement_data(file_bytes, selected_sett_sheet, corner_marks, L_sett, B_sett)
                            if df_sett_angles is not None:
                                st.success(f"✅ Углы по осадкам рассчитаны для {len(df_sett_angles)} циклов.")
                                st.dataframe(df_sett_angles, use_container_width=True)
                                st.session_state.df_sett_angles = df_sett_angles
                            else:
                                st.error("Не удалось рассчитать углы.")
                        else:
                            st.error("Укажите 4 угловые марки.")
                if 'df_sett_angles' in st.session_state:
                    st.dataframe(st.session_state.df_sett_angles, use_container_width=True)
            else:
                st.info("Листов с осадками не найдено.")
                df_sett_angles = None

        with tab2:
            st.subheader("📈 Изменение абсолютных углов наклона по циклам (наклономер)")
            fig = go.Figure()
            for floor in sorted(df_incl['Этаж'].unique()):
                floor_df = df_incl[df_incl['Этаж'] == floor].sort_values('Цикл')
                fig.add_trace(go.Scatter(
                    x=floor_df['Цикл'],
                    y=floor_df['αx_abs'],
                    mode='lines+markers',
                    name=f'Этаж {floor} αx_abs'
                ))
                fig.add_trace(go.Scatter(
                    x=floor_df['Цикл'],
                    y=floor_df['αy_abs'],
                    mode='lines+markers',
                    name=f'Этаж {floor} αy_abs',
                    line=dict(dash='dot')
                ))
            fig.update_layout(
                xaxis_title="Цикл",
                yaxis_title="Абсолютный угол, °",
                template="plotly_white",
                hovermode="x unified"
            )
            st.plotly_chart(fig, use_container_width=True)

            # Сравнение с осадками
            df_sett_angles = st.session_state.get('df_sett_angles', None)
            if df_sett_angles is not None and not df_sett_angles.empty:
                st.subheader("📊 Сравнение углов по осадкам и наклономеру")
                merged = pd.merge(df_incl[df_incl['Этаж'] == 5], df_sett_angles, on='Цикл', how='inner')
                if merged.empty:
                    st.warning("Циклы осадок и наклономера не совпадают по датам. Сравнение по порядку циклов.")
                    last_incl = df_incl[df_incl['Цикл'] == df_incl['Цикл'].iloc[-1]]
                    last_sett = df_sett_angles[df_sett_angles['Цикл'] == df_sett_angles['Цикл'].iloc[-1]]
                    if not last_incl.empty and not last_sett.empty:
                        st.write("Последний цикл наклономера:")
                        st.dataframe(last_incl[['Этаж', 'αx_abs', 'αy_abs']])
                        st.write("Последний цикл осадок:")
                        st.dataframe(last_sett[['Цикл', 'a_град', 'b_град']])
                else:
                    fig2 = go.Figure()
                    fig2.add_trace(go.Scatter(
                        x=merged['Цикл'],
                        y=merged['αx_abs'],
                        mode='lines+markers',
                        name='αx (наклономер)'
                    ))
                    fig2.add_trace(go.Scatter(
                        x=merged['Цикл'],
                        y=merged['a_град'],
                        mode='lines+markers',
                        name='a (осадки)',
                        line=dict(dash='dash')
                    ))
                    fig2.add_trace(go.Scatter(
                        x=merged['Цикл'],
                        y=merged['αy_abs'],
                        mode='lines+markers',
                        name='αy (наклономер)'
                    ))
                    fig2.add_trace(go.Scatter(
                        x=merged['Цикл'],
                        y=merged['b_град'],
                        mode='lines+markers',
                        name='b (осадки)',
                        line=dict(dash='dash')
                    ))
                    fig2.update_layout(
                        xaxis_title="Цикл",
                        yaxis_title="Угол, °",
                        template="plotly_white",
                        hovermode="x unified"
                    )
                    st.plotly_chart(fig2, use_container_width=True)

            # Профиль смещений
            last_cycle = df_incl['Цикл'].iloc[-1]
            profile = df_incl[df_incl['Цикл'] == last_cycle].sort_values('Этаж')
            if not profile.empty:
                st.subheader(f"📊 Профиль смещений на цикле {last_cycle}")
                fig3 = go.Figure()
                fig3.add_trace(go.Scatter(
                    x=profile['Смещение X'],
                    y=profile['Этаж'],
                    mode='lines+markers',
                    name='Смещение X',
                    marker=dict(size=10)
                ))
                fig3.add_trace(go.Scatter(
                    x=profile['Смещение Y'],
                    y=profile['Этаж'],
                    mode='lines+markers',
                    name='Смещение Y',
                    marker=dict(size=10, symbol='square')
                ))
                fig3.update_layout(
                    xaxis_title="Смещение, м",
                    yaxis_title="Этаж",
                    template="plotly_white",
                    yaxis=dict(autorange="reversed"),
                    hovermode="y unified"
                )
                st.plotly_chart(fig3, use_container_width=True)

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

        # ------------------------------------------------------------
        # НОВАЯ ВКЛАДКА "НАКЛОНОМЕР" (информация об устройстве)
        # ------------------------------------------------------------
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

            st.subheader("📏 Построение профилей")
            st.latex(r"M_{yi} = L \cdot \sin \alpha_{yi}, \quad M_y = \sum_{1}^{n} M_{yi}")
            st.markdown("где M_y – суммарное смещение вдоль оси Y, L – длина звена, α_yi – угол наклона i-го звена.")

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
