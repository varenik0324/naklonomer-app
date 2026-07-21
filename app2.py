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
# ПАРСИНГ ДАННЫХ НАКЛОНОМЕРА (адаптирован под структуру вашего файла)
# ------------------------------------------------------------
def parse_inclinometer_data(file_bytes):
    """
    Парсит лист 'Наклономер' из вашего Excel-файла.
    Ищет строки с этажами (5, 15, 27) в первом столбце,
    извлекает пары углов (αx, αy) из всех числовых значений в строке,
    сопоставляет с заголовками циклов (датами).
    """
    xl = pd.ExcelFile(file_bytes)
    # Ищем лист с названием, содержащим "наклономер" или "крен"
    sheet_name = None
    for name in xl.sheet_names:
        if 'наклономер' in name.lower() or 'крен' in name.lower():
            sheet_name = name
            break
    if sheet_name is None:
        st.error("Не найден лист с данными наклономера. Проверьте файл.")
        return None

    df_raw = pd.read_excel(file_bytes, sheet_name=sheet_name, header=None)

    # --- 1. Находим строку с заголовками циклов (она содержит слово "Цикл" и дату) ---
    cycle_header_row = None
    for idx, row in df_raw.iterrows():
        row_str = ' '.join(str(cell) for cell in row if pd.notna(cell))
        if 'Цикл' in row_str and re.search(r'\d{2}\.\d{2}\.\d{4}', row_str):
            cycle_header_row = idx
            break

    if cycle_header_row is None:
        st.warning("Не найдена строка с заголовками циклов. Будем использовать порядковые номера.")
        cycle_labels = [f"Цикл {i+1}" for i in range(len(df_raw.columns)-1)]
    else:
        # Извлекаем даты из заголовков
        cycle_labels = []
        for cell in df_raw.iloc[cycle_header_row, 1:]:
            if pd.notna(cell):
                cell_str = str(cell)
                # Ищем дату в формате ДД.ММ.ГГГГ
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
        # Обрезаем до фактического количества циклов (убираем пустые хвосты)
        cycle_labels = [c for c in cycle_labels if c]  # убираем пустые строки

    # --- 2. Ищем строки с этажами (в первом столбце числа 5, 15, 27) ---
    floor_rows = {}
    for idx, row in df_raw.iterrows():
        if idx <= cycle_header_row:
            continue  # пропускаем строки до заголовка
        first_cell = row[0]
        if pd.notna(first_cell) and isinstance(first_cell, (int, float)) and first_cell in [5, 15, 27]:
            floor_rows[first_cell] = idx
        if len(floor_rows) == 3:
            break

    if len(floor_rows) < 3:
        st.error("Не найдены все строки с этажами (5, 15, 27). Проверьте файл.")
        return None

    # --- 3. Собираем данные ---
    data = []
    for floor in sorted(floor_rows.keys()):
        row_idx = floor_rows[floor]
        # Извлекаем все числовые значения из строки, начиная с колонки 1 (индекс 1)
        values = []
        for cell in df_raw.iloc[row_idx, 1:]:
            if pd.notna(cell) and isinstance(cell, (int, float)):
                values.append(float(cell))
        # Разбиваем на пары (αx, αy)
        pairs = []
        if len(values) % 2 == 0:
            pairs = [(values[i], values[i+1]) for i in range(0, len(values), 2)]
        else:
            # Если нечётное количество, отбрасываем последний
            pairs = [(values[i], values[i+1]) for i in range(0, len(values)-1, 2)]

        # Сопоставляем пары с циклами
        for i, (ax, ay) in enumerate(pairs):
            if i < len(cycle_labels):
                data.append({
                    'Цикл': cycle_labels[i],
                    'Этаж': floor,
                    'αx': ax,
                    'αy': ay
                })
            else:
                # Если пар больше, чем заголовков, добавляем с порядковым номером
                data.append({
                    'Цикл': f"Цикл {i+1}",
                    'Этаж': floor,
                    'αx': ax,
                    'αy': ay
                })

    if not data:
        st.error("Не удалось извлечь данные наклономера. Проверьте структуру листа.")
        return None

    df = pd.DataFrame(data)
    df['Цикл'] = df['Цикл'].astype(str)
    df['Этаж'] = df['Этаж'].astype(int)
    df['αx'] = pd.to_numeric(df['αx'], errors='coerce')
    df['αy'] = pd.to_numeric(df['αy'], errors='coerce')
    # Удаляем строки с NaN (если есть)
    df = df.dropna(subset=['αx', 'αy'])
    return df

# ------------------------------------------------------------
# ПАРСИНГ ДАННЫХ ОСАДОК (работает с листами "Стилобат" и "Высотная часть")
# ------------------------------------------------------------
def parse_settlement_data(file_bytes, sheet_name, corner_marks, L, B):
    """
    Парсит лист с осадками, где строки - марки, столбцы - циклы.
    Ожидается, что в строке с заголовками есть слово "Цикл" и дата.
    Для каждого цикла есть три колонки: Отметка, Осадка, Общая осадка.
    Берём колонку "Общая осадка" (или "Осадка").
    Возвращает DataFrame с колонками: Цикл, a_мм_м, b_мм_м (уклоны по осям).
    """
    df_raw = pd.read_excel(file_bytes, sheet_name=sheet_name, header=None)

    # --- Находим строку с заголовками циклов ---
    cycle_header_row = None
    for idx, row in df_raw.iterrows():
        row_str = ' '.join(str(cell) for cell in row if pd.notna(cell))
        if 'Цикл' in row_str:
            cycle_header_row = idx
            break
    if cycle_header_row is None:
        st.error("Не найдена строка с заголовками циклов в листе осадок.")
        return None

    # --- Определяем колонки для каждого цикла ---
    cycle_cols = {}
    for col_idx, cell in df_raw.iloc[cycle_header_row, :].items():
        if pd.notna(cell):
            cell_str = str(cell).strip()
            if 'Цикл' in cell_str:
                # извлекаем дату
                date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', cell_str)
                if date_match:
                    try:
                        date_obj = pd.to_datetime(date_match.group(1), dayfirst=True)
                        cycle_label = date_obj.strftime('%Y-%m-%d')
                    except:
                        cycle_label = cell_str
                else:
                    cycle_label = cell_str
                # ищем следующие колонки с "Осадка" или "Общая осадка"
                # обычно через 1 или 2 колонки
                for offset in [1, 2, 3]:
                    if col_idx + offset < len(df_raw.columns):
                        next_cell = df_raw.iloc[cycle_header_row, col_idx + offset]
                        if pd.notna(next_cell):
                            next_str = str(next_cell).strip()
                            if 'осадк' in next_str.lower():
                                cycle_cols[cycle_label] = col_idx + offset
                                break
                if cycle_label not in cycle_cols:
                    # если не нашли, пробуем вторую колонку в блоке (иногда "Осадка" идёт сразу после "Отметка")
                    if col_idx + 2 < len(df_raw.columns):
                        cycle_cols[cycle_label] = col_idx + 2  # предположим, что это "Общая осадка"
                    else:
                        cycle_cols[cycle_label] = col_idx + 1

    if not cycle_cols:
        st.error("Не найдены колонки с осадками для циклов.")
        return None

    # --- Собираем данные по маркам ---
    mark_rows = []
    for idx, row in df_raw.iterrows():
        if idx <= cycle_header_row:
            continue
        if pd.notna(row[0]) and isinstance(row[0], (int, float)):
            mark_rows.append(idx)

    if not mark_rows:
        st.error("Не найдены строки с марками.")
        return None

    # --- Извлекаем осадки для угловых марок ---
    data = []
    for cycle_label, col_idx in cycle_cols.items():
        for mark_idx in mark_rows:
            mark_num = int(df_raw.iloc[mark_idx, 0])
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

    # --- Вычисляем уклоны для каждого цикла ---
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
# ОСНОВНАЯ ЛОГИКА ПРИЛОЖЕНИЯ
# ------------------------------------------------------------
uploaded_file = st.file_uploader(
    "Загрузите Excel-файл с данными наклономера и/или осадок",
    type=["xlsx", "xls"],
    help="Файл должен содержать лист с наклономером (строки с этажами 5,15,27) и, опционально, листы с осадками ('Стилобат', 'Высотная часть')."
)

if uploaded_file is not None:
    try:
        file_bytes = uploaded_file.read()
        xl = pd.ExcelFile(file_bytes)
        all_sheets = xl.sheet_names

        # ------------------------------------------------------------
        # 1. Парсинг наклономера (улучшенный)
        # ------------------------------------------------------------
        df_incl = parse_inclinometer_data(file_bytes)
        if df_incl is None:
            st.error("Не удалось загрузить данные наклономера. Проверьте структуру листа.")
            st.stop()

        st.success(f"✅ Данные наклономера загружены. Найдено {len(df_incl)} записей.")
        st.dataframe(df_incl, use_container_width=True)

        cycles = sorted(df_incl['Цикл'].unique())
        if len(cycles) == 0:
            st.error("Нет циклов в данных наклономера.")
            st.stop()

        # --- Боковая панель: параметры наклономера ---
        st.sidebar.header("Параметры наклономера")
        zero_cycle = st.sidebar.selectbox("Нулевой цикл (для вычитания начального угла)", cycles, index=0)
        alpha0_x = st.sidebar.number_input("Начальный угол X (αx0), °", value=0.0, step=0.001, format="%.3f")
        alpha0_y = st.sidebar.number_input("Начальный угол Y (αy0), °", value=0.0, step=0.001, format="%.3f")
        L = st.sidebar.number_input("Высота этажа (L), м", value=3.0, step=0.1, format="%.1f")

        # Расчёт абсолютных углов и смещений
        zero_data = df_incl[df_incl['Цикл'] == zero_cycle][['Этаж', 'αx', 'αy']].rename(columns={'αx': 'αx0_inc', 'αy': 'αy0_inc'})
        df_incl = df_incl.merge(zero_data, on='Этаж', how='left')
        df_incl['αx_abs'] = df_incl['αx'] - df_incl['αx0_inc'] + alpha0_x
        df_incl['αy_abs'] = df_incl['αy'] - df_incl['αy0_inc'] + alpha0_y
        df_incl['Смещение X'] = L * np.sin(np.radians(df_incl['αx_abs']))
        df_incl['Смещение Y'] = L * np.sin(np.radians(df_incl['αy_abs']))

        # ------------------------------------------------------------
        # 2. Парсинг осадок (если есть подходящие листы)
        # ------------------------------------------------------------
        df_sett_angles = None
        sett_sheets = [s for s in all_sheets if 'стилобат' in s.lower() or 'высотн' in s.lower() or 'осадк' in s.lower()]
        if sett_sheets:
            st.sidebar.subheader("Данные осадок")
            selected_sett_sheet = st.sidebar.selectbox("Выберите лист с осадками", sett_sheets)
            corner_marks_str = st.sidebar.text_input(
                "Номера угловых марок (через запятую, в порядке: нижний левый, нижний правый, верхний левый, верхний правый)",
                "1,5,9,13"
            )
            try:
                corner_marks = [int(x.strip()) for x in corner_marks_str.split(',') if x.strip()]
                if len(corner_marks) != 4:
                    st.sidebar.warning("Введите ровно 4 номера марок.")
            except:
                corner_marks = []
            L_sett = st.sidebar.number_input("Длина фундамента L, м", value=70.0, step=1.0)
            B_sett = st.sidebar.number_input("Ширина фундамента B, м", value=18.0, step=1.0)

            if st.sidebar.button("Рассчитать углы по осадкам"):
                if len(corner_marks) == 4:
                    df_sett_angles = parse_settlement_data(file_bytes, selected_sett_sheet, corner_marks, L_sett, B_sett)
                    if df_sett_angles is not None:
                        st.success(f"✅ Углы по осадкам рассчитаны для {len(df_sett_angles)} циклов.")
                        st.dataframe(df_sett_angles, use_container_width=True)
                else:
                    st.sidebar.error("Укажите 4 угловые марки.")

        # ------------------------------------------------------------
        # 3. Графики
        # ------------------------------------------------------------
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

        # Сравнение с осадками, если есть
        if df_sett_angles is not None and not df_sett_angles.empty:
            st.subheader("📊 Сравнение углов по осадкам и наклономеру")
            # Сопоставим по датам (берём этаж 5, т.к. на нём обычно установлен наклономер)
            merged = pd.merge(df_incl[df_incl['Этаж'] == 5], df_sett_angles, on='Цикл', how='inner')
            if merged.empty:
                st.warning("Циклы осадок и наклономера не совпадают по датам. Сравнение по порядку циклов.")
                # покажем последние
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

        # Профиль смещений (последний цикл)
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

        # ------------------------------------------------------------
        # 4. Скачивание отчётов
        # ------------------------------------------------------------
        st.subheader("📥 Скачать отчёт")
        col1, col2, col3 = st.columns(3)
        with col1:
            excel_data = generate_excel_report(df_incl, df_sett_angles, cycles, alpha0_x, alpha0_y, L)
            st.download_button(
                label="📊 Excel",
                data=excel_data,
                file_name=f"inclinometer_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col2:
            pdf_data = generate_pdf_report(df_incl, df_sett_angles, cycles, alpha0_x, alpha0_y, L)
            st.download_button(
                label="📄 PDF",
                data=pdf_data.getvalue(),
                file_name=f"inclinometer_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                mime="application/pdf"
            )
        with col3:
            word_data = generate_word_report(df_incl, df_sett_angles, cycles, alpha0_x, alpha0_y, L)
            st.download_button(
                label="📝 Word",
                data=word_data.getvalue(),
                file_name=f"inclinometer_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    except Exception as e:
        st.error(f"Ошибка обработки: {e}")
        import traceback
        st.code(traceback.format_exc())

else:
    st.info("👆 Загрузите Excel-файл для начала работы.")
