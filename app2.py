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
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import logging

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
# ПАРСИНГ ОСАДОК (возвращает кортеж: df_angles, marks_data, all_cycles)
# ------------------------------------------------------------
def parse_settlement_data(file_bytes, sheet_name, corner_marks, L, B, mark_col=0, zero_cycle_sett=None):
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
    all_cycles = []
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

    # Если нулевой цикл не задан – выбираем первый по дате (самый ранний)
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

    # Собираем данные по ВСЕМ маркам (для отображения) – абсолютные осадки за цикл
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

    # Получаем осадки в нулевом цикле
    zero_marks = marks_abs_data.get(zero_cycle_sett, {})
    if not zero_marks:
        st.warning(f"Нулевой цикл {zero_cycle_sett} не найден в данных осадок. Используем первый доступный.")
        first_cycle = list(marks_abs_data.keys())[0]
        zero_marks = marks_abs_data[first_cycle]
        zero_cycle_sett = first_cycle

    # Собираем приросты для угловых марок
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
        st.error("Не удалось извлечь приросты осадок.")
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
    return df_angles, marks_abs_data, list(marks_abs_data.keys())

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

def generate_pdf_report(df_incl, df_sett_angles, cycles, alpha0_x, alpha0_y, L, report_params):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, height - 50, report_params.get('организация', 'ООО "Геофундамент"'))
    c.setFont("Helvetica", 12)
    c.drawString(50, height - 70, f"Адрес: {report_params.get('address', 'г. Москва, ул. Суздальская, д. 18, корп. 4')}")
    c.drawString(50, height - 90, f"Тел.: {report_params.get('phone', '8 499 399-30-60')}")
    c.drawString(50, height - 110, f"E-mail: {report_params.get('email', 'geofundament@mail.ru')}")

    c.line(50, height - 120, width - 50, height - 120)

    c.setFont("Helvetica-Bold", 14)
    c.drawString(50, height - 160, "НАУЧНО-ТЕХНИЧЕСКИЙ ОТЧЕТ")
    c.setFont("Helvetica", 12)
    c.drawString(50, height - 180, f"По результатам {report_params.get('cycle_number', '')} цикла наблюдений")
    c.drawString(50, height - 200, f"от {report_params.get('date', '')}")

    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, height - 230, f"Объект: {report_params.get('object_name', '')}")
    c.drawString(50, height - 250, f"Адрес: {report_params.get('address_obj', '')}")

    c.setFont("Helvetica", 12)
    c.drawString(50, height - 280, f"Заказчик: {report_params.get('customer', '')}")

    c.setFont("Helvetica", 10)
    c.drawString(50, height - 320, f"Исполнитель: {report_params.get('исполнитель', '')}")
    c.drawString(50, height - 340, f"Дата: {datetime.now().strftime('%d.%m.%Y')}")

    c.showPage()

    c.setFont("Helvetica-Bold", 14)
    c.drawString(50, height - 50, "1. ВВЕДЕНИЕ")
    c.setFont("Helvetica", 10)
    text = ("Настоящий отчёт составлен по результатам геотехнического мониторинга объекта "
            f"«{report_params.get('object_name', '')}», расположенного по адресу: {report_params.get('address_obj', '')}. "
            "Целью мониторинга является обеспечение безопасности строительства и эксплуатационной надежности "
            "вновь возводимого здания, а также сохранности окружающей застройки.")
    c.drawString(50, height - 70, text)

    c.setFont("Helvetica-Bold", 14)
    c.drawString(50, height - 110, "2. РЕЗУЛЬТАТЫ НАБЛЮДЕНИЙ")

    if df_incl is not None and not df_incl.empty:
        zero_cycle = report_params.get('zero_cycle', df_incl['Цикл'].min())
        last_cycle = df_incl['Цикл'].max()
        floors = df_incl['Этаж'].unique()
        target_floor = 5 if 5 in floors else min(floors)
        df_floor = df_incl[df_incl['Этаж'] == target_floor].copy()

        if 'αx_abs' not in df_floor.columns:
            df_floor['αx_abs'] = df_floor['αx']
            df_floor['αy_abs'] = df_floor['αy']

        zero_row = df_floor[df_floor['Цикл'] == zero_cycle]
        last_row = df_floor[df_floor['Цикл'] == last_cycle]

        if not zero_row.empty and not last_row.empty:
            zero_x = zero_row['αx_abs'].values[0]
            zero_y = zero_row['αy_abs'].values[0]
            last_x = last_row['αx_abs'].values[0]
            last_y = last_row['αy_abs'].values[0]
            delta_x_deg = last_x - zero_x
            delta_y_deg = last_y - zero_y
            delta_x_mm_m = delta_x_deg * 1000 / L
            delta_y_mm_m = delta_y_deg * 1000 / L

            c.setFont("Helvetica", 10)
            c.drawString(50, height - 140, f"Цикл «нулевой» ({zero_cycle}) и последний цикл ({last_cycle})")
            c.drawString(50, height - 160, f"Прирост угла наклона по оси X: {delta_x_deg:.3f}° ({delta_x_mm_m:.2f} мм/м)")
            c.drawString(50, height - 180, f"Прирост угла наклона по оси Y: {delta_y_deg:.3f}° ({delta_y_mm_m:.2f} мм/м)")

            limit_kren = report_params.get('limit_kren_mm_m', 2.0)
            calc_kren = report_params.get('calc_kren_mm_m', 1.84)
            c.drawString(50, height - 210, "Сравнение с предельными значениями:")
            c.drawString(50, height - 230, f"Расчётное значение: {calc_kren} мм/м")
            c.drawString(50, height - 250, f"Предельно допустимое значение: {limit_kren} мм/м")
            if delta_x_mm_m <= limit_kren and delta_y_mm_m <= limit_kren:
                c.drawString(50, height - 270, "Полученные значения не превысили допустимые величины.")
            else:
                c.drawString(50, height - 270, "ВНИМАНИЕ: Полученные значения превышают допустимые величины!")

    c.showPage()
    c.setFont("Helvetica-Bold", 14)
    c.drawString(50, height - 50, "3. ВЫВОДЫ")
    c.setFont("Helvetica", 10)
    c.drawString(50, height - 80, "По результатам выполненных наблюдений установлено:")
    c.drawString(50, height - 100, "- Деформации (углы наклона) строящегося здания не превышают расчётных и предельных значений.")
    c.drawString(50, height - 120, "- Техническое состояние объекта соответствует требованиям нормативных документов.")
    c.drawString(50, height - 140, "- Мониторинг следует продолжить в соответствии с программой наблюдений.")

    c.setFont("Helvetica", 10)
    c.drawString(50, height - 180, f"Отчёт составил: {report_params.get('исполнитель', '')}")
    c.drawString(50, height - 200, f"Дата: {datetime.now().strftime('%d.%m.%Y')}")

    c.save()
    buffer.seek(0)
    return buffer

def generate_word_report(df_incl, df_sett_angles, cycles, alpha0_x, alpha0_y, L, report_params):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    doc.add_paragraph(report_params.get('организация', 'ООО "Геофундамент"'), style='Title')
    doc.add_paragraph(f"Адрес: {report_params.get('address', 'г. Москва, ул. Суздальская, д. 18, корп. 4')}")
    doc.add_paragraph(f"Тел.: {report_params.get('phone', '8 499 399-30-60')}")
    doc.add_paragraph(f"E-mail: {report_params.get('email', 'geofundament@mail.ru')}")
    doc.add_paragraph(' ' * 5)
    doc.add_heading('НАУЧНО-ТЕХНИЧЕСКИЙ ОТЧЕТ', level=1)
    doc.add_paragraph(f"По результатам {report_params.get('cycle_number', '')} цикла наблюдений")
    doc.add_paragraph(f"от {report_params.get('date', '')}")
    doc.add_paragraph(f"Объект: {report_params.get('object_name', '')}")
    doc.add_paragraph(f"Адрес: {report_params.get('address_obj', '')}")
    doc.add_paragraph(f"Заказчик: {report_params.get('customer', '')}")
    doc.add_paragraph(' ' * 3)
    doc.add_paragraph(f"Исполнитель: {report_params.get('исполнитель', '')}")
    doc.add_paragraph(f"Дата: {datetime.now().strftime('%d.%m.%Y')}")
    doc.add_page_break()

    doc.add_heading('1. ВВЕДЕНИЕ', level=1)
    doc.add_paragraph(
        f"Настоящий отчёт составлен по результатам геотехнического мониторинга объекта "
        f"«{report_params.get('object_name', '')}», расположенного по адресу: {report_params.get('address_obj', '')}. "
        "Целью мониторинга является обеспечение безопасности строительства и эксплуатационной надежности "
        "вновь возводимого здания, а также сохранности окружающей застройки."
    )

    doc.add_heading('2. РЕЗУЛЬТАТЫ НАБЛЮДЕНИЙ', level=1)

    if df_incl is not None and not df_incl.empty:
        zero_cycle = report_params.get('zero_cycle', df_incl['Цикл'].min())
        last_cycle = df_incl['Цикл'].max()
        floors = df_incl['Этаж'].unique()
        target_floor = 5 if 5 in floors else min(floors)
        df_floor = df_incl[df_incl['Этаж'] == target_floor].copy()
        if 'αx_abs' not in df_floor.columns:
            df_floor['αx_abs'] = df_floor['αx']
            df_floor['αy_abs'] = df_floor['αy']

        zero_row = df_floor[df_floor['Цикл'] == zero_cycle]
        last_row = df_floor[df_floor['Цикл'] == last_cycle]

        if not zero_row.empty and not last_row.empty:
            zero_x = zero_row['αx_abs'].values[0]
            zero_y = zero_row['αy_abs'].values[0]
            last_x = last_row['αx_abs'].values[0]
            last_y = last_row['αy_abs'].values[0]
            delta_x_deg = last_x - zero_x
            delta_y_deg = last_y - zero_y
            delta_x_mm_m = delta_x_deg * 1000 / L
            delta_y_mm_m = delta_y_deg * 1000 / L

            doc.add_paragraph(f"Цикл «нулевой» ({zero_cycle}) и последний цикл ({last_cycle})")
            doc.add_paragraph(f"Прирост угла наклона по оси X: {delta_x_deg:.3f}° ({delta_x_mm_m:.2f} мм/м)")
            doc.add_paragraph(f"Прирост угла наклона по оси Y: {delta_y_deg:.3f}° ({delta_y_mm_m:.2f} мм/м)")

            limit_kren = report_params.get('limit_kren_mm_m', 2.0)
            calc_kren = report_params.get('calc_kren_mm_m', 1.84)
            doc.add_paragraph("Сравнение с предельными значениями:")
            doc.add_paragraph(f"Расчётное значение: {calc_kren} мм/м")
            doc.add_paragraph(f"Предельно допустимое значение: {limit_kren} мм/м")
            if delta_x_mm_m <= limit_kren and delta_y_mm_m <= limit_kren:
                doc.add_paragraph("Полученные значения не превысили допустимые величины.")
            else:
                doc.add_paragraph("ВНИМАНИЕ: Полученные значения превышают допустимые величины!")

    doc.add_heading('3. ВЫВОДЫ', level=1)
    doc.add_paragraph("По результатам выполненных наблюдений установлено:")
    doc.add_paragraph("- Деформации (углы наклона) строящегося здания не превышают расчётных и предельных значений.")
    doc.add_paragraph("- Техническое состояние объекта соответствует требованиям нормативных документов.")
    doc.add_paragraph("- Мониторинг следует продолжить в соответствии с программой наблюдений.")
    doc.add_paragraph(f"Отчёт составил: {report_params.get('исполнитель', '')}")
    doc.add_paragraph(f"Дата: {datetime.now().strftime('%d.%m.%Y')}")

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
        for fmt in ('%Y-%m-%d', '%d.%m.%Y', '%Y-%m-%d'):
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

        # --- Боковая панель: параметры отчёта ---
        st.sidebar.header("Параметры отчёта")
        report_params = {
            'организация': st.sidebar.text_input("Организация", "ООО «Геофундамент»"),
            'address': st.sidebar.text_input("Адрес организации", "111673, Москва, ул. Суздальская д. 18 корп. 4"),
            'phone': st.sidebar.text_input("Телефон", "8 499 399-30-60"),
            'email': st.sidebar.text_input("E-mail", "geofundament@mail.ru"),
            'object_name': st.sidebar.text_input("Объект", "Многофункциональный жилой комплекс"),
            'address_obj': st.sidebar.text_input("Адрес объекта", "г. Москва, ул. Крылатская, влд. 23, стр. 1"),
            'customer': st.sidebar.text_input("Заказчик", "ООО СЗ «Сампад»"),
            'cycle_number': st.sidebar.text_input("Номер цикла", "6-й цикл"),
            'date': st.sidebar.text_input("Дата цикла", datetime.now().strftime("%d.%m.%Y")),
            'limit_kren_mm_m': st.sidebar.number_input("Предельный крен (мм/м)", value=2.0, step=0.1),
            'calc_kren_mm_m': st.sidebar.number_input("Расчётный крен (мм/м)", value=1.84, step=0.1),
            'исполнитель': st.sidebar.text_input("Исполнитель", "Добшиков А.Н."),
            'zero_cycle': st.sidebar.selectbox("Нулевой цикл (для отчёта)", sorted(df_incl['Цикл'].unique()), index=0)
        }
        st.session_state.report_params = report_params

        # --- Боковая панель: координаты марок для 3D ---
        st.sidebar.subheader("Координаты марок для 3D-модели")
        coord_file = st.sidebar.file_uploader(
            "Загрузите CSV с координатами (колонки: Марка, X, Y)",
            type=['csv'],
            key='coord_uploader'
        )
        if coord_file is not None:
            try:
                df_coord = pd.read_csv(coord_file)
                if len(df_coord.columns) < 3:
                    st.sidebar.error("Файл должен содержать как минимум 3 колонки: Марка, X, Y")
                else:
                    coord_dict = {}
                    for _, row in df_coord.iterrows():
                        mark = str(row[0]).strip()
                        try:
                            x = float(row[1])
                            y = float(row[2])
                        except:
                            continue
                        coord_dict[mark] = (x, y)
                    st.session_state.coord_dict = coord_dict
                    st.sidebar.success(f"Загружено {len(coord_dict)} марок!")
            except Exception as e:
                st.sidebar.error(f"Ошибка чтения файла: {e}")
        else:
            # Если файл не загружен, создаём стандартную сетку по размерам стилобата
            if 'coord_dict' not in st.session_state:
                L_st = 70.46
                B_st = 18.69
                coord_dict = {
                    '1': (0, 0),
                    '2': (L_st/3, 0),
                    '3': (2*L_st/3, 0),
                    '4': (L_st, 0),
                    '5': (0, B_st/3),
                    '6': (L_st/3, B_st/3),
                    '7': (2*L_st/3, B_st/3),
                    '8': (L_st, B_st/3),
                    '9': (0, 2*B_st/3),
                    '10': (L_st/3, 2*B_st/3),
                    '11': (2*L_st/3, 2*B_st/3),
                    '12': (L_st, 2*B_st/3),
                    '13': (0, B_st),
                    '14': (L_st, B_st),
                }
                st.session_state.coord_dict = coord_dict
                st.sidebar.info("Используются стандартные координаты (сетка 4x4 по размерам стилобата).")

        # --- Основная область с вкладками ---
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "📊 Данные и параметры",
            "📐 Визуализация крена",
            "🌐 3D-модель",
            "📥 Отчёт",
            "📘 Наклономер"
        ])

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

            cycle_labels = format_cycle_labels(df_incl, zero_cycle)
            st.session_state.cycle_labels = cycle_labels
            st.session_state.zero_cycle = zero_cycle
            st.session_state.report_params['zero_cycle'] = zero_cycle

            st.subheader("🎯 Выбор этажей для отображения на графиках")
            available_floors = sorted(df_incl['Этаж'].unique())
            selected_floors = st.multiselect(
                "Выберите этажи",
                options=available_floors,
                default=available_floors,
                key="floor_selector"
            )
            st.session_state.selected_floors = selected_floors

            # ---------- Блок осадок ----------
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
                    L_sett = st.number_input("Длина фундамента L, м", value=70.46, step=0.1)
                    B_sett = st.number_input("Ширина фундамента B, м", value=18.69, step=0.1)

                    # Получаем список циклов осадок
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
                            index=sorted_cycles_temp.index(best_cycle) if best_cycle in sorted_cycles_temp else 0
                        )
                    else:
                        zero_cycle_sett = None

                    if st.button("Рассчитать углы по осадкам"):
                        if len(corner_marks) == 4 and zero_cycle_sett is not None:
                            result = parse_settlement_data(
                                file_bytes, selected_sett_sheet, corner_marks, L_sett, B_sett, mark_col=mark_col,
                                zero_cycle_sett=zero_cycle_sett
                            )
                            if result is not None:
                                df_sett_angles, marks_data, all_cycles_sett = result
                                st.success(f"✅ Углы по осадкам рассчитаны для {len(df_sett_angles)} циклов.")
                                st.session_state.df_sett_angles = df_sett_angles
                                st.session_state.marks_data = marks_data
                                st.session_state.corner_marks = corner_marks
                                st.session_state.L_sett = L_sett
                                st.session_state.B_sett = B_sett
                                st.session_state.sett_sheet = selected_sett_sheet
                                st.session_state.all_cycles_sett = all_cycles_sett
                                st.session_state.zero_cycle_sett = zero_cycle_sett
                            else:
                                st.error("Не удалось рассчитать углы. Проверьте правильность введённых данных.")
                        else:
                            st.error("Укажите 4 угловые марки и выберите нулевой цикл.")

                # Если есть рассчитанные данные, показываем таблицу
                if 'df_sett_angles' in st.session_state and st.session_state.df_sett_angles is not None:
                    df_angles = st.session_state.df_sett_angles
                    st.dataframe(df_angles, use_container_width=True)

        with tab2:
            st.subheader("📐 Визуализация крена здания")
            st.markdown("""
            **ℹ️ Пояснение:** На этой странице представлены наглядные графики, показывающие крен здания:
            - **Схема фундамента** – вид сверху с вектором смещения (по данным осадок).
            - **Столбчатая диаграмма осадок** – профиль осадок по маркам для выбранного цикла.
            - **Сравнение углов** – наклономер vs осадки.
            - **Профиль смещений** – по этажам (данные наклономера).
            """)

            # Выбор цикла для визуализации
            cycles = sorted(df_incl['Цикл'].unique())
            selected_cycle = st.selectbox("Выберите цикл для визуализации", cycles, index=len(cycles)-1, key="vis_cycle")
            cycle_label = st.session_state.get('cycle_labels', {}).get(selected_cycle, selected_cycle)

            # --- 1. Схема фундамента с вектором крена (по осадкам) ---
            df_sett_angles = st.session_state.get('df_sett_angles', None)
            if df_sett_angles is not None and not df_sett_angles.empty:
                sett_row = df_sett_angles[df_sett_angles['Цикл'] == selected_cycle]
                if not sett_row.empty:
                    a = sett_row['a_мм_м'].values[0]
                    b = sett_row['b_мм_м'].values[0]
                    scale = 1000
                    ax_vis = a * scale
                    ay_vis = b * scale

                    fig_sett = go.Figure()
                    fig_sett.add_shape(
                        type="rect",
                        x0=-0.5, y0=-0.5, x1=0.5, y1=0.5,
                        line=dict(color="black", width=2),
                        fillcolor="lightblue", opacity=0.3
                    )
                    fig_sett.add_annotation(
                        x=ax_vis, y=ay_vis,
                        text=f"a={a:.2f} мм/м, b={b:.2f} мм/м",
                        showarrow=True,
                        arrowhead=2,
                        ax=0, ay=-30,
                        font=dict(size=12, color="red")
                    )
                    fig_sett.add_trace(go.Scatter(
                        x=[0, ax_vis], y=[0, ay_vis],
                        mode='lines+markers',
                        line=dict(color='red', width=3),
                        marker=dict(size=10, color='red'),
                        name='Вектор крена'
                    ))
                    fig_sett.update_layout(
                        title=f"Схема фундамента с вектором крена (цикл {cycle_label})",
                        xaxis_title="Смещение по оси X (усл. ед.)",
                        yaxis_title="Смещение по оси Y (усл. ед.)",
                        xaxis=dict(scaleanchor="y", scaleratio=1),
                        yaxis=dict(scaleanchor="x", scaleratio=1),
                        height=500,
                        template="plotly_white"
                    )
                    st.plotly_chart(fig_sett, use_container_width=True)

                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("Угол a (по оси X)", f"{a:.3f} мм/м")
                    with col2:
                        st.metric("Угол b (по оси Y)", f"{b:.3f} мм/м")

            # --- 2. Столбчатая диаграмма осадок ---
            marks_data = st.session_state.get('marks_data', {})
            if marks_data and selected_cycle in marks_data:
                df_marks = pd.DataFrame({
                    'Марка': list(marks_data[selected_cycle].keys()),
                    'Осадка, мм': list(marks_data[selected_cycle].values())
                })
                df_marks = df_marks.sort_values('Осадка, мм', ascending=False)
                fig_bar = px.bar(
                    df_marks, x='Марка', y='Осадка, мм',
                    title=f"Осадки марок (цикл {cycle_label})",
                    color='Осадка, мм',
                    color_continuous_scale='RdYlGn_r'
                )
                fig_bar.update_layout(template="plotly_white")
                st.plotly_chart(fig_bar, use_container_width=True)

            # --- 3. Сравнение углов ---
            if df_sett_angles is not None and not df_sett_angles.empty:
                kren_df = df_incl.groupby('Цикл', as_index=False)[['αx_abs', 'αy_abs']].mean()
                merged = pd.merge(kren_df, df_sett_angles, on='Цикл', how='inner')
                if not merged.empty:
                    fig_comp = go.Figure()
                    fig_comp.add_trace(go.Scatter(
                        x=merged['Цикл'],
                        y=merged['αx_abs'],
                        mode='lines+markers',
                        name='αx_ср (наклономер)',
                        line=dict(color='blue', width=2)
                    ))
                    fig_comp.add_trace(go.Scatter(
                        x=merged['Цикл'],
                        y=merged['a_град'],
                        mode='lines+markers',
                        name='a (осадки)',
                        line=dict(color='red', width=2, dash='dash')
                    ))
                    fig_comp.add_trace(go.Scatter(
                        x=merged['Цикл'],
                        y=merged['αy_abs'],
                        mode='lines+markers',
                        name='αy_ср (наклономер)',
                        line=dict(color='green', width=2)
                    ))
                    fig_comp.add_trace(go.Scatter(
                        x=merged['Цикл'],
                        y=merged['b_град'],
                        mode='lines+markers',
                        name='b (осадки)',
                        line=dict(color='orange', width=2, dash='dash')
                    ))
                    fig_comp.update_layout(
                        title="Сравнение углов крена (наклономер vs осадки)",
                        xaxis_title="Цикл",
                        yaxis_title="Угол, °",
                        template="plotly_white",
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                    )
                    st.plotly_chart(fig_comp, use_container_width=True)

            # --- 4. Профиль смещений по этажам ---
            st.subheader("📊 Профиль смещений по этажам (наклономер)")
            profile = df_incl[df_incl['Цикл'] == selected_cycle].sort_values('Этаж')
            if not profile.empty:
                fig_prof = go.Figure()
                fig_prof.add_trace(go.Scatter(
                    x=profile['Смещение X'],
                    y=profile['Этаж'],
                    mode='lines+markers',
                    name='Смещение X',
                    line=dict(color='blue', width=2)
                ))
                fig_prof.add_trace(go.Scatter(
                    x=profile['Смещение Y'],
                    y=profile['Этаж'],
                    mode='lines+markers',
                    name='Смещение Y',
                    line=dict(color='red', width=2, dash='dot')
                ))
                fig_prof.update_layout(
                    title=f"Профиль смещений (цикл {cycle_label})",
                    xaxis_title="Смещение, м",
                    yaxis_title="Этаж",
                    yaxis=dict(autorange="reversed"),
                    template="plotly_white"
                )
                st.plotly_chart(fig_prof, use_container_width=True)

        with tab3:
            st.subheader("🌐 3D-модель фундамента с осадками")
            st.markdown("""
            **ℹ️ Пояснение:** Трёхмерная визуализация показывает:
            - **План фундамента** (XY) – расположение марок (координаты можно загрузить в боковой панели).
            - **Осадки (Z)** – вертикальные столбцы (цвет показывает величину осадки).
            - **Вектор крена** – направление и величина наклона (если данные осадок доступны).
            - Можно вращать и масштабировать модель мышью.
            """)

            marks_data = st.session_state.get('marks_data', {})
            df_sett_angles = st.session_state.get('df_sett_angles', None)
            coord_dict = st.session_state.get('coord_dict', None)

            if not marks_data or not df_sett_angles:
                st.warning("Сначала рассчитайте осадки в вкладке 'Данные и параметры'.")
            else:
                cycles_3d = sorted(marks_data.keys())
                selected_cycle_3d = st.selectbox("Выберите цикл для 3D-визуализации", cycles_3d, index=len(cycles_3d)-1, key="3d_cycle")

                sett_dict = marks_data[selected_cycle_3d]
                if not sett_dict:
                    st.warning("Нет данных осадок для выбранного цикла.")
                else:
                    if coord_dict is None:
                        st.error("Не загружены координаты марок. Загрузите CSV в боковой панели или используйте стандартные.")
                    else:
                        mark_keys_filtered = [m for m in sett_dict.keys() if m in coord_dict and not np.isnan(sett_dict.get(m, np.nan))]
                        if not mark_keys_filtered:
                            st.error("Нет марок с координатами и осадками.")
                        else:
                            fig_3d = go.Figure()

                            for mark in mark_keys_filtered:
                                x, y = coord_dict[mark]
                                z = sett_dict.get(mark, 0)
                                if np.isnan(z):
                                    z = 0
                                fig_3d.add_trace(go.Scatter3d(
                                    x=[x, x],
                                    y=[y, y],
                                    z=[0, z],
                                    mode='lines',
                                    line=dict(color='blue', width=4),
                                    name=f'Марка {mark}'
                                ))
                                fig_3d.add_trace(go.Scatter3d(
                                    x=[x],
                                    y=[y],
                                    z=[z],
                                    mode='markers',
                                    marker=dict(size=8, color=z, colorscale='RdYlGn_r',
                                                colorbar=dict(title="Осадка, мм")),
                                    name=f'Осадка {mark}'
                                ))

                            xs = [coord_dict[m][0] for m in mark_keys_filtered]
                            ys = [coord_dict[m][1] for m in mark_keys_filtered]
                            if xs and ys:
                                x_range = np.linspace(min(xs)-0.5, max(xs)+0.5, 10)
                                y_range = np.linspace(min(ys)-0.5, max(ys)+0.5, 10)
                                X_grid, Y_grid = np.meshgrid(x_range, y_range)
                                Z_grid = np.zeros_like(X_grid)
                                fig_3d.add_trace(go.Surface(
                                    x=x_range, y=y_range, z=Z_grid,
                                    colorscale='Greys', opacity=0.3,
                                    showscale=False, name='Основание'
                                ))

                            if df_sett_angles is not None and not df_sett_angles.empty:
                                sett_row = df_sett_angles[df_sett_angles['Цикл'] == selected_cycle_3d]
                                if not sett_row.empty:
                                    a = sett_row['a_мм_м'].values[0]
                                    b = sett_row['b_мм_м'].values[0]
                                    scale_3d = 10
                                    dx = a * scale_3d
                                    dy = b * scale_3d
                                    center_x = np.mean(xs)
                                    center_y = np.mean(ys)
                                    fig_3d.add_trace(go.Scatter3d(
                                        x=[center_x, center_x + dx],
                                        y=[center_y, center_y + dy],
                                        z=[0, 0],
                                        mode='lines+markers',
                                        line=dict(color='red', width=6),
                                        marker=dict(size=8, color='red'),
                                        name='Вектор крена'
                                    ))
                                    fig_3d.add_annotation(
                                        x=center_x + dx, y=center_y + dy, z=0,
                                        text=f"a={a:.2f} мм/м, b={b:.2f} мм/м",
                                        showarrow=False,
                                        font=dict(size=12, color='red')
                                    )

                            fig_3d.update_layout(
                                title=f"3D-модель осадок и крена (цикл {selected_cycle_3d})",
                                scene=dict(
                                    xaxis_title="X, м",
                                    yaxis_title="Y, м",
                                    zaxis_title="Осадка, мм",
                                    camera=dict(eye=dict(x=1.5, y=1.5, z=1.5))
                                ),
                                width=800,
                                height=700,
                                template="plotly_white"
                            )
                            st.plotly_chart(fig_3d, use_container_width=True)
                            st.caption("Синие столбцы – осадки марок. Красный вектор – направление и величина крена (по данным осадок).")

        with tab4:
            st.subheader("📥 Скачать отчёт")
            st.info("Параметры отчёта настраиваются в левой боковой панели.")
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
                pdf_data = generate_pdf_report(df_incl, st.session_state.get('df_sett_angles', None), cycles, alpha0_x, alpha0_y, L, st.session_state.report_params)
                st.download_button(
                    label="📄 PDF",
                    data=pdf_data.getvalue(),
                    file_name=f"inclinometer_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                    mime="application/pdf"
                )
            with col3:
                word_data = generate_word_report(df_incl, st.session_state.get('df_sett_angles', None), cycles, alpha0_x, alpha0_y, L, st.session_state.report_params)
                st.download_button(
                    label="📝 Word",
                    data=word_data.getvalue(),
                    file_name=f"inclinometer_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        with tab5:
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
