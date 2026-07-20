import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import io
import tempfile
from datetime import datetime
import re
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from PIL import Image
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ------------------------------------------------------------
# Настройки страницы
# ------------------------------------------------------------
st.set_page_config(page_title="Анализ наклономера", layout="wide")
st.title("📐 Анализ данных накладного инклинометра (наклономера)")

# ------------------------------------------------------------
# Функции парсинга
# ------------------------------------------------------------
def parse_inclinometer_data(file_bytes):
    """
    Парсинг Excel-файла с данными наклономера.
    Ожидаем структуру: 
        - Первая строка (или несколько) – заголовки с датами циклов.
        - Строки с этажами (5, 15, 27) и значениями углов по осям X и Y.
    Возвращает DataFrame с колонками: Цикл, Этаж, αx, αy.
    """
    xl = pd.ExcelFile(file_bytes)
    sheet_names = xl.sheet_names
    # Ищем лист, содержащий "наклономер" или "крен"
    target_sheet = None
    for name in sheet_names:
        if 'наклономер' in name.lower() or 'крен' in name.lower():
            target_sheet = name
            break
    if target_sheet is None:
        # Если не нашли, берём первый лист
        target_sheet = sheet_names[0]
        st.warning(f"Лист с данными наклономера не найден, использую первый лист: {target_sheet}")

    df_raw = pd.read_excel(file_bytes, sheet_name=target_sheet, header=None)
    # Ищем строки с этажами (числа 5, 15, 27) в первом столбце
    floor_rows = []
    for idx, row in df_raw.iterrows():
        if pd.notna(row[0]) and isinstance(row[0], (int, float)) and row[0] in [5, 15, 27]:
            floor_rows.append(idx)
    if not floor_rows:
        st.error("Не найдены строки с этажами (5, 15, 27). Проверьте структуру файла.")
        return None

    # Ищем заголовки с датами циклов (строки, где есть "Цикл" или дата)
    cycle_header_row = None
    for idx, row in df_raw.iterrows():
        row_str = ' '.join(str(cell) for cell in row if pd.notna(cell))
        if 'Цикл' in row_str or re.search(r'\d{2}\.\d{2}\.\d{4}', row_str):
            cycle_header_row = idx
            break
    if cycle_header_row is None:
        st.warning("Заголовки циклов не найдены, будут использованы порядковые номера.")
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

    # Теперь для каждого этажа собираем значения углов
    data = []
    for floor_idx in floor_rows:
        floor = int(df_raw.iloc[floor_idx, 0])
        # Ищем в этой строке числовые значения, которые могут быть углами
        # Предполагаем, что значения идут парами (αx, αy) для каждого цикла
        values = []
        for cell in df_raw.iloc[floor_idx, 1:]:
            if pd.notna(cell) and isinstance(cell, (int, float)):
                values.append(float(cell))
        # Если значений много, группируем по парам
        # Если число значений чётное, берём пары (x, y)
        # Если нечётное, считаем, что это только x, а y отсутствует
        if len(values) % 2 == 0:
            pairs = [(values[i], values[i+1]) for i in range(0, len(values), 2)]
        else:
            pairs = [(values[i], np.nan) for i in range(len(values))]
        # Сопоставляем с циклами
        for i, (ax, ay) in enumerate(pairs):
            if i < len(cycle_labels):
                data.append({
                    'Цикл': cycle_labels[i] if cycle_labels[i] else f"Цикл {i+1}",
                    'Этаж': floor,
                    'αx': ax,
                    'αy': ay
                })
    if not data:
        st.error("Не удалось извлечь данные. Проверьте структуру файла.")
        return None
    df = pd.DataFrame(data)
    return df

# ------------------------------------------------------------
# Функции для отчётов (упрощённые)
# ------------------------------------------------------------
def generate_excel_report(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Данные')
        # Сводка по этажам
        summary = df.groupby('Этаж').agg({'αx': ['mean', 'min', 'max', 'std'], 'αy': ['mean', 'min', 'max', 'std']})
        summary.to_excel(writer, sheet_name='Сводка')
    return output.getvalue()

def generate_pdf_report(df):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, height - 50, "Отчёт по наклономеру")
    c.setFont("Helvetica", 12)
    c.drawString(50, height - 80, f"Дата: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
    c.drawString(50, height - 100, f"Количество записей: {len(df)}")
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, height - 130, "Сводка по этажам:")
    y = height - 150
    for floor in df['Этаж'].unique():
        floor_data = df[df['Этаж'] == floor]
        c.setFont("Helvetica", 10)
        c.drawString(60, y, f"Этаж {floor}:")
        c.drawString(120, y, f"αx ср. = {floor_data['αx'].mean():.3f}°, αy ср. = {floor_data['αy'].mean():.3f}°")
        y -= 20
        if y < 50:
            c.showPage()
            y = height - 50
    c.save()
    buffer.seek(0)
    return buffer

def generate_word_report(df):
    doc = Document()
    doc.add_heading('Отчёт по наклономеру', level=1)
    doc.add_paragraph(f"Дата: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
    doc.add_paragraph(f"Количество записей: {len(df)}")
    doc.add_heading('Сводка по этажам', level=2)
    for floor in df['Этаж'].unique():
        floor_data = df[df['Этаж'] == floor]
        doc.add_paragraph(f"Этаж {floor}: αx ср. = {floor_data['αx'].mean():.3f}°, αy ср. = {floor_data['αy'].mean():.3f}°")
    doc.add_paragraph("© Геофундамент, 2026").alignment = WD_ALIGN_PARAGRAPH.CENTER
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ------------------------------------------------------------
# Интерфейс
# ------------------------------------------------------------
uploaded_file = st.file_uploader("Загрузите Excel-файл с данными наклономера", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df = parse_inclinometer_data(uploaded_file)
        if df is not None:
            st.success("Данные успешно загружены!")
            st.dataframe(df)

            # График
            fig = go.Figure()
            for floor in df['Этаж'].unique():
                floor_df = df[df['Этаж'] == floor]
                fig.add_trace(go.Scatter(x=floor_df['Цикл'], y=floor_df['αx'], mode='lines+markers', name=f'Этаж {floor} αx'))
                fig.add_trace(go.Scatter(x=floor_df['Цикл'], y=floor_df['αy'], mode='lines+markers', name=f'Этаж {floor} αy'))
            fig.update_layout(title="Изменение углов наклона", xaxis_title="Цикл", yaxis_title="Угол, °")
            st.plotly_chart(fig, use_container_width=True)

            # Скачивание отчётов
            st.subheader("📥 Скачать отчёт")
            col1, col2, col3 = st.columns(3)
            with col1:
                excel_data = generate_excel_report(df)
                st.download_button(label="📊 Excel", data=excel_data, file_name=f"inclinometer_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with col2:
                pdf_data = generate_pdf_report(df)
                st.download_button(label="📄 PDF", data=pdf_data.getvalue(), file_name=f"inclinometer_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf", mime="application/pdf")
            with col3:
                word_data = generate_word_report(df)
                st.download_button(label="📝 Word", data=word_data.getvalue(), file_name=f"inclinometer_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    except Exception as e:
        st.error(f"Ошибка обработки: {e}")
else:
    st.info("Загрузите Excel-файл с данными наклономера для начала работы.")
