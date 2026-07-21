import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import io
import tempfile
import re
from datetime import datetime
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
st.set_page_config(
    page_title="Анализ наклономера",
    page_icon="📐",
    layout="wide"
)
st.title("📐 Анализ данных накладного инклинометра (наклономера)")
st.markdown("Загрузите Excel-файл с данными измерений, укажите параметры, и приложение построит графики, профили и сформирует отчёты.")

# ------------------------------------------------------------
# ПАРСИНГ ДАННЫХ (автоматическое распознавание структуры)
# ------------------------------------------------------------
def parse_inclinometer_data(file_bytes):
    """
    Парсит Excel-файл с данными наклономера.
    Ожидает структуру: строки с этажами (5, 15, 27) и столбцы с циклами (даты или номера).
    Возвращает DataFrame с колонками: Цикл, Этаж, αx, αy.
    """
    xl = pd.ExcelFile(file_bytes)
    sheet_names = xl.sheet_names

    target_sheet = None
    for name in sheet_names:
        if 'наклономер' in name.lower() or 'крен' in name.lower():
            target_sheet = name
            break
    if target_sheet is None:
        target_sheet = sheet_names[0]
        st.info(f"Лист с данными наклономера не найден, используется первый лист: {target_sheet}")

    df_raw = pd.read_excel(file_bytes, sheet_name=target_sheet, header=None)

    # --- 1. Находим строки с этажами (5, 15, 27) ---
    floor_rows = []
    for idx, row in df_raw.iterrows():
        if pd.notna(row[0]) and isinstance(row[0], (int, float)) and row[0] in [5, 15, 27]:
            floor_rows.append(idx)
    if not floor_rows:
        st.error("Не найдены строки с этажами (5, 15, 27). Проверьте структуру файла.")
        return None

    # --- 2. Находим строку с заголовками циклов ---
    cycle_header_row = None
    for idx, row in df_raw.iterrows():
        row_str = ' '.join(str(cell) for cell in row if pd.notna(cell))
        if 'Цикл' in row_str or re.search(r'\d{2}\.\d{2}\.\d{4}', row_str):
            cycle_header_row = idx
            break

    if cycle_header_row is None:
        st.warning("Заголовки циклов не найдены, будут использованы порядковые номера.")
        num_cycles = len(df_raw.columns) - 1
        cycle_labels = [f"Цикл {i+1}" for i in range(num_cycles)]
    else:
        cycle_labels = []
        for cell in df_raw.iloc[cycle_header_row, 1:]:
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

    # --- 3. Собираем данные по этажам ---
    data = []
    for floor_idx in floor_rows:
        floor = int(df_raw.iloc[floor_idx, 0])
        values = []
        for cell in df_raw.iloc[floor_idx, 1:]:
            if pd.notna(cell) and isinstance(cell, (int, float)):
                values.append(float(cell))

        pairs = []
        if len(values) % 2 == 0:
            pairs = [(values[i], values[i+1]) for i in range(0, len(values), 2)]
        else:
            pairs = [(values[i], np.nan) for i in range(len(values))]

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
    df['Цикл'] = df['Цикл'].astype(str)
    df['Этаж'] = df['Этаж'].astype(int)
    df['αx'] = pd.to_numeric(df['αx'], errors='coerce')
    df['αy'] = pd.to_numeric(df['αy'], errors='coerce')
    return df

# ------------------------------------------------------------
# ГЕНЕРАЦИЯ ОТЧЁТОВ (с учётом расчётов)
# ------------------------------------------------------------
def generate_excel_report(df, cycles, alpha0_x, alpha0_y, L):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Данные с абсолютными углами и смещениями
        df.to_excel(writer, index=False, sheet_name='Данные')
        # Сводка по этажам (последний цикл)
        summary = df[df['Цикл'] == df['Цикл'].iloc[-1]].copy()
        summary['Смещение X, м'] = L * np.sin(np.radians(summary['αx_abs']))
        summary['Смещение Y, м'] = L * np.sin(np.radians(summary['αy_abs']))
        summary[['Этаж', 'αx_abs', 'αy_abs', 'Смещение X, м', 'Смещение Y, м']].to_excel(writer, sheet_name='Профиль', index=False)
        # Параметры
        params = pd.DataFrame({
            'Параметр': ['Начальный угол X, °', 'Начальный угол Y, °', 'Высота этажа, м', 'Нулевой цикл'],
            'Значение': [alpha0_x, alpha0_y, L, cycles[0]]
        })
        params.to_excel(writer, sheet_name='Параметры', index=False)
    return output.getvalue()

def generate_pdf_report(df, cycles, alpha0_x, alpha0_y, L):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, height - 50, "Отчёт по наклономеру")
    c.setFont("Helvetica", 12)
    c.drawString(50, height - 80, f"Дата: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
    c.drawString(50, height - 100, f"Параметры: αx0 = {alpha0_x:.3f}°, αy0 = {alpha0_y:.3f}°, L = {L} м")
    c.drawString(50, height - 120, f"Нулевой цикл: {cycles[0]}")
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, height - 150, "Профиль смещений (последний цикл):")
    y = height - 170
    last_cycle = df['Цикл'].iloc[-1]
    profile = df[df['Цикл'] == last_cycle]
    for _, row in profile.iterrows():
        c.setFont("Helvetica", 10)
        c.drawString(60, y, f"Этаж {row['Этаж']}: смещ. X = {row['Смещение X']:.3f} м, Y = {row['Смещение Y']:.3f} м")
        y -= 20
        if y < 50:
            c.showPage()
            y = height - 50
    c.save()
    buffer.seek(0)
    return buffer

def generate_word_report(df, cycles, alpha0_x, alpha0_y, L):
    doc = Document()
    doc.add_heading('Отчёт по наклономеру', level=1)
    doc.add_paragraph(f"Дата: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
    doc.add_paragraph(f"Параметры: αx0 = {alpha0_x:.3f}°, αy0 = {alpha0_y:.3f}°, L = {L} м")
    doc.add_paragraph(f"Нулевой цикл: {cycles[0]}")
    doc.add_heading('Профиль смещений (последний цикл)', level=2)
    last_cycle = df['Цикл'].iloc[-1]
    profile = df[df['Цикл'] == last_cycle]
    for _, row in profile.iterrows():
        doc.add_paragraph(f"Этаж {row['Этаж']}: смещ. X = {row['Смещение X']:.3f} м, Y = {row['Смещение Y']:.3f} м")
    doc.add_paragraph("© Геофундамент, 2026").alignment = WD_ALIGN_PARAGRAPH.CENTER
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ------------------------------------------------------------
# ИНТЕРФЕЙС STREAMLIT
# ------------------------------------------------------------
uploaded_file = st.file_uploader(
    "Загрузите Excel-файл с данными наклономера",
    type=["xlsx", "xls"],
    help="Ожидается структура: строки с этажами (5, 15, 27), столбцы с циклами (даты или номера), внутри – пары углов αx, αy."
)

if uploaded_file is not None:
    try:
        df = parse_inclinometer_data(uploaded_file)
        if df is None:
            st.stop()

        # Получаем уникальные циклы
        cycles = sorted(df['Цикл'].unique())
        if len(cycles) == 0:
            st.error("Нет циклов в данных.")
            st.stop()

        # -----------------------------------------------------------------
        # БОКОВАЯ ПАНЕЛЬ – настройки
        # -----------------------------------------------------------------
        st.sidebar.header("Параметры расчёта")
        default_cycle = cycles[0]
        zero_cycle = st.sidebar.selectbox("Нулевой цикл (начальный угол)", cycles, index=0)
        alpha0_x = st.sidebar.number_input("Начальный угол X (αx0), °", value=0.0, step=0.001, format="%.3f")
        alpha0_y = st.sidebar.number_input("Начальный угол Y (αy0), °", value=0.0, step=0.001, format="%.3f")
        L = st.sidebar.number_input("Высота этажа (L), м", value=3.0, step=0.1, format="%.1f", help="Используется для расчёта горизонтального смещения M = L * sin(α)")

        # -----------------------------------------------------------------
        # РАСЧЁТ АБСОЛЮТНЫХ УГЛОВ И СМЕЩЕНИЙ
        # -----------------------------------------------------------------
        # Выбираем нулевой цикл для вычитания
        zero_data = df[df['Цикл'] == zero_cycle][['Этаж', 'αx', 'αy']].rename(columns={'αx': 'αx0_inc', 'αy': 'αy0_inc'})
        # Присоединяем нулевые значения к основному DataFrame
        df = df.merge(zero_data, on='Этаж', how='left')
        # Абсолютный угол = (измеренный угол – нулевой угол инклинометра) + начальный угол фундамента
        df['αx_abs'] = df['αx'] - df['αx0_inc'] + alpha0_x
        df['αy_abs'] = df['αy'] - df['αy0_inc'] + alpha0_y
        # Смещения
        df['Смещение X'] = L * np.sin(np.radians(df['αx_abs']))
        df['Смещение Y'] = L * np.sin(np.radians(df['αy_abs']))

        st.success(f"✅ Данные успешно загружены! Найдено {len(df)} записей.")
        st.dataframe(df, use_container_width=True)

        # ---------- График абсолютных углов ----------
        st.subheader("📈 Изменение абсолютных углов наклона по циклам")
        fig = go.Figure()
        for floor in sorted(df['Этаж'].unique()):
            floor_df = df[df['Этаж'] == floor].sort_values('Цикл')
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

        # ---------- Профиль смещений (последний цикл) ----------
        last_cycle = df['Цикл'].iloc[-1]
        profile = df[df['Цикл'] == last_cycle].sort_values('Этаж')
        if not profile.empty:
            st.subheader(f"📊 Профиль смещений на цикле {last_cycle}")
            fig2 = go.Figure()
            fig2.add_trace(go.Scatter(
                x=profile['Смещение X'],
                y=profile['Этаж'],
                mode='lines+markers',
                name='Смещение X',
                marker=dict(size=10)
            ))
            fig2.add_trace(go.Scatter(
                x=profile['Смещение Y'],
                y=profile['Этаж'],
                mode='lines+markers',
                name='Смещение Y',
                marker=dict(size=10, symbol='square')
            ))
            fig2.update_layout(
                xaxis_title="Смещение, м",
                yaxis_title="Этаж",
                template="plotly_white",
                yaxis=dict(autorange="reversed"),
                hovermode="y unified"
            )
            st.plotly_chart(fig2, use_container_width=True)

        # ---------- Скачивание отчётов ----------
        st.subheader("📥 Скачать отчёт")
        col1, col2, col3 = st.columns(3)
        with col1:
            excel_data = generate_excel_report(df, cycles, alpha0_x, alpha0_y, L)
            st.download_button(
                label="📊 Excel",
                data=excel_data,
                file_name=f"inclinometer_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col2:
            pdf_data = generate_pdf_report(df, cycles, alpha0_x, alpha0_y, L)
            st.download_button(
                label="📄 PDF",
                data=pdf_data.getvalue(),
                file_name=f"inclinometer_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                mime="application/pdf"
            )
        with col3:
            word_data = generate_word_report(df, cycles, alpha0_x, alpha0_y, L)
            st.download_button(
                label="📝 Word",
                data=word_data.getvalue(),
                file_name=f"inclinometer_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    except Exception as e:
        st.error(f"Ошибка обработки: {e}")
        st.stop()

else:
    st.info("👆 Загрузите Excel-файл для начала работы.")
