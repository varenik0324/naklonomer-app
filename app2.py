import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import plotly.io as pio
import io
import re
import json
import sqlite3
import requests
import os
import sys
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import logging

# ========== НАСТРОЙКИ ==========
BOT_TOKEN = "8538186715:AAG7XsBxp6TAy2lalWQ6_KkBkrUIEZCqxuw"  # ЗАМЕНИ
CHAT_ID = "1278271780"
logging.basicConfig(filename='app_errors.log', level=logging.ERROR)

def send_telegram(message):
    try:
        url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"
        payload = {"chat_id": CHAT_ID, "text": f"📩 {message}", "parse_mode": "HTML"}
        r = requests.post(url, json=payload, timeout=5)
        return r.status_code == 200
    except Exception as e:
        logging.error(f"Telegram: {e}")
        return False

# ========== ПАРСИНГ НАКЛОНОМЕРА (УЛУЧШЕННЫЙ) ==========
def parse_inclinometer_data(file_bytes, sheet_name=None, manual_floor_rows=None):
    """
    Автоматически определяет лист, строки с этажами, столбцы с углами.
    """
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    all_sheets = xl.sheet_names

    if sheet_name is None:
        # Ищем лист по ключевым словам
        for name in all_sheets:
            if 'наклономер' in name.lower() or 'крен' in name.lower() or 'инклинометр' in name.lower():
                sheet_name = name
                break
        if sheet_name is None:
            # Если не найден, берём первый непустой лист с большим количеством чисел
            for name in all_sheets:
                df_sample = pd.read_excel(file_bytes, sheet_name=name, header=None, nrows=30)
                num_count = df_sample.map(lambda x: isinstance(x, (int, float))).sum().sum()
                if num_count > 50:
                    sheet_name = name
                    break
        if sheet_name is None:
            st.error("Не удалось определить лист с наклономером. Выберите вручную.")
            return None

    df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)
    total_rows, total_cols = df_raw.shape

    # ---- Автоопределение строки с заголовками циклов ----
    cycle_header_row = None
    for idx, row in df_raw.iterrows():
        row_str = ' '.join(str(cell) for cell in row if pd.notna(cell))
        if 'Цикл' in row_str and re.search(r'\d{2}\.\d{2}\.\d{4}', row_str):
            cycle_header_row = idx
            break

    if cycle_header_row is None:
        st.warning("Не найдены заголовки циклов. Используем порядковые номера.")
        cycle_labels = None
    else:
        cycle_labels = []
        for col in range(1, total_cols):
            cell = df_raw.iloc[cycle_header_row, col]
            if pd.notna(cell):
                match = re.search(r'(\d{2}\.\d{2}\.\d{4})', str(cell))
                if match:
                    try:
                        date_obj = pd.to_datetime(match.group(1), dayfirst=True)
                        cycle_labels.append(date_obj.strftime('%Y-%m-%d'))
                    except:
                        cycle_labels.append(str(cell))
                else:
                    cycle_labels.append(str(cell))
        cycle_labels = [c for c in cycle_labels if c]

    # ---- Автоопределение строк с этажами ----
    if manual_floor_rows is not None:
        floor_rows = manual_floor_rows
    else:
        # Поиск строк, где есть числа 5, 15, 27 и много числовых данных
        floor_rows = {}
        for idx in range(cycle_header_row + 1 if cycle_header_row else 0, total_rows):
            row = df_raw.iloc[idx]
            num_count = sum(1 for v in row if pd.notna(v) and isinstance(v, (int, float)))
            if num_count < 6:
                continue
            found_floor = None
            for cell in row:
                if pd.isna(cell):
                    continue
                cell_str = str(cell).strip()
                match = re.search(r'\b(5|15|27)\b', cell_str)
                if match:
                    found_floor = int(match.group(1))
                    break
            if found_floor:
                floor_rows[found_floor] = idx
            if len(floor_rows) == 3:
                break

        if len(floor_rows) < 2:
            st.warning("Не удалось автоопределить строки этажей. Используйте ручной ввод.")
            return None

    # Сбор данных
    max_pairs = 0
    for floor_idx in floor_rows.values():
        values = []
        for col in range(1, total_cols):
            cell = df_raw.iloc[floor_idx, col]
            if pd.notna(cell) and isinstance(cell, (int, float)):
                values.append(float(cell))
        pairs = [(values[i], values[i+1]) for i in range(0, len(values)-1, 2)]
        max_pairs = max(max_pairs, len(pairs))

    if not cycle_labels:
        cycle_labels = [f"Цикл {i+1}" for i in range(max_pairs)]
    elif len(cycle_labels) < max_pairs:
        cycle_labels += [f"Цикл {i+1}" for i in range(len(cycle_labels), max_pairs)]

    data = []
    for floor, floor_idx in floor_rows.items():
        values = []
        for col in range(1, total_cols):
            cell = df_raw.iloc[floor_idx, col]
            if pd.notna(cell) and isinstance(cell, (int, float)):
                values.append(float(cell))
        pairs = [(values[i], values[i+1]) for i in range(0, len(values)-1, 2)]
        for i, (ax, ay) in enumerate(pairs):
            if i < len(cycle_labels):
                data.append({'Цикл': cycle_labels[i], 'Этаж': floor, 'αx': ax, 'αy': ay})

    if not data:
        st.error("Не удалось извлечь данные.")
        return None

    df = pd.DataFrame(data)
    df['Цикл'] = df['Цикл'].astype(str)
    df['Этаж'] = df['Этаж'].astype(int)
    df['αx'] = pd.to_numeric(df['αx'], errors='coerce')
    df['αy'] = pd.to_numeric(df['αy'], errors='coerce')
    df = df.dropna(subset=['αx', 'αy'])
    return df

# ========== ПАРСИНГ ОСАДОК (УЛУЧШЕННЫЙ) ==========
def parse_settlement_data(file_bytes, sheet_name=None, corner_marks=None, L=70, B=18):
    """
    Автоопределение листа, столбцов с осадками.
    """
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    all_sheets = xl.sheet_names
    if sheet_name is None:
        for name in all_sheets:
            if 'стилобат' in name.lower() or 'высотн' in name.lower() or 'осадк' in name.lower():
                sheet_name = name
                break
        if sheet_name is None:
            st.error("Не найден лист с осадками.")
            return None

    df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)
    # Поиск строки с заголовками циклов
    cycle_header_row = None
    for idx, row in df_raw.iterrows():
        row_str = ' '.join(str(cell) for cell in row if pd.notna(cell))
        if 'Цикл' in row_str:
            cycle_header_row = idx
            break
    if cycle_header_row is None:
        st.error("Не найдены заголовки циклов в листе осадок.")
        return None

    # Определяем столбцы с осадками
    cycle_cols = {}
    for col_idx, cell in df_raw.iloc[cycle_header_row, :].items():
        if pd.notna(cell) and 'Цикл' in str(cell):
            date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', str(cell))
            if date_match:
                try:
                    cycle_label = pd.to_datetime(date_match.group(1), dayfirst=True).strftime('%Y-%m-%d')
                except:
                    cycle_label = str(cell)
            else:
                cycle_label = str(cell)
            # Ищем столбец с осадками (обычно через 2-3 колонки)
            for offset in [1, 2, 3]:
                if col_idx + offset < len(df_raw.columns):
                    next_cell = df_raw.iloc[cycle_header_row, col_idx + offset]
                    if pd.notna(next_cell) and 'осадк' in str(next_cell).lower():
                        cycle_cols[cycle_label] = col_idx + offset
                        break
            if cycle_label not in cycle_cols:
                cycle_cols[cycle_label] = col_idx + 2  # по умолчанию

    if not cycle_cols:
        st.error("Не найдены колонки с осадками.")
        return None

    # Автоопределение марок (первый столбец с числами после заголовка)
    mark_rows = []
    for idx in range(cycle_header_row + 1, len(df_raw)):
        cell_val = df_raw.iloc[idx, 0]
        if pd.notna(cell_val) and (isinstance(cell_val, (int, float)) or (isinstance(cell_val, str) and cell_val.strip())):
            mark_rows.append(idx)

    if not mark_rows:
        st.error("Не найдены марки.")
        return None

    # Если corner_marks не заданы – берём первые 4 марки
    if corner_marks is None:
        corner_marks = [str(df_raw.iloc[idx, 0]) for idx in mark_rows[:4]]
    else:
        corner_marks = [str(m) for m in corner_marks]

    # Сбор данных
    data = []
    for cycle, col_idx in cycle_cols.items():
        for mark_idx in mark_rows:
            mark = str(df_raw.iloc[mark_idx, 0]).strip()
            if mark not in corner_marks:
                continue
            settlement = df_raw.iloc[mark_idx, col_idx]
            if pd.notna(settlement) and isinstance(settlement, (int, float)):
                data.append({'Цикл': cycle, 'Марка': mark, 'Осадка_мм': settlement})

    if not data:
        st.error("Не удалось извлечь осадки.")
        return None

    df_sett = pd.DataFrame(data)

    # Расчёт углов по осадкам
    marks_order = corner_marks
    results = []
    for cycle in df_sett['Цикл'].unique():
        cycle_data = df_sett[df_sett['Цикл'] == cycle]
        s = {}
        for mark in marks_order:
            val = cycle_data[cycle_data['Марка'] == mark]['Осадка_мм']
            s[mark] = val.iloc[0] if not val.empty else np.nan
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
    return pd.DataFrame(results)

# ========== ГЕНЕРАЦИЯ ОТЧЁТОВ ==========
def generate_excel_report(df_incl, df_sett, stats, params):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if df_incl is not None:
            df_incl.to_excel(writer, index=False, sheet_name='Наклономер')
        if df_sett is not None:
            df_sett.to_excel(writer, index=False, sheet_name='Осадки_углы')
        if stats:
            pd.DataFrame(stats).to_excel(writer, sheet_name='Статистика')
        pd.DataFrame(params).to_excel(writer, sheet_name='Параметры')
    return output.getvalue()

def generate_html_report(df_incl, df_sett, stats, params):
    # Создаём интерактивный HTML с графиками Plotly
    fig1 = go.Figure()
    if df_incl is not None:
        for floor in df_incl['Этаж'].unique():
            floor_df = df_incl[df_incl['Этаж'] == floor]
            fig1.add_trace(go.Scatter(x=floor_df['Цикл'], y=floor_df['αx_abs'], name=f'αx этаж {floor}'))
    html = pio.to_html(fig1, include_plotlyjs='cdn', full_html=False)
    # Добавляем таблицы
    html += "<h2>Статистика</h2>" + pd.DataFrame(stats).to_html()
    return html

def generate_pdf_report(df_incl, df_sett, stats, params):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    y = height - 50
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, y, "Отчёт по мониторингу")
    y -= 30
    c.setFont("Helvetica", 10)
    for key, val in params.items():
        c.drawString(50, y, f"{key}: {val}")
        y -= 15
    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer

def generate_word_report(df_incl, df_sett, stats, params):
    doc = Document()
    doc.add_heading("Отчёт по мониторингу", level=1)
    for key, val in params.items():
        doc.add_paragraph(f"{key}: {val}")
    doc.add_heading("Статистика", level=2)
    if stats:
        for k, v in stats.items():
            doc.add_paragraph(f"{k}: {v}")
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ========== БАЗА ДАННЫХ ==========
def init_db():
    conn = sqlite3.connect('monitoring_history.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS results
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  date TEXT, object_name TEXT, cycle TEXT,
                  avg_ax REAL, avg_ay REAL, max_ax REAL, max_ay REAL,
                  json_data TEXT)''')
    conn.commit()
    conn.close()

def save_to_db(object_name, cycle, df_incl):
    conn = sqlite3.connect('monitoring_history.db')
    c = conn.cursor()
    avg_ax = df_incl['αx_abs'].mean()
    avg_ay = df_incl['αy_abs'].mean()
    max_ax = df_incl['αx_abs'].max()
    max_ay = df_incl['αy_abs'].max()
    json_data = df_incl.to_json(orient='records')
    c.execute("INSERT INTO results (date, object_name, cycle, avg_ax, avg_ay, max_ax, max_ay, json_data) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
              (datetime.now().isoformat(), object_name, cycle, avg_ax, avg_ay, max_ax, max_ay, json_data))
    conn.commit()
    conn.close()

def load_history():
    conn = sqlite3.connect('monitoring_history.db')
    df = pd.read_sql_query("SELECT * FROM results ORDER BY id DESC", conn)
    conn.close()
    return df

# ========== ОСНОВНОЙ ИНТЕРФЕЙС ==========
st.set_page_config(page_title="Анализ наклономера + осадки", layout="wide")
st.title("📐 Анализ данных накладного инклинометра и осадок фундамента")

# Боковая панель
with st.sidebar:
    st.header("Настройки")
    theme = st.selectbox("Тема", ["Светлая", "Тёмная", "Корпоративная"], index=0)
    if theme == "Светлая":
        template = "plotly_white"
    elif theme == "Тёмная":
        template = "plotly_dark"
    else:
        template = "seaborn"
    st.session_state.template = template

    st.header("Параметры отчёта")
    report_params = {
        'организация': st.text_input("Организация", "ООО «Геофундамент»"),
        'address': st.text_input("Адрес", "111673, Москва, ул. Суздальская д. 18 корп. 4"),
        'object_name': st.text_input("Объект", "Многофункциональный жилой комплекс"),
        'customer': st.text_input("Заказчик", "ООО СЗ «Сампад»"),
        'cycle_number': st.text_input("Цикл", "6-й цикл"),
        'date': st.text_input("Дата", datetime.now().strftime("%d.%m.%Y")),
        'исполнитель': st.text_input("Исполнитель", "Добшиков А.Н."),
    }

uploaded_file = st.file_uploader("Загрузите Excel-файл", type=["xlsx", "xls", "csv", "json"])

if uploaded_file:
    file_bytes = uploaded_file.read()
    # Определяем формат
    if uploaded_file.name.endswith('.csv'):
        df_raw = pd.read_csv(io.BytesIO(file_bytes))
    elif uploaded_file.name.endswith('.json'):
        df_raw = pd.read_json(io.BytesIO(file_bytes))
    else:
        df_raw = None

    # Парсинг наклономера (авто)
    df_incl = parse_inclinometer_data(file_bytes)
    if df_incl is None:
        st.error("Не удалось распарсить данные. Попробуйте вручную выбрать лист и строки.")
        st.stop()

    # Параметры расчёта
    st.session_state.df_incl = df_incl
    cycles = sorted(df_incl['Цикл'].unique())
    zero_cycle = st.selectbox("Нулевой цикл", cycles, index=0)
    alpha0_x = st.number_input("αx0, °", value=0.0, step=0.001)
    alpha0_y = st.number_input("αy0, °", value=0.0, step=0.001)
    L = st.number_input("Высота этажа, м", value=3.0)

    # Расчёт абсолютных углов
    zero_data = df_incl[df_incl['Цикл'] == zero_cycle][['Этаж', 'αx', 'αy']].rename(columns={'αx': 'αx0_inc', 'αy': 'αy0_inc'})
    df_incl = df_incl.merge(zero_data, on='Этаж', how='left')
    df_incl['αx_abs'] = df_incl['αx'] - df_incl['αx0_inc'] + alpha0_x
    df_incl['αy_abs'] = df_incl['αy'] - df_incl['αy0_inc'] + alpha0_y
    df_incl['Смещение X'] = L * np.sin(np.radians(df_incl['αx_abs']))
    df_incl['Смещение Y'] = L * np.sin(np.radians(df_incl['αy_abs']))

    # Автоопределение осадок (если есть)
    df_sett = parse_settlement_data(file_bytes)
    if df_sett is not None:
        st.session_state.df_sett = df_sett

    # Статистика
    stats = {
        'Средний αx_abs': df_incl['αx_abs'].mean(),
        'Макс αx_abs': df_incl['αx_abs'].max(),
        'Средний αy_abs': df_incl['αy_abs'].mean(),
        'Макс αy_abs': df_incl['αy_abs'].max(),
        'Среднее смещение X, м': df_incl['Смещение X'].mean(),
        'Среднее смещение Y, м': df_incl['Смещение Y'].mean(),
    }

    # Вкладки
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 Данные", "📈 Графики", "📐 Профиль", "📊 Статистика", "📥 Отчёт"])

    with tab1:
        st.dataframe(df_incl)
        if df_sett is not None:
            st.subheader("Данные осадок")
            st.dataframe(df_sett)

    with tab2:
        st.subheader("Изменение углов по времени")
        fig = go.Figure()
        for floor in df_incl['Этаж'].unique():
            floor_df = df_incl[df_incl['Этаж'] == floor]
            fig.add_trace(go.Scatter(x=floor_df['Цикл'], y=floor_df['αx_abs'], name=f'αx эт.{floor}'))
            fig.add_trace(go.Scatter(x=floor_df['Цикл'], y=floor_df['αy_abs'], name=f'αy эт.{floor}', line=dict(dash='dot')))
        fig.update_layout(template=st.session_state.template)
        st.plotly_chart(fig, use_container_width=True)

    with tab3:
        st.subheader("Профиль смещений по этажам")
        selected_cycles = st.multiselect("Выберите циклы", df_incl['Цикл'].unique(), default=[df_incl['Цикл'].unique()[-1]])
        fig2 = go.Figure()
        for cycle in selected_cycles:
            profile = df_incl[df_incl['Цикл'] == cycle].sort_values('Этаж')
            fig2.add_trace(go.Scatter(x=profile['Смещение X'], y=profile['Этаж'], mode='lines+markers', name=f'X {cycle}'))
            fig2.add_trace(go.Scatter(x=profile['Смещение Y'], y=profile['Этаж'], mode='lines+markers', name=f'Y {cycle}', line=dict(dash='dot')))
        fig2.update_layout(yaxis=dict(autorange='reversed'), template=st.session_state.template)
        st.plotly_chart(fig2, use_container_width=True)

    with tab4:
        st.subheader("Детальная статистика")
        st.dataframe(pd.DataFrame(stats, index=[0]))
        if df_sett is not None:
            st.subheader("Осадки: сравнение с нормами")
            # Пример сравнения
            max_sett = df_sett['a_град'].max() if 'a_град' in df_sett else 0
            limit = 0.001  # из СП
            if max_sett < limit:
                st.success(f"Максимальный угол по осадкам {max_sett:.5f}° < {limit}° (норма) ✅")
            else:
                st.error(f"Превышение! {max_sett:.5f}° > {limit}°")

    with tab5:
        st.subheader("Скачать отчёт")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            excel = generate_excel_report(df_incl, df_sett, stats, report_params)
            st.download_button("📊 Excel", data=excel, file_name="report.xlsx")
        with col2:
            pdf = generate_pdf_report(df_incl, df_sett, stats, report_params)
            st.download_button("📄 PDF", data=pdf.getvalue(), file_name="report.pdf")
        with col3:
            word = generate_word_report(df_incl, df_sett, stats, report_params)
            st.download_button("📝 Word", data=word.getvalue(), file_name="report.docx")
        with col4:
            html = generate_html_report(df_incl, df_sett, stats, report_params)
            st.download_button("🌐 HTML", data=html, file_name="report.html")

        # Сохранение в историю
        if st.button("Сохранить в историю"):
            init_db()
            save_to_db(report_params['object_name'], report_params['cycle_number'], df_incl)
            st.success("Сохранено в историю!")

        # Показать историю
        if st.checkbox("Показать историю"):
            history = load_history()
            st.dataframe(history)

        # Уведомление о превышениях
        if st.button("Отправить уведомление в Telegram"):
            msg = f"Объект: {report_params['object_name']}\nЦикл: {report_params['cycle_number']}\nМакс. αx: {stats['Макс αx_abs']:.3f}°"
            if send_telegram(msg):
                st.success("Уведомление отправлено!")
            else:
                st.error("Ошибка отправки.")

else:
    st.info("👆 Загрузите файл для начала работы.")
