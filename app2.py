import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import openpyxl
import re
from datetime import datetime

st.set_page_config(page_title="Автоматический расчёт осадок", layout="wide")
st.title("🌧️ Генератор отчёта по осадкам")
st.markdown("Загрузите Excel-файл с данными мониторинга (формат как в примерах)")

uploaded_file = st.file_uploader("Выберите файл .xlsx", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names
    st.sidebar.write("Найдены листы:", sheet_names)

    default_sheet = None
    for name in sheet_names:
        if name in ["Окружайка", "реконструкции", "Инженерные коммуникации"]:
            default_sheet = name
            break
    if not default_sheet:
        default_sheet = sheet_names[0]

    sheet = st.selectbox("Выберите лист с данными марок", sheet_names, index=sheet_names.index(default_sheet) if default_sheet else 0)
    df_raw = pd.read_excel(uploaded_file, sheet_name=sheet, header=None)

    st.subheader("Исходные данные (первые строки)")
    st.dataframe(df_raw.head(10))

    header_row_idx = None
    for i, row in df_raw.iterrows():
        if isinstance(row.iloc[0], str) and "№ марки" in row.iloc[0]:
            header_row_idx = i
            break

    if header_row_idx is None:
        st.error("Не найдена строка с заголовком '№ марки'. Проверьте формат файла.")
        st.stop()

    header_row = df_raw.iloc[header_row_idx]
    cycle_columns = {}
    current_cycle = None
    for col_idx, val in enumerate(header_row):
        if pd.isna(val):
            continue
        val_str = str(val)
        if "цикл" in val_str.lower():
            match = re.search(r'(\d+)-й цикл\s+(\d{2}\.\d{2}\.\d{4})', val_str, re.IGNORECASE)
            if match:
                cycle_num = int(match.group(1))
                date_str = match.group(2)
                if col_idx+1 < len(header_row) and "Отметка" in str(header_row.iloc[col_idx+1]):
                    cycle_columns[cycle_num] = {
                        'date': datetime.strptime(date_str, '%d.%m.%Y'),
                        'col_mark': col_idx,
                        'col_sediment': col_idx+1,
                        'col_total': col_idx+2
                    }

    if not cycle_columns:
        st.error("Не удалось найти столбцы с циклами. Проверьте структуру.")
        st.stop()

    st.write("Найдены циклы:", {k: v['date'].strftime('%Y-%m-%d') for k,v in cycle_columns.items()})

    data_start = header_row_idx + 1
    marks_data = []
    for i in range(data_start, len(df_raw)):
        row = df_raw.iloc[i]
        first_val = row.iloc[0]
        if pd.isna(first_val):
            continue
        if isinstance(first_val, str) and ("Таблица" in first_val or "Ведомость" in first_val):
            break
        mark = str(first_val).strip()
        if mark == '' or mark == 'nan':
            continue
        mark_dict = {'Марка': mark}
        for cycle_num, cols in cycle_columns.items():
            mark_value = row.iloc[cols['col_mark']] if cols['col_mark'] < len(row) else np.nan
            mark_dict[f'Цикл_{cycle_num}'] = mark_value if not pd.isna(mark_value) else np.nan
        marks_data.append(mark_dict)

    if not marks_data:
        st.error("Не найдены данные марок. Проверьте формат.")
        st.stop()

    df_marks = pd.DataFrame(marks_data)
    for col in df_marks.columns:
        if col.startswith('Цикл_'):
            df_marks[col] = pd.to_numeric(df_marks[col], errors='coerce')

    st.subheader("Данные марок (отметки по циклам)")
    st.dataframe(df_marks)

    zero_cycle = min(cycle_columns.keys())
    if zero_cycle not in df_marks.columns:
        st.error("Нет нулевого цикла в данных.")
        st.stop()

    for cycle_num in sorted(cycle_columns.keys()):
        if cycle_num == zero_cycle:
            continue
        col_current = f'Цикл_{cycle_num}'
        col_zero = f'Цикл_{zero_cycle}'
        df_marks[f'Осадка_{cycle_num}'] = (df_marks[col_current] - df_marks[col_zero]) * 1000

    df_marks['Макс_осадка_мм'] = df_marks[[f'Осадка_{c}' for c in sorted(cycle_columns.keys()) if c != zero_cycle]].max(axis=1)

    max_allowed = st.number_input("Максимально допустимая осадка (мм)", value=10.0, step=0.5)

    df_marks['Превышение'] = df_marks['Макс_осадка_мм'] > max_allowed

    st.subheader("Графики осадок по маркам")
    marks_to_plot = st.multiselect("Выберите марки для графика", df_marks['Марка'].tolist(), default=df_marks['Марка'].tolist()[:5])

    if marks_to_plot:
        plot_df = df_marks[df_marks['Марка'].isin(marks_to_plot)]
        cycles = sorted([c for c in cycle_columns.keys() if c != zero_cycle])
        long_data = []
        for _, row in plot_df.iterrows():
            mark = row['Марка']
            for c in cycles:
                sed = row.get(f'Осадка_{c}', np.nan)
                if not np.isnan(sed):
                    long_data.append({'Марка': mark, 'Цикл': c, 'Осадка_мм': sed})
        df_long = pd.DataFrame(long_data)
        if not df_long.empty:
            fig = px.line(df_long, x='Цикл', y='Осадка_мм', color='Марка', title="Изменение осадок по циклам")
            st.plotly_chart(fig, use_container_width=True)

            fig2 = px.bar(plot_df, x='Марка', y='Макс_осадка_мм', color='Марка', title="Максимальная осадка за период")
            st.plotly_chart(fig2, use_container_width=True)

    st.subheader("📄 Генерация отчёта")
    if st.button("Сформировать отчёт (DOCX)"):
        doc = Document()
        title = doc.add_heading('ВЕДОМОСТЬ ОСАДОК', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"Лист: {sheet}")
        doc.add_paragraph(f"Дата формирования: {datetime.now().strftime('%d.%m.%Y %H:%M')}")

        table_cols = ['Марка']
        for c in sorted(cycle_columns.keys()):
            table_cols.append(f'Цикл_{c}')
        for c in sorted(cycle_columns.keys()):
            if c != zero_cycle:
                table_cols.append(f'Осадка_{c}')
        table_cols.append('Макс_осадка_мм')

        table = doc.add_table(rows=1, cols=len(table_cols))
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        for i, col_name in enumerate(table_cols):
            hdr_cells[i].text = col_name

        for _, row in df_marks.iterrows():
            row_cells = table.add_row().cells
            for i, col in enumerate(table_cols):
                val = row.get(col, '')
                if isinstance(val, float):
                    if np.isnan(val):
                        val = ''
                    else:
                        val = f"{val:.3f}"
                row_cells[i].text = str(val)

        doc.add_paragraph()
        if df_marks['Превышение'].any():
            doc.add_paragraph("⚠️ Обнаружены превышения допустимой осадки!", style='List Bullet')
        else:
            doc.add_paragraph("✅ Все значения не превышают допустимую осадку.", style='List Bullet')

        doc_bytes = BytesIO()
        doc.save(doc_bytes)
        doc_bytes.seek(0)

        st.download_button(
            label="Скачать отчёт DOCX",
            data=doc_bytes,
            file_name=f"Отчет_осадки_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

        st.success("Отчёт сгенерирован!")

else:
    st.info("Загрузите Excel-файл для начала работы.")
