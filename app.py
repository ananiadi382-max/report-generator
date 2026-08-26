import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta
from openpyxl.styles import PatternFill, Font, Alignment

st.set_page_config(page_title="Генератор отчетов Статусы", layout="centered")
st.title("📊 Генератор отчета «Статусы»")
st.write("Загрузите файл «Выгрузка.xlsx» и получите отчет «Статусы»")

uploaded_file = st.file_uploader("Выберите файл выгрузки (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    # Читаем данные
    df = pd.read_excel(uploaded_file, header=0)
    
    with st.expander("👀 Превью загруженных данных"):
        st.dataframe(df.head(10))
    
    # Находим нужные колонки
    status_col = None
    responsible_col = None
    date_col = None
    
    for col in df.columns:
        col_str = str(col).strip()
        if "Стадия" in col_str:
            status_col = col
        if "Ответственный" in col_str:
            responsible_col = col
        if "Дата изменения" in col_str:
            date_col = col
    
    if status_col is None or responsible_col is None:
        st.error(f"❌ Не найдены нужные колонки. Найдены: {list(df.columns)}")
        st.stop()
    
    # Создаем копию с нужными колонками
    df_clean = df[[status_col, responsible_col]].copy()
    if date_col is not None:
        df_clean["Дата изменения"] = df[date_col]
    
    df_clean.columns = ["Стадия", "Ответственный"] + (["Дата изменения"] if date_col is not None else [])
    
    # Объединяем статусы
    df_clean["Стадия"] = df_clean["Стадия"].replace(
        {
            "Аккаунты_Менеджер назначен": "Менеджер назначен",
            "Аккаунты_Менеджер назначен ": "Менеджер назначен",
        }
    )
    
    # Строим сводную таблицу
    pivot = pd.crosstab(df_clean["Стадия"], df_clean["Ответственный"])
    
    # Добавляем строку "Без изменений >10 дней"
    if date_col is not None:
        df_clean["Дата изменения"] = pd.to_datetime(df_clean["Дата изменения"], errors="coerce")
        today = datetime.now()
        
        old_leads = df_clean[
            (df_clean["Дата изменения"].notna()) & 
            ((today - df_clean["Дата изменения"]) > timedelta(days=10))
        ]
        
        old_count = old_leads.groupby("Ответственный").size()
        pivot.loc["Без изменений >10 дней"] = old_count.reindex(pivot.columns, fill_value=0)
    
    # Добавляем столбец "Итого"
    pivot["Итого"] = pivot.sum(axis=1)
    
    # Добавляем строку "Итого"
    total_row = pivot.sum(axis=0)
    pivot.loc["Итого"] = total_row
    
    pivot.index.name = "Статус лида"
    
    st.subheader("📋 Сгенерированный отчет")
    st.dataframe(pivot, use_container_width=True)
    
    # --- СОЗДАНИЕ EXCEL С ФОРМАТИРОВАНИЕМ ---
    output = BytesIO()
    
    temp_output = BytesIO()
    with pd.ExcelWriter(temp_output, engine="openpyxl") as writer:
        pivot.to_excel(writer, sheet_name="Статусы")
    
    from openpyxl import load_workbook
    
    temp_output.seek(0)
    wb = load_workbook(temp_output)
    ws = wb["Статусы"]
    
    # Ширина колонок
    ws.column_dimensions["A"].width = 37
    for col in ["B", "C", "D", "E", "F"]:
        ws.column_dimensions[col].width = 15
    
    # Цвета
    color_highlight = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    color_total = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
    bold_font = Font(bold=True)
    center_alignment = Alignment(horizontal="center", vertical="center")
    
    # Применяем форматирование
    for row in range(2, ws.max_row + 1):
        cell_value = ws.cell(row=row, column=1).value
        
        if cell_value == "Итого":
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=row, column=col)
                cell.fill = color_total
                cell.font = bold_font
        
        elif cell_value in ["Менеджер назначен", "Пора звонить", "Пора звонить (холодняк)"]:
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col).fill = color_highlight
    
    # Выравнивание по центру
    for row in range(1, ws.max_row + 1):
        for col in range(2, ws.max_column + 1):
            ws.cell(row=row, column=col).alignment = center_alignment
    
    wb.save(output)
    output.seek(0)
    
    st.download_button(
        label="⬇️ Скачать отчет Статусы.xlsx",
        data=output,
        file_name="Статусы.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.success("✅ Готово! Нажмите кнопку выше, чтобы скачать файл.")
