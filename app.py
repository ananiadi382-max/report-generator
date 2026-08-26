import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta

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
    
    # --- ФИКС №1: Объединяем статусы ---
    # Заменяем "Аккаунты_Менеджер назначен" на "Менеджер назначен"
    df_clean["Стадия"] = df_clean["Стадия"].replace(
        "Аккаунты_Менеджер назначен", 
        "Менеджер назначен"
    )
    
    # --- ФИКС №2: Правильный расчет "Без изменений >10 дней" ---
    # Подсчет по статусам и менеджерам (БЕЗ учета даты)
    pivot = pd.crosstab(df_clean["Стадия"], df_clean["Ответственный"], margins=True, margins_name="Итого")
    
    # Добавляем строку "Без изменений >10 дней" (только если есть дата)
    if date_col is not None:
        # Преобразуем даты
        df_clean["Дата изменения"] = pd.to_datetime(df_clean["Дата изменения"], errors="coerce")
        today = datetime.now()
        
        # Берем ТОЛЬКО те строки, где дата есть и она старше 10 дней
        old_leads = df_clean[
            (df_clean["Дата изменения"].notna()) & 
            ((today - df_clean["Дата изменения"]) > timedelta(days=10))
        ]
        
        old_count = old_leads.groupby("Ответственный").size()
        pivot.loc["Без изменений >10 дней"] = old_count.reindex(pivot.columns[:-1], fill_value=0)
        pivot.loc["Без изменений >10 дней", "Итого"] = old_count.sum()
    
    pivot.index.name = "Статус лида"
    
    st.subheader("📋 Сгенерированный отчет")
    st.dataframe(pivot, use_container_width=True)
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pivot.to_excel(writer, sheet_name="Статусы")
    output.seek(0)
    
    st.download_button(
        label="⬇️ Скачать отчет Статусы.xlsx",
        data=output,
        file_name="Статусы.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.success("✅ Готово! Нажмите кнопку выше, чтобы скачать файл.")
