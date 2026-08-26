import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta

st.set_page_config(page_title="Генератор отчетов Статусы", layout="centered")
st.title("📊 Генератор отчета «Статусы»")
st.write("Загрузите файл «Выгрузка.xlsx» и получите отчет «Статусы»")

uploaded_file = st.file_uploader("Выберите файл выгрузки (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    # Читаем данные — пропускаем первую строку с датой
    df = pd.read_excel(uploaded_file, header=1)
    
    with st.expander("👀 Превью загруженных данных"):
        st.dataframe(df.head(10))
    
    required_cols = ["Ответственный", "Стадия"]
    if not all(col in df.columns for col in required_cols):
        st.error(f"❌ В файле нет нужных колонок. Найдены: {list(df.columns)}")
        st.stop()
    
    pivot = pd.crosstab(df["Стадия"], df["Ответственный"], margins=True, margins_name="Итого")
    
    if "Дата изменения" in df.columns:
        df["Дата изменения"] = pd.to_datetime(df["Дата изменения"], errors="coerce")
        today = datetime.now()
        old_leads = df[(today - df["Дата изменения"]) > timedelta(days=10)]
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