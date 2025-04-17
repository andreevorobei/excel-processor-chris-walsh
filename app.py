# app.py

import streamlit as st
import pandas as pd
from excel_processor_app_en import process_excel  # Импорт твоей функции

st.set_page_config(page_title="Excel Processor", layout="centered")

st.title("🏗 Excel-фильтр по PIN")

uploaded_file = st.file_uploader("Загрузите Excel-файл", type=["xlsx"])
pin_input = st.text_input("Введите PIN успешного кейса")

if uploaded_file and pin_input:
    try:
        result_df = process_excel(uploaded_file, pin_input)
        st.success(f"Найдено совпадений: {len(result_df)}")
        st.dataframe(result_df)

        @st.cache_data
        def convert_df_to_excel(df):
            from io import BytesIO
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        excel_bytes = convert_df_to_excel(result_df)

        st.download_button(
            label="📥 Скачать результат в Excel",
            data=excel_bytes,
            file_name="filtered_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Ошибка: {e}")
