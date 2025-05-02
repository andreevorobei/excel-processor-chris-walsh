# app.py

import streamlit as st
import pandas as pd
from excel_processor_app_en import process_excel  # Импорт обновленной функции

st.set_page_config(page_title="Excel Processor", layout="centered")

st.title("🏗 Excel-фильтр данных")

st.write("""
### Инструкция по использованию:
1. Загрузите Excel-файл с данными
2. Нажмите кнопку "Обработать данные"
3. Система автоматически:
   - Отфильтрует записи по OVACLS (оставит только 201-212, 234, 295)
   - Рассчитает SCORE для всех записей
   - Сгруппирует данные по OVACLS и NEIGHBORHOOD
   - Оставит только записи, значение SCORE которых выше среднего в группе
   - Отсортирует результаты по OVACLS и NEIGHBORHOOD
4. Вы можете скачать отфильтрованные записи в формате Excel
""")

uploaded_file = st.file_uploader("Загрузите Excel-файл", type=["xlsx"])

if uploaded_file:
    if st.button("Обработать данные"):
        try:
            with st.spinner('Обработка данных...'):
                result_df = process_excel(uploaded_file)
            
            if len(result_df) > 0:
                # Группировка для отображения статистики
                groups_stats = result_df.groupby('OVACLS').agg(
                    Записей=('SCORE', 'count')
                ).reset_index()
                
                st.success(f"Найдено {len(result_df)} записей с оценкой выше средней в их группах")
                
                # Показываем статистику по группам
                st.subheader("Статистика по группам")
                st.dataframe(groups_stats)
                
                # Показываем результаты
                st.subheader("Отфильтрованные записи")
                st.dataframe(result_df)
                
                # Кнопка для скачивания
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
            else:
                st.warning("Не найдено записей, удовлетворяющих критериям фильтрации")
                
        except Exception as e:
            st.error(f"Ошибка при обработке файла: {e}")