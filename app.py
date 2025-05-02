# app.py

import streamlit as st
import pandas as pd
from excel_processor_app_en import process_excel  # Импорт обновленной функции

st.set_page_config(page_title="Excel Processor", layout="wide")

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

# Создаем две колонки для верхней части интерфейса
col1, col2 = st.columns([2, 1])

with col1:
    uploaded_file = st.file_uploader("Загрузите Excel-файл", type=["xlsx"])

with col2:
    if uploaded_file:
        if st.button("Обработать данные", use_container_width=True):
            try:
                with st.spinner('Обработка данных...'):
                    result_df = process_excel(uploaded_file)
                
                if len(result_df) > 0:
                    # Группировка для отображения статистики
                    groups_stats = result_df.groupby('OVACLS').agg(
                        Записей=('SCORE', 'count')
                    ).reset_index()
                    
                    st.success(f"Найдено {len(result_df)} записей с оценкой выше средней в их группах")
                    
                    # Создаем вкладки для разных типов данных
                    tab1, tab2, tab3 = st.tabs(["Результаты", "Статистика", "Подробный лог"])
                    
                    with tab1:
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
                    
                    with tab2:
                        # Показываем статистику по группам
                        st.subheader("Статистика по группам OVACLS")
                        st.dataframe(groups_stats)
                        
                        # Показываем более детальную статистику по группам
                        st.subheader("Детальная статистика по группам")
                        detailed_stats = result_df.groupby(['OVACLS', 'NEIGHBORHOOD']).agg(
                            Количество=('SCORE', 'count'),
                            Среднее_SCORE=('SCORE', 'mean'),
                            Мин_SCORE=('SCORE', 'min'),
                            Макс_SCORE=('SCORE', 'max')
                        ).reset_index()
                        st.dataframe(detailed_stats)
                    
                    with tab3:
                        # Показываем лог обработки
                        st.subheader("Подробный лог обработки данных")
                        if 'process_log' in st.session_state:
                            st.text_area("Лог обработки:", st.session_state['process_log'], height=600)
                        else:
                            st.info("Лог обработки недоступен")
                else:
                    st.warning("Не найдено записей, удовлетворяющих критериям фильтрации")
                    
                    # Показываем лог в любом случае, чтобы понять причину
                    if 'process_log' in st.session_state:
                        st.subheader("Подробный лог обработки данных")
                        st.text_area("Лог обработки:", st.session_state['process_log'], height=400)
                    
            except Exception as e:
                st.error(f"Ошибка при обработке файла: {e}")
                
                # Показываем лог в случае ошибки для диагностики
                if 'process_log' in st.session_state:
                    st.subheader("Лог обработки (содержит ошибки)")
                    st.text_area("Лог обработки:", st.session_state['process_log'], height=400)

# Если файл загружен, показываем исходные данные внизу
if uploaded_file and 'df' not in st.session_state:
    try:
        # Просто читаем данные для отображения, без обработки
        df = pd.read_excel(uploaded_file)
        st.session_state['df'] = df
    except Exception as e:
        st.error(f"Ошибка при чтении файла: {e}")

if 'df' in st.session_state:
    with st.expander("Исходные данные"):
        st.dataframe(st.session_state['df'].head(1000))  # Ограничиваем для производительности