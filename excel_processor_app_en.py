import streamlit as st
import pandas as pd
import os
from pathlib import Path
import traceback
import io

# Настраиваем страницу
st.set_page_config(page_title="Excel Processor", layout="wide")

# Заголовок приложения
st.title("Excel Processor for Chris Walsh Law")

# Функция для загрузки Excel файла
def load_excel_file(uploaded_file):
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            st.session_state['df'] = df
            st.session_state['file_name'] = uploaded_file.name
            
            # Проверка обязательных колонок
            required_columns = ["PIN", "OVACLS", "NEIGHBORHOOD", "BUILDING_SQ_FT", "CURRENT_BUILDING"]
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                st.error(f"ERROR: Missing the following columns: {', '.join(missing_columns)}")
                return False
            else:
                st.success(f"File '{uploaded_file.name}' successfully loaded. Number of rows: {len(df)}")
                return True
        except Exception as e:
            st.error(f"ERROR reading file: {str(e)}")
            st.error(traceback.format_exc())
            return False
    return False

# Функция для вычисления SCORE и фильтрации данных
def process_data(pin, filter_percent):
    try:
        if 'df' not in st.session_state:
            st.error("Please upload an Excel file first.")
            return False
        
        df = st.session_state['df']
        
        # Поиск строки по PIN
        # Проверяем различные форматы PIN
        main_row = df[df["PIN"] == pin]
        
        if len(main_row) == 0 and pin.isdigit():
            main_row = df[df["PIN"] == int(pin)]
        
        if len(main_row) == 0:
            main_row = df[df["PIN"].astype(str) == str(pin)]
        
        if len(main_row) == 0:
            st.error(f"PIN '{pin}' not found in the data.")
            return False
        
        if len(main_row) > 1:
            st.warning(f"Multiple rows found with PIN '{pin}'. The first one will be used.")
        
        main_row = main_row.iloc[0]
        
        # Расчет SCORE для каждой строки
        st.info("Calculating SCORE for each row...")
        
        # Проверка обязательных колонок
        required_columns = ["CURRENT_BUILDING", "BUILDING_SQ_FT"]
        for col in required_columns:
            if col not in df.columns:
                st.error(f"Missing required column: {col}")
                return False
        
        # Обработка нулевых значений в BUILDING_SQ_FT
        zero_sq_ft = df[df["BUILDING_SQ_FT"] == 0].index
        if len(zero_sq_ft) > 0:
            st.warning(f"Found {len(zero_sq_ft)} rows with BUILDING_SQ_FT = 0. Using small value instead.")
            df.loc[zero_sq_ft, "BUILDING_SQ_FT"] = 0.0001
        
        # Преобразование в числовой формат
        df["CURRENT_BUILDING"] = pd.to_numeric(df["CURRENT_BUILDING"], errors='coerce')
        df["BUILDING_SQ_FT"] = pd.to_numeric(df["BUILDING_SQ_FT"], errors='coerce')
        
        # Расчет колонки SCORE
        df["SCORE"] = df.apply(
            lambda row: round(row["CURRENT_BUILDING"] / row["BUILDING_SQ_FT"], 2) 
            if row["BUILDING_SQ_FT"] > 0 else 0, 
            axis=1
        )
        
        # Получение значения SCORE для основной строки
        main_score = df.loc[main_row.name, "SCORE"]
        
        # Получение критериев фильтрации
        main_building_sq_ft = main_row["BUILDING_SQ_FT"]
        main_ovacls = main_row["OVACLS"]
        main_neighborhood = main_row["NEIGHBORHOOD"]
        
        # Вычисление диапазона для фильтрации
        percent_factor = filter_percent / 100
        min_sq_ft = main_building_sq_ft * (1 - percent_factor)
        max_sq_ft = main_building_sq_ft * (1 + percent_factor)
        
        st.info(f"SCORE for the main row (PIN={pin}): {main_score}")
        st.info(f"Using filter percent: ±{filter_percent}%")
        st.info(f"BUILDING_SQ_FT range for filtering: {min_sq_ft:.2f} - {max_sq_ft:.2f}")
        
        # Применение фильтров
        filtered_df = df[
            (df["BUILDING_SQ_FT"] >= min_sq_ft) & 
            (df["BUILDING_SQ_FT"] <= max_sq_ft) &
            (df["SCORE"] < main_score) &
            (df["OVACLS"] == main_ovacls) &
            (df["NEIGHBORHOOD"] == main_neighborhood)
        ]
        
        # Создание строки для основного PIN
        main_row_dict = {"SCORE": main_score}
        for col in main_row.index:
            if col != "SCORE":
                main_row_dict[col] = main_row[col]
                
        main_row_df = pd.DataFrame([main_row_dict])
        
        # Убеждаемся, что порядок колонок совпадает
        if len(filtered_df) > 0:
            ordered_cols = list(filtered_df.columns)
            for col in ordered_cols:
                if col not in main_row_df.columns:
                    main_row_df[col] = None
            main_row_df = main_row_df[ordered_cols]
        
        # Объединяем данные
        result_df = pd.concat([main_row_df, filtered_df], ignore_index=True)
        
        # Убеждаемся, что SCORE находится на первом месте
        if "SCORE" in result_df.columns:
            cols = ["SCORE"] + [col for col in result_df.columns if col != "SCORE"]
            result_df = result_df[cols]
        
        st.session_state['result_df'] = result_df
        st.success(f"Filtering completed. Found {len(filtered_df)} matching records.")
        return True
        
    except Exception as e:
        st.error(f"ERROR during processing: {str(e)}")
        st.error(traceback.format_exc())
        return False

# Основной интерфейс
st.sidebar.header("File Selection")
uploaded_file = st.sidebar.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

# Если файл загружен, отображаем его и форму ввода
if uploaded_file is not None:
    if 'df' not in st.session_state or st.sidebar.button("Reload File"):
        load_excel_file(uploaded_file)
    
    if 'df' in st.session_state:
        # Создаем боковую панель для ввода параметров
        st.sidebar.header("Input Parameters")
        pin = st.sidebar.text_input("Enter successful case PIN")
        filter_percent = st.sidebar.number_input("Enter percent to filter (±)", min_value=1.0, max_value=100.0, value=15.0, step=1.0)
        
        # Кнопка для расчета и фильтрации
        if st.sidebar.button("Score and filter"):
            if process_data(pin, filter_percent):
                st.subheader("Filtered Data")
                result_df = st.session_state['result_df']
                
                # Отображаем результаты с выделением первой строки
                st.dataframe(result_df.style.apply(
                    lambda x: ['background-color: yellow' if i==0 else '' for i in range(len(x))], axis=0
                ))
                
                # Создаем кнопку для экспорта с прямым скачиванием
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False)
                output.seek(0)
                
                file_name = Path(st.session_state['file_name']).stem
                st.download_button(
                    label="Export filtered data",
                    data=output,
                    file_name=f"{file_name}_processed.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
        # Отображаем исходные данные
        if 'df' in st.session_state:
            st.subheader("Original Data")
            st.dataframe(st.session_state['df'].head(1000))  # Ограничиваем для производительности

# Информация о статусе внизу страницы
st.sidebar.text("")
st.sidebar.text("")
st.sidebar.text("App Status: Ready")