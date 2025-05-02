# excel_processor_app_en.py

import pandas as pd
import streamlit as st

# Список разрешенных OVACLS кодов
ALLOWED_OVACLS = [201, 202, 203, 204, 205, 206, 207, 208, 209, 210, 211, 212, 234, 295]

def process_excel(uploaded_file, pin=None):
    """
    Обрабатывает Excel-файл по новым требованиям:
    1. Фильтрует по разрешенным OVACLS
    2. Рассчитывает SCORE
    3. Группирует по OVACLS и NEIGHBORHOOD
    4. Оставляет записи выше среднего SCORE в каждой группе
    5. Сортирует по OVACLS
    """
    try:
        # Загружаем данные
        df = pd.read_excel(uploaded_file)
        
        # Проверка обязательных колонок
        required_columns = ["OVACLS", "NEIGHBORHOOD", "BUILDING_SQ_FT", "CURRENT_BUILDING"]
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            raise ValueError(f"Missing columns: {', '.join(missing_columns)}")
        
        # 1. Фильтруем только разрешенные OVACLS
        # Конвертируем OVACLS в числовой формат для корректной фильтрации
        try:
            df["OVACLS"] = pd.to_numeric(df["OVACLS"], errors='coerce')
        except Exception:
            # Продолжаем даже если конвертация не удалась полностью
            pass
        
        # Фильтруем только разрешенные OVACLS
        df = df[df["OVACLS"].isin(ALLOWED_OVACLS)]
        
        if len(df) == 0:
            raise ValueError("No records found with allowed OVACLS codes")
        
        # 2. Преобразование в числовой формат и расчет SCORE
        df["CURRENT_BUILDING"] = pd.to_numeric(df["CURRENT_BUILDING"], errors='coerce')
        df["BUILDING_SQ_FT"] = pd.to_numeric(df["BUILDING_SQ_FT"], errors='coerce')
        
        # Обработка нулевых значений в BUILDING_SQ_FT
        zero_sq_ft = df[df["BUILDING_SQ_FT"] == 0].index
        if len(zero_sq_ft) > 0:
            df.loc[zero_sq_ft, "BUILDING_SQ_FT"] = 0.0001
        
        # Расчет колонки SCORE
        df["SCORE"] = df.apply(
            lambda row: round(row["CURRENT_BUILDING"] / row["BUILDING_SQ_FT"], 2) 
            if row["BUILDING_SQ_FT"] > 0 else 0, 
            axis=1
        )
        
        # 3. Группировка по OVACLS и NEIGHBORHOOD
        df['GROUP'] = df['OVACLS'].astype(str) + '_' + df['NEIGHBORHOOD'].astype(str)
        group_avg_scores = df.groupby('GROUP')['SCORE'].mean().reset_index()
        group_avg_scores.rename(columns={'SCORE': 'AVG_SCORE'}, inplace=True)
        
        # Добавляем средний SCORE группы к каждой записи
        df = pd.merge(df, group_avg_scores, on='GROUP', how='left')
        
        # 4. Фильтруем записи выше среднего SCORE для каждой группы
        df_above_avg = df[df['SCORE'] > df['AVG_SCORE']]
        
        # 5. Сортировка по OVACLS и NEIGHBORHOOD
        df_above_avg = df_above_avg.sort_values(by=['OVACLS', 'NEIGHBORHOOD', 'SCORE'], ascending=[True, True, False])
        
        # Исключаем служебные колонки из результата
        if 'GROUP' in df_above_avg.columns:
            df_above_avg = df_above_avg.drop(columns=['GROUP'])
        
        if 'AVG_SCORE' in df_above_avg.columns:
            df_above_avg = df_above_avg.drop(columns=['AVG_SCORE'])
        
        # Убеждаемся, что SCORE находится на первом месте
        if "SCORE" in df_above_avg.columns:
            cols = ["SCORE"] + [col for col in df_above_avg.columns if col != "SCORE"]
            df_above_avg = df_above_avg[cols]
        
        return df_above_avg
        
    except Exception as e:
        raise Exception(f"Error processing Excel file: {str(e)}")