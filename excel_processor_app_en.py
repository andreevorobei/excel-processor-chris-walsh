# excel_processor_app_en.py

import pandas as pd
import streamlit as st
import io
import logging

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
    # Настраиваем логгер для вывода в StringIO (чтобы потом показать в Streamlit)
    log_output = io.StringIO()
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
        stream=log_output
    )
    
    logger = logging.getLogger("excel_processor")
    
    logger.info("=== Начало обработки файла ===")
    
    try:
        # Загружаем данные
        logger.info("Загрузка данных из Excel-файла...")
        df = pd.read_excel(uploaded_file)
        logger.info(f"Загружено {len(df)} строк данных.")
        
        # Проверка обязательных колонок
        required_columns = ["OVACLS", "NEIGHBORHOOD", "BUILDING_SQ_FT", "CURRENT_BUILDING"]
        logger.info(f"Проверка наличия обязательных колонок: {required_columns}")
        
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            logger.error(f"Отсутствуют обязательные колонки: {', '.join(missing_columns)}")
            raise ValueError(f"Missing columns: {', '.join(missing_columns)}")
        else:
            logger.info("Все обязательные колонки найдены.")
        
        # Выводим информацию о первых строках для проверки данных
        logger.info("Первые 5 строк данных:")
        for i, row in df.head(5).iterrows():
            logger.info(f"Строка {i+1}: {dict(row)}")
        
        # 1. Фильтруем только разрешенные OVACLS
        # Конвертируем OVACLS в числовой формат для корректной фильтрации
        logger.info(f"Фильтрация по разрешенным OVACLS: {ALLOWED_OVACLS}")
        logger.info("Конвертация OVACLS в числовой формат...")
        
        # Анализ уникальных значений OVACLS до конвертации
        unique_ovacls_before = df['OVACLS'].unique()
        logger.info(f"Уникальные значения OVACLS до конвертации: {unique_ovacls_before}")
        
        try:
            df["OVACLS"] = pd.to_numeric(df["OVACLS"], errors='coerce')
            logger.info("Конвертация OVACLS выполнена успешно.")
        except Exception as e:
            logger.warning(f"Проблема при конвертации OVACLS: {str(e)}. Продолжаем работу...")
        
        # Анализ уникальных значений OVACLS после конвертации
        unique_ovacls_after = df['OVACLS'].unique()
        logger.info(f"Уникальные значения OVACLS после конвертации: {unique_ovacls_after}")
        
        # Фильтруем только разрешенные OVACLS
        initial_count = len(df)
        df = df[df["OVACLS"].isin(ALLOWED_OVACLS)]
        filtered_count = len(df)
        removed_count = initial_count - filtered_count
        
        logger.info(f"Отфильтровано по OVACLS: оставлено {filtered_count} из {initial_count} записей (удалено {removed_count}).")
        
        if len(df) == 0:
            logger.error("После фильтрации не осталось записей с разрешенными OVACLS!")
            raise ValueError("No records found with allowed OVACLS codes")
        
        # 2. Преобразование в числовой формат и расчет SCORE
        logger.info("Конвертация CURRENT_BUILDING и BUILDING_SQ_FT в числовой формат...")
        
        df["CURRENT_BUILDING"] = pd.to_numeric(df["CURRENT_BUILDING"], errors='coerce')
        df["BUILDING_SQ_FT"] = pd.to_numeric(df["BUILDING_SQ_FT"], errors='coerce')
        logger.info("Конвертация выполнена успешно.")
        
        # Обработка нулевых значений в BUILDING_SQ_FT
        zero_sq_ft = df[df["BUILDING_SQ_FT"] == 0].index
        if len(zero_sq_ft) > 0:
            logger.warning(f"Найдено {len(zero_sq_ft)} строк с BUILDING_SQ_FT = 0. Заменяем на 0.0001.")
            df.loc[zero_sq_ft, "BUILDING_SQ_FT"] = 0.0001
        
        # Расчет колонки SCORE
        logger.info("Расчет SCORE для каждой строки...")
        
        df["SCORE"] = df.apply(
            lambda row: round(row["CURRENT_BUILDING"] / row["BUILDING_SQ_FT"], 2) 
            if row["BUILDING_SQ_FT"] > 0 else 0, 
            axis=1
        )
        logger.info("Расчет SCORE завершен.")
        
        # Примеры расчета SCORE
        logger.info("Примеры расчета SCORE (первые 5 строк):")
        for i, row in df.head(5).iterrows():
            logger.info(f"Строка {i+1}: CURRENT_BUILDING={row['CURRENT_BUILDING']}, BUILDING_SQ_FT={row['BUILDING_SQ_FT']}, SCORE={row['SCORE']}")
        
        # 3. Группировка по OVACLS и NEIGHBORHOOD
        logger.info("Группировка по OVACLS и NEIGHBORHOOD...")
        
        df['GROUP'] = df['OVACLS'].astype(str) + '_' + df['NEIGHBORHOOD'].astype(str)
        
        # Выводим список созданных групп
        unique_groups = df['GROUP'].unique()
        logger.info(f"Создано {len(unique_groups)} уникальных групп.")
        logger.info(f"Примеры групп: {unique_groups[:min(5, len(unique_groups))]}")
        
        # Расчет среднего SCORE для каждой группы
        logger.info("Расчет среднего SCORE для каждой группы...")
        group_avg_scores = df.groupby('GROUP')['SCORE'].mean().reset_index()
        group_avg_scores.rename(columns={'SCORE': 'AVG_SCORE'}, inplace=True)
        
        # Выводим средние SCORE для групп
        logger.info("Средние значения SCORE по группам:")
        for i, row in group_avg_scores.head(10).iterrows():
            logger.info(f"Группа {row['GROUP']}: Средний SCORE = {row['AVG_SCORE']}")
        
        # Добавляем средний SCORE группы к каждой записи
        df = pd.merge(df, group_avg_scores, on='GROUP', how='left')
        logger.info("Добавлен средний SCORE группы к каждой записи.")
        
        # 4. Фильтруем записи выше среднего SCORE для каждой группы
        logger.info("Фильтрация записей выше среднего SCORE для каждой группы...")
        before_filter_count = len(df)
        df_above_avg = df[df['SCORE'] > df['AVG_SCORE']]
        after_filter_count = len(df_above_avg)
        removed_filter_count = before_filter_count - after_filter_count
        
        logger.info(f"Фильтрация по среднему SCORE: оставлено {after_filter_count} из {before_filter_count} записей (удалено {removed_filter_count}).")
        
        # Статистика по группам после фильтрации
        group_counts = df_above_avg.groupby(['OVACLS', 'NEIGHBORHOOD']).size().reset_index(name='COUNT')
        logger.info("Количество записей в группах после фильтрации:")
        for i, row in group_counts.iterrows():
            logger.info(f"OVACLS={row['OVACLS']}, NEIGHBORHOOD={row['NEIGHBORHOOD']}: {row['COUNT']} записей")
        
        # 5. Сортировка по OVACLS и NEIGHBORHOOD
        logger.info("Сортировка результатов по OVACLS, NEIGHBORHOOD и SCORE...")
        df_above_avg = df_above_avg.sort_values(by=['OVACLS', 'NEIGHBORHOOD', 'SCORE'], ascending=[True, True, False])
        logger.info("Сортировка завершена.")
        
        # Исключаем служебные колонки из результата
        if 'GROUP' in df_above_avg.columns:
            df_above_avg = df_above_avg.drop(columns=['GROUP'])
            logger.info("Удалена служебная колонка GROUP.")
        
        if 'AVG_SCORE' in df_above_avg.columns:
            df_above_avg = df_above_avg.drop(columns=['AVG_SCORE'])
            logger.info("Удалена служебная колонка AVG_SCORE.")
        
        # Убеждаемся, что SCORE находится на первом месте
        if "SCORE" in df_above_avg.columns:
            cols = ["SCORE"] + [col for col in df_above_avg.columns if col != "SCORE"]
            df_above_avg = df_above_avg[cols]
            logger.info("Колонка SCORE перемещена на первое место.")
        
        logger.info(f"Обработка завершена успешно. Итоговый результат: {len(df_above_avg)} записей.")
        logger.info("=== Конец обработки файла ===")
        
        # Сохраняем лог в сессии Streamlit для отображения
        st.session_state['process_log'] = log_output.getvalue()
        
        return df_above_avg
        
    except Exception as e:
        logger.error(f"Ошибка при обработке файла: {str(e)}")
        logger.info("=== Обработка файла завершена с ошибкой ===")
        
        # Сохраняем лог в сессии Streamlit для отображения
        st.session_state['process_log'] = log_output.getvalue()
        
        raise Exception(f"Error processing Excel file: {str(e)}")