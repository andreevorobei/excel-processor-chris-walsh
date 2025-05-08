# excel_processor_app_en.py

import pandas as pd
import streamlit as st
import io
import logging
import numpy as np

# Список разрешенных OVACLS кодов
ALLOWED_OVACLS = [201, 202, 203, 204, 205, 206, 207, 208, 209, 210, 211, 212, 234, 295]

def process_excel(uploaded_file, pin=None):
    """
    Processes Excel file according to new requirements:
    1. Filters by allowed OVACLS codes
    2. Calculates SCORE
    3. Groups by OVACLS and NEIGHBORHOOD
    4. Keeps records with SCORE above average for their group
    5. Sorts by OVACLS
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
    
    logger.info("=== File processing started ===")
    
    try:
        # Загружаем данные
        logger.info("Loading data from Excel file...")
        df = pd.read_excel(uploaded_file)
        logger.info(f"Loaded {len(df)} rows of data.")
        
        # Проверка обязательных колонок
        required_columns = ["OVACLS", "NEIGHBORHOOD", "BUILDING_SQ_FT", "CURRENT_BUILDING"]
        logger.info(f"Checking required columns: {required_columns}")
        
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            logger.error(f"Missing required columns: {', '.join(missing_columns)}")
            raise ValueError(f"Missing columns: {', '.join(missing_columns)}")
        else:
            logger.info("All required columns found.")
        
        # Выводим информацию о первых строках для проверки данных
        logger.info("First 5 rows of data:")
        for i, row in df.head(5).iterrows():
            logger.info(f"Row {i+1}: {dict(row)}")
        
        # 1. Фильтруем только разрешенные OVACLS
        # Конвертируем OVACLS в числовой формат для корректной фильтрации
        logger.info(f"Filtering by allowed OVACLS codes: {ALLOWED_OVACLS}")
        logger.info("Converting OVACLS to numeric format...")
        
        # Анализ уникальных значений OVACLS до конвертации
        unique_ovacls_before = df['OVACLS'].unique()
        logger.info(f"Unique OVACLS values before conversion: {unique_ovacls_before}")
        
        try:
            df["OVACLS"] = pd.to_numeric(df["OVACLS"], errors='coerce')
            logger.info("OVACLS conversion completed successfully.")
        except Exception as e:
            logger.warning(f"Issue with OVACLS conversion: {str(e)}. Continuing...")
        
        # Анализ уникальных значений OVACLS после конвертации
        unique_ovacls_after = df['OVACLS'].unique()
        logger.info(f"Unique OVACLS values after conversion: {unique_ovacls_after}")
        
        # Фильтруем только разрешенные OVACLS
        initial_count = len(df)
        df = df[df["OVACLS"].isin(ALLOWED_OVACLS)]
        filtered_count = len(df)
        removed_count = initial_count - filtered_count
        
        logger.info(f"Filtered by OVACLS: kept {filtered_count} out of {initial_count} records (removed {removed_count}).")
        
        if len(df) == 0:
            logger.error("No records left after filtering by allowed OVACLS codes!")
            raise ValueError("No records found with allowed OVACLS codes")
        
        # Анализ распределения по OVACLS
        ovacls_counts = df.groupby('OVACLS').size().reset_index(name='Count')
        logger.info("Distribution of filtered records by OVACLS:")
        for i, row in ovacls_counts.iterrows():
            logger.info(f"OVACLS = {row['OVACLS']}: {row['Count']} records")
        
        # 2. Преобразование в числовой формат и расчет SCORE
        logger.info("Converting CURRENT_BUILDING and BUILDING_SQ_FT to numeric format...")
        
        df["CURRENT_BUILDING"] = pd.to_numeric(df["CURRENT_BUILDING"], errors='coerce')
        df["BUILDING_SQ_FT"] = pd.to_numeric(df["BUILDING_SQ_FT"], errors='coerce')
        logger.info("Conversion completed successfully.")
        
        # Обработка нулевых значений в BUILDING_SQ_FT
        zero_sq_ft = df[df["BUILDING_SQ_FT"] == 0].index
        if len(zero_sq_ft) > 0:
            logger.warning(f"Found {len(zero_sq_ft)} rows with BUILDING_SQ_FT = 0. Replacing with 0.0001.")
            df.loc[zero_sq_ft, "BUILDING_SQ_FT"] = 0.0001
        
        # Анализ данных по OVACLS 201
        ovacls_201_records = df[df['OVACLS'] == 201]
        logger.info(f"Analysis of records with OVACLS 201 (total: {len(ovacls_201_records)} records)")
        if len(ovacls_201_records) > 0:
            logger.info(f"Average CURRENT_BUILDING: {ovacls_201_records['CURRENT_BUILDING'].mean()}")
            logger.info(f"Average BUILDING_SQ_FT: {ovacls_201_records['BUILDING_SQ_FT'].mean()}")
            logger.info(f"Number of records with CURRENT_BUILDING = 0: {len(ovacls_201_records[ovacls_201_records['CURRENT_BUILDING'] == 0])}")
            logger.info(f"Number of records with BUILDING_SQ_FT = 0 or NaN: {len(ovacls_201_records[(ovacls_201_records['BUILDING_SQ_FT'] == 0) | (ovacls_201_records['BUILDING_SQ_FT'].isna())])}")
            
            # Выборка некоторых записей для анализа
            logger.info("Examples of records with OVACLS 201:")
            for i, row in ovacls_201_records.head(min(5, len(ovacls_201_records))).iterrows():
                logger.info(f"Record {i}: CURRENT_BUILDING={row['CURRENT_BUILDING']}, BUILDING_SQ_FT={row['BUILDING_SQ_FT']}, CLASS_DECRIPTION={row.get('CLASS_DECRIPTION', 'N/A')}")
        else:
            logger.info("No records with OVACLS 201 found.")
        
        # Расчет колонки SCORE
        logger.info("Calculating SCORE for each row...")
        
        df["SCORE"] = df.apply(
            lambda row: round(row["CURRENT_BUILDING"] / row["BUILDING_SQ_FT"], 2) 
            if row["BUILDING_SQ_FT"] > 0 else 0, 
            axis=1
        )
        logger.info("SCORE calculation completed.")
        
        # Примеры расчета SCORE
        logger.info("Examples of SCORE calculation (first 5 rows):")
        for i, row in df.head(5).iterrows():
            logger.info(f"Row {i+1}: CURRENT_BUILDING={row['CURRENT_BUILDING']}, BUILDING_SQ_FT={row['BUILDING_SQ_FT']}, SCORE={row['SCORE']}")
        
        # Статистика по SCORE для разных OVACLS
        logger.info("SCORE statistics by OVACLS codes:")
        score_stats = df.groupby('OVACLS')['SCORE'].agg(['mean', 'min', 'max', 'count']).reset_index()
        for i, row in score_stats.iterrows():
            logger.info(f"OVACLS={row['OVACLS']}: mean={row['mean']:.2f}, min={row['min']:.2f}, max={row['max']:.2f}, count={row['count']}")
        
        # 3. Группировка по OVACLS и NEIGHBORHOOD
        logger.info("Grouping by OVACLS and NEIGHBORHOOD...")
        
        df['GROUP'] = df['OVACLS'].astype(str) + '_' + df['NEIGHBORHOOD'].astype(str)
        
        # Выводим список созданных групп
        unique_groups = df['GROUP'].unique()
        logger.info(f"Created {len(unique_groups)} unique groups.")
        logger.info(f"Example groups: {unique_groups[:min(5, len(unique_groups))]}")
        
        # Собираем статистику BEFORE (после группировки по OVACLS и NEIGHBORHOOD)
        # Статистика по группам OVACLS
        before_ovacls_stats = df.groupby('OVACLS').size().reset_index(name='Records_Before')
        
        # Детальная статистика по группам OVACLS и NEIGHBORHOOD
        before_detailed_stats = df.groupby(['OVACLS', 'NEIGHBORHOOD']).size().reset_index(name='Count_Before')
        
        # Расчет среднего SCORE для каждой группы
        logger.info("Calculating average SCORE for each group...")
        group_avg_scores = df.groupby('GROUP')['SCORE'].mean().reset_index()
        group_avg_scores.rename(columns={'SCORE': 'AVG_SCORE'}, inplace=True)
        
        # Выводим средние SCORE для групп
        logger.info("Average SCORE values by group:")
        for i, row in group_avg_scores.head(10).iterrows():
            logger.info(f"Group {row['GROUP']}: Average SCORE = {row['AVG_SCORE']}")
        
        # Анализ групп с нулевыми средними SCORE
        zero_avg_groups = group_avg_scores[group_avg_scores['AVG_SCORE'] == 0]
        logger.info(f"Found {len(zero_avg_groups)} groups with zero average SCORE:")
        for i, row in zero_avg_groups.iterrows():
            group_name = row['GROUP']
            ovacls_value = group_name.split('_')[0]
            neighborhood_value = group_name.split('_')[1]
            records_count = len(df[df['GROUP'] == group_name])
            logger.info(f"Group {group_name}: OVACLS={ovacls_value}, NEIGHBORHOOD={neighborhood_value}, record count={records_count}")
            
            # Дополнительный анализ записей в группах с нулевым средним
            group_records = df[df['GROUP'] == group_name]
            zero_score_count = len(group_records[group_records['SCORE'] == 0])
            logger.info(f"  - Records with SCORE=0: {zero_score_count} out of {records_count} ({zero_score_count/records_count*100:.1f}%)")
            
            if len(group_records) > 0:
                cb_sum = group_records['CURRENT_BUILDING'].sum()
                bs_sum = group_records['BUILDING_SQ_FT'].sum()
                logger.info(f"  - CURRENT_BUILDING sum: {cb_sum}, BUILDING_SQ_FT sum: {bs_sum}")
        
        # Добавляем средний SCORE группы к каждой записи
        df = pd.merge(df, group_avg_scores, on='GROUP', how='left')
        logger.info("Average SCORE for group added to each record.")
        
        # 4. Фильтруем записи выше среднего SCORE для каждой группы
        logger.info("Filtering records with SCORE above average for their group...")
        before_filter_count = len(df)
        df_above_avg = df[df['SCORE'] > df['AVG_SCORE']]
        after_filter_count = len(df_above_avg)
        removed_filter_count = before_filter_count - after_filter_count
        
        logger.info(f"Filtered by average SCORE: kept {after_filter_count} out of {before_filter_count} records (removed {removed_filter_count}).")
        
        # Статистика по группам после фильтрации
        group_counts = df_above_avg.groupby(['OVACLS', 'NEIGHBORHOOD']).size().reset_index(name='COUNT')
        logger.info("Number of records in groups after filtering:")
        for i, row in group_counts.iterrows():
            logger.info(f"OVACLS={row['OVACLS']}, NEIGHBORHOOD={row['NEIGHBORHOOD']}: {row['COUNT']} records")
        
        # Анализ отфильтрованных OVACLS
        removed_ovacls = set(df['OVACLS'].unique()) - set(df_above_avg['OVACLS'].unique())
        logger.info(f"OVACLS codes completely excluded after filtering: {removed_ovacls}")
        
        if 201 in removed_ovacls:
            logger.info("ATTENTION: OVACLS 201 is completely excluded after filtering!")
            logger.info("Reason: all records have SCORE=0, equal to the group average. No records satisfy the condition SCORE > AVG_SCORE.")
        
        # 5. Сортировка по OVACLS и NEIGHBORHOOD
        logger.info("Sorting results by OVACLS, NEIGHBORHOOD, and SCORE...")
        df_above_avg = df_above_avg.sort_values(by=['OVACLS', 'NEIGHBORHOOD', 'SCORE'], ascending=[True, True, False])
        logger.info("Sorting completed.")
        
        # ИЗМЕНЕНИЕ: Сохраняем колонки GROUP и AVG_SCORE
        logger.info("Preserving GROUP and AVG_SCORE columns in the result.")
        
        # Убеждаемся, что SCORE, GROUP и AVG_SCORE находятся в начале таблицы
        if all(col in df_above_avg.columns for col in ["SCORE", "GROUP", "AVG_SCORE"]):
            # Перемещаем важные колонки в начало
            cols = ["SCORE", "GROUP", "AVG_SCORE"] + [col for col in df_above_avg.columns if col not in ["SCORE", "GROUP", "AVG_SCORE"]]
            df_above_avg = df_above_avg[cols]
            logger.info("Columns SCORE, GROUP, and AVG_SCORE moved to the beginning of the table.")
        
        # Собираем объединенную статистику BEFORE и AFTER
        # Статистика по OVACLS
        after_ovacls_stats = df_above_avg.groupby('OVACLS').size().reset_index(name='Records_After')
        
        # Объединяем Before и After статистику по OVACLS
        combined_ovacls_stats = pd.merge(before_ovacls_stats, after_ovacls_stats, on='OVACLS', how='outer').fillna(0)
        combined_ovacls_stats[['Records_Before', 'Records_After']] = combined_ovacls_stats[['Records_Before', 'Records_After']].astype(int)
        combined_ovacls_stats = combined_ovacls_stats.sort_values('OVACLS')
        
        # Детальная статистика по OVACLS и NEIGHBORHOOD для AFTER
        after_detailed_stats = df_above_avg.groupby(['OVACLS', 'NEIGHBORHOOD']).agg({
            'SCORE': ['count', 'mean', 'min', 'max']
        })
        
        # Преобразуем многоуровневые колонки в плоские
        after_detailed_stats.columns = ['Count_After', 'Average_SCORE', 'Min_SCORE', 'Max_SCORE']
        after_detailed_stats = after_detailed_stats.reset_index()
        
        # Объединяем Before и After детальную статистику
        combined_detailed_stats = pd.merge(before_detailed_stats, after_detailed_stats, on=['OVACLS', 'NEIGHBORHOOD'], how='outer')
        combined_detailed_stats = combined_detailed_stats.fillna(0)
        combined_detailed_stats['Count_Before'] = combined_detailed_stats['Count_Before'].astype(int)
        combined_detailed_stats['Count_After'] = combined_detailed_stats['Count_After'].astype(int)
        combined_detailed_stats = combined_detailed_stats.sort_values(['OVACLS', 'NEIGHBORHOOD'])
        
        # Сохраняем объединенную статистику в сессии
        st.session_state['combined_ovacls_stats'] = combined_ovacls_stats
        st.session_state['combined_detailed_stats'] = combined_detailed_stats
        
        logger.info(f"Processing completed successfully. Final result: {len(df_above_avg)} records.")
        logger.info("=== File processing finished ===")
        
        # Сохраняем лог в сессии Streamlit для отображения
        st.session_state['process_log'] = log_output.getvalue()
        
        return df_above_avg
        
    except Exception as e:
        logger.error(f"Error processing file: {str(e)}")
        logger.info("=== File processing finished with errors ===")
        
        # Сохраняем лог в сессии Streamlit для отображения
        st.session_state['process_log'] = log_output.getvalue()
        
        raise Exception(f"Error processing Excel file: {str(e)}")