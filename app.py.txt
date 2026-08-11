import streamlit as st
import pandas as pd
import io
from datetime import datetime

# ------------------ НАСТРОЙКИ СТРАНИЦЫ ------------------
st.set_page_config(page_title="АССО Калькулятор тендеров", layout="wide")
st.title("📊 АССО - Калькулятор тендеров (223-ФЗ, делимые лоты)")
st.markdown("---")

# ------------------ ФУНКЦИИ РАСЧЕТА ------------------

def get_konvert(days):
    if days > 175:
        return 1.08
    elif days > 110:
        return 1.045
    elif days > 80:
        return 1.03
    elif days > 44:
        return 1.02
    elif days > 29:
        return 1.01
    else:
        return 1.0

def calculate_transport(df, total_transport):
    if df.empty or total_transport <= 0:
        return df
    
    if 'Вес' in df.columns and 'Объем' in df.columns:
        df['Транспортный_коэф'] = (df['Вес'] / df['Вес'].sum()) * 0.7 + (df['Объем'] / df['Объем'].sum()) * 0.3
    else:
        df['Транспортный_коэф'] = 1 / len(df)
    
    min_transport_per_item = 1500
    transport_list = []
    for idx, row in df.iterrows():
        base = total_transport * row['Транспортный_коэф']
        if len(df) > 10:
            base = max(base, min_transport_per_item * 0.3)
        else:
            base = max(base, min_transport_per_item)
        transport_list.append(round(base, 2))
    
    df['Транспорт_на_позицию_с_НДС'] = transport_list
    df['Транспорт_на_позицию_без_НДС'] = round(df['Транспорт_на_позицию_с_НДС'] / 1.22, 2)
    return df

def calculate_prices(df, course_usd, course_eur, days, nds_rate, total_transport):
    if df.empty:
        return df
    
    df['Валюта'] = df.apply(lambda row: 'USD' if 'USD' in str(row.get('Валюта', '')) else 'EUR', axis=1)
    konvert = get_konvert(days)
    
    df['Вход_руб'] = df.apply(
        lambda row: round(row['Цена_вход'] * (course_usd if row['Валюта'] == 'USD' else course_eur) * konvert, 2),
        axis=1
    )
    
    df = calculate_transport(df, total_transport)
    
    df['Выход_без_НДС'] = df.apply(
        lambda row: round((row['Вход_руб'] * row['Наценка']) + row['Транспорт_на_позицию_без_НДС'], 2),
        axis=1
    )
    df['Выход_с_НДС'] = round(df['Выход_без_НДС'] * (1 + nds_rate / 100), 2)
    df['Сумма_с_НДС'] = round(df['Выход_с_НДС'] * df['Кол-во'], 2)
    
    if 'НМЦК_за_ед' in df.columns:
        df['Проверка'] = df.apply(
            lambda row: '✅ OK' if row['Выход_с_НДС'] <= row['НМЦК_за_ед'] else '⚠️ ПРЕВЫШЕНИЕ НМЦК!',
            axis=1
        )
    else:
        df['Проверка'] = '✅ НМЦК не задана'
    
    df['Прибыль_руб'] = round(df['Сумма_с_НДС'] - (df['Вход_руб'] * df['Кол-во'] * (1 + nds_rate/100)), 2)
    df['Маржа_%'] = round((df['Прибыль_руб'] / df['Сумма_с_НДС']) * 100, 2)
    
    return df

# ------------------ БОКОВАЯ ПАНЕЛЬ ------------------
with st.sidebar:
    st.header("⚙️ Параметры расчета")
    course_usd = st.number_input("Курс USD (₽)", value=82.0, step=0.5)
    course_eur = st.number_input("Курс EUR (₽)", value=95.0, step=0.5)
    days = st.number_input("Срок оферты (дней)", value=30, step=1)
    st.markdown("---")
    total_transport = st.number_input("Общие транспортные расходы (₽, с НДС)", value=20000, step=1000)
    nds_rate = st.selectbox("Ставка НДС (%)", [0, 5, 7, 22], index=3)
    st.markdown("---")
    st.caption(f"Конверт при сроке {days} дней: **{get_konvert(days)}**")

# ------------------ ЗАГРУЗКА ДАННЫХ ------------------
st.subheader("📂 Шаг 1. Загрузите файл анализа АССО.xlsx")
uploaded_file = st.file_uploader("Выберите файл Excel", type=['xlsx', 'xls'])

if uploaded_file:
    try:
        df_raw = pd.read_excel(uploaded_file, sheet_name='АНАЛИЗ ЦЕН', header=8)
        
        required_cols = {
            '№ п/п': '№',
            'Наименование позиций заказчика': 'Наименование',
            'Кол-во': 'Кол-во',
            'Цена за единицу РАБОТЫ с НДС, вход': 'Цена_вход',
            'НМЦК за ед.': 'НМЦК_за_ед',
            'Вес позиции, строки': 'Вес',
            'Объем': 'Объем',
            'Аналог': 'Аналог'
        }
        
        df_clean = pd.DataFrame()
        for old, new in required_cols.items():
            if old in df_raw.columns:
                df_clean[new] = df_raw[old]
            else:
                df_clean[new] = None
        
        df_clean = df_clean.dropna(subset=['Наименование', 'Кол-во', 'Цена_вход'])
        
        if df_clean.empty:
            st.error("❌ Не удалось найти данные в файле. Проверьте структуру.")
        else:
            st.success(f"✅ Загружено {len(df_clean)} позиций")
            
            with st.expander("📋 Предпросмотр загруженных данных"):
                st.dataframe(df_clean)
            
            st.subheader("🔢 Шаг 2. Введите наценки для каждой позиции")
            st.info("Наценки зависят от бренда, конкурентов, глубины скидки поставщика.")
            
            df_edit = df_clean.copy()
            if 'Наценка' not in df_edit.columns:
                df_edit['Наценка'] = 1.1
            
            edited_df = st.data_editor(
                df_edit[['№', 'Наименование', 'Кол-во', 'Цена_вход', 'Наценка', 'НМЦК_за_ед']],
                num_rows="dynamic",
                column_config={
                    "Наценка": st.column_config.NumberColumn(
                        "Наценка (коэф.)",
                        help="1.1 = 10%, 1.3 = 30%",
                        min_value=0.0,
                        max_value=5.0,
                        step=0.01,
                        format="%.2f"
                    )
                },
                use_container_width=True
            )
            
            if st.button("🚀 Рассчитать цены", type="primary"):
                with st.spinner("Выполняется расчет..."):
                    result_df = calculate_prices(
                        edited_df, course_usd, course_eur, days, nds_rate, total_transport
                    )
                    
                    st.subheader("📊 Шаг 3. Результат расчета")
                    
                    output_cols = [
                        '№', 'Наименование', 'Кол-во', 'Цена_вход', 'Валюта',
                        'Наценка', 'Вход_руб', 'Транспорт_на_позицию_без_НДС',
                        'Выход_без_НДС', 'Выход_с_НДС', 'Сумма_с_НДС',
                        'НМЦК_за_ед', 'Проверка', 'Прибыль_руб', 'Маржа_%'
                    ]
                    display_df = result_df[output_cols].copy()
                    
                    def highlight_warnings(val):
                        if isinstance(val, str) and 'ПРЕВЫШЕНИЕ' in val:
                            return 'background-color: #FF6B6B; color: white'
                        return ''
                    
                    st.dataframe(
                        display_df.style.applymap(highlight_warnings, subset=['Проверка']),
                        use_container_width=True,
                        height=400
                    )
                    
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Общая сумма (с НДС)", f"{result_df['Сумма_с_НДС'].sum():,.2f} ₽")
                    with col2:
                        st.metric("Количество позиций", len(result_df))
                    with col3:
                        st.metric("Средняя маржа", f"{result_df['Маржа_%'].mean():.1f}%")
                    with col4:
                        errors = result_df[result_df['Проверка'].str.contains('ПРЕВЫШЕНИЕ', na=False)]
                        st.metric("Ошибок по НМЦК", len(errors), delta_color="off")
                    
                    if len(errors) > 0:
                        st.error(f"⚠️ Обнаружено {len(errors)} позиций с превышением НМЦК!")
                    
                    st.subheader("💾 Шаг 4. Скачать результат")
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        display_df.to_excel(writer, sheet_name='Результат', index=False)
                        
                        summary = pd.DataFrame({
                            'Параметр': ['Общая сумма с НДС', 'Кол-во позиций', 'Средняя маржа', 'Ошибок по НМЦК'],
                            'Значение': [
                                result_df['Сумма_с_НДС'].sum(),
                                len(result_df),
                                f"{result_df['Маржа_%'].mean():.1f}%",
                                len(errors)
                            ]
                        })
                        summary.to_excel(writer, sheet_name='Сводка', index=False)
                    
                    st.download_button(
                        label="📥 Скачать Excel с результатами",
                        data=output.getvalue(),
                        file_name=f"АССО_расчет_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.success("✅ Расчет выполнен! Файл готов для передачи в тендерный отдел.")
                    
    except Exception as e:
        st.error(f"❌ Ошибка при обработке файла: {str(e)}")
        st.info("Убедитесь, что файл имеет структуру, как в вашем примере (лист 'АНАЛИЗ ЦЕН', заголовки в 9-й строке).")
else:
    st.info("👆 Загрузите ваш файл АНАЛИЗ АССО.xlsx, чтобы начать работу")

st.markdown("---")
st.caption("📌 Версия 1.0. Разработано для отдела продаж.")