import streamlit as st
import pandas as pd
import numpy as np

# ------------------------------------------------------------
# 1. Парсер Excel-файла (для строго фиксированной структуры)
# ------------------------------------------------------------
def parse_excel(file):
    """
    Читает файл с жёсткой структурой:
      - Строка 0: названия городов (повторяются блоки по 10 колонок)
      - Строка 1: названия полей внутри каждого блока
      - Строка 2 и далее: данные (номенклатура, характеристика, затем значения для каждого города)
    Возвращает DataFrame с колонками:
      ['Номенклатура', 'Характеристика', 'Город', 'Класс', 'Количество', 'Остаток',
       'Свободный', 'В пути', 'Итоговый', 'Ср_продажа', 'Себестоимость_продажи',
       'Остаток_по_себест', 'Расчет_к_заказу']
    """
    df_raw = pd.read_excel(file, header=None)
    
    # Строки заголовков
    city_row = df_raw.iloc[0]
    field_row = df_raw.iloc[1]
    data = df_raw.iloc[2:].reset_index(drop=True)
    
    # Находим индексы, где начинаются города (непустые ячейки в city_row)
    city_indices = []
    for i in range(2, len(city_row)):
        val = city_row[i]
        if pd.notna(val) and str(val).strip() != '':
            city_indices.append(i)
    
    # Ожидаемые названия полей (в том порядке, как они идут в блоке)
    expected_fields = [
        'Класс шт',
        'Количество',
        'Остаток на конец периода, шт',
        'Свободный остаток',
        'Товары в пути',
        'Итоговый остаток, шт',
        'Ср. продажа в день за период, шт',
        'Себестоимость продажи',
        'Остаток на конец периода по себест., руб',
        'Расчет к заказу в днях'
    ]
    
    records = []
    
    for idx, city_start in enumerate(city_indices):
        # Определяем конец блока
        if idx + 1 < len(city_indices):
            city_end = city_indices[idx + 1]
        else:
            city_end = len(city_row)
        
        city_name = str(city_row[city_start]).strip()
        
        # Находим индексы полей внутри блока
        field_map = {}
        for field in expected_fields:
            for col in range(city_start, city_end):
                if pd.notna(field_row[col]) and str(field_row[col]).strip() == field:
                    field_map[field] = col
                    break
        
        # Если не все поля найдены, пропускаем город (ошибка структуры)
        if len(field_map) < len(expected_fields):
            continue
        
        # Проходим по строкам данных
        for row_idx in range(len(data)):
            row = data.iloc[row_idx]
            nomen = row[0] if pd.notna(row[0]) else ''
            char = row[1] if pd.notna(row[1]) else ''
            
            # Пропускаем строки "Итого" или пустые (без номенклатуры)
            if str(nomen).strip() == 'Итого' or pd.isna(row[0]):
                continue
            
            # Извлекаем значения
            class_val = row[field_map['Класс шт']]
            quantity = row[field_map['Количество']]
            ost = row[field_map['Остаток на конец периода, шт']]
            free = row[field_map['Свободный остаток']]
            in_transit = row[field_map['Товары в пути']]
            itog = row[field_map['Итоговый остаток, шт']]
            avg_sale = row[field_map['Ср. продажа в день за период, шт']]
            cost_sale = row[field_map['Себестоимость продажи']]
            ost_cost = row[field_map['Остаток на конец периода по себест., руб']]
            calc_order = row[field_map['Расчет к заказу в днях']]
            
            records.append({
                'Номенклатура': nomen,
                'Характеристика': char,
                'Город': city_name,
                'Класс': class_val,
                'Количество': quantity,
                'Остаток': ost,
                'Свободный': free,
                'В пути': in_transit,
                'Итоговый': itog,
                'Ср_продажа': avg_sale,
                'Себестоимость_продажи': cost_sale,
                'Остаток_по_себест': ost_cost,
                'Расчет_к_заказу': calc_order
            })
    
    df = pd.DataFrame(records)
    # Приводим числовые колонки к float (очищаем от возможных строк)
    numeric_cols = ['Количество', 'Остаток', 'Свободный', 'В пути', 'Итоговый',
                    'Ср_продажа', 'Себестоимость_продажи', 'Остаток_по_себест', 'Расчет_к_заказу']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    return df


# ------------------------------------------------------------
# 2. Расчёт KPI и стоп-факторов
# ------------------------------------------------------------
def calculate_kpi(df_june, df_july,
                  target_in_stock=90,
                  target_turnover=35,
                  limit_lost_profit=5,
                  limit_dead_stock=10):
    """
    Принимает два DataFrame (июнь, июль) и пороговые значения.
    Возвращает словарь с результатами.
    """
    # --- In-Stock Retail (по категории в целом) ---
    cities = df_july['Город'].unique()
    total_cities = len(cities)
    cities_with_stock = df_july[df_july['Остаток'] > 0]['Город'].nunique()
    in_stock_pct = (cities_with_stock / total_cities) * 100 if total_cities > 0 else 0
    in_stock_ok = in_stock_pct >= target_in_stock

    # --- Оборачиваемость ---
    cost_sales_july = df_july['Себестоимость_продажи'].sum()
    opening_stock = df_june['Остаток_по_себест'].sum()   # остаток на начало июля = конец июня
    closing_stock = df_july['Остаток_по_себест'].sum()
    avg_stock = (opening_stock + closing_stock) / 2
    days_in_month = 31  # для июля
    if cost_sales_july > 0:
        turnover = (avg_stock / cost_sales_july) * days_in_month
    else:
        turnover = float('inf')
    turnover_ok = turnover <= target_turnover

    # --- Упущенная прибыль (стоп-фактор) ---
    # Приближение: если остаток = 0, то недопроданное = Ср_продажа * 31
    df_july['Недопроданное'] = df_july.apply(
        lambda row: row['Ср_продажа'] * days_in_month if row['Остаток'] == 0 else 0,
        axis=1
    )
    total_lost_qty = df_july['Недопроданное'].sum()
    total_sales_qty = df_july['Количество'].sum()
    total_demand = total_sales_qty + total_lost_qty
    lost_profit_pct = (total_lost_qty / total_demand * 100) if total_demand > 0 else 0
    lost_profit_ok = lost_profit_pct <= limit_lost_profit

    # --- Неликвид (стоп-фактор) ---
    # Группируем продажи по товарам за июнь и июль
    sales_june = df_june.groupby(['Номенклатура', 'Характеристика'])['Количество'].sum().reset_index()
    sales_july = df_july.groupby(['Номенклатура', 'Характеристика'])['Количество'].sum().reset_index()
    merged = pd.merge(sales_june, sales_july, on=['Номенклатура', 'Характеристика'],
                      how='outer', suffixes=('_june', '_july')).fillna(0)
    dead_items = merged[(merged['Количество_june'] == 0) & (merged['Количество_july'] == 0)]
    dead_keys = set(zip(dead_items['Номенклатура'], dead_items['Характеристика']))

    dead_stock_cost = 0
    for _, row in df_july.iterrows():
        if (row['Номенклатура'], row['Характеристика']) in dead_keys:
            dead_stock_cost += row['Остаток_по_себест']

    total_stock_cost = df_july['Остаток_по_себест'].sum()
    dead_stock_pct = (dead_stock_cost / total_stock_cost * 100) if total_stock_cost > 0 else 0
    dead_stock_ok = dead_stock_pct <= limit_dead_stock

    # --- Итоговая премия ---
    stop_factors_ok = lost_profit_ok and dead_stock_ok
    if not stop_factors_ok:
        bonus = 0
        bonus_detail = "0 руб. (нарушен стоп-фактор)"
    else:
        if in_stock_ok and turnover_ok:
            bonus = 50000
            bonus_detail = "50 000 руб. (оба KPI выполнены)"
        elif in_stock_ok or turnover_ok:
            bonus = 25000
            bonus_detail = "25 000 руб. (выполнен только один KPI)"
        else:
            bonus = 0
            bonus_detail = "0 руб. (не выполнены оба KPI)"

    result = {
        'in_stock_pct': in_stock_pct,
        'in_stock_ok': in_stock_ok,
        'turnover': turnover,
        'turnover_ok': turnover_ok,
        'lost_profit_pct': lost_profit_pct,
        'lost_profit_ok': lost_profit_ok,
        'dead_stock_pct': dead_stock_pct,
        'dead_stock_ok': dead_stock_ok,
        'bonus': bonus,
        'bonus_detail': bonus_detail,
        'stop_factors_ok': stop_factors_ok,
        'cities_with_stock': cities_with_stock,
        'total_cities': total_cities,
        'cost_sales_july': cost_sales_july,
        'opening_stock': opening_stock,
        'closing_stock': closing_stock,
        'avg_stock': avg_stock,
        'total_lost_qty': total_lost_qty,
        'total_sales_qty': total_sales_qty,
        'dead_stock_cost': dead_stock_cost,
        'total_stock_cost': total_stock_cost,
        'dead_items_count': len(dead_keys)
    }
    return result


# ------------------------------------------------------------
# 3. Интерфейс Streamlit
# ------------------------------------------------------------
st.set_page_config(page_title="KPI Категорийного менеджера", layout="wide")
st.title("📊 Оценка эффективности категорийного менеджера")

# Боковая панель с настройками
st.sidebar.header("⚙️ Настройки целевых показателей")
target_in_stock = st.sidebar.number_input("In‑Stock Retail (цель, %)", min_value=0, max_value=100, value=90)
target_turnover = st.sidebar.number_input("Оборачиваемость (цель, дни)", min_value=1, value=35)
limit_lost_profit = st.sidebar.number_input("Упущенная прибыль (лимит, %)", min_value=0, max_value=100, value=5)
limit_dead_stock = st.sidebar.number_input("Неликвид (лимит, %)", min_value=0, max_value=100, value=10)

st.markdown("---")
st.write("Загрузите два Excel-файла: **за июнь** (предыдущий месяц) и **за июль** (отчётный месяц).")
st.caption("Структура файлов должна строго соответствовать образцу (первые две строки – заголовки городов и полей).")

col1, col2 = st.columns(2)
with col1:
    file_june = st.file_uploader("📁 Файл за ИЮНЬ", type=["xlsx"])
with col2:
    file_july = st.file_uploader("📁 Файл за ИЮЛЬ", type=["xlsx"])

if st.button("🚀 Рассчитать KPI"):
    if file_june is None or file_july is None:
        st.error("Загрузите оба файла!")
    else:
        with st.spinner("Обработка данных..."):
            try:
                df_june = parse_excel(file_june)
                df_july = parse_excel(file_july)

                if df_june.empty or df_july.empty:
                    st.error("Не удалось распознать структуру файлов. Проверьте формат.")
                    st.stop()

                result = calculate_kpi(df_june, df_july,
                                       target_in_stock, target_turnover,
                                       limit_lost_profit, limit_dead_stock)

                st.success("✅ Расчёт выполнен!")

                # --- Общий итог по премии ---
                st.header("💰 Итоговая премия")
                if result['bonus'] > 0:
                    st.success(f"**{result['bonus']:,} руб.**")
                else:
                    st.error(f"**{result['bonus']} руб.**")
                st.caption(result['bonus_detail'])

                # --- KPI ---
                st.subheader("🎯 Драйверные KPI")
                col1, col2 = st.columns(2)
                with col1:
                    st.metric(
                        label="In‑Stock Retail",
                        value=f"{result['in_stock_pct']:.1f}%",
                        delta=f"Цель: {target_in_stock}%",
                        delta_color="normal" if result['in_stock_ok'] else "inverse"
                    )
                    st.caption(f"Городов с наличием: {result['cities_with_stock']} из {result['total_cities']}")
                with col2:
                    st.metric(
                        label="Оборачиваемость (дни)",
                        value=f"{result['turnover']:.1f}" if result['turnover'] != float('inf') else "∞",
                        delta=f"Цель: ≤ {target_turnover}",
                        delta_color="normal" if result['turnover_ok'] else "inverse"
                    )
                    st.caption(f"Себестоимость продаж: {result['cost_sales_july']:,.0f} руб.")

                # --- Стоп-факторы ---
                st.subheader("⛔ Стоп-факторы (обнуляют премию)")
                col1, col2 = st.columns(2)
                with col1:
                    st.metric(
                        label="Упущенная прибыль (доля недопроданных штук)",
                        value=f"{result['lost_profit_pct']:.2f}%",
                        delta=f"Лимит: ≤ {limit_lost_profit}%",
                        delta_color="normal" if result['lost_profit_ok'] else "inverse"
                    )
                    st.caption(f"Недопроданных штук: {result['total_lost_qty']:,.0f} из {result['total_sales_qty'] + result['total_lost_qty']:,.0f}")
                with col2:
                    st.metric(
                        label="Неликвид (доля в остатке)",
                        value=f"{result['dead_stock_pct']:.2f}%",
                        delta=f"Лимит: ≤ {limit_dead_stock}%",
                        delta_color="normal" if result['dead_stock_ok'] else "inverse"
                    )
                    st.caption(f"Товаров без продаж за 2 месяца: {result['dead_items_count']}")

                # --- Детали ---
                with st.expander("📋 Детальный расчёт"):
                    st.write("**Оборачиваемость**")
                    st.write(f"- Остаток на начало июля (конец июня): {result['opening_stock']:,.0f} руб.")
                    st.write(f"- Остаток на конец июля: {result['closing_stock']:,.0f} руб.")
                    st.write(f"- Средний остаток: {result['avg_stock']:,.0f} руб.")
                    st.write(f"- Себестоимость продаж за июль: {result['cost_sales_july']:,.0f} руб.")
                    st.write(f"- Оборачиваемость: {result['turnover']:.2f} дней")

                    st.write("**Неликвид**")
                    st.write(f"- Остаток неликвидов по себестоимости: {result['dead_stock_cost']:,.0f} руб.")
                    st.write(f"- Общий остаток: {result['total_stock_cost']:,.0f} руб.")
                    st.write(f"- Доля неликвида: {result['dead_stock_pct']:.2f}%")

            except Exception as e:
                st.error(f"Ошибка при обработке: {e}")
                st.stop()

st.markdown("---")
st.caption("Приложение автоматически определяет структуру Excel-файлов. Ожидается, что первые две строки содержат заголовки городов и полей, а данные начинаются с третьей строки.")
