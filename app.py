import streamlit as st
import pandas as pd
import numpy as np
from calendar import monthrange

# ------------------------------------------------------------
# 1. Парсер Excel-файла (без изменений)
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
    
    city_row = df_raw.iloc[0]
    field_row = df_raw.iloc[1]
    data = df_raw.iloc[2:].reset_index(drop=True)
    
    city_indices = []
    for i in range(2, len(city_row)):
        val = city_row[i]
        if pd.notna(val) and str(val).strip() != '':
            city_indices.append(i)
    
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
        if idx + 1 < len(city_indices):
            city_end = city_indices[idx + 1]
        else:
            city_end = len(city_row)
        
        city_name = str(city_row[city_start]).strip()
        
        field_map = {}
        for field in expected_fields:
            for col in range(city_start, city_end):
                if pd.notna(field_row[col]) and str(field_row[col]).strip() == field:
                    field_map[field] = col
                    break
        
        if len(field_map) < len(expected_fields):
            continue
        
        for row_idx in range(len(data)):
            row = data.iloc[row_idx]
            nomen = row[0] if pd.notna(row[0]) else ''
            char = row[1] if pd.notna(row[1]) else ''
            
            if str(nomen).strip() == 'Итого' or pd.isna(row[0]):
                continue
            
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
    numeric_cols = ['Количество', 'Остаток', 'Свободный', 'В пути', 'Итоговый',
                    'Ср_продажа', 'Себестоимость_продажи', 'Остаток_по_себест', 'Расчет_к_заказу']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    return df


# ------------------------------------------------------------
# 2. Расчёт KPI (с учётом количества дней во втором месяце)
# ------------------------------------------------------------
def calculate_kpi(df_month1, df_month2,
                  days_in_month2,
                  target_in_stock=90,
                  target_turnover=35,
                  limit_lost_profit=5,
                  limit_dead_stock=10):
    """
    df_month1 – данные за первый месяц (предыдущий)
    df_month2 – данные за второй месяц (отчётный)
    days_in_month2 – количество дней во втором месяце
    """
    # --- In-Stock Retail ---
    cities = df_month2['Город'].unique()
    total_cities = len(cities)
    cities_with_stock = df_month2[df_month2['Остаток'] > 0]['Город'].nunique()
    in_stock_pct = (cities_with_stock / total_cities) * 100 if total_cities > 0 else 0
    in_stock_ok = in_stock_pct >= target_in_stock

    # --- Оборачиваемость ---
    cost_sales_month2 = df_month2['Себестоимость_продажи'].sum()
    opening_stock = df_month1['Остаток_по_себест'].sum()   # остаток на начало отчётного месяца
    closing_stock = df_month2['Остаток_по_себест'].sum()
    avg_stock = (opening_stock + closing_stock) / 2
    if cost_sales_month2 > 0:
        turnover = (avg_stock / cost_sales_month2) * days_in_month2
    else:
        turnover = float('inf')
    turnover_ok = turnover <= target_turnover

    # --- Упущенная прибыль (в штуках) ---
    df_month2['Недопроданное'] = df_month2.apply(
        lambda row: row['Ср_продажа'] * days_in_month2 if row['Остаток'] == 0 else 0,
        axis=1
    )
    total_lost_qty = df_month2['Недопроданное'].sum()
    total_sales_qty = df_month2['Количество'].sum()
    total_demand = total_sales_qty + total_lost_qty
    lost_profit_pct = (total_lost_qty / total_demand * 100) if total_demand > 0 else 0
    lost_profit_ok = lost_profit_pct <= limit_lost_profit

    # --- Неликвид (товары без продаж за два месяца) ---
    sales_m1 = df_month1.groupby(['Номенклатура', 'Характеристика'])['Количество'].sum().reset_index()
    sales_m2 = df_month2.groupby(['Номенклатура', 'Характеристика'])['Количество'].sum().reset_index()
    merged = pd.merge(sales_m1, sales_m2, on=['Номенклатура', 'Характеристика'],
                      how='outer', suffixes=('_m1', '_m2')).fillna(0)
    dead_items = merged[(merged['Количество_m1'] == 0) & (merged['Количество_m2'] == 0)]
    dead_keys = set(zip(dead_items['Номенклатура'], dead_items['Характеристика']))

    dead_stock_cost = 0
    for _, row in df_month2.iterrows():
        if (row['Номенклатура'], row['Характеристика']) in dead_keys:
            dead_stock_cost += row['Остаток_по_себест']

    total_stock_cost = df_month2['Остаток_по_себест'].sum()
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
        'cost_sales_month2': cost_sales_month2,
        'opening_stock': opening_stock,
        'closing_stock': closing_stock,
        'avg_stock': avg_stock,
        'total_lost_qty': total_lost_qty,
        'total_sales_qty': total_sales_qty,
        'dead_stock_cost': dead_stock_cost,
        'total_stock_cost': total_stock_cost,
        'dead_items_count': len(dead_keys),
        'days_in_month2': days_in_month2
    }
    return result


# ------------------------------------------------------------
# 3. Интерфейс Streamlit
# ------------------------------------------------------------
st.set_page_config(page_title="KPI Категорийного менеджера", layout="wide")
st.title("📊 Оценка эффективности категорийного менеджера")

# Словарь месяцев
month_names = {
    1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
    5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",
    9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"
}

# Боковая панель с настройками
st.sidebar.header("⚙️ Настройки целевых показателей")
target_in_stock = st.sidebar.number_input("In‑Stock Retail (цель, %)", min_value=0, max_value=100, value=90)
target_turnover = st.sidebar.number_input("Оборачиваемость (цель, дни)", min_value=1, value=35)
limit_lost_profit = st.sidebar.number_input("Упущенная прибыль (лимит, %)", min_value=0, max_value=100, value=5)
limit_dead_stock = st.sidebar.number_input("Неликвид (лимит, %)", min_value=0, max_value=100, value=10)

st.markdown("---")
st.write("Загрузите два Excel-файла: **первый** (предыдущий месяц) и **второй** (отчётный месяц).")
st.caption("Структура файлов должна строго соответствовать образцу (первые две строки – заголовки городов и полей).")

# Выбор месяцев
col_m1, col_m2 = st.columns(2)
with col_m1:
    month1_num = st.selectbox("Выберите первый месяц (предыдущий)", options=list(month_names.keys()), format_func=lambda x: month_names[x], index=5)  # июнь по умолчанию
with col_m2:
    month2_num = st.selectbox("Выберите второй месяц (отчётный)", options=list(month_names.keys()), format_func=lambda x: month_names[x], index=6)  # июль по умолчанию

# Получаем количество дней для второго месяца
days_in_month2 = monthrange(2026, month2_num)[1]  # используем фиксированный год (например, 2026), но если нужен текущий, можно взять из даты, но для простоты оставим так.

st.write(f"**Количество дней во втором месяце ({month_names[month2_num]}):** {days_in_month2}")

col1, col2 = st.columns(2)
with col1:
    file_month1 = st.file_uploader(f"📁 Файл за {month_names[month1_num]} (Месяц 1)", type=["xlsx"])
with col2:
    file_month2 = st.file_uploader(f"📁 Файл за {month_names[month2_num]} (Месяц 2)", type=["xlsx"])

if st.button("🚀 Рассчитать KPI"):
    if file_month1 is None or file_month2 is None:
        st.error("Загрузите оба файла!")
    else:
        with st.spinner("Обработка данных..."):
            try:
                df_m1 = parse_excel(file_month1)
                df_m2 = parse_excel(file_month2)

                if df_m1.empty or df_m2.empty:
                    st.error("Не удалось распознать структуру файлов. Проверьте формат.")
                    st.stop()

                result = calculate_kpi(df_m1, df_m2, days_in_month2,
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
                    st.caption(f"Себестоимость продаж: {result['cost_sales_month2']:,.0f} руб.")

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
                    st.write(f"**Оборачиваемость (за {month_names[month2_num]}, {days_in_month2} дней)**")
                    st.write(f"- Остаток на начало месяца (конец первого месяца): {result['opening_stock']:,.0f} руб.")
                    st.write(f"- Остаток на конец месяца: {result['closing_stock']:,.0f} руб.")
                    st.write(f"- Средний остаток: {result['avg_stock']:,.0f} руб.")
                    st.write(f"- Себестоимость продаж за месяц: {result['cost_sales_month2']:,.0f} руб.")
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
