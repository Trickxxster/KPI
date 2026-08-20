import streamlit as st
import pandas as pd
import numpy as np
import datetime
import difflib
import re

# ------------------------------------------------------------
# 1. Парсер Excel-файла
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
# 2. Нормализация строк для сопоставления
# ------------------------------------------------------------
def normalize_text(text):
    """Приводит строку к единому формату: нижний регистр, удаление лишних пробелов и спецсимволов."""
    if pd.isna(text):
        return ''
    text = str(text).lower().strip()
    text = re.sub(r'[^a-zа-яё0-9\s]', '', text)
    return ' '.join(text.split())


# ------------------------------------------------------------
# 3. Загрузка прайс-листа (только Excel)
# ------------------------------------------------------------
def load_prices(file):
    """
    Загружает файл с ценами (Excel с несколькими листами).
    Ожидает колонки: 'Номенклатура', 'Характеристика', 'Себестоимость', 'РРЦ'.
    Если названия колонок отличаются, пытается угадать.
    Игнорирует строки, где Номенклатура не является товаром (пустые или подзаголовки).
    """
    xls = pd.ExcelFile(file)
    sheets = [pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names]
    
    all_data = []
    for df in sheets:
        if df.empty:
            continue
        col_map = {}
        for col in df.columns:
            col_lower = col.lower().strip()
            if 'номенклатур' in col_lower or 'название' in col_lower or 'товар' in col_lower:
                col_map['Номенклатура'] = col
            elif 'характеристик' in col_lower or 'модель' in col_lower or 'артикул' in col_lower:
                col_map['Характеристика'] = col
            elif 'себестоим' in col_lower or 'закуп' in col_lower:
                col_map['Себестоимость'] = col
            elif 'ррц' in col_lower or 'розничн' in col_lower or 'цена' in col_lower:
                col_map['РРЦ'] = col
        if len(col_map) < 4:
            if len(df.columns) >= 4:
                col_map = {
                    'Номенклатура': df.columns[0],
                    'Характеристика': df.columns[1],
                    'Себестоимость': df.columns[2],
                    'РРЦ': df.columns[3]
                }
            else:
                continue
        
        df_renamed = df.rename(columns=col_map)
        df_sub = df_renamed[['Номенклатура', 'Характеристика', 'Себестоимость', 'РРЦ']].copy()
        for col in ['Себестоимость', 'РРЦ']:
            df_sub[col] = pd.to_numeric(df_sub[col], errors='coerce')
        df_sub = df_sub.dropna(subset=['Себестоимость', 'РРЦ'], how='all')
        df_sub = df_sub[df_sub['Номенклатура'].notna()]
        df_sub = df_sub[df_sub['Номенклатура'].str.strip() != '']
        all_data.append(df_sub)
    
    if not all_data:
        return pd.DataFrame()
    
    df_prices = pd.concat(all_data, ignore_index=True)
    df_prices = df_prices.drop_duplicates(subset=['Номенклатура', 'Характеристика'], keep='first')
    return df_prices


# ------------------------------------------------------------
# 4. Сопоставление товаров с прайс-листом
# ------------------------------------------------------------
def match_prices(df_products, df_prices, threshold=0.7):
    if df_prices.empty:
        return {}
    
    price_keys = []
    price_indices = []
    for idx, row in df_prices.iterrows():
        key = normalize_text(row['Номенклатура']) + ' ' + normalize_text(row.get('Характеристика', ''))
        price_keys.append(key)
        price_indices.append(idx)
    
    unique_products = df_products[['Номенклатура', 'Характеристика']].drop_duplicates()
    match_dict = {}
    
    for _, prod_row in unique_products.iterrows():
        search_key = normalize_text(prod_row['Номенклатура']) + ' ' + normalize_text(prod_row.get('Характеристика', ''))
        if search_key.strip() == '':
            continue
        matches = difflib.get_close_matches(search_key, price_keys, n=1, cutoff=threshold)
        if matches:
            best_match = matches[0]
            idx = price_keys.index(best_match)
            matched_row = df_prices.iloc[price_indices[idx]]
            match_dict[(prod_row['Номенклатура'], prod_row['Характеристика'])] = (
                matched_row['Себестоимость'],
                matched_row['РРЦ']
            )
    
    return match_dict


# ------------------------------------------------------------
# 5. Расчёт In‑Stock Retail по городам
# ------------------------------------------------------------
def calculate_in_stock_retail_by_city_sales(df_month1, df_month2, min_avg_sales=0.01):
    df_combined = pd.concat([df_month1, df_month2], ignore_index=True)
    sales_by_city_sku = df_combined.groupby(['Город', 'Номенклатура', 'Характеристика'])['Ср_продажа'].sum().reset_index()
    sales_by_city_sku['Продаваемый'] = sales_by_city_sku['Ср_продажа'] > min_avg_sales
    
    cities = df_month2['Город'].unique()
    city_shares = {}
    
    for city in cities:
        selling_items = sales_by_city_sku[(sales_by_city_sku['Город'] == city) & (sales_by_city_sku['Продаваемый'] == True)]
        total_selling = len(selling_items)
        if total_selling == 0:
            city_shares[city] = None
            continue
        
        city_data = df_month2[df_month2['Город'] == city]
        present = 0
        for _, row in selling_items.iterrows():
            if ((city_data['Номенклатура'] == row['Номенклатура']) & 
                (city_data['Характеристика'] == row['Характеристика']) & 
                (city_data['Остаток'] > 0)).any():
                present += 1
        city_shares[city] = (present / total_selling) * 100 if total_selling > 0 else None
    
    return city_shares


# ------------------------------------------------------------
# 6. Основной расчёт KPI (обновлён с учётом срока хранения)
# ------------------------------------------------------------
def calculate_kpi(df_month1, df_month2,
                  days_in_period2,
                  target_in_stock=90,
                  target_turnover=35,
                  limit_lost_profit=5,
                  limit_dead_stock=10,
                  max_storage_days=180,
                  min_avg_sales=0.01,
                  price_dict=None,
                  use_prices=False):
    """
    Рассчитывает KPI.
    max_storage_days – порог срока хранения, выше которого товар считается неликвидом.
    """
    # --- In-Stock Retail ---
    city_in_stock = calculate_in_stock_retail_by_city_sales(df_month1, df_month2, min_avg_sales)
    in_stock_values = [v for v in city_in_stock.values() if v is not None]
    in_stock_pct = np.mean(in_stock_values) if in_stock_values else 0
    in_stock_ok = in_stock_pct >= target_in_stock

    # --- Оборачиваемость ---
    opening_stock_qty = df_month1['Остаток'].sum()
    closing_stock_qty = df_month2['Остаток'].sum()
    avg_stock_qty = (opening_stock_qty + closing_stock_qty) / 2
    avg_daily_sales = df_month2['Ср_продажа'].sum()
    if avg_daily_sales > 0:
        turnover_days = avg_stock_qty / avg_daily_sales
    else:
        turnover_days = float('inf')
    turnover_ok = turnover_days <= target_turnover

    # --- Упущенная прибыль ---
    df_month2['Недопроданное'] = df_month2.apply(
        lambda row: row['Ср_продажа'] * days_in_period2 
                     if (row['Остаток'] == 0 and row['Ср_продажа'] > min_avg_sales) 
                     else 0,
        axis=1
    )
    total_lost_qty = df_month2['Недопроданное'].sum()
    total_sales_qty = df_month2['Количество'].sum()
    total_demand = total_sales_qty + total_lost_qty
    lost_profit_pct = (total_lost_qty / total_demand * 100) if total_demand > 0 else 0
    lost_profit_ok = lost_profit_pct <= limit_lost_profit

    # Упущенная прибыль в деньгах (если есть цены)
    lost_profit_revenue = None
    lost_profit_margin = None
    if use_prices and price_dict:
        if 'Закупочная_цена' not in df_month2.columns:
            df_month2['Закупочная_цена'] = df_month2.apply(
                lambda row: price_dict.get((row['Номенклатура'], row['Характеристика']), (None, None))[0],
                axis=1
            )
            df_month2['РРЦ'] = df_month2.apply(
                lambda row: price_dict.get((row['Номенклатура'], row['Характеристика']), (None, None))[1],
                axis=1
            )
        lost_rows = df_month2[df_month2['Недопроданное'] > 0]
        if not lost_rows.empty:
            lost_revenue = (lost_rows['Недопроданное'] * lost_rows['РРЦ']).sum()
            lost_margin = (lost_rows['Недопроданное'] * (lost_rows['РРЦ'] - lost_rows['Закупочная_цена'])).sum()
            lost_profit_revenue = lost_revenue
            lost_profit_margin = lost_margin

    # --- Неликвид (по сроку хранения) ---
    # Группируем по (Город, Номенклатура, Характеристика)
    grouped_m2 = df_month2.groupby(['Город', 'Номенклатура', 'Характеристика']).agg({
        'Остаток': 'sum',
        'Ср_продажа': 'sum'
    }).reset_index()
    
    grouped_m1 = df_month1.groupby(['Город', 'Номенклатура', 'Характеристика'])['Остаток'].sum().reset_index()
    grouped_m1.rename(columns={'Остаток': 'Остаток_период1'}, inplace=True)
    
    merged = pd.merge(grouped_m2, grouped_m1, on=['Город', 'Номенклатура', 'Характеристика'], how='left').fillna(0)
    merged['Средний_остаток'] = (merged['Остаток'] + merged['Остаток_период1']) / 2
    merged['Среднедневные_продажи'] = merged['Ср_продажа']
    
    merged['Срок_хранения'] = merged.apply(
        lambda row: row['Средний_остаток'] / row['Среднедневные_продажи'] 
                    if row['Среднедневные_продажи'] > 0 else float('inf'),
        axis=1
    )
    
    dead_items_by_city = merged[merged['Срок_хранения'] > max_storage_days]
    dead_stock_qty = dead_items_by_city['Остаток'].sum()
    total_stock_qty = df_month2['Остаток'].sum()
    dead_stock_pct = (dead_stock_qty / total_stock_qty * 100) if total_stock_qty > 0 else 0
    dead_stock_ok = dead_stock_pct <= limit_dead_stock
    
    # Для отображения проблемных позиций
    dead_keys = set(zip(dead_items_by_city['Номенклатура'], dead_items_by_city['Характеристика']))

    # --- Премия ---
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

    # --- Детализация по городам ---
    city_metrics = {}
    cities = df_month2['Город'].unique()
    for city in cities:
        city_df = df_month2[df_month2['Город'] == city]
        city_df1 = df_month1[df_month1['Город'] == city]
        
        open_stock = city_df1['Остаток'].sum()
        close_stock = city_df['Остаток'].sum()
        avg_stock = (open_stock + close_stock) / 2
        avg_sales = city_df['Ср_продажа'].sum()
        turnover_city = avg_stock / avg_sales if avg_sales > 0 else float('inf')
        
        in_stock_city = city_in_stock.get(city, None)
        
        city_lost = city_df.apply(
            lambda row: row['Ср_продажа'] * days_in_period2 
                         if (row['Остаток'] == 0 and row['Ср_продажа'] > min_avg_sales) 
                         else 0,
            axis=1
        ).sum()
        city_sales = city_df['Количество'].sum()
        city_demand = city_sales + city_lost
        lost_pct_city = (city_lost / city_demand * 100) if city_demand > 0 else 0
        
        # Неликвид по городу (используем те же dead_keys)
        dead_city = city_df[city_df.apply(lambda row: (row['Номенклатура'], row['Характеристика']) in dead_keys, axis=1)]['Остаток'].sum()
        total_stock_city = city_df['Остаток'].sum()
        dead_pct_city = (dead_city / total_stock_city * 100) if total_stock_city > 0 else 0
        
        city_metrics[city] = {
            'in_stock_pct': in_stock_city,
            'turnover_days': turnover_city,
            'lost_profit_pct': lost_pct_city,
            'lost_profit_qty': city_lost,
            'dead_stock_pct': dead_pct_city,
            'dead_stock_qty': dead_city,
            'total_sales': city_sales,
            'total_stock': total_stock_city
        }

    result = {
        'in_stock_pct': in_stock_pct,
        'in_stock_ok': in_stock_ok,
        'turnover_days': turnover_days,
        'turnover_ok': turnover_ok,
        'lost_profit_pct': lost_profit_pct,
        'lost_profit_ok': lost_profit_ok,
        'dead_stock_pct': dead_stock_pct,
        'dead_stock_ok': dead_stock_ok,
        'bonus': bonus,
        'bonus_detail': bonus_detail,
        'stop_factors_ok': stop_factors_ok,
        'opening_stock_qty': opening_stock_qty,
        'closing_stock_qty': closing_stock_qty,
        'avg_stock_qty': avg_stock_qty,
        'avg_daily_sales': avg_daily_sales,
        'total_lost_qty': total_lost_qty,
        'total_sales_qty': total_sales_qty,
        'dead_stock_qty': dead_stock_qty,
        'total_stock_qty': total_stock_qty,
        'dead_items_count': len(dead_keys),
        'days_in_period2': days_in_period2,
        'city_metrics': city_metrics,
        'lost_profit_revenue': lost_profit_revenue,
        'lost_profit_margin': lost_profit_margin,
        'price_dict_used': price_dict is not None and use_prices,
        'dead_keys': dead_keys,
        'max_storage_days': max_storage_days
    }
    return result


# ------------------------------------------------------------
# 7. Детальный расчёт оборачиваемости
# ------------------------------------------------------------
def compute_detailed_turnover(df_month1, df_month2, days_in_period2, target_turnover):
    merged = pd.merge(
        df_month1[['Номенклатура', 'Характеристика', 'Город', 'Остаток']],
        df_month2[['Номенклатура', 'Характеристика', 'Город', 'Остаток', 'Ср_продажа', 'Количество']],
        on=['Номенклатура', 'Характеристика', 'Город'],
        how='outer',
        suffixes=('_нач', '_кон')
    ).fillna(0)

    merged['Средний_остаток'] = (merged['Остаток_нач'] + merged['Остаток_кон']) / 2
    merged['Оборачиваемость_дни'] = merged.apply(
        lambda row: row['Средний_остаток'] / row['Ср_продажа'] if row['Ср_продажа'] > 0 else float('inf'),
        axis=1
    )
    merged = merged[
        (merged['Остаток_нач'] > 0) | 
        (merged['Остаток_кон'] > 0) | 
        (merged['Ср_продажа'] > 0) |
        (merged['Количество'] > 0)
    ]
    merged['Превышение'] = merged['Оборачиваемость_дни'] > target_turnover
    merged.rename(columns={
        'Остаток_нач': 'Остаток_на_начало_периода_шт',
        'Остаток_кон': 'Остаток_на_конец_периода_шт',
        'Ср_продажа': 'Среднедневные_продажи_шт',
        'Количество': 'Продано_за_период_шт'
    }, inplace=True)
    merged = merged.sort_values('Оборачиваемость_дни', ascending=False)
    return merged


# ------------------------------------------------------------
# 8. Интерфейс Streamlit
# ------------------------------------------------------------
st.set_page_config(page_title="KPI Категорийного менеджера", layout="wide")
st.title("📊 Оценка эффективности категорийного менеджера")

# Боковая панель
st.sidebar.header("⚙️ Настройки целевых показателей")
target_in_stock = st.sidebar.number_input("In‑Stock Retail (цель, %)", min_value=0, max_value=100, value=90)
target_turnover = st.sidebar.number_input("Оборачиваемость (цель, дни)", min_value=1, value=35)
limit_lost_profit = st.sidebar.number_input("Упущенная прибыль (лимит, %)", min_value=0, max_value=100, value=5)
limit_dead_stock = st.sidebar.number_input("Неликвид (лимит, %)", min_value=0, max_value=100, value=10)
max_storage_days = st.sidebar.number_input(
    "Максимальный срок хранения (дни) для неликвида",
    min_value=1, value=180, step=10,
    help="Товар считается неликвидом, если его расчётный срок хранения превышает это значение."
)
min_avg_sales = st.sidebar.number_input(
    "Минимальный порог среднедневных продаж (для определения продаваемого товара)",
    min_value=0.0, value=0.01, step=0.01, format="%.3f"
)

# --- Выбор дат периодов (по умолчанию 2026 год) ---
st.sidebar.markdown("---")
st.sidebar.subheader("📅 Периоды")

st.sidebar.markdown("**Период 1 (предыдущий)**")
col1_start, col1_end = st.sidebar.columns(2)
with col1_start:
    start_date1 = st.date_input("Начало", value=datetime.date(2026, 6, 15), key="start1")
with col1_end:
    end_date1 = st.date_input("Окончание", value=datetime.date(2026, 7, 14), key="end1")

st.sidebar.markdown("**Период 2 (отчётный)**")
col2_start, col2_end = st.sidebar.columns(2)
with col2_start:
    start_date2 = st.date_input("Начало", value=datetime.date(2026, 7, 15), key="start2")
with col2_end:
    end_date2 = st.date_input("Окончание", value=datetime.date(2026, 8, 14), key="end2")

if start_date1 > end_date1:
    st.sidebar.error("Дата начала периода 1 не может быть позже даты окончания.")
if start_date2 > end_date2:
    st.sidebar.error("Дата начала периода 2 не может быть позже даты окончания.")

days_in_period2 = (end_date2 - start_date2).days + 1
st.sidebar.caption(f"Дней во втором периоде: {days_in_period2}")

# Загрузка прайс-листа
st.sidebar.markdown("---")
st.sidebar.subheader("💰 Цены для расчёта упущенной прибыли")
use_prices = st.sidebar.checkbox("Использовать цены (Excel)", value=False)
price_file = None
if use_prices:
    price_file = st.sidebar.file_uploader("Загрузите файл с ценами", type=["xlsx", "xls"])
    st.sidebar.caption("Ожидаются колонки: Номенклатура, Характеристика, Себестоимость, РРЦ")

st.markdown("---")
st.write("Загрузите два Excel-файла, соответствующие выбранным периодам.")
st.caption(f"**Период 1:** {start_date1.strftime('%d.%m.%Y')} – {end_date1.strftime('%d.%m.%Y')}  |  **Период 2:** {start_date2.strftime('%d.%m.%Y')} – {end_date2.strftime('%d.%m.%Y')}")

col1, col2 = st.columns(2)
with col1:
    file_month1 = st.file_uploader(f"📁 Файл за период 1", type=["xlsx"])
with col2:
    file_month2 = st.file_uploader(f"📁 Файл за период 2", type=["xlsx"])

if st.button("🚀 Рассчитать KPI"):
    if file_month1 is None or file_month2 is None:
        st.error("Загрузите оба файла!")
    else:
        with st.spinner("Обработка данных..."):
            try:
                df_m1 = parse_excel(file_month1)
                df_m2 = parse_excel(file_month2)

                # Валидация
                if df_m1.empty:
                    st.error("Первый файл не содержит данных или имеет неверную структуру.")
                    st.stop()
                if df_m2.empty:
                    st.error("Второй файл не содержит данных или имеет неверную структуру.")
                    st.stop()
                if df_m1['Номенклатура'].isna().all() or (df_m1['Номенклатура'].str.strip() == '').all():
                    st.error("В первом файле не найдено ни одного товара (столбец 'Номенклатура' пуст).")
                    st.stop()
                if df_m2['Номенклатура'].isna().all() or (df_m2['Номенклатура'].str.strip() == '').all():
                    st.error("Во втором файле не найдено ни одного товара (столбец 'Номенклатура' пуст).")
                    st.stop()

                # Загрузка прайса
                price_dict = None
                if use_prices and price_file is not None:
                    df_prices = load_prices(price_file)
                    if df_prices.empty:
                        st.warning("Файл с ценами пуст или имеет неверный формат. Расчёт будет выполнен без учёта цен.")
                    else:
                        price_dict = match_prices(df_m2, df_prices, threshold=0.7)
                        if not price_dict:
                            st.warning("Не удалось сопоставить ни одного товара с прайс-листом. Расчёт продолжится без цен.")
                elif use_prices and price_file is None:
                    st.warning("Цены включены, но файл не загружен. Расчёт будет выполнен без цен.")

                if price_dict:
                    df_m2['Закупочная_цена'] = df_m2.apply(
                        lambda row: price_dict.get((row['Номенклатура'], row['Характеристика']), (None, None))[0],
                        axis=1
                    )
                    df_m2['РРЦ'] = df_m2.apply(
                        lambda row: price_dict.get((row['Номенклатура'], row['Характеристика']), (None, None))[1],
                        axis=1
                    )

                # Основной расчёт
                result = calculate_kpi(
                    df_m1, df_m2, days_in_period2,
                    target_in_stock, target_turnover,
                    limit_lost_profit, limit_dead_stock,
                    max_storage_days, min_avg_sales,
                    price_dict, use_prices and price_dict is not None
                )

                detailed_turnover = compute_detailed_turnover(df_m1, df_m2, days_in_period2, target_turnover)

                # Вкладки
                tab1, tab2, tab3 = st.tabs(["📈 Основные KPI", "🔍 Детали оборачиваемости", "🏙️ Аналитика по городам"])

                with tab1:
                    st.success("✅ Расчёт выполнен!")

                    st.header("💰 Итоговая премия")
                    if result['bonus'] > 0:
                        st.success(f"**{result['bonus']:,} руб.**")
                    else:
                        st.error(f"**{result['bonus']} руб.**")
                    st.caption(result['bonus_detail'])

                    st.subheader("🎯 Драйверные KPI")
                    col_a, col_b = st.columns(2)
                    with col_a:
                        st.metric(
                            label="In‑Stock Retail",
                            value=f"{result['in_stock_pct']:.1f}%",
                            delta=f"Цель: {target_in_stock}%",
                            delta_color="normal" if result['in_stock_ok'] else "inverse",
                            help="Доля продаваемых товаров, которые есть в наличии на конец периода. Среднее по городам."
                        )
                        progress_in = min(result['in_stock_pct'] / target_in_stock, 1.0) if target_in_stock > 0 else 0
                        st.progress(progress_in, text=f"{result['in_stock_pct']:.1f}% от цели")
                    with col_b:
                        st.metric(
                            label="Оборачиваемость (дни)",
                            value=f"{result['turnover_days']:.1f}" if result['turnover_days'] != float('inf') else "∞",
                            delta=f"Цель: ≤ {target_turnover}",
                            delta_color="normal" if result['turnover_ok'] else "inverse",
                            help="Средний срок хранения запаса в днях по всей сети."
                        )
                        if result['turnover_days'] != float('inf'):
                            progress_turn = max(0, min(1 - (result['turnover_days'] / (target_turnover * 2)), 1))
                            st.progress(progress_turn, text=f"{result['turnover_days']:.1f} дн. (норма ≤ {target_turnover})")
                        else:
                            st.progress(0, text="Нет продаж")

                    st.subheader("⛔ Стоп-факторы (обнуляют премию)")
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric(
                            label="Упущенная прибыль",
                            value=f"{result['lost_profit_pct']:.2f}%",
                            delta=f"Лимит: ≤ {limit_lost_profit}%",
                            delta_color="normal" if result['lost_profit_ok'] else "inverse",
                            help="Доля недопроданных штук от общего спроса из-за дефицита."
                        )
                        st.caption(f"Недопроданных штук: {result['total_lost_qty']:,.0f} из {result['total_sales_qty'] + result['total_lost_qty']:,.0f}")
                        if result['price_dict_used']:
                            st.caption(f"Упущенная выручка: {result['lost_profit_revenue']:,.0f} руб." if result['lost_profit_revenue'] else "—")
                            st.caption(f"Упущенная маржа: {result['lost_profit_margin']:,.0f} руб." if result['lost_profit_margin'] else "—")
                    with col2:
                        st.metric(
                            label="Неликвид",
                            value=f"{result['dead_stock_pct']:.2f}%",
                            delta=f"Лимит: ≤ {limit_dead_stock}%",
                            delta_color="normal" if result['dead_stock_ok'] else "inverse",
                            help=f"Доля остатка товаров со сроком хранения > {max_storage_days} дней."
                        )
                        st.caption(f"Товаров с превышением срока хранения: {result['dead_items_count']}")

                    # Структура запасов
                    if result['total_stock_qty'] > 0:
                        dead_share = result['dead_stock_pct']
                        st.write("**Структура запасов:**")
                        st.progress(dead_share / 100, text=f"Неликвид: {dead_share:.1f}%")
                        st.caption(f"Оборотный запас: {100 - dead_share:.1f}%")
                    else:
                        st.info("Нет остатков для анализа.")

                with tab2:
                    st.subheader("🔍 Детальная оборачиваемость по товарам и городам")
                    st.caption("Показаны только позиции, у которых оборачиваемость превышает целевой норматив и есть остаток или продажи.")

                    over_df = detailed_turnover[detailed_turnover['Превышение'] == True].copy()

                    if over_df.empty:
                        st.info("✅ Все товары укладываются в норматив оборачиваемости. Отличная работа!")
                    else:
                        over_df['Оборачиваемость_дни'] = over_df['Оборачиваемость_дни'].apply(lambda x: f"{x:.1f}" if x != float('inf') else "∞")
                        over_df['Среднедневные_продажи_шт'] = over_df['Среднедневные_продажи_шт'].round(2)
                        over_df = over_df.sort_values('Оборачиваемость_дни', ascending=False, key=lambda x: x.replace('∞', '9999').astype(float))

                        st.dataframe(
                            over_df[['Номенклатура', 'Характеристика', 'Город',
                                     'Остаток_на_начало_периода_шт', 'Остаток_на_конец_периода_шт',
                                     'Среднедневные_продажи_шт', 'Оборачиваемость_дни']],
                            use_container_width=True,
                            column_config={
                                "Оборачиваемость_дни": st.column_config.TextColumn("Оборачиваемость (дни)"),
                            }
                        )
                        st.caption(f"Всего проблемных позиций: {len(over_df)}")
                        csv = over_df.to_csv(index=False).encode('utf-8')
                        st.download_button(
                            label="📥 Скачать отчёт по оборачиваемости (CSV)",
                            data=csv,
                            file_name=f"detected_turnover_issues_{start_date2.strftime('%Y%m%d')}_{end_date2.strftime('%Y%m%d')}.csv",
                            mime="text/csv",
                        )

                with tab3:
                    st.subheader("🏙️ Аналитика по городам")
                    city_metrics = result['city_metrics']
                    if not city_metrics:
                        st.info("Нет данных по городам.")
                    else:
                        rows = []
                        for city, metrics in city_metrics.items():
                            rows.append({
                                'Город': city,
                                'In-Stock Retail, %': round(metrics['in_stock_pct'], 1) if metrics['in_stock_pct'] is not None else '—',
                                'Оборачиваемость, дни': round(metrics['turnover_days'], 1) if metrics['turnover_days'] != float('inf') else '∞',
                                'Упущенная прибыль, %': round(metrics['lost_profit_pct'], 2),
                                'Недопроданно, шт': int(metrics['lost_profit_qty']),
                                'Неликвид, %': round(metrics['dead_stock_pct'], 2),
                                'Неликвид, шт': int(metrics['dead_stock_qty']),
                                'Продажи, шт': int(metrics['total_sales']),
                                'Остаток, шт': int(metrics['total_stock'])
                            })
                        df_city = pd.DataFrame(rows)
                        st.dataframe(df_city, use_container_width=True)

                        csv_city = df_city.to_csv(index=False).encode('utf-8')
                        st.download_button(
                            label="📥 Скачать аналитику по городам (CSV)",
                            data=csv_city,
                            file_name=f"city_analytics_{start_date2.strftime('%Y%m%d')}_{end_date2.strftime('%Y%m%d')}.csv",
                            mime="text/csv",
                        )

                        st.subheader("⚠️ Проблемные позиции")
                        
                        # Дефицит
                        deficit_df = df_m2[(df_m2['Остаток'] == 0) & (df_m2['Ср_продажа'] > min_avg_sales)].copy()
                        if not deficit_df.empty:
                            deficit_df['Недопроданное'] = deficit_df['Ср_продажа'] * days_in_period2
                            if 'РРЦ' in deficit_df.columns:
                                deficit_df['Сумма в РРЦ'] = deficit_df['Недопроданное'] * deficit_df['РРЦ']
                            else:
                                deficit_df['Сумма в РРЦ'] = None
                            st.write("**Дефицит (товар закончился, но продавался):**")
                            display_cols = ['Номенклатура', 'Характеристика', 'Город', 'Ср_продажа', 'Недопроданное']
                            if 'Сумма в РРЦ' in deficit_df.columns:
                                display_cols.append('Сумма в РРЦ')
                            st.dataframe(
                                deficit_df[display_cols],
                                use_container_width=True,
                                column_config={
                                    "Недопроданное": st.column_config.NumberColumn("Недопроданно, шт", format="%.0f"),
                                    "Сумма в РРЦ": st.column_config.NumberColumn("Сумма в РРЦ, руб", format="%.0f") if 'Сумма в РРЦ' in deficit_df.columns else None,
                                }
                            )
                        else:
                            st.success("✅ Нет дефицита по продаваемым товарам.")

                        # Неликвид (по сроку хранения)
                        dead_keys = result.get('dead_keys', set())
                        dead_stock_positions = df_m2[df_m2.apply(lambda row: (row['Номенклатура'], row['Характеристика']) in dead_keys and row['Остаток'] > 0, axis=1)].copy()
                        if not dead_stock_positions.empty:
                            if 'Закупочная_цена' in dead_stock_positions.columns:
                                dead_stock_positions['Сумма в закупке'] = dead_stock_positions['Остаток'] * dead_stock_positions['Закупочная_цена']
                            else:
                                dead_stock_positions['Сумма в закупке'] = None
                            st.write(f"**Неликвид (срок хранения > {max_storage_days} дней, остаток > 0):**")
                            display_cols = ['Номенклатура', 'Характеристика', 'Город', 'Остаток']
                            if 'Сумма в закупке' in dead_stock_positions.columns:
                                display_cols.append('Сумма в закупке')
                            st.dataframe(
                                dead_stock_positions[display_cols],
                                use_container_width=True,
                                column_config={
                                    "Остаток": st.column_config.NumberColumn("Остаток, шт", format="%.0f"),
                                    "Сумма в закупке": st.column_config.NumberColumn("Сумма в закупке, руб", format="%.0f") if 'Сумма в закупке' in dead_stock_positions.columns else None,
                                }
                            )
                        else:
                            st.success("✅ Нет неликвида.")

            except Exception as e:
                st.error(f"Ошибка при обработке: {e}")
                st.stop()

st.markdown("---")
st.caption("Приложение автоматически определяет структуру Excel-файлов. Ожидается, что первые две строки содержат заголовки городов и полей, а данные начинаются с третьей строки.")
