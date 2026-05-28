import os
import requests
import json
import pandas as pd
from datetime import datetime, timedelta
from typing import Dict, List, Optional
from dotenv import load_dotenv

load_dotenv()


class OzonFinanceAPI:
    """Класс для работы с финансовыми API Ozon (v1)"""
    
    def __init__(self):
        self.client_id = os.getenv('OZON_CLIENT_ID')
        self.api_key = os.getenv('OZON_API_KEY')
        self.base_url = os.getenv('OZON_API_URL', 'https://api-seller.ozon.ru')
        
        if not self.client_id or not self.api_key:
            raise ValueError("OZON_CLIENT_ID и OZON_API_KEY должны быть указаны в .env файле")
        
        self.headers = {
            'Client-Id': self.client_id,
            'Api-Key': self.api_key,
            'Content-Type': 'application/json'
        }
    
    def get_accrual_types(self) -> List[Dict]:
        """POST /v1/finance/accrual/types - Получение списка типов начислений"""
        url = f"{self.base_url}/v1/finance/accrual/types"
        
        try:
            response = requests.post(url, headers=self.headers, json={})
            response.raise_for_status()
            data = response.json()
            return data.get('accrual_types', [])
        except requests.exceptions.RequestException as e:
            print(f"Ошибка при получении типов начислений: {e}")
            return []
    
    def get_posting_numbers_by_period(self, date_from: datetime, date_to: datetime, limit: int = 100) -> List[str]:
        """Получение списка номеров отправлений за период"""
        url = f"{self.base_url}/v3/posting/fbs/list"
        
        payload = {
            "filter": {
                "since": date_from.strftime("%Y-%m-%dT%H:%M:%SZ"),
                "to": date_to.strftime("%Y-%m-%dT%H:%M:%SZ")
            },
            "limit": limit
        }
        
        all_posting_numbers = []
        offset = 0
        
        while True:
            payload["offset"] = offset
            
            try:
                response = requests.post(url, headers=self.headers, json=payload)
                response.raise_for_status()
                
                data = response.json()
                postings = data.get('result', {}).get('postings', [])
                
                if not postings:
                    break
                    
                posting_numbers = [p.get('posting_number') for p in postings if p.get('posting_number')]
                all_posting_numbers.extend(posting_numbers)
                
                if len(postings) < limit:
                    break
                    
                offset += limit
                
            except requests.exceptions.RequestException as e:
                print(f"Ошибка при получении списка отправлений: {e}")
                break
        
        print(f"  Найдено отправлений: {len(all_posting_numbers)}")
        return all_posting_numbers
    
    def get_accrual_postings(self, posting_numbers: List[str]) -> List[Dict]:
        """POST /v1/finance/accrual/postings - Получение начислений по отправлениям"""
        if not posting_numbers:
            return []
        
        url = f"{self.base_url}/v1/finance/accrual/postings"
        
        all_results = []
        
        for i in range(0, len(posting_numbers), 200):
            batch = posting_numbers[i:i+200]
            
            payload = {
                "posting_numbers": batch,
                "limit": len(batch)
            }
            
            print(f"  Загружаем батч {i//200 + 1}: {len(batch)} отправлений")
            
            try:
                response = requests.post(url, headers=self.headers, json=payload)
                response.raise_for_status()
                
                data = response.json()
                posting_accruals = data.get('posting_accruals', [])
                all_results.extend(posting_accruals)
                print(f"    Получено записей: {len(posting_accruals)}")
                
            except requests.exceptions.RequestException as e:
                print(f"    Ошибка: {e}")
                if e.response is not None:
                    print(f"    Ответ: {e.response.text[:200]}")
        
        return all_results


def load_cost_data(file_path: str):
    """Загружает файл с себестоимостью и категориями"""
    print(f"\n📂 Загрузка файла с себестоимостью: {file_path}")
    
    if file_path.endswith('.xlsx') or file_path.endswith('.xls'):
        df = pd.read_excel(file_path, header=1)
    elif file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
    else:
        raise ValueError("Неподдерживаемый формат файла. Используйте .xlsx, .xls или .csv")
    
    print(f"   ✅ Загружено строк: {len(df)}")
    
    # Определяем колонки
    sku_col = None
    cost_col = None
    category_col = None
    
    for col in df.columns:
        col_str = str(col).strip()
        if 'Артикул Озон' in col_str:
            sku_col = col
        if 'Трансферная цена комплект' in col_str:
            cost_col = col
        if 'Категория' in col_str and 'цен' not in col_str:
            category_col = col
    
    print(f"   🏷️  SKU колонка: {sku_col}")
    print(f"   💰 Себестоимость (комплект): {cost_col}")
    print(f"   📁 Категория колонка: {category_col}")
    
    if sku_col:
        # Нормализуем SKU: удаляем пробелы, приводим к строке
        df[sku_col] = df[sku_col].astype(str).str.strip()
        df[sku_col] = df[sku_col].str.replace('.0', '', regex=False)
        df[sku_col] = df[sku_col].str.replace(r'\.0$', '', regex=True)
        
        # Удаляем пустые
        df = df[df[sku_col].notna()]
        df = df[df[sku_col] != 'nan']
        df = df[df[sku_col] != '']
        
        # Выводим примеры SKU
        sample_skus = df[sku_col].head(10).tolist()
        print(f"   🔍 Примеры SKU из файла (после нормализации): {sample_skus}")
    
    if cost_col:
        df[cost_col] = pd.to_numeric(df[cost_col], errors='coerce')
    
    print(f"   ✅ После очистки осталось строк: {len(df)}")
    
    return df, sku_col, cost_col, category_col


def process_financial_data(posting_accruals: List[Dict], accrual_types: List[Dict], 
                           cost_df: pd.DataFrame = None, sku_col: str = None, 
                           cost_col: str = None, category_col: str = None) -> pd.DataFrame:
    """
    Преобразует финансовые данные в плоскую таблицу для Excel
    """
    rows = []
    
    type_dict = {t['id']: t for t in accrual_types} if accrual_types else {}
    
    # Создаем словарь для быстрого поиска себестоимости и категории
    cost_dict = {}
    category_dict = {}
    if cost_df is not None and sku_col:
        for _, row in cost_df.iterrows():
            sku = str(row[sku_col]) if pd.notna(row[sku_col]) else None
            if sku and sku != 'nan':
                sku = sku.strip()
                if cost_col and pd.notna(row[cost_col]):
                    cost_dict[sku] = float(row[cost_col])
                if category_col and pd.notna(row[category_col]):
                    category_dict[sku] = row[category_col]
    
    print(f"   📊 Загружено себестоимость для {len(cost_dict)} SKU")
    
    # Собираем все SKU из продаж для отладки
    sales_skus = set()
    for posting_accrual in posting_accruals:
        for accrual in posting_accrual.get('accruals', []):
            if accrual.get('type_id') == 69:
                sku = accrual.get('sku', '')
                if sku:
                    sku_str = str(sku).strip()
                    if sku_str.endswith('.0'):
                        sku_str = sku_str[:-2]
                    sales_skus.add(sku_str)
    
    print(f"   🔍 SKU из продаж Ozon (первые 10): {list(sales_skus)[:10]}")
    
    # Проверяем соответствие
    found_count = 0
    for sku in sales_skus:
        if sku in cost_dict:
            found_count += 1
    
    print(f"   ✅ Найдено себестоимость для {found_count} из {len(sales_skus)} SKU")
    
    if found_count < len(sales_skus):
        missing = [sku for sku in sales_skus if sku not in cost_dict]
        print(f"   ⚠️ Не найдены SKU: {missing[:10]}")
    
    # Сначала соберем все операции по отправлениям
    posting_expenses = {}
    
    for posting_accrual in posting_accruals:
        posting_number = posting_accrual.get('posting_number', '')
        accruals = posting_accrual.get('accruals', [])
        
        if posting_number not in posting_expenses:
            posting_expenses[posting_number] = {'other_expenses': 0, 'sales': []}
        
        for accrual in accruals:
            type_id = accrual.get('type_id')
            amount = float(accrual.get('accrued', {}).get('amount', 0))
            
            if type_id != 69:
                posting_expenses[posting_number]['other_expenses'] += amount
            else:
                sku = accrual.get('sku', '')
                quantity = accrual.get('quantity', 1)
                seller_price = float(accrual.get('seller_price', {}).get('amount', 0)) if accrual.get('seller_price') else None
                
                # Нормализуем SKU для поиска
                sku_normalized = str(sku).strip()
                if sku_normalized.endswith('.0'):
                    sku_normalized = sku_normalized[:-2]
                
                posting_expenses[posting_number]['sales'].append({
                    'sku': sku_normalized,
                    'sku_original': sku,
                    'quantity': quantity,
                    'seller_price': seller_price,
                    'commission': amount,
                    'accrual_date': accrual.get('accrual_date', '')
                })
    
    # Создаем строки для всех операций
    for posting_accrual in posting_accruals:
        posting_number = posting_accrual.get('posting_number', '')
        accruals = posting_accrual.get('accruals', [])
        total_other_expenses = posting_expenses[posting_number]['other_expenses']
        total_sales_count = sum(s['quantity'] for s in posting_expenses[posting_number]['sales'])
        
        for accrual in accruals:
            type_id = accrual.get('type_id')
            sku_original = accrual.get('sku', '')
            quantity = accrual.get('quantity', 1)
            amount = float(accrual.get('accrued', {}).get('amount', 0))
            seller_price = float(accrual.get('seller_price', {}).get('amount', 0)) if accrual.get('seller_price') else None
            
            type_info = type_dict.get(type_id, {})
            
            # Нормализуем SKU для поиска себестоимости
            sku_normalized = str(sku_original).strip()
            if sku_normalized.endswith('.0'):
                sku_normalized = sku_normalized[:-2]
            
            # Получаем себестоимость комплекта и категорию
            cost_kit = cost_dict.get(sku_normalized, 0)
            category = category_dict.get(sku_normalized, 'Неизвестно')
            
            if type_id == 69:
                share_of_expenses = (quantity / total_sales_count) * total_other_expenses if total_sales_count > 0 else 0
                gross_revenue = seller_price * quantity if seller_price else 0
                total_cost = cost_kit
                gross_profit = gross_revenue - total_cost
                net_income = gross_revenue + amount + share_of_expenses
                net_profit = net_income - total_cost
                margin = (gross_profit / gross_revenue * 100) if gross_revenue > 0 else 0
                net_margin = (net_profit / gross_revenue * 100) if gross_revenue > 0 else 0
                
                row = {
                    'Номер отправления': posting_number,
                    'Тип операции ID': type_id,
                    'Тип операции (код)': type_info.get('name', 'SaleCommission'),
                    'Тип операции (описание)': type_info.get('description', 'Вознаграждение за продажу'),
                    'Сумма': amount,
                    'Валюта': accrual.get('accrued', {}).get('currency', 'RUB'),
                    'Дата начисления': accrual.get('accrual_date', ''),
                    'SKU': sku_original,
                    'Категория': category,
                    'Количество': quantity,
                    'Цена продавца за шт.': seller_price,
                    'Себестоимость (за комплект)': round(cost_kit, 2),
                    'Выручка (брутто)': gross_revenue,
                    'Себестоимость всего': round(total_cost, 2),
                    'Валовая прибыль': round(gross_profit, 2),
                    'Валовая маржа %': round(margin, 2),
                    'Комиссия Ozon': amount,
                    'Прочие расходы (доля)': round(share_of_expenses, 2),
                    'Чистый доход (после расходов)': round(net_income, 2),
                    'Чистая прибыль': round(net_profit, 2),
                    'Чистая маржа %': round(net_margin, 2),
                    'Валюта цены': 'RUB',
                }
            else:
                row = {
                    'Номер отправления': posting_number,
                    'Тип операции ID': type_id,
                    'Тип операции (код)': type_info.get('name', ''),
                    'Тип операции (описание)': type_info.get('description', ''),
                    'Сумма': amount,
                    'Валюта': accrual.get('accrued', {}).get('currency', 'RUB'),
                    'Дата начисления': accrual.get('accrual_date', ''),
                    'SKU': sku_original,
                    'Категория': category,
                    'Количество': quantity,
                    'Цена продавца за шт.': seller_price if seller_price else None,
                    'Себестоимость (за комплект)': None,
                    'Выручка (брутто)': None,
                    'Себестоимость всего': None,
                    'Валовая прибыль': None,
                    'Валовая маржа %': None,
                    'Комиссия Ozon': None,
                    'Прочие расходы (доля)': None,
                    'Чистый доход (после расходов)': None,
                    'Чистая прибыль': None,
                    'Чистая маржа %': None,
                    'Валюта цены': accrual.get('seller_price', {}).get('currency', 'RUB') if accrual.get('seller_price') else None,
                }
            rows.append(row)
    
    return pd.DataFrame(rows)


def save_to_excel(df: pd.DataFrame, filename: str = None):
    """Сохраняет DataFrame в Excel с форматированием"""
    if filename is None:
        filename = f'ozon_finance_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    
    if df.empty:
        print("Нет данных для сохранения")
        return
    
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        # 1. Основной лист с детализацией
        df.to_excel(writer, sheet_name='Все операции', index=False)
        
        # 2. Только продажи с расчетом прибыли
        sales_df = df[df['Тип операции ID'] == 69].copy()
        if not sales_df.empty:
            sales_df.to_excel(writer, sheet_name='Продажи по позициям', index=False)
        
        # 3. Сводка по SKU
        sku_data = []
        for sku in sales_df['SKU'].unique():
            sku_rows = sales_df[sales_df['SKU'] == sku]
            
            quantity = sku_rows['Количество'].sum()
            revenue = sku_rows['Выручка (брутто)'].sum()
            cost_total = sku_rows['Себестоимость всего'].sum()
            gross_profit = sku_rows['Валовая прибыль'].sum()
            commission = sku_rows['Комиссия Ozon'].sum()
            other_expenses = df[(df['SKU'] == sku) & (df['Тип операции ID'] != 69)]['Сумма'].sum()
            net_income = revenue + commission + other_expenses
            net_profit = net_income - cost_total
            category = sku_rows['Категория'].iloc[0] if not sku_rows.empty else 'Неизвестно'
            avg_price = sku_rows['Цена продавца за шт.'].mean()
            avg_cost = sku_rows['Себестоимость (за комплект)'].mean()
            sales_count = sku_rows['Номер отправления'].nunique()
            
            sku_data.append({
                'SKU': sku,
                'Категория': category,
                'Количество (шт)': quantity,
                'Выручка': round(revenue, 2),
                'Себестоимость (комплект)': round(cost_total, 2),
                'Валовая прибыль': round(gross_profit, 2),
                'Валовая маржа %': round(gross_profit / revenue * 100, 2) if revenue > 0 else 0,
                'Комиссии и расходы': round(-(commission + other_expenses), 2),
                'Чистая прибыль': round(net_profit, 2),
                'Чистая маржа %': round(net_profit / revenue * 100, 2) if revenue > 0 else 0,
                'Ср. цена продажи': round(avg_price, 2),
                'Ср. себестоимость (комплект)': round(avg_cost, 2),
                'Кол-во продаж': sales_count
            })
        
        sku_summary = pd.DataFrame(sku_data)
        sku_summary = sku_summary.sort_values('Чистая прибыль', ascending=False)
        sku_summary.to_excel(writer, sheet_name='Сводка по SKU', index=False)
        
        # 4. Сводка по категориям
        category_data = []
        for category in sales_df['Категория'].unique():
            cat_rows = sales_df[sales_df['Категория'] == category]
            cat_all = df[df['Категория'] == category]
            
            revenue = cat_rows['Выручка (брутто)'].sum()
            cost_total = cat_rows['Себестоимость всего'].sum()
            gross_profit = cat_rows['Валовая прибыль'].sum()
            commission = cat_rows['Комиссия Ozon'].sum()
            other_expenses = cat_all[cat_all['Тип операции ID'] != 69]['Сумма'].sum()
            net_profit = revenue + commission + other_expenses - cost_total
            quantity = cat_rows['Количество'].sum()
            
            category_data.append({
                'Категория': category,
                'Количество продаж': quantity,
                'Выручка': round(revenue, 2),
                'Себестоимость': round(cost_total, 2),
                'Валовая прибыль': round(gross_profit, 2),
                'Валовая маржа %': round(gross_profit / revenue * 100, 2) if revenue > 0 else 0,
                'Чистая прибыль': round(net_profit, 2),
                'Чистая маржа %': round(net_profit / revenue * 100, 2) if revenue > 0 else 0
            })
        
        category_summary = pd.DataFrame(category_data)
        category_summary = category_summary.sort_values('Чистая прибыль', ascending=False)
        category_summary.to_excel(writer, sheet_name='Сводка по категориям', index=False)
        
        # 5. Топ убыточных товаров
        lossers = sku_summary[sku_summary['Чистая прибыль'] < 0].copy()
        if not lossers.empty:
            lossers = lossers.sort_values('Чистая прибыль', ascending=True)
            lossers.to_excel(writer, sheet_name='Убыточные товары', index=False)
        
        # 6. Сводка по типам операций
        summary_by_type = df.groupby(['Тип операции ID', 'Тип операции (код)', 'Тип операции (описание)']).agg({
            'Сумма': 'sum',
            'Номер отправления': 'count'
        }).rename(columns={'Номер отправления': 'Количество операций'}).reset_index()
        summary_by_type = summary_by_type.sort_values('Сумма')
        summary_by_type.to_excel(writer, sheet_name='Сводка по типам', index=False)
        
        # 7. Сводка по отправлениям
        posting_data = []
        for posting_number in df['Номер отправления'].unique():
            posting_df = df[df['Номер отправления'] == posting_number]
            revenue = posting_df['Выручка (брутто)'].sum()
            commission = posting_df['Комиссия Ozon'].sum()
            other_expenses = posting_df[posting_df['Тип операции ID'] != 69]['Сумма'].sum()
            cost_total = posting_df['Себестоимость всего'].sum()
            net_income = revenue + commission + other_expenses
            net_profit = net_income - cost_total
            
            posting_data.append({
                'Номер отправления': posting_number,
                'Выручка': round(revenue, 2),
                'Себестоимость': round(cost_total, 2),
                'Комиссии': round(commission, 2),
                'Прочие расходы': round(other_expenses, 2),
                'Чистый доход': round(net_income, 2),
                'Чистая прибыль': round(net_profit, 2),
                'Количество операций': len(posting_df)
            })
        
        posting_summary = pd.DataFrame(posting_data)
        posting_summary = posting_summary.sort_values('Чистая прибыль', ascending=False)
        posting_summary.to_excel(writer, sheet_name='Сводка по отправлениям', index=False)
        
        # 8. Общая статистика
        total_revenue = sales_df['Выручка (брутто)'].sum()
        total_cost = sales_df['Себестоимость всего'].sum()
        total_commission = sales_df['Комиссия Ozon'].sum()
        total_other_expenses = df[df['Тип операции ID'] != 69]['Сумма'].sum()
        final_income = total_revenue + total_commission + total_other_expenses
        final_profit = final_income - total_cost
        
        summary_stats = pd.DataFrame([
            ['Общая выручка (брутто)', f'{total_revenue:,.2f} ₽'],
            ['Общая себестоимость', f'{total_cost:,.2f} ₽'],
            ['Валовая прибыль', f'{total_revenue - total_cost:,.2f} ₽'],
            ['Валовая маржинальность', f'{((total_revenue - total_cost) / total_revenue * 100):.2f}%' if total_revenue > 0 else '0%'],
            ['Всего комиссий Ozon', f'{abs(total_commission):,.2f} ₽'],
            ['Прочие расходы (логистика, штрафы)', f'{abs(total_other_expenses):,.2f} ₽'],
            ['ИТОГОВЫЙ ДОХОД (до себестоимости)', f'{final_income:,.2f} ₽'],
            ['ЧИСТАЯ ПРИБЫЛЬ (после всех расходов)', f'{final_profit:,.2f} ₽'],
            ['Чистая маржинальность', f'{(final_profit / total_revenue * 100):.2f}%' if total_revenue > 0 else '0%']
        ], columns=['Показатель', 'Значение'])
        summary_stats.to_excel(writer, sheet_name='Общая статистика', index=False)
        
        # Автоматическая настройка ширины колонок
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
    
    print(f"\n✅ Данные сохранены в {filename}")
    print(f"   Всего операций: {len(df)}")
    print(f"   Продаж: {len(sales_df) if not sales_df.empty else 0}")


def get_date_from_input(prompt: str, allow_empty: bool = False) -> Optional[datetime]:
    """Получает дату от пользователя с проверкой формата"""
    while True:
        date_str = input(prompt).strip()
        
        if allow_empty and not date_str:
            return None
        
        formats = [
            "%Y-%m-%d",
            "%d.%m.%Y",
            "%d/%m/%Y",
            "%Y%m%d"
        ]
        
        for fmt in formats:
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue
        
        print("❌ Неверный формат даты. Используйте: ГГГГ-ММ-ДД, ДД.ММ.ГГГГ, ДД/ММ/ГГГГ или ГГГГММДД")


def main():
    """Основная функция"""
    print("\n" + "="*60)
    print("   ПОЛУЧЕНИЕ ФИНАНСОВЫХ ДАННЫХ OZON С РАСЧЕТОМ ПРИБЫЛИ")
    print("="*60)
    
    # Запрос файла с себестоимостью
    cost_file = input("\n📁 Введите путь к файлу с себестоимостью (Enter для пропуска): ").strip()
    
    cost_df = None
    sku_col = None
    cost_col = None
    category_col = None
    
    if cost_file and os.path.exists(cost_file):
        cost_df, sku_col, cost_col, category_col = load_cost_data(cost_file)
    elif cost_file:
        print(f"❌ Файл не найден: {cost_file}")
    
    print("\nВведите период для анализа:")
    print("  1 - Последние 30 дней")
    print("  2 - Последние 7 дней")
    print("  3 - Текущий месяц")
    print("  4 - Прошлый месяц")
    print("  5 - Вчера")
    print("  6 - Сегодня")
    print("  7 - Конкретный день")
    print("  8 - Произвольный период")
    
    choice = input("\nВыберите вариант (1-8) или нажмите Enter для варианта 1: ").strip()
    
    now = datetime.now()
    today = datetime(now.year, now.month, now.day, 23, 59, 59)
    yesterday = today - timedelta(days=1)
    
    if choice == '2':
        date_to = now
        date_from = now - timedelta(days=7)
    elif choice == '3':
        date_to = now
        date_from = datetime(now.year, now.month, 1)
    elif choice == '4':
        if now.month == 1:
            date_from = datetime(now.year - 1, 12, 1)
        else:
            date_from = datetime(now.year, now.month - 1, 1)
        date_to = datetime(now.year, now.month, 1) - timedelta(days=1)
    elif choice == '5':
        date_from = yesterday
        date_to = yesterday
    elif choice == '6':
        date_from = today
        date_to = today
    elif choice == '7':
        print("\nВведите дату:")
        selected_date = get_date_from_input("Дата (например, 2024-01-15): ")
        if selected_date:
            date_from = datetime(selected_date.year, selected_date.month, selected_date.day)
            date_to = datetime(selected_date.year, selected_date.month, selected_date.day, 23, 59, 59)
        else:
            print("❌ Дата не указана, используем последние 30 дней")
            date_to = now
            date_from = now - timedelta(days=30)
    elif choice == '8':
        print("\nВведите даты в формате ГГГГ-ММ-ДД")
        while True:
            date_from_str = input("Дата начала: ").strip()
            date_to_str = input("Дата окончания: ").strip()
            
            date_from = get_date_from_input(date_from_str, allow_empty=False)
            date_to = get_date_from_input(date_to_str, allow_empty=False)
            
            if date_from and date_to:
                if date_from <= date_to:
                    date_from = datetime(date_from.year, date_from.month, date_from.day)
                    date_to = datetime(date_to.year, date_to.month, date_to.day, 23, 59, 59)
                    break
                else:
                    print("❌ Дата начала должна быть меньше или равна дате окончания")
    else:
        date_to = now
        date_from = now - timedelta(days=30)
    
    print(f"\n✅ Выбран период: {date_from.date()} - {date_to.date()}")
    print(f"📅 Количество дней: {(date_to - date_from).days + 1}")
    
    save_types = input("\nСохранить список типов начислений в JSON? (y/N): ").strip().lower()
    save_raw = input("Сохранить сырые финансовые данные в JSON? (y/N): ").strip().lower()
    
    api = OzonFinanceAPI()
    
    print(f"\n=== Получение финансовых данных ===\n")
    
    print("1. Получение типов начислений...")
    accrual_types = api.get_accrual_types()
    if accrual_types:
        print(f"   ✅ Получено {len(accrual_types)} типов")
        if save_types == 'y':
            with open('accrual_types.json', 'w', encoding='utf-8') as f:
                json.dump(accrual_types, f, ensure_ascii=False, indent=2)
    
    print("\n2. Получение номеров отправлений...")
    posting_numbers = api.get_posting_numbers_by_period(date_from, date_to)
    
    if not posting_numbers:
        print("   ⚠️ Нет отправлений за указанный период")
        return
    
    print("\n3. Получение финансовых данных...")
    posting_accruals = api.get_accrual_postings(posting_numbers)
    
    if not posting_accruals:
        print("   ⚠️ Нет финансовых данных")
        return
    
    print(f"\n   ✅ Получено отправлений с финансовыми данными: {len(posting_accruals)}")
    
    print("\n4. Обработка и сохранение данных...")
    df = process_financial_data(posting_accruals, accrual_types, cost_df, sku_col, cost_col, category_col)
    
    if df.empty:
        print("   ⚠️ Нет данных для сохранения")
        return
    
    save_to_excel(df)
    
    if save_raw == 'y':
        with open('posting_accruals.json', 'w', encoding='utf-8') as f:
            json.dump(posting_accruals, f, ensure_ascii=False, indent=2, default=str)
        print("   ✅ Сырые данные сохранены в posting_accruals.json")
    
    print("\n✅ Готово!")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠️ Программа прервана пользователем")
    except Exception as e:
        print(f"\n❌ Ошибка: {e}")
