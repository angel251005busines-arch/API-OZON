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
        
        try:
            response = requests.post(url, headers=self.headers, json=payload)
            response.raise_for_status()
            
            data = response.json()
            postings = data.get('result', {}).get('postings', [])
            posting_numbers = [p.get('posting_number') for p in postings if p.get('posting_number')]
            
            print(f"  Найдено отправлений: {len(posting_numbers)}")
            return posting_numbers
            
        except requests.exceptions.RequestException as e:
            print(f"Ошибка при получении списка отправлений: {e}")
            return []
    
    def get_accrual_postings(self, posting_numbers: List[str]) -> List[Dict]:
        """
        POST /v1/finance/accrual/postings
        Получение начислений по конкретным отправлениям
        """
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
                # КЛЮЧЕВОЕ ИСПРАВЛЕНИЕ: данные в поле 'posting_accruals'
                posting_accruals = data.get('posting_accruals', [])
                all_results.extend(posting_accruals)
                print(f"    Получено записей: {len(posting_accruals)}")
                
            except requests.exceptions.RequestException as e:
                print(f"    Ошибка: {e}")
                if e.response is not None:
                    print(f"    Ответ: {e.response.text[:200]}")
        
        return all_results


def process_financial_data(posting_accruals: List[Dict], accrual_types: List[Dict]) -> pd.DataFrame:
    """
    Преобразует финансовые данные в плоскую таблицу для Excel
    """
    rows = []
    
    # Создаем словарь для быстрого поиска описаний типов операций
    type_dict = {t['id']: t for t in accrual_types} if accrual_types else {}
    
    for posting_accrual in posting_accruals:
        posting_number = posting_accrual.get('posting_number', '')
        accruals = posting_accrual.get('accruals', [])
        
        for accrual in accruals:
            type_id = accrual.get('type_id')
            type_info = type_dict.get(type_id, {})
            
            row = {
                'Номер отправления': posting_number,
                'Тип операции ID': type_id,
                'Тип операции (код)': type_info.get('name', ''),
                'Тип операции (описание)': type_info.get('description', ''),
                'Сумма': float(accrual.get('accrued', {}).get('amount', 0)),
                'Валюта': accrual.get('accrued', {}).get('currency', 'RUB'),
                'Дата начисления': accrual.get('accrual_date', ''),
                'SKU': accrual.get('sku', ''),
                'Количество': accrual.get('quantity', 1),
                'Цена продавца': float(accrual.get('seller_price', {}).get('amount', 0)) if accrual.get('seller_price') else None,
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
        # Основной лист с детализацией
        df.to_excel(writer, sheet_name='Финансовые операции', index=False)
        
        # Лист со сводкой по типам операций
        summary_by_type = df.groupby(['Тип операции ID', 'Тип операции (код)', 'Тип операции (описание)']).agg({
            'Сумма': 'sum',
            'Номер отправления': 'count'
        }).rename(columns={'Номер отправления': 'Количество операций'}).reset_index()
        summary_by_type.to_excel(writer, sheet_name='Сводка по типам', index=False)
        
        # Лист со сводкой по отправлениям
        summary_by_posting = df.groupby('Номер отправления').agg({
            'Сумма': 'sum',
            'Тип операции ID': 'count'
        }).rename(columns={'Тип операции ID': 'Количество операций', 'Сумма': 'Итоговая сумма'}).reset_index()
        summary_by_posting.to_excel(writer, sheet_name='Сводка по отправлениям', index=False)
        
        # Лист со сводкой по SKU (только продажи, тип 69)
        sales_df = df[df['Тип операции ID'] == 69].copy()
        if not sales_df.empty:
            summary_by_sku = sales_df.groupby('SKU').agg({
                'Количество': 'sum',
                'Сумма': 'sum',
                'Цена продавца': 'first',
                'Номер отправления': 'count'
            }).rename(columns={'Номер отправления': 'Количество продаж'}).reset_index()
            summary_by_sku.to_excel(writer, sheet_name='Сводка по SKU', index=False)
        
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
    print(f"   Всего отправлений с операциями: {df['Номер отправления'].nunique()}")


def main():
    """Основная функция"""
    api = OzonFinanceAPI()
    
    # Задаем период (последние 30 дней)
    date_to = datetime.now()
    date_from = date_to - timedelta(days=30)
    
    print(f"\n=== Получение финансовых данных за период {date_from.date()} - {date_to.date()} ===\n")
    
    # 1. Получаем типы начислений
    print("1. Получение типов начислений...")
    accrual_types = api.get_accrual_types()
    if accrual_types:
        print(f"   ✅ Получено {len(accrual_types)} типов")
        
        # Сохраняем типы в JSON
        with open('accrual_types.json', 'w', encoding='utf-8') as f:
            json.dump(accrual_types, f, ensure_ascii=False, indent=2)
        print(f"   ✅ Сохранены в accrual_types.json")
    
    # 2. Получаем номера отправлений
    print("\n2. Получение номеров отправлений...")
    posting_numbers = api.get_posting_numbers_by_period(date_from, date_to)
    
    if not posting_numbers:
        print("   ⚠️ Нет отправлений за указанный период")
        return
    
    # 3. Получаем финансовые данные
    print("\n3. Получение финансовых данных...")
    posting_accruals = api.get_accrual_postings(posting_numbers)
    
    if not posting_accruals:
        print("   ⚠️ Нет финансовых данных")
        return
    
    print(f"\n   ✅ Получено отправлений с финансовыми данными: {len(posting_accruals)}")
    
    # 4. Обрабатываем и сохраняем в Excel
    print("\n4. Обработка и сохранение данных...")
    df = process_financial_data(posting_accruals, accrual_types)
    save_to_excel(df)
    
    # 5. Сохраняем сырые данные в JSON
    with open('posting_accruals.json', 'w', encoding='utf-8') as f:
        json.dump(posting_accruals, f, ensure_ascii=False, indent=2, default=str)
    print("   ✅ Сырые данные сохранены в posting_accruals.json")
    
    # 6. Выводим статистику
    print("\n=== СТАТИСТИКА ПО ТИПАМ ОПЕРАЦИЙ ===")
    
    # Группируем по типам
    stats = df.groupby(['Тип операции ID', 'Тип операции (описание)']).agg({
        'Сумма': 'sum',
        'Номер отправления': 'count'
    }).rename(columns={'Сумма': 'Сумма, ₽', 'Номер отправления': 'Кол-во операций'}).reset_index()
    stats = stats.sort_values('Сумма, ₽')
    
    for _, row in stats.iterrows():
        print(f"  {row['Тип операции (описание)']}: {row['Сумма, ₽']:,.2f} ₽ ({row['Кол-во операций']} шт.)")
    
    # Общая сумма
    total = df['Сумма'].sum()
    print(f"\n  ИТОГО: {total:,.2f} ₽")


if __name__ == "__main__":
    main()
