import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import numpy as np
from collections import defaultdict
import os
import shutil
import re
from datetime import timedelta
import pandas as pd
import numpy as np
from collections import defaultdict
import os
from datetime import timedelta
import re

def merge_excel_tables():
    # Проверяем существование файлов
    files = {
        'orders': 'Статистика по заказам за период.xlsx',
        'materials1': 'Расход материалов на заказ 1.xlsx', 
        'materials2': 'Расход материалов на заказ 2.xlsx',
        'stock': 'Склад доступный.xlsx'
    }
    
    # Проверяем какие файлы существуют
    existing_files = {}
    for key, filename in files.items():
        if os.path.exists(filename):
            existing_files[key] = filename
            print(f"Найден файл: {filename}")
        else:
            print(f"Файл не найден: {filename}")
    
    if 'orders' not in existing_files:
        print("Ошибка: Файл со статистикой заказов не найден!")
        return None
    
    # Чтение файла заказов
    orders_df = read_orders_file(existing_files['orders'])
    class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Объединение таблиц заказов")
        self.root.geometry("600x400")
        
        self.orders_file = tk.StringVar()
        self.materials1_file = tk.StringVar()
        self.materials2_file = tk.StringVar()
        self.stock_file = tk.StringVar()
        
        self.create_widgets()
    
    def create_widgets(self):
        title_label = tk.Label(self.root, text="Объединение таблиц Excel", 
                              font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        instruction = tk.Label(self.root, 
                              text="Выберите 4 файла Excel для объединения:\n"
                                   "1. Статистика по заказам\n"
                                   "2. Расход материалов 1\n" 
                                   "3. Расход материалов 2\n"
                                   "4. Склад доступный",
                              justify=tk.LEFT)
        instruction.pack(pady=10)
        
        self.create_file_selector("Статистика по заказам:", self.orders_file)
        self.create_file_selector("Расход материалов 1:", self.materials1_file)
        self.create_file_selector("Расход материалов 2:", self.materials2_file)
        self.create_file_selector("Склад доступный:", self.stock_file)
        
        merge_btn = tk.Button(self.root, text="Объединить таблицы", 
                             command=self.merge_tables, bg="green", fg="white",
                             font=("Arial", 12))
        merge_btn.pack(pady=20)
        
        self.progress = ttk.Progressbar(self.root, mode='indeterminate')
        self.progress.pack(fill=tk.X, padx=20)
        
        self.status_label = tk.Label(self.root, text="Готов к работе")
        self.status_label.pack(pady=5)
    
    def create_file_selector(self, label_text, file_var):
        frame = tk.Frame(self.root)
        frame.pack(fill=tk.X, padx=20, pady=5)
        
        label = tk.Label(frame, text=label_text, width=20, anchor='w')
        label.pack(side=tk.LEFT)
        
        entry = tk.Entry(frame, textvariable=file_var, width=40)
        entry.pack(side=tk.LEFT, padx=5)
        
        browse_btn = tk.Button(frame, text="Обзор", 
                              command=lambda: self.browse_file(file_var))
        browse_btn.pack(side=tk.LEFT)
    
    def browse_file(self, file_var):
        filename = filedialog.askopenfilename(
            title="Выберите файл Excel",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            file_var.set(filename)
    
    def merge_tables(self):
        files = {
            'orders': self.orders_file.get(),
            'materials1': self.materials1_file.get(), 
            'materials2': self.materials2_file.get(),
            'stock': self.stock_file.get()
        }
        
        for file_type, file_path in files.items():
            if not file_path or not os.path.exists(file_path):
                messagebox.showerror("Ошибка", f"Файл не выбран или не существует: {file_type}")
                return
        
        try:
            self.status_label.config(text="Объединение таблиц...")
            self.progress.start()
            
            original_files = {
                'orders': 'Статистика по заказам за период.xlsx',
                'materials1': 'Расход материалов на заказ 1.xlsx',
                'materials2': 'Расход материалов на заказ 2.xlsx', 
                'stock': 'Склад доступный.xlsx'
            }
            
            for file_type, original_name in original_files.items():
                shutil.copy2(files[file_type], original_name)
            
            result = merge_excel_tables()
            
            self.progress.stop()
            
            if result is not None:
                self.status_label.config(text="Готово! Файл создан: Объединенная_статистика_заказов.xlsx")
                messagebox.showinfo("Успех", "Таблицы успешно объединены!\n\n"
                                          "Файл сохранен как: 'Объединенная_статистика_заказов.xlsx'\n"
                                          "в той же папке, где находится программа.")
            else:
                messagebox.showerror("Ошибка", "Не удалось объединить таблицы")
                
        except Exception as e:
            self.progress.stop()
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{str(e)}")
        finally:
            for original_name in original_files.values():
                if os.path.exists(original_name):
                    os.remove(original_name)

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMergerApp(root)
    root.mainloop()
    
    if orders_df is None or orders_df.empty:
        print("Не удалось прочитать данные заказов")
        return None
    
    # Обработка таблиц материалов - ПЕРЕПРОВЕРКА ВСЕХ МАТЕРИАЛОВ
    print("\n=== ПЕРЕПРОВЕРКА ВСЕХ МАТЕРИАЛОВ ===")
    materials_dict = {}
    
    if 'materials1' in existing_files:
        materials_df1 = pd.read_excel(existing_files['materials1'], header=None)
        materials_dict.update(process_materials_from_columns(materials_df1))
    
    if 'materials2' in existing_files:
        materials_df2 = pd.read_excel(existing_files['materials2'], header=None)
        materials_dict.update(process_materials_from_columns(materials_df2))
    
    # Детальная проверка материалов
    print("\n=== ДЕТАЛЬНАЯ ПРОВЕРКА МАТЕРИАЛОВ ===")
    total_materials = 0
    for order_id, materials in materials_dict.items():
        print(f"Заказ {order_id}: {len(materials)} материалов")
        for material, quantity in materials.items():
            print(f"  - {material}: {quantity}")
        total_materials += len(materials)
    
    print(f"Всего материалов во всех заказах: {total_materials}")
    
    # Чтение данных о наличии на складе
    stock_data = {}
    if 'stock' in existing_files:
        print("\n=== ЧТЕНИЕ ДАННЫХ СКЛАДА ===")
        stock_data = read_stock_file(existing_files['stock'])
    
    # Обработка заказов с одинаковыми номерами
    processed_orders = process_duplicate_orders(orders_df)
    
    # Создание основной таблицы заказов с ПЕРЕПРОВЕРКОЙ всех материалов
    final_df = create_final_table_with_verification(processed_orders, materials_dict, orders_df, existing_files)
    
    # Создание анализа потребности материалов с группировкой по заказам
    materials_analysis_df = create_materials_analysis_with_orders(materials_dict, stock_data)
    
    # Сохранение результата (листы "Заказы" и "Потребность материалов")
    save_final_report_with_grouping(final_df, materials_analysis_df)
    
    return final_df

def read_orders_file(filename):
    """Чтение и обработка файла заказов"""
    try:
        # Читаем файл без пропуска строк и находим заголовок
        df_raw = pd.read_excel(filename, header=None)
        
        # Найдем строку с заголовком "Номер заказа"
        header_row = None
        for i in range(min(10, len(df_raw))):
            row_values = [str(x) for x in df_raw.iloc[i].values if pd.notna(x)]
            if any('Номер заказа' in str(x) for x in row_values):
                header_row = i
                break
        
        if header_row is not None:
            df = pd.read_excel(filename, skiprows=header_row)
            print(f"Найден заголовок в строке {header_row + 1}")
        else:
            # Используем стандартный подход
            df = pd.read_excel(filename, skiprows=2)
            print("Использован стандартный заголовок")
        
        # Очистка данных
        df = df.dropna(subset=['Номер заказа'])
        df['Номер заказа'] = df['Номер заказа'].astype(str).str.strip()
        
        # Форматирование стоимости заказа
        if 'Стоимость заказа' in df.columns:
            df['Стоимость заказа'] = df['Стоимость заказа'].apply(format_cost)
        
        # Форматирование площади заказа
        if 'Площадь заказа' in df.columns:
            df['Площадь заказа'] = df['Площадь заказа'].apply(format_area)
        
        print(f"Успешно прочитано {len(df)} заказов")
        print(f"Колонки: {list(df.columns)}")
        
        return df
        
    except Exception as e:
        print(f"Ошибка при чтении файла заказов: {e}")
        return None

def read_stock_file(filename):
    """Чтение файла с наличием материалов на складе с автоопределением структуры"""
    try:
        print(f"Чтение файла склада: {filename}")
        
        # Пробуем разные подходы к чтению файла
        stock_data = {}
        
        # Подход 1: Чтение с автоопределением заголовка
        df_raw = pd.read_excel(filename, header=None)
        
        # Ищем заголовок таблицы
        header_row = None
        name_col = None
        quantity_col = None
        
        # Ключевые слова для поиска колонок
        name_keywords = ['наименование', 'материал', 'название', 'name', 'material']
        quantity_keywords = ['количество', 'кол-во', 'остаток', 'доступно', 'quantity', 'stock', 'available']
        
        for i in range(min(15, len(df_raw))):  # Проверяем первые 15 строк
            row = df_raw.iloc[i]
            for col_idx, cell in enumerate(row):
                if pd.notna(cell) and isinstance(cell, str):
                    cell_lower = cell.lower().strip()
                    
                    # Поиск колонки с наименованием
                    if any(keyword in cell_lower for keyword in name_keywords) and name_col is None:
                        name_col = col_idx
                        print(f"Найдена колонка наименований: {col_idx} ('{cell}')")
                    
                    # Поиск колонки с количеством
                    if any(keyword in cell_lower for keyword in quantity_keywords) and quantity_col is None:
                        quantity_col = col_idx
                        print(f"Найдена колонка количества: {col_idx} ('{cell}')")
            
            # Если нашли обе колонки, запоминаем строку заголовка
            if name_col is not None and quantity_col is not None:
                header_row = i
                print(f"Найден заголовок таблицы в строке {header_row + 1}")
                break
        
        # Если не нашли заголовок, пробуем альтернативные подходы
        if header_row is None:
            print("Не удалось автоматически определить заголовок таблицы склада")
            print("Пробуем альтернативные методы...")
            
            # Подход 2: Ищем данные в первых столбцах
            for i in range(len(df_raw)):
                row = df_raw.iloc[i]
                # Предполагаем, что первые два столбца содержат наименование и количество
                if len(row) >= 2 and pd.notna(row[0]) and pd.notna(row[1]):
                    if isinstance(row[0], str) and (isinstance(row[1], (int, float)) or (isinstance(row[1], str) and row[1].replace(',', '.').replace(' ', '').replace('.', '').isdigit())):
                        name_col = 0
                        quantity_col = 1
                        header_row = i - 1 if i > 0 else 0
                        print(f"Предположительная структура: колонка 0 - наименование, колонка 1 - количество")
                        break
        
        # Извлекаем данные
        if name_col is not None and quantity_col is not None:
            start_row = header_row + 1 if header_row is not None else 0
            
            for i in range(start_row, len(df_raw)):
                row = df_raw.iloc[i]
                
                if len(row) > max(name_col, quantity_col):
                    material_name = row[name_col]
                    quantity = row[quantity_col]
                    
                    if pd.notna(material_name) and isinstance(material_name, str) and material_name.strip():
                        material_name_clean = material_name.strip()
                        
                        if pd.notna(quantity):
                            try:
                                if isinstance(quantity, (int, float)):
                                    qty_value = float(quantity)
                                else:
                                    qty_str = str(quantity).replace(',', '.').replace(' ', '')
                                    qty_value = float(qty_str) if qty_str.replace('.', '').isdigit() else 0
                                
                                if qty_value > 0:
                                    stock_data[material_name_clean] = qty_value
                                    print(f"  Добавлен материал на складе: {material_name_clean} = {qty_value}")
                            except (ValueError, TypeError) as e:
                                print(f"  Ошибка преобразования количества: {quantity}, ошибка: {e}")
        
        print(f"Успешно прочитано материалов на складе: {len(stock_data)}")
        
        # Если ничего не нашли, пробуем стандартное чтение
        if not stock_data:
            print("Пробуем стандартное чтение DataFrame...")
            try:
                df_standard = pd.read_excel(filename)
                print(f"Колонки в файле: {list(df_standard.columns)}")
                
                # Пробуем найти подходящие колонки
                for col in df_standard.columns:
                    if any(keyword in col.lower() for keyword in name_keywords):
                        name_col = col
                    if any(keyword in col.lower() for keyword in quantity_keywords):
                        quantity_col = col
                
                if name_col and quantity_col:
                    for _, row in df_standard.iterrows():
                        if pd.notna(row[name_col]) and pd.notna(row[quantity_col]):
                            material_name = str(row[name_col]).strip()
                            try:
                                qty_value = float(row[quantity_col])
                                if qty_value > 0:
                                    stock_data[material_name] = qty_value
                                    print(f"  Добавлен материал на складе: {material_name} = {qty_value}")
                            except (ValueError, TypeError):
                                pass
            except Exception as e:
                print(f"Ошибка при стандартном чтении файла склада: {e}")
        
        return stock_data
        
    except Exception as e:
        print(f"Ошибка при чтении файла наличия на складе: {e}")
        return {}

def format_cost(value):
    """Форматирование стоимости: убираем пробелы, заменяем запятую на точку"""
    if pd.isna(value):
        return value
    
    # Преобразуем в строку
    str_value = str(value)
    
    # Убираем пробелы (разделители тысяч)
    str_value = str_value.replace(' ', '')
    str_value = str_value.replace('\u202f', '')  # Убираем thin space
    str_value = str_value.replace('\xa0', '')    # Убираем non-breaking space
    
    # Заменяем запятую на точку (если есть)
    str_value = str_value.replace(',', '.')
    
    try:
        # Пробуем преобразовать в float
        num_value = float(str_value)
        # Форматируем с двумя знаками после запятой
        return round(num_value, 2)
    except ValueError:
        # Если не получается, возвращаем оригинальное значение
        return value

def format_area(value):
    """Форматирование площади: убираем пробелы, заменяем запятую на точку"""
    if pd.isna(value):
        return value
    
    # Преобразуем в строку
    str_value = str(value)
    
    # Убираем пробелы (разделители тысяч)
    str_value = str_value.replace(' ', '')
    str_value = str_value.replace('\u202f', '')  # Убираем thin space
    str_value = str_value.replace('\xa0', '')    # Убираем non-breaking space
    
    # Заменяем запятую на точку (если есть)
    str_value = str_value.replace(',', '.')
    
    try:
        # Пробуем преобразовать в float
        num_value = float(str_value)
        return num_value
    except ValueError:
        # Если не получается, возвращаем оригинальное значение
        return value

def process_materials_from_columns(materials_df):
    """Обработка таблицы материалов с использованием колонок 'Наименование' и количеств БЕЗ единиц измерения"""
    materials_dict = {}
    current_order = None
    i = 0
    
    print(f"Обработка файла материалов: {len(materials_df)} строк")
    
    # Найдем заголовок таблицы с колонками
    header_row = None
    name_col = None
    quantity_col = None
    
    for i in range(min(20, len(materials_df))):  # Проверяем первые 20 строк
        row = materials_df.iloc[i]
        for col_idx, cell in enumerate(row):
            if pd.notna(cell) and isinstance(cell, str):
                if 'Наименование' in cell:
                    name_col = col_idx
                if 'Кол-во c отходами' in cell or 'Количество' in cell:
                    quantity_col = col_idx
        
        # Если нашли нужные колонки, запоминаем строку заголовка
        if name_col is not None and quantity_col is not None:
            header_row = i
            print(f"Найдены колонки: Наименование={name_col}, Количество={quantity_col}")
            break
    
    if header_row is None:
        print("Не найдены заголовки таблицы материалов")
        return materials_dict
    
    # Обрабатываем строки после заголовка
    i = header_row + 1
    while i < len(materials_df):
        row = materials_df.iloc[i]
        
        # Поиск нового заказа
        order_found = False
        for col in range(min(5, len(row))):
            if pd.notna(row[col]) and isinstance(row[col], str):
                cell_value = str(row[col]).strip()
                if 'Расход материалов на заказ' in cell_value:
                    # Извлекаем номер заказа
                    numbers = re.findall(r'\d+', cell_value)
                    if numbers:
                        current_order = numbers[0]
                        print(f"Найден заказ: {current_order}")
                        order_found = True
                        break
        
        if order_found:
            i += 1
            continue
        
        # Обработка строк с материалами
        if current_order and name_col < len(row) and quantity_col < len(row):
            material_name = row[name_col] if name_col < len(row) else None
            quantity = row[quantity_col] if quantity_col < len(row) else None
            
            if (pd.notna(material_name) and isinstance(material_name, str) and 
                material_name.strip() and not any(keyword in material_name.lower() 
                for keyword in ['расход', 'материал', 'заказ', 'наименование', 'артикул', 'количество', 'ед. изм.'])):
                
                if pd.notna(quantity):
                    try:
                        # Преобразуем количество в число
                        if isinstance(quantity, (int, float)):
                            qty_value = float(quantity)
                        else:
                            qty_str = str(quantity).replace(',', '.').replace(' ', '')
                            qty_value = float(qty_str) if qty_str.replace('.', '').isdigit() else 0
                        
                        if qty_value > 0:
                            # Сохраняем название материала БЕЗ единиц измерения
                            material_name_clean = material_name.strip()
                            
                            if current_order not in materials_dict:
                                materials_dict[current_order] = {}
                            
                            materials_dict[current_order][material_name_clean] = qty_value
                            print(f"  Добавлен материал: {material_name_clean} = {qty_value}")
                    except (ValueError, TypeError) as e:
                        print(f"  Ошибка преобразования количества: {quantity}, ошибка: {e}")
        
        i += 1
    
    print(f"Обработано заказов с материалами: {len(materials_dict)}")
    return materials_dict

def process_duplicate_orders(orders_df):
    """Обработка заказов с одинаковыми номерами"""
    processed_orders = []
    orders_seen = {}
    
    for _, row in orders_df.iterrows():
        order_num = str(row['Номер заказа']).strip()
        product_type = row['Тип продукции']
        
        if order_num in orders_seen:
            # Объединяем типы продукции через запятую
            existing_idx = orders_seen[order_num]
            existing_type = processed_orders[existing_idx]['Тип продукции']
            
            if pd.notna(product_type) and pd.notna(existing_type):
                existing_types = str(existing_type).split(', ')
                if str(product_type) not in existing_types:
                    processed_orders[existing_idx]['Тип продукции'] = f"{existing_type}, {product_type}"
        else:
            # Добавляем новую запись
            new_order = row.to_dict()
            processed_orders.append(new_order)
            orders_seen[order_num] = len(processed_orders) - 1
    
    print(f"После обработки дубликатов: {len(processed_orders)} уникальных заказов")
    return processed_orders

def create_final_table_with_verification(processed_orders, materials_dict, original_orders_df, existing_files):
    """Создание финальной таблицы с ПЕРЕПРОВЕРКОЙ всех материалов"""
    
    # Собираем ВСЕ уникальные наименования материалов из ВСЕХ заказов
    all_materials = set()
    for order_materials in materials_dict.values():
        all_materials.update(order_materials.keys())
    
    all_materials = sorted(list(all_materials))
    print(f"\n=== ВСЕГО УНИКАЛЬНЫХ МАТЕРИАЛОВ: {len(all_materials)} ===")
    for material in all_materials[:20]:  # Показываем первые 20
        print(f"  - {material}")
    if len(all_materials) > 20:
        print(f"  ... и еще {len(all_materials) - 20} материалов")
    
    # Создаем DataFrame
    final_data = []
    
    # Счетчики для проверки
    orders_with_materials = 0
    orders_without_materials = 0
    
    for order in processed_orders:
        order_num = str(order['Номер заказа']).strip()
        order_data = order.copy()
        
        # Добавляем данные о материалах
        if order_num in materials_dict:
            orders_with_materials += 1
            for material in all_materials:
                quantity = materials_dict[order_num].get(material, 0)
                # Заменяем 0 на пустую строку для лучшего восприятия
                order_data[material] = quantity if quantity != 0 else ""
            
            # ПРОВЕРКА: выводим информацию о материалах для этого заказа
            order_materials_count = len([q for q in materials_dict[order_num].values() if q > 0])
            print(f"Заказ {order_num}: добавлено {order_materials_count} материалов")
        else:
            orders_without_materials += 1
            print(f"ПРЕДУПРЕЖДЕНИЕ: Заказ {order_num} не найден в материалах!")
            # Заполняем пустыми значениями
            for material in all_materials:
                order_data[material] = ""
        
        final_data.append(order_data)
    
    print(f"\n=== СВОДКА ПО МАТЕРИАЛАМ ===")
    print(f"Заказов с материалыми: {orders_with_materials}")
    print(f"Заказов без материалов: {orders_without_materials}")
    
    # Создаем финальный DataFrame
    final_columns = list(processed_orders[0].keys()) + all_materials
    final_df = pd.DataFrame(final_data, columns=final_columns)
    
    # Удаляем лишние столбцы (Unnamed и Пусто)
    columns_to_keep = []
    for col in final_df.columns:
        if not col.startswith('Unnamed:') and not col.startswith('Пусто') and col != '№':
            columns_to_keep.append(col)
    
    final_df = final_df[columns_to_keep]
    
    # ДОПОЛНИТЕЛЬНАЯ ПРОВЕРКА: сравниваем с исходными данными
    print("\n=== ПРОВЕРКА СООТВЕТСТВИЯ МАТЕРИАЛОВ ===")
    verify_materials_coverage(final_df, materials_dict, existing_files)
    
    # Проверка на потерю заказов
    original_orders = set(str(x).strip() for x in original_orders_df['Номер заказа'] if pd.notna(x))
    final_orders = set(str(x).strip() for x in final_df['Номер заказа'] if pd.notna(x))
    
    lost_orders = original_orders - final_orders
    if lost_orders:
        print(f"Предупреждение: потеряны заказов: {len(lost_orders)}")
        for lost in list(lost_orders)[:10]:  # Показываем первые 10
            print(f"  - {lost}")
    else:
        print("Все заказы успешно обработаны")
    
    print(f"Итоговый файл: {len(final_df)} строк, {len(final_df.columns)} колонок")
    print(f"Уникальных материалов: {len(all_materials)}")
    
    return final_df

def verify_materials_coverage(final_df, materials_dict, existing_files):
    """Проверка покрытия материалов - убеждаемся, что ВСЕ материалы из исходников включены"""
    print("Проверка покрытия материалов...")
    
    # Снова читаем исходные файлы для проверки
    all_original_materials = set()
    
    for file_key in ['materials1', 'materials2']:
        if file_key in existing_files:
            print(f"Проверка файла: {existing_files[file_key]}")
            df = pd.read_excel(existing_files[file_key], header=None)
            
            # Ищем заголовки
            header_row = None
            name_col = None
            
            for i in range(min(20, len(df))):
                row = df.iloc[i]
                for col_idx, cell in enumerate(row):
                    if pd.notna(cell) and isinstance(cell, str):
                        if 'Наименование' in cell:
                            name_col = col_idx
                            header_row = i
                            break
                if name_col is not None:
                    break
            
            if header_row is None:
                continue
                
            # Собираем материалы
            current_order = None
            for i in range(header_row + 1, len(df)):
                row = df.iloc[i]
                
                # Поиск заказа
                for col in range(min(5, len(row))):
                    if pd.notna(row[col]) and isinstance(row[col], str):
                        cell_value = str(row[col]).strip()
                        if 'Расход материалов на заказ' in cell_value:
                            numbers = re.findall(r'\d+', cell_value)
                            if numbers:
                                current_order = numbers[0]
                            break
                
                # Сбор материалов
                if current_order and name_col < len(row):
                    material_name = row[name_col] if name_col < len(row) else None
                    
                    if (pd.notna(material_name) and isinstance(material_name, str) and 
                        material_name.strip() and not any(keyword in material_name.lower() 
                        for keyword in ['расход', 'материал', 'заказ', 'наименование', 'артикул', 'количество', 'ед. изм.'])):
                        
                        material_name_clean = material_name.strip()
                        all_original_materials.add(material_name_clean)
    
    print(f"Всего материалов в исходниках: {len(all_original_materials)}")
    
    # Проверяем, какие материалы из исходников есть в финальной таблице
    final_materials = set(final_df.columns) - set(['Номер заказа', 'Клиент', 'Тип продукции', 'Площадь заказа', 
                                                 'Дата создания', 'Дата завершения', 'Стоимость заказа', 
                                                 'Состояние заказа', 'Комментарий'])
    
    missing_materials = all_original_materials - final_materials
    if missing_materials:
        print(f"ПРЕДУПРЕЖДЕНИЕ: Не хватает материалов в финальной таблице: {len(missing_materials)}")
        for material in sorted(list(missing_materials))[:20]:  # Показываем первые 20
            print(f"  - {material}")
    else:
        print("✓ Все материалы из исходников присутствуют в финальной таблице")

def create_materials_analysis_with_orders(materials_dict, stock_data):
    """Создание анализа потребности материалов с отдельными колонками для каждого заказа"""
    print("\n=== СОЗДАНИЕ АНАЛИЗА ПОТРЕБНОСТИ МАТЕРИАЛОВ С ГРУППИРОВКОЙ ===")
    
    # Собираем все уникальные заказы
    all_orders = sorted(list(materials_dict.keys()))
    print(f"Всего уникальных заказов: {len(all_orders)}")
    
    # Собираем все уникальные материалы
    all_materials = set()
    for materials in materials_dict.values():
        all_materials.update(materials.keys())
    
    all_materials = sorted(list(all_materials))
    
    # Создаем структуру данных для анализа
    analysis_data = []
    
    for material in all_materials:
        material_row = {'Материал': material}
        
        # Добавляем данные по каждому заказу
        total_required = 0
        for order in all_orders:
            quantity = materials_dict[order].get(material, 0)
            material_row[order] = quantity
            total_required += quantity
        
        # Добавляем итоговые колонки
        material_row['Требуется всего'] = total_required
        material_row['На складе'] = stock_data.get(material, 0)
        # ИЗМЕНЕНИЕ: Баланс = На складе - Требуется всего
        material_row['Баланс'] = material_row['На складе'] - total_required
        
        analysis_data.append(material_row)
    
    # Создаем DataFrame с правильным порядком колонок
    columns_order = ['Материал'] + all_orders + ['Требуется всего', 'На складе', 'Баланс']
    analysis_df = pd.DataFrame(analysis_data, columns=columns_order)
    
    # Сортируем по убыванию потребности
    analysis_df = analysis_df.sort_values('Требуется всего', ascending=False)
    
    print(f"Создан анализ для {len(analysis_df)} материалов с {len(all_orders)} заказами")
    
    return analysis_df

def save_final_report_with_grouping(final_df, materials_analysis_df):
    """Сохранение финального отчета с группировкой колонок заказов"""
    try:
        # Используем xlsxwriter для форматирования и группировки
        with pd.ExcelWriter('Объединенная_статистика_заказов.xlsx', engine='xlsxwriter') as writer:
            # Лист 1: Заказы
            final_df.to_excel(writer, sheet_name='Заказы', index=False)
            
            # Лист 2: Потребность материалов
            materials_analysis_df.to_excel(writer, sheet_name='Потребность материалов', index=False)
            
            # Получаем workbook для форматирования
            workbook = writer.book
            
            # Форматирование листа "Заказы"
            worksheet_orders = writer.sheets['Заказы']
            number_format = workbook.add_format({'num_format': '#,##0.00'})
            date_format = workbook.add_format({'num_format': 'dd.mm.yyyy'})
            text_format = workbook.add_format()
            
            # Определяем основные колонки
            main_columns = ['Номер заказа', 'Клиент', 'Тип продукции', 'Площадь заказа', 
                           'Дата создания', 'Дата завершения', 'Стоимость заказа', 
                           'Состояние заказа', 'Комментарий']
            
            # Форматируем колонки листа "Заказы"
            for col_idx, col_name in enumerate(final_df.columns):
                if col_name == 'Стоимость заказа' or col_name == 'Площадь заказа':
                    worksheet_orders.set_column(col_idx, col_idx, None, number_format)
                elif 'Дата' in col_name:
                    worksheet_orders.set_column(col_idx, col_idx, None, date_format)
                elif col_name not in main_columns:
                    worksheet_orders.set_column(col_idx, col_idx, None, number_format)
                else:
                    worksheet_orders.set_column(col_idx, col_idx, None, text_format)
                
                max_len = max(final_df[col_name].astype(str).str.len().max(), len(col_name)) + 2
                worksheet_orders.set_column(col_idx, col_idx, min(max_len, 50))
            
            # Форматирование листа "Потребность материалов"
            worksheet_analysis = writer.sheets['Потребность материалов']
            
            # Определяем индексы колонок для группировки
            # Колонка "Материал" = 0
            # Колонки заказов = от 1 до len(materials_analysis_df.columns) - 4
            # "Требуется всего" = len(materials_analysis_df.columns) - 3
            # "На складе" = len(materials_analysis_df.columns) - 2  
            # "Баланс" = len(materials_analysis_df.columns) - 1
            
            num_order_columns = len(materials_analysis_df.columns) - 4  # минус 4 основные колонки
            
            # Создаем форматы
            number_format = workbook.add_format({'num_format': '#,##0.00'})
            red_number_format = workbook.add_format({'num_format': '#,##0.00', 'font_color': 'red'})
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D7E4BC',
                'border': 1
            })
            
            # Форматируем колонки
            worksheet_analysis.set_column('A:A', 40)  # Материал - широкая колонка
            
            # Колонки заказов (группируемые) - числовой формат
            if num_order_columns > 0:
                start_col = 1  # B
                end_col = num_order_columns  # последняя колонка заказов
                worksheet_analysis.set_column(start_col, end_col, 12, number_format)  # Колонки заказов
            
            # Итоговые колонки
            total_col = num_order_columns + 1  # Колонка "Требуется всего"
            stock_col = num_order_columns + 2  # Колонка "На складе"  
            balance_col = num_order_columns + 3  # Колонка "Баланс"
            
            worksheet_analysis.set_column(total_col, total_col, 18, number_format)  # Требуется всего
            worksheet_analysis.set_column(stock_col, stock_col, 15, number_format)  # На складе
            worksheet_analysis.set_column(balance_col, balance_col, 15, number_format)  # Баланс
            
            # Добавляем условное форматирование для колонки "Баланс" (отрицательные значения красным)
            # ИЗМЕНЕНИЕ: теперь отрицательный баланс означает недостачу (Склад - Потребность < 0)
            last_row = len(materials_analysis_df)
            worksheet_analysis.conditional_format(
                balance_col, 1, balance_col, last_row,
                {
                    'type': 'cell',
                    'criteria': 'less than',
                    'value': 0,
                    'format': red_number_format
                }
            )
            
            # ГРУППИРОВКА: группируем колонки заказов
            if num_order_columns > 0:
                # Устанавливаем уровень для группировки (level 1)
                worksheet_analysis.set_column(1, num_order_columns, None, None, {'level': 1})
                # Сворачиваем группировку
                worksheet_analysis.outline_settings(1, True, False, True, True)
            
            # Добавляем заголовки
            for col_idx, col_name in enumerate(materials_analysis_df.columns):
                worksheet_analysis.write(0, col_idx, col_name, header_format)
        
        print("\n✓ Файл успешно сохранен как 'Объединенная_статистика_заказов.xlsx'")
        print("✓ Созданы листы:")
        print("  - 'Заказы': полная таблица заказов с материалами")
        print("  - 'Потребность материалов': анализ с группировкой по заказам")
        print("✓ Особенности листа 'Потребность материалов':")
        print("  - Колонки заказов сгруппированы и свернуты (можно развернуть по плюсику)")
        print("  - Колонка 'Баланс' рассчитывается как: На складе - Требуется всего")
        print("  - Отрицательные значения в колонке 'Баланс' выделены красным (недостача)")
        print("  - Положительные значения в колонке 'Баланс' - излишек материалов")
        print("✓ Применено форматирование:")
        print("  - Числовые значения: разделители тысяч, 2 знака после запятой")
        print("  - Даты: формат дд.мм.гггг")
        
    except Exception as e:
        print(f"Ошибка при сохранении файла: {e}")
        # Резервное сохранение без форматирования
        with pd.ExcelWriter('Объединенная_статистика_заказов.xlsx', engine='openpyxl') as writer:
            final_df.to_excel(writer, sheet_name='Заказы', index=False)
            materials_analysis_df.to_excel(writer, sheet_name='Потребность материалов', index=False)
        print("Файл сохранен без расширенного форматирования и группировки")

# Запуск скрипта
if __name__ == "__main__":
    result_df = merge_excel_tables()
