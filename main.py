import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import numpy as np
from collections import defaultdict
import os
import shutil
import re
from datetime import timedelta

# ВАЖНО: Вставьте сюда ВЕСЬ ваш код из uag.py
# Но убедитесь, что отступы правильные!
# Начните с функции merge_excel_tables() и всех вспомогательных функций

class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Объединение таблиц заказов")
        self.root.geometry("600x400")
        
        self.orders_file = tk.StringVar()
        self.materials1_file = tk.StringVar()
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
            
            # Импортируем и запускаем вашу функцию
            from uag import merge_excel_tables
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
