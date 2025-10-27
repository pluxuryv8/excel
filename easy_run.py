#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EASY RUN - Простой интерфейс для Excel Analytics
Просто вставь данные и получи готовый отчёт!
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import os
import sys
import subprocess
import tempfile
from pathlib import Path

class ExcelAnalyticsGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("📊 Excel Analytics PRO - Генератор отчётов")
        self.root.geometry("800x600")
        
        # Стиль
        style = ttk.Style()
        style.theme_use('clam')
        
        # Основной контейнер
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Заголовок
        title_label = ttk.Label(main_frame, text="Excel Analytics PRO", 
                               font=('Arial', 20, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=10)
        
        subtitle_label = ttk.Label(main_frame, 
                                  text="Вставь данные в два столбца (№ и значение) и получи профессиональный отчёт!",
                                  font=('Arial', 10))
        subtitle_label.grid(row=1, column=0, columnspan=3, pady=5)
        
        # Табы для нескольких выборок
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        # Создаём первую вкладку
        self.tabs = []
        self.add_data_tab("Выборка 1")
        
        # Кнопки управления вкладками
        tab_frame = ttk.Frame(main_frame)
        tab_frame.grid(row=3, column=0, columnspan=3, pady=5)
        
        ttk.Button(tab_frame, text="➕ Добавить выборку", 
                  command=self.add_tab).grid(row=0, column=0, padx=5)
        ttk.Button(tab_frame, text="➖ Удалить выборку", 
                  command=self.remove_tab).grid(row=0, column=1, padx=5)
        
        # Выбор папки для сохранения
        ttk.Label(main_frame, text="Папка для сохранения:").grid(row=4, column=0, sticky=tk.W, pady=10)
        
        self.output_path = tk.StringVar(value=str(Path.home() / "Desktop"))
        output_entry = ttk.Entry(main_frame, textvariable=self.output_path, width=50)
        output_entry.grid(row=4, column=1, padx=5)
        
        ttk.Button(main_frame, text="📁 Выбрать", 
                  command=self.choose_folder).grid(row=4, column=2)
        
        # Кнопки действий
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=5, column=0, columnspan=3, pady=20)
        
        self.generate_btn = ttk.Button(action_frame, text="🚀 СОЗДАТЬ ОТЧЁТ", 
                                      command=self.generate_report,
                                      style='Accent.TButton')
        self.generate_btn.grid(row=0, column=0, padx=10)
        
        ttk.Button(action_frame, text="📋 Пример данных", 
                  command=self.show_example).grid(row=0, column=1, padx=10)
        
        ttk.Button(action_frame, text="🗑️ Очистить всё", 
                  command=self.clear_all).grid(row=0, column=2, padx=10)
        
        # Статус
        self.status_var = tk.StringVar(value="Готов к работе")
        status_label = ttk.Label(main_frame, textvariable=self.status_var, 
                                relief=tk.SUNKEN, anchor=tk.W)
        status_label.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # Настройка весов для растягивания
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # Стиль для кнопки
        style.configure('Accent.TButton', font=('Arial', 12, 'bold'))
    
    def add_data_tab(self, name):
        """Добавляет новую вкладку для данных"""
        tab_frame = ttk.Frame(self.notebook)
        self.notebook.add(tab_frame, text=name)
        
        # Текстовое поле для данных
        text_widget = scrolledtext.ScrolledText(tab_frame, width=60, height=20, 
                                               font=('Courier', 10))
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Подсказка
        text_widget.insert('1.0', "# Вставь сюда данные в формате:\n# № <tab или пробел> значение\n\n1\t12.45\n2\t15.67\n3\t14.23\n4\t13.89\n5\t15.12\n")
        
        self.tabs.append({
            'frame': tab_frame,
            'text': text_widget,
            'name': name
        })
    
    def add_tab(self):
        """Добавляет новую вкладку"""
        num = len(self.tabs) + 1
        self.add_data_tab(f"Выборка {num}")
        self.notebook.select(len(self.tabs) - 1)
    
    def remove_tab(self):
        """Удаляет текущую вкладку"""
        if len(self.tabs) > 1:
            current = self.notebook.index(self.notebook.select())
            self.notebook.forget(current)
            self.tabs.pop(current)
        else:
            messagebox.showwarning("Внимание", "Должна остаться хотя бы одна выборка!")
    
    def choose_folder(self):
        """Выбор папки для сохранения"""
        folder = filedialog.askdirectory(initialdir=self.output_path.get())
        if folder:
            self.output_path.set(folder)
    
    def show_example(self):
        """Показывает пример данных"""
        example = """ПРИМЕР ДАННЫХ:

Формат: номер <пробел или tab> значение

1    100.71
2    100.56  
3    98.97
4    100.63
5    100.58
6    100.87
7    100.78
8    102.51
9    99.97
10   101.11

Можно использовать:
- Точку или запятую для дробной части
- Пробел или табуляцию как разделитель
- Комментарии начинаются с #"""
        
        messagebox.showinfo("Пример данных", example)
    
    def clear_all(self):
        """Очищает все поля"""
        if messagebox.askyesno("Подтверждение", "Очистить все данные?"):
            for tab in self.tabs:
                tab['text'].delete('1.0', tk.END)
            self.status_var.set("Все данные очищены")
    
    def generate_report(self):
        """Генерирует отчёт"""
        try:
            self.status_var.set("Подготовка данных...")
            self.generate_btn.config(state='disabled')
            
            # Создаём временные файлы для каждой выборки
            temp_files = []
            
            for i, tab in enumerate(self.tabs):
                data = tab['text'].get('1.0', tk.END).strip()
                if not data or data.startswith("# Вставь сюда"):
                    continue
                
                # Создаём временный файл
                temp_file = tempfile.NamedTemporaryFile(mode='w', suffix='.txt', 
                                                       delete=False, encoding='utf-8')
                temp_file.write(data)
                temp_file.close()
                temp_files.append(temp_file.name)
            
            if not temp_files:
                messagebox.showerror("Ошибка", "Нет данных для обработки!")
                return
            
            self.status_var.set(f"Обработка {len(temp_files)} выборок...")
            
            # Путь к основному скрипту
            script_path = os.path.join(os.path.dirname(__file__), 'report.py')
            
            # Папка для результатов
            output_dir = os.path.join(self.output_path.get(), 'Excel_Report')
            os.makedirs(output_dir, exist_ok=True)
            
            # Запускаем основной скрипт
            cmd = [sys.executable, script_path] + temp_files
            
            # Меняем рабочую директорию для вывода
            original_dir = os.getcwd()
            os.chdir(self.output_path.get())
            
            result = subprocess.run(cmd, capture_output=True, text=True)
            
            os.chdir(original_dir)
            
            # Удаляем временные файлы
            for f in temp_files:
                try:
                    os.unlink(f)
                except:
                    pass
            
            if result.returncode == 0:
                self.status_var.set("✅ Отчёт успешно создан!")
                output_file = os.path.join(self.output_path.get(), 'out', 'report_pro.xlsx')
                
                message = f"Отчёт успешно создан!\n\nФайл: {output_file}\n\nОткрыть папку с отчётом?"
                
                if messagebox.askyesno("Успех!", message):
                    # Открываем папку с результатом
                    if sys.platform == 'win32':
                        os.startfile(os.path.dirname(output_file))
                    elif sys.platform == 'darwin':
                        subprocess.run(['open', os.path.dirname(output_file)])
                    else:
                        subprocess.run(['xdg-open', os.path.dirname(output_file)])
            else:
                self.status_var.set("❌ Ошибка при создании отчёта")
                messagebox.showerror("Ошибка", f"Ошибка при создании отчёта:\n\n{result.stderr}")
                
        except Exception as e:
            self.status_var.set("❌ Произошла ошибка")
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n\n{str(e)}")
        finally:
            self.generate_btn.config(state='normal')


def main():
    # Проверяем наличие основного скрипта
    script_path = os.path.join(os.path.dirname(__file__), 'report.py')
    if not os.path.exists(script_path):
        messagebox.showerror("Ошибка", "Не найден файл report.py!\n\nУбедитесь, что easy_run.py находится в той же папке, что и report.py")
        return
    
    root = tk.Tk()
    app = ExcelAnalyticsGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()
