#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SUPER EASY - Максимально простой запуск
Просто запусти и следуй инструкциям!
"""

import os
import sys
import tempfile
import subprocess
from pathlib import Path

def print_header():
    print("\n" + "="*60)
    print("🚀 EXCEL ANALYTICS PRO - СУПЕР ПРОСТОЙ ЗАПУСК")
    print("="*60 + "\n")

def get_data_simple():
    """Простой ввод данных через консоль"""
    print("📊 ВСТАВЬ ДАННЫЕ (или напиши 'demo' для демо-данных):")
    print("Формат: номер пробел значение")
    print("Например:")
    print("1 12.45")
    print("2 15.67")
    print("...")
    print("\n⏹️  Когда закончишь, нажми Enter два раза\n")
    
    lines = []
    empty_count = 0
    
    while True:
        line = input()
        
        if line.lower() == 'demo':
            # Демо данные
            return """1 100.71
2 100.56
3 98.97
4 100.63
5 100.58
6 100.87
7 100.78
8 102.51
9 99.97
10 101.11
11 100.02"""
        
        if not line:
            empty_count += 1
            if empty_count >= 2:
                break
        else:
            empty_count = 0
            lines.append(line)
    
    return '\n'.join(lines)

def main():
    print_header()
    
    datasets = []
    dataset_count = 1
    
    while True:
        print(f"\n📈 ВЫБОРКА {dataset_count}:")
        data = get_data_simple()
        
        if data.strip():
            datasets.append(data)
            dataset_count += 1
            
            another = input("\n➕ Добавить ещё одну выборку? (да/нет): ").lower()
            if another not in ['да', 'д', 'yes', 'y']:
                break
        else:
            print("⚠️  Данные не введены!")
            continue
    
    if not datasets:
        print("\n❌ Нет данных для обработки!")
        return
    
    print(f"\n✅ Загружено выборок: {len(datasets)}")
    
    # Выбор папки
    print("\n📁 Куда сохранить отчёт?")
    print("1. На Рабочий стол (по умолчанию)")
    print("2. В текущую папку")
    print("3. Указать путь")
    
    choice = input("\nВыбор (1/2/3): ").strip() or "1"
    
    if choice == "1":
        output_dir = str(Path.home() / "Desktop" / "Excel_Report")
    elif choice == "2":
        output_dir = os.path.join(os.getcwd(), "Excel_Report")
    else:
        custom_path = input("Введи путь к папке: ").strip()
        output_dir = os.path.join(custom_path, "Excel_Report")
    
    os.makedirs(output_dir, exist_ok=True)
    
    print(f"\n📂 Результат будет сохранён в: {output_dir}")
    print("\n⏳ Создаю отчёт...")
    
    # Создаём временные файлы
    temp_files = []
    for i, data in enumerate(datasets):
        temp_file = tempfile.NamedTemporaryFile(mode='w', suffix='.txt', 
                                               delete=False, encoding='utf-8')
        temp_file.write(data)
        temp_file.close()
        temp_files.append(temp_file.name)
    
    # Запускаем основной скрипт
    script_path = os.path.join(os.path.dirname(__file__), 'report.py')
    
    # Меняем рабочую директорию
    original_dir = os.getcwd()
    os.chdir(os.path.dirname(output_dir))
    
    cmd = [sys.executable, script_path] + temp_files
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    os.chdir(original_dir)
    
    # Удаляем временные файлы
    for f in temp_files:
        try:
            os.unlink(f)
        except:
            pass
    
    if result.returncode == 0:
        output_file = os.path.join(os.path.dirname(output_dir), 'out', 'report_pro.xlsx')
        print("\n" + "="*60)
        print("✅ ГОТОВО!")
        print("="*60)
        print(f"\n📊 Отчёт создан: {output_file}")
        
        # Открываем папку
        if sys.platform == 'win32':
            os.startfile(os.path.dirname(output_file))
        elif sys.platform == 'darwin':
            subprocess.run(['open', os.path.dirname(output_file)])
        else:
            subprocess.run(['xdg-open', os.path.dirname(output_file)])
            
        print("\n🎉 Папка с отчётом открыта!")
    else:
        print("\n❌ Ошибка при создании отчёта:")
        print(result.stderr)
    
    input("\n\nНажми Enter для выхода...")

if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n👋 До встречи!")
    except Exception as e:
        print(f"\n❌ Ошибка: {e}")
        input("\nНажми Enter для выхода...")
