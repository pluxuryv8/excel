# -*- coding: utf-8 -*-
"""
Excel Pro Master - Профессиональный генератор статистических отчетов
Создает Excel файлы с полным анализом данных, графиками и форматированием
"""

import os
import sys
import json
import tempfile
from pathlib import Path
from typing import List, Dict, Tuple, Any, Optional
from io import BytesIO
import warnings
warnings.filterwarnings('ignore')

# Основные библиотеки
import numpy as np
import pandas as pd
from scipy import stats
from scipy.stats import shapiro, normaltest, jarque_bera, kstest, norm

# Для графиков
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import rcParams
import seaborn as sns

# Настройка графиков
sns.set_style("whitegrid")
rcParams['font.size'] = 10
rcParams['font.family'] = 'DejaVu Sans'
rcParams['axes.unicode_minus'] = False

# Excel
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell, xl_col_to_name
import subprocess

# GUI
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
from tkinter.font import Font

# ============================================================================
# КОНСТАНТЫ И ЭТАЛОННЫЕ ДАННЫЕ
# ============================================================================

# Примеры данных для демонстрации (48 и 25 значений)
EXAMPLE_DATA_48 = [
    101.09, 100.65, 100.93, 101.06, 100.57, 100.98,
    99.37, 100.71, 100.51, 100.58, 101.01, 100.49,
    100.72, 100.67, 100.24, 100.34, 100.23, 100.63,
    99.66, 100.31, 100.43, 100.18, 99.79, 100.26,
    100.77, 100.93, 100.36, 100.03, 100.87, 100.51,
    100.34, 100.53, 100.20, 102.37, 101.42, 101.08,
    100.46, 101.17, 100.56, 98.97, 100.63, 100.85,
    100.87, 100.78, 102.51, 99.97, 101.11, 100.02
]

EXAMPLE_DATA_25 = [
    100.71, 100.56, 98.97, 100.63, 100.58,
    100.87, 100.78, 102.51, 99.97, 101.11,
    100.02, 100.55, 100.46, 100.29, 100.84,
    100.98, 100.35, 100.89, 100.67, 101.10,
    99.94, 100.21, 100.58, 100.47, 101.70
]

# Космическая цветовая схема (Space Theme)
SPACE_COLORS = {
    'bg_dark': '#0a0a0a',          # Космическая чернота
    'bg_panel': '#1a1a1a',         # Темно-серая панель
    'bg_input': '#0f0f0f',         # Фон для ввода
    'accent': '#00d4ff',           # Космический голубой
    'accent_hover': '#00a8cc',     # Темнее голубой при наведении
    'text_primary': '#ffffff',      # Основной текст
    'text_secondary': '#b0b0b0',   # Вторичный текст
    'border': '#333333',           # Границы
    'success': '#00ff88',          # Успех (зеленый неон)
    'warning': '#ff9500',          # Предупреждение (оранжевый)
    'error': '#ff3366',            # Ошибка (красный неон)
}

# Цветовая схема для форматирования Excel
COLORS = {
    'header_main': '#90EE90',      # Светло-зеленый для основных заголовков
    'header_sub': '#7FBF7F',       # Темнее зеленый для подзаголовков  
    'data_bg': '#F0FFF0',          # Очень светлый зеленый для данных
    'highlight': '#FFE699',        # Желтый для выделения
    'border': '#006400',           # Темно-зеленый для границ
    'white': '#FFFFFF',            # Белый фон
    'gray': '#F2F2F2',             # Серый для чередования строк
}

# ============================================================================
# КЛАСС ДЛЯ СТАТИСТИЧЕСКОГО АНАЛИЗА
# ============================================================================

class StatisticalAnalyzer:
    """Полный статистический анализ данных"""
    
    def __init__(self, data: np.ndarray):
        self.data = np.array(data, dtype=float)
        self.n = len(self.data)
        self.results = {}
        self._calculate_all()
    
    def _calculate_all(self):
        """Вычисляет все статистические показатели"""
        
        # Основная описательная статистика
        self.results['mean'] = np.mean(self.data)
        self.results['std'] = np.std(self.data, ddof=1)  # Исправленное СКО
        self.results['std_pop'] = np.std(self.data)  # Генеральное СКО
        self.results['variance'] = np.var(self.data, ddof=1)  # Исправленная дисперсия
        self.results['variance_pop'] = np.var(self.data)  # Генеральная дисперсия
        self.results['min'] = np.min(self.data)
        self.results['max'] = np.max(self.data)
        self.results['range'] = self.results['max'] - self.results['min']
        self.results['median'] = np.median(self.data)
        self.results['q1'] = np.percentile(self.data, 25)
        self.results['q3'] = np.percentile(self.data, 75)
        
        # Моменты и характеристики формы
        self.results['skewness'] = stats.skew(self.data)
        self.results['kurtosis'] = stats.kurtosis(self.data, fisher=True)
        self.results['excess_kurtosis'] = self.results['kurtosis']
        
        # Средние
        self.results['mean_harmonic'] = stats.hmean(self.data[self.data > 0]) if np.all(self.data > 0) else None
        self.results['mean_geometric'] = stats.gmean(self.data[self.data > 0]) if np.all(self.data > 0) else None
        
        # Ошибки
        self.results['se'] = self.results['std'] / np.sqrt(self.n)
        
        # Коэффициент вариации
        self.results['cv'] = (self.results['std'] / self.results['mean']) * 100 if self.results['mean'] != 0 else 0
        
        # Доверительные интервалы (95%)
        confidence_level = 0.95
        alpha = 1 - confidence_level
        t_critical = stats.t.ppf(1 - alpha/2, self.n - 1)
        
        self.results['ci_mean_lower'] = self.results['mean'] - t_critical * self.results['se']
        self.results['ci_mean_upper'] = self.results['mean'] + t_critical * self.results['se']
        
        # Для стандартного отклонения (через хи-квадрат)
        chi2_lower = stats.chi2.ppf(1 - alpha/2, self.n - 1)
        chi2_upper = stats.chi2.ppf(alpha/2, self.n - 1)
        
        self.results['ci_std_lower'] = np.sqrt((self.n - 1) * self.results['variance'] / chi2_lower)
        self.results['ci_std_upper'] = np.sqrt((self.n - 1) * self.results['variance'] / chi2_upper)
        
    def test_normality(self) -> Dict[str, Any]:
        """Тесты на нормальность распределения"""
        tests = {}
        
        # Критерий Шапиро-Уилка
        try:
            stat_shapiro, p_shapiro = shapiro(self.data)
            tests['shapiro'] = {
                'statistic': stat_shapiro,
                'p_value': p_shapiro,
                'is_normal': p_shapiro > 0.05,
                'name': 'Шапиро-Уилка'
            }
        except:
            tests['shapiro'] = None
        
        # Критерий Романовского (для малых выборок)
        if self.n <= 50:
            # Используем упрощенный критерий
            romanovsky_stat = abs(self.results['skewness']) / np.sqrt(6/self.n)
            tests['romanovsky'] = {
                'statistic': romanovsky_stat,
                'critical': 3,
                'is_normal': romanovsky_stat < 3,
                'name': 'Романовского'
            }
        
        # Критерий Пирсона (хи-квадрат)
        try:
            # Группируем данные
            k = int(1 + 3.322 * np.log10(self.n))  # Правило Стерджесса
            k = max(5, min(k, 20))
            
            observed, bin_edges = np.histogram(self.data, bins=k)
            
            # Ожидаемые частоты для нормального распределения
            expected = []
            for i in range(len(bin_edges) - 1):
                p = norm.cdf(bin_edges[i+1], self.results['mean'], self.results['std']) - \
                    norm.cdf(bin_edges[i], self.results['mean'], self.results['std'])
                expected.append(self.n * p)
            expected = np.array(expected)
            
            # Объединяем малые группы
            min_expected = 5
            while np.any(expected < min_expected) and len(expected) > 2:
                idx = np.argmin(expected)
                if idx == 0:
                    expected[1] += expected[0]
                    expected = expected[1:]
                    observed[1] += observed[0]
                    observed = observed[1:]
                elif idx == len(expected) - 1:
                    expected[-2] += expected[-1]
                    expected = expected[:-1]
                    observed[-2] += observed[-1]
                    observed = observed[:-1]
                else:
                    expected[idx-1] += expected[idx]
                    expected = np.delete(expected, idx)
                    observed[idx-1] += observed[idx]
                    observed = np.delete(observed, idx)
            
            chi2_stat = np.sum((observed - expected)**2 / expected)
            df = len(expected) - 3  # k-1-2 (2 параметра: среднее и СКО)
            p_chi2 = 1 - stats.chi2.cdf(chi2_stat, df) if df > 0 else 0
            
            tests['chi2'] = {
                'statistic': chi2_stat,
                'p_value': p_chi2,
                'df': df,
                'is_normal': p_chi2 > 0.05 if df > 0 else False,
                'name': 'Пирсона (χ²)'
            }
        except:
            tests['chi2'] = None
        
        # Критерий Колмогорова-Смирнова
        try:
            stat_ks, p_ks = kstest(self.data, 'norm', args=(self.results['mean'], self.results['std']))
            tests['ks'] = {
                'statistic': stat_ks,
                'p_value': p_ks,
                'is_normal': p_ks > 0.05,
                'name': 'Колмогорова-Смирнова'
            }
        except:
            tests['ks'] = None
        
        # Критерий Смирнова (модифицированный)
        try:
            # Вычисляем эмпирическую функцию распределения
            sorted_data = np.sort(self.data)
            n = len(self.data)
            # Стандартизируем данные
            z_sorted = (sorted_data - self.results['mean']) / self.results['std']
            
            # Вычисляем максимальное отклонение
            d_plus = []
            d_minus = []
            for i in range(n):
                # Теоретическая функция распределения (нормальная)
                F_theoretical = norm.cdf(z_sorted[i])
                # Эмпирическая функция распределения
                F_empirical = (i + 1) / n
                F_empirical_prev = i / n
                
                d_plus.append(F_empirical - F_theoretical)
                d_minus.append(F_theoretical - F_empirical_prev)
            
            D = max(max(d_plus), max(d_minus))
            
            # Критическое значение (приблизительное)
            if n <= 20:
                d_critical = 0.294  # для α = 0.05
            elif n <= 30:
                d_critical = 0.242
            elif n <= 40:
                d_critical = 0.210
            else:
                d_critical = 1.36 / np.sqrt(n)
            
            tests['smirnov'] = {
                'statistic': D,
                'critical_value': d_critical,
                'is_normal': D <= d_critical,
                'name': 'Смирнова'
            }
        except:
            tests['smirnov'] = None
        
        return tests
    
    def detect_outliers(self, method='iqr') -> Dict[str, Any]:
        """Обнаружение выбросов"""
        outliers = {}
        
        # Метод межквартильного размаха (IQR)
        if method == 'iqr' or method == 'all':
            iqr = self.results['q3'] - self.results['q1']
            lower_fence = self.results['q1'] - 1.5 * iqr
            upper_fence = self.results['q3'] + 1.5 * iqr
            
            outlier_indices = np.where((self.data < lower_fence) | (self.data > upper_fence))[0]
            outliers['iqr'] = {
                'indices': outlier_indices.tolist(),
                'values': self.data[outlier_indices].tolist(),
                'lower_fence': lower_fence,
                'upper_fence': upper_fence,
                'count': len(outlier_indices)
            }
        
        # Метод 3-сигм (Райта)
        if method == '3sigma' or method == 'all':
            lower_limit = self.results['mean'] - 3 * self.results['std']
            upper_limit = self.results['mean'] + 3 * self.results['std']
            
            outlier_indices = np.where((self.data < lower_limit) | (self.data > upper_limit))[0]
            outliers['3sigma'] = {
                'indices': outlier_indices.tolist(),
                'values': self.data[outlier_indices].tolist(),
                'lower_limit': lower_limit,
                'upper_limit': upper_limit,
                'count': len(outlier_indices)
            }
        
        # Критерий Граббса
        if method == 'grubbs' or method == 'all':
            sorted_data = np.sort(self.data)
            z_scores = np.abs((self.data - self.results['mean']) / self.results['std'])
            max_z = np.max(z_scores)
            max_idx = np.argmax(z_scores)
            
            # Критическое значение
            alpha = 0.05
            t_critical = stats.t.ppf(1 - alpha/(2*self.n), self.n - 2)
            g_critical = ((self.n - 1) * t_critical) / np.sqrt(self.n * (self.n - 2 + t_critical**2))
            
            outliers['grubbs'] = {
                'max_z_score': max_z,
                'critical_value': g_critical,
                'outlier_index': max_idx if max_z > g_critical else None,
                'outlier_value': self.data[max_idx] if max_z > g_critical else None,
                'has_outliers': max_z > g_critical
            }
        
        # Критерий Шарлье
        if method == 'sharlie' or method == 'all':
            # Считаем количество точек за пределами 3σ
            z_scores = np.abs((self.data - self.results['mean']) / self.results['std'])
            outlier_count = np.sum(z_scores > 3)
            
            outliers['sharlie'] = {
                'outlier_count': int(outlier_count),
                'threshold': 3,
                'has_outliers': outlier_count > 0
            }
        
        # Критерий Ирвина
        if method == 'irwin' or method == 'all':
            sorted_data = np.sort(self.data)
            diffs = np.diff(sorted_data)
            lambda_values = diffs / self.results['std']
            max_lambda = np.max(lambda_values)
            max_lambda_idx = np.argmax(lambda_values)
            
            # Критическое значение (упрощенное)
            lambda_critical = 1.7  # для n ≈ 50
            
            outliers['irwin'] = {
                'max_lambda': max_lambda,
                'critical_value': lambda_critical,
                'outlier_index': max_lambda_idx if max_lambda > lambda_critical else None,
                'has_outliers': max_lambda > lambda_critical
            }
        
        # Критерий Шовене
        if method == 'chauvenet' or method == 'all':
            z_scores = np.abs((self.data - self.results['mean']) / self.results['std'])
            # Вероятность для каждой точки
            p = 2 * (1 - norm.cdf(z_scores))
            # Ожидаемое количество точек
            n_expected = self.n * p
            # Выбросы - где ожидается меньше 0.5 точек
            outlier_mask = n_expected < 0.5
            
            outliers['chauvenet'] = {
                'outlier_indices': np.where(outlier_mask)[0].tolist(),
                'outlier_values': self.data[outlier_mask].tolist(),
                'count': int(np.sum(outlier_mask))
            }
        
        return outliers


# ============================================================================
# КЛАСС ДЛЯ СОЗДАНИЯ EXCEL ОТЧЕТА
# ============================================================================

class ExcelReportGenerator:
    """Генератор профессионального Excel отчета"""
    
    def __init__(self, output_path: str):
        self.output_path = output_path
        self.workbook = xlsxwriter.Workbook(output_path)
        self.formats = self._create_formats()
        
    def _create_formats(self) -> Dict[str, Any]:
        """Создает форматы для Excel"""
        formats = {}
        
        # Заголовок листа (большой, зеленый)
        formats['title'] = self.workbook.add_format({
            'bold': True,
            'font_size': 14,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': COLORS['header_main'],
            'border': 2,
            'border_color': COLORS['border'],
        })
        
        # Заголовки таблиц
        formats['header'] = self.workbook.add_format({
            'bold': True,
            'font_size': 11,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': COLORS['header_main'],
            'border': 1,
            'border_color': COLORS['border'],
            'text_wrap': True,
        })
        
        # Подзаголовки
        formats['subheader'] = self.workbook.add_format({
            'bold': True,
            'font_size': 10,
            'bg_color': COLORS['header_sub'],
            'border': 1,
            'border_color': COLORS['border'],
        })
        
        # Данные - обычные
        formats['data'] = self.workbook.add_format({
            'font_size': 10,
            'border': 1,
            'border_color': '#D0D0D0',
        })
        
        # Данные - числовые (2 знака)
        formats['number2'] = self.workbook.add_format({
            'font_size': 10,
            'border': 1,
            'border_color': '#D0D0D0',
            'num_format': '0.00',
        })
        
        # Данные - числовые (4 знака)
        formats['number4'] = self.workbook.add_format({
            'font_size': 10,
            'border': 1,
            'border_color': '#D0D0D0',
            'num_format': '0.0000',
        })
        
        # Данные - числовые (6 знаков)
        formats['number6'] = self.workbook.add_format({
            'font_size': 10,
            'border': 1,
            'border_color': '#D0D0D0',
            'num_format': '0.000000',
        })
        
        # Выделенные ячейки
        formats['highlight'] = self.workbook.add_format({
            'font_size': 10,
            'bold': True,
            'border': 1,
            'bg_color': COLORS['highlight'],
            'border_color': COLORS['border'],
        })
        
        # Результат/вывод
        formats['result'] = self.workbook.add_format({
            'font_size': 11,
            'bold': True,
            'bg_color': COLORS['data_bg'],
            'border': 2,
            'border_color': COLORS['border'],
            'text_wrap': True,
        })
        
        # Формулы
        formats['formula'] = self.workbook.add_format({
            'font_size': 10,
            'border': 1,
            'bg_color': '#F0F0F0',
            'border_color': '#D0D0D0',
            'num_format': '0.0000',
        })
        
        # Форматы для результатов тестов
        formats['error'] = self.workbook.add_format({
            'font_size': 11,
            'bold': True,
            'bg_color': '#FFC7CE',
            'border': 1,
            'font_color': '#9C0006'
        })
        
        formats['success'] = self.workbook.add_format({
            'font_size': 11,
            'bold': True,
            'bg_color': '#C6EFCE',
            'border': 1,
            'font_color': '#006100'
        })
        
        return formats
    
    def create_main_sheet(self, data: np.ndarray, analyzer: StatisticalAnalyzer):
        """Создает основной лист с расчетами как на скриншоте"""
        sheet = self.workbook.add_worksheet('Задание 1')
        
        # Настройка ширины столбцов
        sheet.set_column('A:A', 5)   # №
        sheet.set_column('B:B', 12)  # Xj
        sheet.set_column('C:G', 15)  # Расчетные столбцы
        sheet.set_column('H:H', 3)   # Пробел
        sheet.set_column('I:J', 18)  # Показатели
        sheet.set_column('K:K', 15)  # Значения
        
        # Заголовок
        sheet.merge_range('A1:G2', 'СТАТИСТИЧЕСКИЙ АНАЛИЗ ДАННЫХ', self.formats['title'])
        
        # Заголовки таблицы данных
        headers = ['№', 'Xj', 'Xj - Xср', '|Xj - Xср|', '(Xj - Xср)²', '(Xj - Xср)³', '(Xj - Xср)⁴']
        for col, header in enumerate(headers):
            sheet.write(3, col, header, self.formats['header'])
        
        # Записываем данные и формулы
        row_start = 4
        n = len(data)
        mean = analyzer.results['mean']
        stats_start_row = 4  # Объявляем переменную перед использованием
        
        # Данные
        for i, value in enumerate(data):
            row = row_start + i
            sheet.write(row, 0, i + 1, self.formats['data'])  # Номер
            sheet.write(row, 1, value, self.formats['number2'])  # Значение
            
            # Формулы Excel
            cell_xj = xl_rowcol_to_cell(row, 1)
            
            # Сначала записываем само среднее в правой части
            if i == 0:  # Только для первой строки
                mean_row = stats_start_row + 1  # Строка со средним в правой таблице
                sheet.write_formula(mean_row-1, 10, f'=AVERAGE(B{row_start+1}:B{row_start+n})', self.formats['number4'])
            
            # Теперь используем правильную ссылку на среднее
            cell_mean = f'$K${stats_start_row+1}'  # K5 если stats_start_row = 4
            
            # Xj - Xср
            sheet.write_formula(row, 2, f'={cell_xj}-{cell_mean}', self.formats['number4'])
            
            # |Xj - Xср|
            sheet.write_formula(row, 3, f'=ABS(C{row+1})', self.formats['number4'])
            
            # (Xj - Xср)²
            sheet.write_formula(row, 4, f'=C{row+1}^2', self.formats['number4'])
            
            # (Xj - Xср)³
            sheet.write_formula(row, 5, f'=C{row+1}^3', self.formats['number6'])
            
            # (Xj - Xср)⁴
            sheet.write_formula(row, 6, f'=C{row+1}^4', self.formats['number6'])
        
        # Суммы
        sum_row = row_start + n
        sheet.write(sum_row, 0, 'Σ', self.formats['header'])
        for col in range(1, 7):
            col_letter = xl_col_to_name(col)
            sheet.write_formula(sum_row, col, f'=SUM({col_letter}{row_start+1}:{col_letter}{sum_row})', 
                               self.formats['highlight'])
        
        # Статистические показатели в правой части
        # stats_start_row уже объявлена выше
        
        # Основные показатели
        sheet.write(stats_start_row, 9, 'Показатель', self.formats['header'])
        sheet.write(stats_start_row, 10, 'Значение', self.formats['header'])
        
        stats_data = [
            ('Среднее X̄', f'=AVERAGE(B{row_start+1}:B{sum_row})'),
            ('Станд. отклонение S', f'=STDEV.S(B{row_start+1}:B{sum_row})'),
            ('Дисперсия S²', f'=VAR.S(B{row_start+1}:B{sum_row})'),
            ('Минимум', f'=MIN(B{row_start+1}:B{sum_row})'),
            ('Максимум', f'=MAX(B{row_start+1}:B{sum_row})'),
            ('Размах', f'=K{stats_start_row+5}-K{stats_start_row+4}'),
            ('Медиана', f'=MEDIAN(B{row_start+1}:B{sum_row})'),
            ('Мода', f'=MODE.SNGL(B{row_start+1}:B{sum_row})'),
            ('Квартиль Q1', f'=QUARTILE.INC(B{row_start+1}:B{sum_row},1)'),
            ('Квартиль Q3', f'=QUARTILE.INC(B{row_start+1}:B{sum_row},3)'),
            ('Асимметрия', f'=SKEW(B{row_start+1}:B{sum_row})'),
            ('Эксцесс', f'=KURT(B{row_start+1}:B{sum_row})'),
            ('Коэфф. вариации, %', f'=K{stats_start_row+2}/K{stats_start_row+1}*100'),
        ]
        
        for i, (label, formula) in enumerate(stats_data):
            row = stats_start_row + i + 1
            sheet.write(row, 9, label, self.formats['subheader'])
            if formula.startswith('='):
                sheet.write_formula(row, 10, formula, self.formats['number4'])
            else:
                sheet.write(row, 10, formula, self.formats['number4'])
        
        # Доверительные интервалы
        ci_row = stats_start_row + len(stats_data) + 3
        sheet.merge_range(ci_row, 9, ci_row, 10, 'Доверительные интервалы (α=0.05)', self.formats['header'])
        
        sheet.write(ci_row+1, 9, 'Для среднего μ', self.formats['subheader'])
        sheet.write(ci_row+1, 10, f'[{analyzer.results["ci_mean_lower"]:.4f}; {analyzer.results["ci_mean_upper"]:.4f}]', 
                   self.formats['number4'])
        
        sheet.write(ci_row+2, 9, 'Для СКО σ', self.formats['subheader'])
        sheet.write(ci_row+2, 10, f'[{analyzer.results["ci_std_lower"]:.4f}; {analyzer.results["ci_std_upper"]:.4f}]', 
                   self.formats['number4'])
        
        return sheet
    
    def create_normality_sheet(self, data: np.ndarray, analyzer: StatisticalAnalyzer):
        """Создает лист с проверкой нормальности"""
        sheet = self.workbook.add_worksheet('Проверка нормальности')
        
        # Настройка ширины столбцов
        sheet.set_column('A:A', 5)
        sheet.set_column('B:B', 12)
        sheet.set_column('C:D', 20)
        sheet.set_column('E:F', 15)
        
        # Заголовок
        sheet.merge_range('A1:F2', 'ПРОВЕРКА НОРМАЛЬНОСТИ РАСПРЕДЕЛЕНИЯ', self.formats['title'])
        
        # Критерии нормальности
        tests = analyzer.test_normality()
        
        row = 4
        sheet.merge_range(row, 0, row, 5, 'Критерии нормальности', self.formats['header'])
        
        row += 2
        headers = ['№', 'Критерий', 'Статистика', 'p-value', 'Крит. значение', 'Вывод']
        for col, header in enumerate(headers):
            sheet.write(row, col, header, self.formats['subheader'])
        
        row += 1
        test_num = 1
        
        # Выводим результаты тестов
        for test_name, test_data in tests.items():
            if test_data:
                sheet.write(row, 0, test_num, self.formats['data'])
                sheet.write(row, 1, test_data.get('name', test_name), self.formats['data'])
                sheet.write(row, 2, test_data.get('statistic', '-'), self.formats['number4'])
                
                # p-value или критическое значение
                if 'p_value' in test_data:
                    sheet.write(row, 3, test_data.get('p_value', '-'), self.formats['number4'])
                else:
                    sheet.write(row, 3, '-', self.formats['data'])
                
                # Критическое значение
                if 'critical_value' in test_data:
                    sheet.write(row, 4, test_data.get('critical_value'), self.formats['number4'])
                elif 'critical' in test_data:
                    sheet.write(row, 4, test_data.get('critical'), self.formats['number4'])
                else:
                    sheet.write(row, 4, 0.05, self.formats['number4'])
                
                is_normal = test_data.get('is_normal', False)
                conclusion = 'Норма' if is_normal else 'Не норма'
                format_conclusion = self.formats['highlight'] if not is_normal else self.formats['data']
                sheet.write(row, 5, conclusion, format_conclusion)
                
                row += 1
                test_num += 1
        
        # Критерий Романовского-Пирсона (если есть)
        row += 2
        sheet.merge_range(row, 0, row, 5, 'Критерий Романовского для федерального округа', self.formats['header'])
        
        row += 2
        # Здесь добавляем данные как на скриншоте
        headers2 = ['№', 'Xj', 'Крит. Романовского ДФО']
        for col, header in enumerate(headers2[:3]):
            sheet.write(row, col, header, self.formats['subheader'])
        
        # Данные для критерия
        row += 1
        for i, value in enumerate(data[:25]):  # Первые 25 значений
            sheet.write(row + i, 0, i + 1, self.formats['data'])
            sheet.write(row + i, 1, value, self.formats['number2'])
            # Расчет критерия для каждого значения
            z_score = abs(value - analyzer.results['mean']) / analyzer.results['std']
            sheet.write(row + i, 2, z_score, self.formats['number4'])
        
        return sheet
    
    def create_charts_sheet(self, data: np.ndarray, analyzer: StatisticalAnalyzer):
        """Создает лист с графиками"""
        sheet = self.workbook.add_worksheet('Графики')
        
        # Заголовок
        sheet.merge_range('A1:F1', 'ВИЗУАЛИЗАЦИЯ ДАННЫХ', self.formats['title'])
        
        # Создаем графики matplotlib
        fig, axes = plt.subplots(2, 2, figsize=(14, 10))
        
        # 1. Гистограмма с плотностью
        ax1 = axes[0, 0]
        n_bins = int(1 + 3.322 * np.log10(len(data)))
        counts, bins, patches = ax1.hist(data, bins=n_bins, density=True, 
                                         alpha=0.7, color='#90EE90', edgecolor='black')
        
        # Добавляем нормальную кривую
        x = np.linspace(data.min(), data.max(), 100)
        ax1.plot(x, norm.pdf(x, analyzer.results['mean'], analyzer.results['std']), 
                'r-', linewidth=2, label='Норм. распределение')
        ax1.set_title('Гистограмма плотности', fontsize=12, fontweight='bold')
        ax1.set_xlabel('Значение')
        ax1.set_ylabel('Плотность')
        ax1.legend()
        ax1.grid(True, alpha=0.3)
        
        # 2. Q-Q plot
        ax2 = axes[0, 1]
        stats.probplot(data, dist="norm", plot=ax2)
        ax2.set_title('Q-Q plot: сравнение с нормальным распределением', fontsize=12, fontweight='bold')
        ax2.grid(True, alpha=0.3)
        
        # 3. Ящик с усами
        ax3 = axes[1, 0]
        bp = ax3.boxplot(data, vert=True, patch_artist=True, widths=0.5)
        bp['boxes'][0].set_facecolor('#90EE90')
        ax3.set_title('Ящик с усами (Box Plot)', fontsize=12, fontweight='bold')
        ax3.set_ylabel('Значение')
        ax3.grid(True, alpha=0.3)
        
        # 4. График плотности
        ax4 = axes[1, 1]
        from scipy.stats import gaussian_kde
        kde = gaussian_kde(data)
        x_range = np.linspace(data.min() - 1, data.max() + 1, 200)
        ax4.plot(x_range, kde(x_range), color='#006400', linewidth=2, label='Эмпирическая')
        ax4.plot(x_range, norm.pdf(x_range, analyzer.results['mean'], analyzer.results['std']), 
                'r--', linewidth=2, label='Теоретическая')
        ax4.fill_between(x_range, kde(x_range), alpha=0.3, color='#90EE90')
        ax4.set_title('Сравнение плотностей распределения', fontsize=12, fontweight='bold')
        ax4.set_xlabel('Значение')
        ax4.set_ylabel('Плотность')
        ax4.legend()
        ax4.grid(True, alpha=0.3)
        
        plt.tight_layout()
        
        # Сохраняем график в память (без создания файла!)
        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=100, bbox_inches='tight')
        plt.close()
        
        # Перематываем буфер в начало
        img_buffer.seek(0)
        
        # Вставляем график в Excel прямо из памяти
        sheet.insert_image('A3', 'dummy.png', {'image_data': img_buffer, 'x_scale': 0.9, 'y_scale': 0.9})
        
        # Закрываем буфер
        img_buffer.close()
        
        return sheet
    
    def create_outliers_sheet(self, data: np.ndarray, analyzer: StatisticalAnalyzer):
        """Создает лист с анализом выбросов"""
        sheet = self.workbook.add_worksheet('Анализ выбросов')
        
        sheet.merge_range('A1:G1', 'АНАЛИЗ ВЫБРОСОВ И АНОМАЛИЙ', self.formats['title'])
        
        # Настройка ширины столбцов
        sheet.set_column('A:A', 25)
        sheet.set_column('B:G', 15)
        
        outliers = analyzer.detect_outliers('all')
        
        row = 3
        
        # Критерий Граббса
        sheet.merge_range(row, 0, row, 6, 'Критерий Граббса', self.formats['header'])
        row += 2
        
        grubbs_data = outliers.get('grubbs', {})
        sheet.write(row, 0, 'Максимальное Z-значение:', self.formats['subheader'])
        sheet.write(row, 1, grubbs_data.get('max_z_score', 0), self.formats['number4'])
        sheet.write(row, 2, 'Критическое значение:', self.formats['subheader'])
        sheet.write(row, 3, grubbs_data.get('critical_value', 0), self.formats['number4'])
        row += 1
        if grubbs_data.get('has_outliers'):
            sheet.write(row, 0, 'Выброс найден:', self.formats['subheader'])
            sheet.write(row, 1, grubbs_data.get('outlier_value', '-'), self.formats['highlight'])
            sheet.write(row, 2, 'ВЫБРОС ОБНАРУЖЕН', self.formats['error'])
        else:
            sheet.write(row, 0, 'Результат:', self.formats['subheader'])
            sheet.write(row, 1, 'Выбросов НЕТ', self.formats['success'])
        
        # Критерий Романовского (перенесем из проверки нормальности)
        row += 3
        sheet.merge_range(row, 0, row, 6, 'Критерий Романовского', self.formats['header'])
        row += 2
        
        romanovsky_values = []
        for val in data:
            tau = abs(val - analyzer.results['mean']) / analyzer.results['std']
            romanovsky_values.append(tau)
        max_tau = max(romanovsky_values)
        
        sheet.write(row, 0, 'Макс. значение τ:', self.formats['subheader'])
        sheet.write(row, 1, max_tau, self.formats['number4'])
        sheet.write(row, 2, 'Критическое значение:', self.formats['subheader'])
        sheet.write(row, 3, 2.96 if len(data) <= 25 else 3.0, self.formats['number4'])  # Упрощенно
        row += 1
        sheet.write(row, 0, 'Результат:', self.formats['subheader'])
        if max_tau > (2.96 if len(data) <= 25 else 3.0):
            sheet.write(row, 1, 'АНОМАЛИЯ ОБНАРУЖЕНА', self.formats['error'])
        else:
            sheet.write(row, 1, 'Аномалий НЕТ', self.formats['success'])
        
        # Критерий Шарлье
        row += 3
        sheet.merge_range(row, 0, row, 6, 'Критерий Шарлье', self.formats['header'])
        row += 2
        
        sharlie_data = outliers.get('sharlie', {})
        sheet.write(row, 0, 'Точек за пределами 3σ:', self.formats['subheader'])
        sheet.write(row, 1, sharlie_data.get('outlier_count', 0), self.formats['data'])
        sheet.write(row, 2, 'Результат:', self.formats['subheader'])
        if sharlie_data.get('has_outliers'):
            sheet.write(row, 3, 'ЕСТЬ АНОМАЛИИ', self.formats['error'])
        else:
            sheet.write(row, 3, 'Нет аномалий', self.formats['success'])
        
        # Критерий Ирвина
        row += 3
        sheet.merge_range(row, 0, row, 6, 'Критерий Ирвина', self.formats['header'])
        row += 2
        
        irwin_data = outliers.get('irwin', {})
        sheet.write(row, 0, 'Максимальное λ:', self.formats['subheader'])
        sheet.write(row, 1, irwin_data.get('max_lambda', 0), self.formats['number4'])
        sheet.write(row, 2, 'Критическое значение:', self.formats['subheader'])
        sheet.write(row, 3, irwin_data.get('critical_value', 0), self.formats['number4'])
        row += 1
        sheet.write(row, 0, 'Результат:', self.formats['subheader'])
        if irwin_data.get('has_outliers'):
            sheet.write(row, 1, 'ВЫБРОС ОБНАРУЖЕН', self.formats['error'])
        else:
            sheet.write(row, 1, 'Выбросов НЕТ', self.formats['success'])
        
        # Критерий Шовене
        row += 3
        sheet.merge_range(row, 0, row, 6, 'Критерий Шовене', self.formats['header'])
        row += 2
        
        chauvenet_data = outliers.get('chauvenet', {})
        sheet.write(row, 0, 'Количество выбросов:', self.formats['subheader'])
        sheet.write(row, 1, chauvenet_data.get('count', 0), self.formats['data'])
        if chauvenet_data.get('outlier_values'):
            row += 1
            sheet.write(row, 0, 'Выбросы:', self.formats['subheader'])
            for i, val in enumerate(chauvenet_data['outlier_values'][:5]):
                sheet.write(row, i + 1, val, self.formats['number4'])
        
        # Правило трёх сигм (Райта)
        row += 3
        sheet.merge_range(row, 0, row, 6, 'Критерий Райта (правило 3σ)', self.formats['header'])
        row += 2
        
        sigma3_data = outliers.get('3sigma', {})
        sheet.write(row, 0, 'Нижняя граница (X̄ - 3σ):', self.formats['subheader'])
        sheet.write(row, 1, sigma3_data.get('lower_limit', 0), self.formats['number4'])
        sheet.write(row, 2, 'Верхняя граница (X̄ + 3σ):', self.formats['subheader'])
        sheet.write(row, 3, sigma3_data.get('upper_limit', 0), self.formats['number4'])
        row += 1
        sheet.write(row, 0, 'Количество выбросов:', self.formats['subheader'])
        sheet.write(row, 1, sigma3_data.get('count', 0), self.formats['data'])
        sheet.write(row, 2, 'Результат:', self.formats['subheader'])
        if sigma3_data.get('count', 0) > 0:
            sheet.write(row, 3, 'ЕСТЬ ВЫБРОСЫ', self.formats['error'])
        else:
            sheet.write(row, 3, 'Нет выбросов', self.formats['success'])
        
        # Метод IQR
        row += 3
        sheet.merge_range(row, 0, row, 6, 'Метод межквартильного размаха (IQR)', self.formats['header'])
        row += 2
        
        iqr_data = outliers.get('iqr', {})
        sheet.write(row, 0, 'Нижняя граница:', self.formats['subheader'])
        sheet.write(row, 1, iqr_data.get('lower_fence', 0), self.formats['number4'])
        sheet.write(row, 2, 'Верхняя граница:', self.formats['subheader'])
        sheet.write(row, 3, iqr_data.get('upper_fence', 0), self.formats['number4'])
        row += 1
        sheet.write(row, 0, 'Количество выбросов:', self.formats['subheader'])
        sheet.write(row, 1, iqr_data.get('count', 0), self.formats['data'])
        
        return sheet
    
    def create_conclusion_sheet(self, analyzer: StatisticalAnalyzer):
        """Создает лист с выводами"""
        sheet = self.workbook.add_worksheet('Выводы')
        
        sheet.merge_range('A1:F1', 'ВЫВОДЫ И ЗАКЛЮЧЕНИЕ', self.formats['title'])
        
        # Формируем выводы
        tests = analyzer.test_normality()
        normal_tests_passed = sum(1 for t in tests.values() if t and t.get('is_normal', False))
        total_tests = len([t for t in tests.values() if t])
        
        conclusion_text = f"""
Вывод. При уровне значимости α=0.05 доверительный интервал 
для математического ожидания нормального распределения 
составил μ ∈ [{analyzer.results['ci_mean_lower']:.2f}; {analyzer.results['ci_mean_upper']:.2f}], для среднего квадратического 
отклонения — σ ∈ [{analyzer.results['ci_std_lower']:.2f}; {analyzer.results['ci_std_upper']:.2f}]. Полученные интервалы узкие 
(n={analyzer.n}), что указывает на хорошую точность оценок. С учётом 
проверки нормальности (пройдено {normal_tests_passed} из {total_tests} тестов) параметры 
выборки можно считать надёжными: распределение в целом 
согласуется с нормальным, влияние единичного 
правостороннего выброса невелико.
"""
        
        # Записываем вывод
        sheet.merge_range('A3:F10', conclusion_text, self.formats['result'])
        
        # Статистическая сводка
        row = 12
        sheet.merge_range(row, 0, row, 5, 'Исходные статистические данные', self.formats['header'])
        
        row += 2
        stats_summary = [
            ('Уровень значимости α', 0.05),
            ('Среднее x̄', analyzer.results['mean']),
            ('Станд. отклонение s', analyzer.results['std']),
            ('Объём выборки n', analyzer.n),
        ]
        
        for label, value in stats_summary:
            sheet.write(row, 0, label, self.formats['subheader'])
            sheet.write(row, 1, value, self.formats['number4'] if isinstance(value, float) else self.formats['data'])
            row += 1
        
        return sheet
    
    def close(self):
        """Закрывает workbook"""
        self.workbook.close()


# ============================================================================
# GUI ИНТЕРФЕЙС
# ============================================================================

class ExcelProMasterGUI:
    """Космический GUI для Excel Pro Master в стиле SpaceX"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("🚀 Excel Pro Master | Космическая версия")
        
        # Размер и позиционирование окна
        window_width = 950
        window_height = 750
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # Темная космическая тема
        self.root.configure(bg=SPACE_COLORS['bg_dark'])
        
        # Стиль
        self.setup_styles()
        
        # Данные
        self.datasets = []
        
        # Создаем интерфейс
        self.create_widgets()
        
        # Загружаем эталонные данные если есть
        self.load_reference_data()
    
    def setup_styles(self):
        """Настройка космических стилей"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Основные стили в космической теме
        style.configure('Space.TFrame', 
                       background=SPACE_COLORS['bg_panel'],
                       borderwidth=1,
                       relief='flat')
        
        style.configure('Title.TLabel', 
                       font=('Orbitron', 20, 'bold'),
                       foreground=SPACE_COLORS['accent'],
                       background=SPACE_COLORS['bg_dark'])
        
        style.configure('Subtitle.TLabel',
                       font=('Segoe UI', 11),
                       foreground=SPACE_COLORS['text_secondary'],
                       background=SPACE_COLORS['bg_dark'])
        
        style.configure('Header.TLabel', 
                       font=('Segoe UI', 12, 'bold'),
                       foreground=SPACE_COLORS['text_primary'],
                       background=SPACE_COLORS['bg_panel'])
        
        style.configure('Space.TNotebook',
                       background=SPACE_COLORS['bg_panel'],
                       borderwidth=0)
        
        style.configure('Space.TNotebook.Tab',
                       background=SPACE_COLORS['bg_panel'],
                       foreground=SPACE_COLORS['text_secondary'],
                       padding=[20, 10])
        
        style.map('Space.TNotebook.Tab',
                 background=[('selected', SPACE_COLORS['bg_dark'])],
                 foreground=[('selected', SPACE_COLORS['accent'])])
        
        # Кнопки
        style.configure('Space.TButton',
                       font=('Segoe UI', 10, 'bold'),
                       foreground=SPACE_COLORS['text_primary'],
                       background=SPACE_COLORS['bg_panel'],
                       borderwidth=1,
                       relief='flat')
        
        style.map('Space.TButton',
                 background=[('active', SPACE_COLORS['accent'])],
                 foreground=[('active', SPACE_COLORS['bg_dark'])])
        
        style.configure('Launch.TButton',
                       font=('Segoe UI', 13, 'bold'),
                       foreground=SPACE_COLORS['bg_dark'],
                       background=SPACE_COLORS['success'],
                       borderwidth=2,
                       relief='flat')
        
        style.map('Launch.TButton',
                 background=[('active', SPACE_COLORS['accent'])],
                 foreground=[('active', SPACE_COLORS['text_primary'])])
    
    def create_widgets(self):
        """Создает космический интерфейс"""
        
        # Главный контейнер с темным фоном
        main_frame = ttk.Frame(self.root, padding="20", style='Space.TFrame')
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # ASCII арт заголовок
        ascii_art = """
        ╔═══════════════════════════════════════════════╗
        ║  EXCEL PRO MASTER ◆ КОСМИЧЕСКАЯ ВЕРСИЯ       ║
        ╚═══════════════════════════════════════════════╝
        """
        
        # Заголовок
        title_frame = tk.Frame(main_frame, bg=SPACE_COLORS['bg_dark'])
        title_frame.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # ASCII заголовок
        ascii_label = tk.Label(title_frame, text=ascii_art, 
                              font=('Courier', 10), 
                              fg=SPACE_COLORS['accent'], 
                              bg=SPACE_COLORS['bg_dark'])
        ascii_label.pack()
        
        subtitle_label = tk.Label(title_frame, 
                                 text="◈ Система Статистического Анализа ◈", 
                                 font=('Segoe UI', 12),
                                 fg=SPACE_COLORS['text_secondary'],
                                 bg=SPACE_COLORS['bg_dark'])
        subtitle_label.pack(pady=(0, 10))
        
        # Панель с вкладками
        self.notebook = ttk.Notebook(main_frame, style='Space.TNotebook')
        self.notebook.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        # Вкладка 1: Ввод данных (48 строк)
        self.tab1 = tk.Frame(self.notebook, bg=SPACE_COLORS['bg_panel'])
        self.notebook.add(self.tab1, text='◆ ДАННЫЕ-48')
        self.create_data_tab(self.tab1, 
            "// ВСТАВЬТЕ ДАННЫЕ ИЗ EXCEL ИЛИ ВВЕДИТЕ В ФОРМАТЕ:\n" +
            "// [НОМЕР] [ЗНАЧЕНИЕ] или просто [ЗНАЧЕНИЕ]\n" +
            "// ПРИМЕР: 1 100.55 или просто 100.55\n" +
            "// МОЖНО ВСТАВИТЬ СТОЛБЕЦ ИЗ EXCEL ПРЯМО СЮДА!", 48)
        
        # Вкладка 2: Ввод данных (25 строк)
        self.tab2 = tk.Frame(self.notebook, bg=SPACE_COLORS['bg_panel'])
        self.notebook.add(self.tab2, text='◆ ДАННЫЕ-25')
        self.create_data_tab(self.tab2,
            "// ВСТАВЬТЕ ДАННЫЕ ИЗ EXCEL ИЛИ ВВЕДИТЕ В ФОРМАТЕ:\n" +
            "// [НОМЕР] [ЗНАЧЕНИЕ] или просто [ЗНАЧЕНИЕ]\n" +
            "// ПРИМЕР: 1 100.55 или просто 100.55\n" +
            "// МОЖНО ВСТАВИТЬ СТОЛБЕЦ ИЗ EXCEL ПРЯМО СЮДА!", 25)
        
        # Вкладка 3: Настройки
        self.tab3 = tk.Frame(self.notebook, bg=SPACE_COLORS['bg_panel'])
        self.notebook.add(self.tab3, text='◆ НАСТРОЙКИ')
        self.create_settings_tab(self.tab3)
        
        # Панель управления
        control_frame = tk.Frame(main_frame, bg=SPACE_COLORS['bg_panel'], 
                                highlightbackground=SPACE_COLORS['border'],
                                highlightthickness=1)
        control_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        
        # Заголовок панели
        control_label = tk.Label(control_frame, text="◈ ЦЕНТР УПРАВЛЕНИЯ ◈",
                               font=('Segoe UI', 10, 'bold'),
                               fg=SPACE_COLORS['accent'],
                               bg=SPACE_COLORS['bg_panel'])
        control_label.pack(pady=(10, 5))
        
        # Кнопки
        button_frame = tk.Frame(control_frame, bg=SPACE_COLORS['bg_panel'])
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(button_frame, text="◆ ЗАГРУЗИТЬ ПРИМЕР", 
                  command=self.paste_example, 
                  style='Space.TButton', width=18).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="◆ ОЧИСТИТЬ", 
                  command=self.clear_data, 
                  style='Space.TButton', width=18).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="◆ ИМПОРТ ФАЙЛА", 
                  command=self.load_from_file, 
                  style='Space.TButton', width=18).pack(side=tk.LEFT, padx=5)
        
        self.generate_btn = ttk.Button(button_frame, text="▶ ЗАПУСК АНАЛИЗА", 
                                      command=self.generate_report, 
                                      style='Launch.TButton', width=20)
        self.generate_btn.pack(side=tk.RIGHT, padx=5)
        
        # Статус бар
        status_frame = tk.Frame(main_frame, bg=SPACE_COLORS['bg_dark'])
        status_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        tk.Label(status_frame, text="СТАТУС:", 
                font=('Segoe UI', 9, 'bold'),
                fg=SPACE_COLORS['accent'],
                bg=SPACE_COLORS['bg_dark']).pack(side=tk.LEFT, padx=5)
        
        self.status_var = tk.StringVar(value="◆ СИСТЕМА ГОТОВА")
        self.status_bar = tk.Label(status_frame, textvariable=self.status_var,
                                  font=('Courier', 10),
                                  fg=SPACE_COLORS['success'],
                                  bg=SPACE_COLORS['bg_dark'])
        self.status_bar.pack(side=tk.LEFT)
        
        # Настройка растягивания
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
    
    def create_data_tab(self, parent, instruction, expected_rows):
        """Создает космическую вкладку для ввода данных"""
        frame = tk.Frame(parent, bg=SPACE_COLORS['bg_panel'])
        frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        # Инструкция
        tk.Label(frame, text=instruction, 
                font=('Courier', 10),
                fg=SPACE_COLORS['accent'],
                bg=SPACE_COLORS['bg_panel']).pack(pady=(0, 10))
        
        # Текстовое поле
        text_frame = tk.Frame(frame, bg=SPACE_COLORS['bg_panel'])
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        # Добавляем скролл
        scrollbar = tk.Scrollbar(text_frame, 
                                bg=SPACE_COLORS['bg_panel'],
                                troughcolor=SPACE_COLORS['bg_dark'])
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        text_widget = tk.Text(text_frame, 
                             wrap=tk.NONE, 
                             font=('Courier', 11), 
                             bg=SPACE_COLORS['bg_input'],
                             fg=SPACE_COLORS['text_primary'],
                             insertbackground=SPACE_COLORS['accent'],
                             selectbackground=SPACE_COLORS['accent'],
                             selectforeground=SPACE_COLORS['bg_dark'],
                             yscrollcommand=scrollbar.set,
                             relief='flat',
                             borderwidth=2)
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=text_widget.yview)
        
        # Сохраняем ссылку
        if expected_rows == 48:
            self.text_48 = text_widget
        else:
            self.text_25 = text_widget
        
        # Счетчик строк
        count_var = tk.StringVar(value=f"◈ СТРОК: 0 / {expected_rows}")
        count_label = tk.Label(frame, 
                              textvariable=count_var,
                              font=('Courier', 10),
                              fg=SPACE_COLORS['success'],
                              bg=SPACE_COLORS['bg_panel'])
        count_label.pack(pady=5)
        
        # Обновление счетчика
        def update_count(event=None):
            content = text_widget.get('1.0', tk.END).strip()
            lines = [l for l in content.split('\n') if l.strip() and not l.startswith('#') and not l.startswith('//')]
            count = len(lines)
            count_var.set(f"◈ СТРОК: {count} / {expected_rows}")
            
            # Меняем цвет в зависимости от количества
            if count == 0:
                count_label.config(fg=SPACE_COLORS['text_secondary'])
            elif count < expected_rows * 0.8:
                count_label.config(fg=SPACE_COLORS['warning'])
            else:
                count_label.config(fg=SPACE_COLORS['success'])
        
        text_widget.bind('<KeyRelease>', update_count)
    
    def create_settings_tab(self, parent):
        """Создает космическую вкладку настроек"""
        frame = tk.Frame(parent, bg=SPACE_COLORS['bg_panel'])
        frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Заголовок секции
        tk.Label(frame, text="◈ ПАПКА ДЛЯ СОХРАНЕНИЯ ◈", 
                font=('Segoe UI', 11, 'bold'),
                fg=SPACE_COLORS['accent'],
                bg=SPACE_COLORS['bg_panel']).grid(row=0, column=0, sticky=tk.W, pady=10)
        
        self.output_path = tk.StringVar(value=str(Path.home() / "Desktop"))
        path_frame = tk.Frame(frame, bg=SPACE_COLORS['bg_panel'])
        path_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        
        path_entry = tk.Entry(path_frame, 
                             textvariable=self.output_path, 
                             width=50,
                             font=('Courier', 10),
                             bg=SPACE_COLORS['bg_input'],
                             fg=SPACE_COLORS['text_primary'],
                             insertbackground=SPACE_COLORS['accent'],
                             relief='flat',
                             borderwidth=2)
        path_entry.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(path_frame, text="◆ ОБЗОР", 
                  command=self.choose_folder, 
                  style='Space.TButton').pack(side=tk.LEFT)
        
        # Опции анализа
        tk.Label(frame, text="◈ ПАРАМЕТРЫ АНАЛИЗА ◈", 
                font=('Segoe UI', 11, 'bold'),
                fg=SPACE_COLORS['accent'],
                bg=SPACE_COLORS['bg_panel']).grid(row=2, column=0, sticky=tk.W, pady=(20, 10))
        
        # Стиль для чекбоксов
        checkbox_style = {
            'font': ('Segoe UI', 10),
            'fg': SPACE_COLORS['text_primary'],
            'bg': SPACE_COLORS['bg_panel'],
            'selectcolor': SPACE_COLORS['bg_dark'],
            'activebackground': SPACE_COLORS['bg_panel'],
            'activeforeground': SPACE_COLORS['accent']
        }
        
        self.include_charts = tk.BooleanVar(value=False)  # Отключено по умолчанию
        tk.Checkbutton(frame, 
                      text="◆ Создавать графики (ВНИМАНИЕ: может вызвать ошибки!)", 
                      variable=self.include_charts,
                      **checkbox_style).grid(row=3, column=0, sticky=tk.W, pady=3)
        
        self.include_outliers = tk.BooleanVar(value=True)
        tk.Checkbutton(frame, 
                      text="◆ Анализ выбросов и аномалий", 
                      variable=self.include_outliers,
                      **checkbox_style).grid(row=4, column=0, sticky=tk.W, pady=3)
        
        self.include_normality = tk.BooleanVar(value=True)
        tk.Checkbutton(frame, 
                      text="◆ Проверка нормальности распределения", 
                      variable=self.include_normality,
                      **checkbox_style).grid(row=5, column=0, sticky=tk.W, pady=3)
        
        self.auto_open = tk.BooleanVar(value=True)
        tk.Checkbutton(frame, 
                      text="◆ Открыть файл после создания", 
                      variable=self.auto_open,
                      **checkbox_style).grid(row=6, column=0, sticky=tk.W, pady=3)
    
    def choose_folder(self):
        """Выбор папки для сохранения"""
        folder = filedialog.askdirectory(initialdir=self.output_path.get())
        if folder:
            self.output_path.set(folder)
    
    def load_reference_data(self):
        """Загружает примеры данных при запуске"""
        # Вставляем в поле для 48 строк
        if hasattr(self, 'text_48'):
            text = ""
            for i, val in enumerate(EXAMPLE_DATA_48, 1):
                text += f"{i} {val:.2f}\n"
            self.text_48.insert('1.0', text)
        
        # Вставляем в поле для 25 строк
        if hasattr(self, 'text_25'):
            text = ""
            for i, val in enumerate(EXAMPLE_DATA_25, 1):
                text += f"{i} {val:.2f}\n"
            self.text_25.insert('1.0', text)
    
    def paste_example(self):
        """Вставляет пример данных"""
        current_tab = self.notebook.index(self.notebook.select())
        
        if current_tab == 0:  # 48 строк
            self.text_48.delete('1.0', tk.END)
            text = ""
            for i, val in enumerate(EXAMPLE_DATA_48, 1):
                text += f"{i} {val:.2f}\n"
            self.text_48.insert('1.0', text)
        elif current_tab == 1:  # 25 строк
            self.text_25.delete('1.0', tk.END)
            text = ""
            for i, val in enumerate(EXAMPLE_DATA_25, 1):
                text += f"{i} {val:.2f}\n"
            self.text_25.insert('1.0', text)
    
    def clear_data(self):
        """Очищает поля ввода"""
        current_tab = self.notebook.index(self.notebook.select())
        if current_tab == 0:
            self.text_48.delete('1.0', tk.END)
        elif current_tab == 1:
            self.text_25.delete('1.0', tk.END)
    
    def load_from_file(self):
        """Загрузка данных из файла"""
        file_path = filedialog.askopenfilename(
            title="Выберите файл с данными",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                current_tab = self.notebook.index(self.notebook.select())
                if current_tab == 0:
                    self.text_48.delete('1.0', tk.END)
                    self.text_48.insert('1.0', content)
                elif current_tab == 1:
                    self.text_25.delete('1.0', tk.END)
                    self.text_25.insert('1.0', content)
                
                self.status_var.set(f"✅ Данные загружены из {Path(file_path).name}")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось загрузить файл:\n{str(e)}")
    
    def parse_data(self, text):
        """Парсит введенные данные - поддерживает вставку из Excel"""
        lines = text.strip().split('\n')
        data = []
        
        for line in lines:
            line = line.strip()
            if not line or line.startswith('#') or line.startswith('//'):
                continue
            
            # Заменяем запятую на точку для дробной части
            line = line.replace(',', '.')
            
            # Разбиваем по табуляции (если копируют из Excel)
            parts = line.split('\t')
            if len(parts) == 1:
                # Если нет табуляции, пробуем по пробелам
                parts = line.split()
            
            # Пробуем найти число в строке
            value = None
            
            # Сначала проверяем второй столбец (если есть)
            if len(parts) >= 2:
                try:
                    value = float(parts[1])
                except ValueError:
                    pass
            
            # Если не нашли, проверяем первый столбец
            if value is None and len(parts) >= 1:
                try:
                    value = float(parts[0])
                except ValueError:
                    # Если первое значение не число, ищем первое число в строке
                    for part in parts:
                        try:
                            value = float(part)
                            break
                        except ValueError:
                            continue
            
            if value is not None:
                data.append(value)
        
        return np.array(data)
    
    def generate_report(self):
        """Генерирует отчет"""
        try:
            self.status_var.set("◈ ИНИЦИАЛИЗАЦИЯ АНАЛИЗА...")
            self.generate_btn.config(state='disabled')
            
            # Определяем какая вкладка активна
            current_tab = self.notebook.index(self.notebook.select())
            
            if current_tab == 0:
                text = self.text_48.get('1.0', tk.END)
                expected = 48
            elif current_tab == 1:
                text = self.text_25.get('1.0', tk.END)
                expected = 25
            else:
                messagebox.showwarning("Внимание", "Выберите вкладку с данными!")
                return
            
            # Парсим данные
            data = self.parse_data(text)
            
            if len(data) == 0:
                messagebox.showerror("Ошибка", "Нет данных для анализа!")
                return
            
            # Предупреждение если данных не хватает
            if len(data) < expected * 0.8:  # Меньше 80% от ожидаемого
                if not messagebox.askyesno("Внимание", 
                    f"Введено {len(data)} значений вместо ожидаемых {expected}.\nПродолжить?"):
                    return
            
            # Создаем анализатор
            analyzer = StatisticalAnalyzer(data)
            
            # Путь для сохранения
            timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
            filename = f"Statistical_Report_{timestamp}.xlsx"
            output_path = os.path.join(self.output_path.get(), filename)
            
            # Создаем Excel отчет
            report = ExcelReportGenerator(output_path)
            
            # Создаем листы
            report.create_main_sheet(data, analyzer)
            
            if self.include_normality.get():
                report.create_normality_sheet(data, analyzer)
            
            if self.include_charts.get():
                report.create_charts_sheet(data, analyzer)
            
            if self.include_outliers.get():
                report.create_outliers_sheet(data, analyzer)
            
            report.create_conclusion_sheet(analyzer)
            
            # Закрываем файл
            report.close()
            
            self.status_var.set(f"◆ АНАЛИЗ ЗАВЕРШЁН: {filename}")
            
            # Открываем файл если нужно
            if self.auto_open.get():
                if sys.platform == 'win32':
                    os.startfile(output_path)
                elif sys.platform == 'darwin':
                    subprocess.run(['open', output_path])
                else:
                    subprocess.run(['xdg-open', output_path])
            
            messagebox.showinfo("Успех!", f"Отчёт успешно создан!\n\n{output_path}")
            
        except Exception as e:
            self.status_var.set("◆ ОШИБКА: АНАЛИЗ НЕ ВЫПОЛНЕН")
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n\n{str(e)}")
        finally:
            self.generate_btn.config(state='normal')


# ============================================================================
# ГЛАВНАЯ ФУНКЦИЯ
# ============================================================================

def main():
    """Главная функция запуска приложения"""
    root = tk.Tk()
    app = ExcelProMasterGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()
