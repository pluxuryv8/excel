#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WEB INTERFACE - Красивый веб-интерфейс для Excel Analytics
Запусти и открой в браузере!
"""

import os
import sys
import tempfile
import subprocess
import webbrowser
from pathlib import Path
from flask import Flask, render_template_string, request, send_file, jsonify
import json

app = Flask(__name__)

HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <title>Excel Analytics PRO</title>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
        * { box-sizing: border-box; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            margin: 0;
            padding: 20px;
            background: #f5f5f5;
        }
        .container {
            max-width: 1000px;
            margin: 0 auto;
            background: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        h1 {
            color: #2c3e50;
            text-align: center;
            margin-bottom: 10px;
        }
        .subtitle {
            text-align: center;
            color: #7f8c8d;
            margin-bottom: 30px;
        }
        .data-input {
            margin-bottom: 20px;
        }
        .data-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 10px;
        }
        .data-title {
            font-weight: bold;
            color: #34495e;
        }
        textarea {
            width: 100%;
            min-height: 300px;
            padding: 15px;
            border: 2px solid #ddd;
            border-radius: 5px;
            font-family: 'Courier New', monospace;
            font-size: 14px;
            resize: vertical;
        }
        textarea:focus {
            outline: none;
            border-color: #3498db;
        }
        .buttons {
            display: flex;
            gap: 10px;
            margin-top: 20px;
            flex-wrap: wrap;
        }
        button {
            padding: 12px 24px;
            border: none;
            border-radius: 5px;
            font-size: 16px;
            cursor: pointer;
            transition: all 0.3s;
        }
        .btn-primary {
            background: #3498db;
            color: white;
            flex: 1;
            min-width: 200px;
        }
        .btn-primary:hover {
            background: #2980b9;
        }
        .btn-secondary {
            background: #95a5a6;
            color: white;
        }
        .btn-secondary:hover {
            background: #7f8c8d;
        }
        .btn-success {
            background: #27ae60;
            color: white;
        }
        .btn-danger {
            background: #e74c3c;
            color: white;
        }
        .status {
            margin-top: 20px;
            padding: 15px;
            border-radius: 5px;
            text-align: center;
            display: none;
        }
        .status.success {
            background: #d4edda;
            color: #155724;
            border: 1px solid #c3e6cb;
        }
        .status.error {
            background: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
        }
        .status.loading {
            background: #cce5ff;
            color: #004085;
            border: 1px solid #b8daff;
        }
        .example {
            background: #f8f9fa;
            padding: 15px;
            border-radius: 5px;
            margin-top: 20px;
            font-family: monospace;
            white-space: pre-line;
            color: #495057;
            border: 1px solid #dee2e6;
        }
        .loader {
            border: 4px solid #f3f3f3;
            border-top: 4px solid #3498db;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin: 20px auto;
        }
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        #datasets {
            display: flex;
            flex-direction: column;
            gap: 20px;
        }
        .dataset {
            border: 2px solid #e0e0e0;
            border-radius: 8px;
            padding: 20px;
            position: relative;
        }
        .remove-btn {
            position: absolute;
            top: 10px;
            right: 10px;
            background: #e74c3c;
            color: white;
            border: none;
            border-radius: 50%;
            width: 30px;
            height: 30px;
            cursor: pointer;
            font-size: 18px;
        }
        .remove-btn:hover {
            background: #c0392b;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 Excel Analytics PRO</h1>
        <p class="subtitle">Вставь данные и получи профессиональный Excel-отчёт с формулами!</p>
        
        <div id="datasets">
            <div class="dataset" id="dataset-1">
                <div class="data-header">
                    <span class="data-title">📈 Выборка 1</span>
                    <button class="remove-btn" onclick="removeDataset(1)" style="display:none;">✕</button>
                </div>
                <textarea id="data-1" placeholder="Вставь данные сюда...&#10;&#10;Формат:&#10;1    12.45&#10;2    15.67&#10;3    14.23&#10;...&#10;&#10;Или нажми 'Демо данные' для примера"></textarea>
            </div>
        </div>
        
        <div class="buttons">
            <button class="btn-secondary" onclick="addDataset()">➕ Добавить выборку</button>
            <button class="btn-secondary" onclick="loadDemo()">📋 Демо данные</button>
            <button class="btn-danger" onclick="clearAll()">🗑️ Очистить</button>
        </div>
        
        <div class="buttons">
            <button class="btn-primary" onclick="generateReport()">
                🚀 СОЗДАТЬ ОТЧЁТ
            </button>
        </div>
        
        <div id="status" class="status"></div>
        
        <div class="example">
<strong>Как использовать:</strong>
1. Вставь данные в формате: номер &lt;пробел&gt; значение
2. Можно добавить несколько выборок
3. Нажми "СОЗДАТЬ ОТЧЁТ"
4. Скачай готовый Excel-файл

<strong>Пример данных:</strong>
1    100.71
2    100.56
3    98.97
4    100.63
5    100.58</div>
    </div>
    
    <script>
        let datasetCount = 1;
        
        function addDataset() {
            datasetCount++;
            const html = `
                <div class="dataset" id="dataset-${datasetCount}">
                    <div class="data-header">
                        <span class="data-title">📈 Выборка ${datasetCount}</span>
                        <button class="remove-btn" onclick="removeDataset(${datasetCount})">✕</button>
                    </div>
                    <textarea id="data-${datasetCount}" placeholder="Вставь данные сюда..."></textarea>
                </div>
            `;
            document.getElementById('datasets').insertAdjacentHTML('beforeend', html);
            updateRemoveButtons();
        }
        
        function removeDataset(id) {
            document.getElementById(`dataset-${id}`).remove();
            updateRemoveButtons();
        }
        
        function updateRemoveButtons() {
            const datasets = document.querySelectorAll('.dataset');
            datasets.forEach(dataset => {
                const btn = dataset.querySelector('.remove-btn');
                btn.style.display = datasets.length > 1 ? 'block' : 'none';
            });
        }
        
        function loadDemo() {
            const demoData = `1    100.71
2    100.56
3    98.97
4    100.63
5    100.58
6    100.87
7    100.78
8    102.51
9    99.97
10   101.11
11   100.02`;
            
            document.getElementById('data-1').value = demoData;
            
            if (datasetCount > 1) return;
            
            addDataset();
            const demoData2 = `1    100.55
2    100.46
3    100.29
4    100.84
5    100.98
6    100.35
7    100.89
8    100.67
9    101.10
10   99.94
11   100.21
12   100.58
13   100.47
14   101.70`;
            document.getElementById('data-2').value = demoData2;
        }
        
        function clearAll() {
            if (confirm('Очистить все данные?')) {
                document.querySelectorAll('textarea').forEach(ta => ta.value = '');
            }
        }
        
        function showStatus(message, type) {
            const status = document.getElementById('status');
            status.className = 'status ' + type;
            status.innerHTML = message;
            status.style.display = 'block';
        }
        
        async function generateReport() {
            const datasets = [];
            document.querySelectorAll('textarea').forEach((ta, index) => {
                const data = ta.value.trim();
                if (data) {
                    datasets.push(data);
                }
            });
            
            if (datasets.length === 0) {
                showStatus('❌ Нет данных для обработки!', 'error');
                return;
            }
            
            showStatus('<div class="loader"></div>Создаю отчёт... Это займёт несколько секунд', 'loading');
            
            try {
                const response = await fetch('/generate', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({datasets: datasets})
                });
                
                if (response.ok) {
                    const blob = await response.blob();
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = 'Excel_Report.xlsx';
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    showStatus('✅ Отчёт создан и загружен!', 'success');
                } else {
                    const error = await response.text();
                    showStatus('❌ Ошибка: ' + error, 'error');
                }
            } catch (e) {
                showStatus('❌ Ошибка соединения', 'error');
            }
        }
    </script>
</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/generate', methods=['POST'])
def generate():
    try:
        data = request.json
        datasets = data.get('datasets', [])
        
        if not datasets:
            return "Нет данных", 400
        
        # Создаём временные файлы
        temp_files = []
        for dataset in datasets:
            temp_file = tempfile.NamedTemporaryFile(mode='w', suffix='.txt', 
                                                   delete=False, encoding='utf-8')
            temp_file.write(dataset)
            temp_file.close()
            temp_files.append(temp_file.name)
        
        # Запускаем основной скрипт
        script_path = os.path.join(os.path.dirname(__file__), 'report.py')
        output_dir = tempfile.mkdtemp()
        
        os.chdir(output_dir)
        cmd = [sys.executable, script_path] + temp_files
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        # Удаляем временные файлы
        for f in temp_files:
            try:
                os.unlink(f)
            except:
                pass
        
        if result.returncode == 0:
            output_file = os.path.join(output_dir, 'out', 'report_pro.xlsx')
            if os.path.exists(output_file):
                return send_file(output_file, as_attachment=True, 
                               download_name='Excel_Report.xlsx',
                               mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        
        return f"Ошибка генерации: {result.stderr}", 500
        
    except Exception as e:
        return str(e), 500

def main():
    print("\n" + "="*60)
    print("🌐 EXCEL ANALYTICS PRO - ВЕБ-ИНТЕРФЕЙС")
    print("="*60 + "\n")
    print("Запускаю веб-сервер...")
    print("\n🚀 Открой в браузере: http://localhost:5000")
    print("\nДля остановки нажми Ctrl+C\n")
    
    # Автоматически открываем браузер
    webbrowser.open('http://localhost:5000')
    
    # Запускаем сервер
    app.run(debug=False, port=5000)

if __name__ == '__main__':
    # Проверяем Flask
    try:
        import flask
    except ImportError:
        print("❌ Нужно установить Flask!")
        print("Выполни: pip install flask")
        sys.exit(1)
    
    main()
