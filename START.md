# 🚀 БЫСТРЫЙ СТАРТ - Excel Analytics PRO

## 1️⃣ Установка (один раз)

```bash
cd "/Users/mihailzarov/ексель"
python3 -m pip install -r requirements.txt
```

## 2️⃣ Запуск

### Команда:
```bash
python3 report.py твои_данные1.txt твои_данные2.txt
```

### Готовые примеры:

```bash
# Вариант 1: Полная выборка (48 точек)
python3 report.py data/full_data.txt

# Вариант 2: Два округа 
python3 report.py data/region1_dfo.txt data/region2_pfo.txt

# Вариант 3: Всё вместе (рекомендуется для полного отчёта)
python3 report.py data/full_data.txt data/region1_dfo.txt data/region2_pfo.txt
```

## 3️⃣ Результат

Открой файл: **`out/report_pro.xlsx`**

Там будет:
- ✅ Все расчёты через ФОРМУЛЫ Excel
- ✅ Профессиональное оформление
- ✅ Графики уже вставлены
- ✅ Готово к печати/сдаче

## 4️⃣ Свои данные

Создай текстовый файл:
```
1    12.45
2    15.67
3    14.23
...
```

Запусти:
```bash
python3 report.py мои_данные.txt
```

---

**💡 Фишка:** Можешь изменить числа прямо в Excel на листе `Data_*` — все формулы пересчитаются автоматически!

**🎯 Готово!** Клепай работы на автомате 🔥
