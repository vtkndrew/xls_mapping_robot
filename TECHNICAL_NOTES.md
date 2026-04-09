# Технические заметки к реализации

## Ключевые решения

### 1. Двойное открытие файла
```python
# Чтение значений (включая вычисленные формулы)
wb_read = openpyxl.load_workbook(input_file, data_only=True)

# Запись с сохранением формул
wb_write = openpyxl.load_workbook(input_file, data_only=False)
```

**Причина:** openpyxl с `data_only=True` позволяет читать вычисленные значения формул, но при сохранении теряет сами формулы. Поэтому используется два экземпляра.

### 2. Валидация данных

#### Лист "Шаблон пров и атриб"
```python
def normalize_template_value(value, row_num, column_name, sheet_name, logger):
    # 1. Проверка на None
    # 2. Удаление точек: "70.606" → "70606"
    # 3. Проверка regex: только цифры
    # 4. Логирование невалидных значений
```

#### Лист "Проводки по СП анализ"
```python
def parse_provodki_account(value, row_num, column_name, sheet_name, logger):
    # 1. Проверка количества подчеркиваний (0 или 1)
    # 2. Разделение: "70606_3141302" → ("70606", "3141302")
    # 3. Удаление точек из обеих частей
    # 4. Валидация каждой части
```

### 3. Создание ключей сравнения
```python
comparison_key = f"{dt_account}{dt_ofr}{kt_account}{kt_ofr}"
# Пример: "706063141302474230"
```

Пустые символы ОФР просто становятся пустыми строками в ключе, что не влияет на сравнение.

### 4. Структура данных для мэппинга

#### Словарь для шаблона
```python
template_keys = {
    "706063141302474230": [
        {
            'row': 10,
            'pokaza': "СП-123",
            'dt_col': "A",
            'kt_col': "B"
        },
        # ... другие строки с таким же ключом
    ]
}
```

#### Словарь для проводок
```python
provodki_keys = {
    "706063141302474230": [
        {
            'row': 25,
            'dt_col': "C",
            'kt_col': "D"
        },
        # ... другие строки с таким же ключом
    ]
}
```

### 5. Формат результатов для записи
```python
results = [
    {
        'sheet': 'Шаблон пров и атриб',
        'row': 10,
        'dt_mapping_col': 15,  # Индекс столбца
        'kt_mapping_col': 16,
        'dt_value': "='Проводки по СП анализ'!C25, ='Проводки по СП анализ'!C30",
        'kt_value': "='Проводки по СП анализ'!D25, ='Проводки по СП анализ'!D30"
    }
]
```

## Оптимизации

### 1. Использование defaultdict
```python
from collections import defaultdict
template_keys = defaultdict(list)
```
Автоматически создаёт пустой список при первом обращении к ключу.

### 2. Индексация столбцов
Столбцы индексируются один раз в начале обработки каждого листа, затем используются числовые индексы для быстрого доступа.

### 3. Батчинг результатов
Все результаты собираются в памяти, затем записываются одним проходом, минимизируя операции I/O.

## Обработка edge cases

### 1. Пустые значения символов ОФР
```python
if dt_ofr_norm is None:
    dt_ofr_norm = ""  # Допустимо для символов ОФР
```

### 2. Множественные совпадения
```python
# Все ссылки объединяются через запятую
dt_refs = [ref1, ref2, ref3]
result = ", ".join(dt_refs)
# "='Проводки'!A1, ='Проводки'!A2, ='Проводки'!A3"
```

### 3. Развёрнутые проводки
```python
# Парсинг номеров строк
expanded_str = "5, 10, 23"
row_numbers = [int(part.strip()) for part in expanded_str.split(',') if part.strip().isdigit()]
# [5, 10, 23]
```

### 4. Уникальные значения "№ Показа"
```python
pokaza_values = []
for template_info in template_rows:
    pokaza = str(template_info['pokaza']).strip()
    if pokaza and pokaza not in pokaza_values:
        pokaza_values.append(pokaza)
```

## Логирование

### Структура лога
```python
timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
message = f"[{timestamp}] ОШИБКА | Лист: {sheet} | Строка: {row} | " \
          f"Столбец: {column} | Значение: {value} | Причина: {reason}"
```

### Типы сообщений
- **ОШИБКА** - критические проблемы валидации
- **ПРЕДУПРЕЖДЕНИЕ** - некритические проблемы

## Статистика

### Собираемые метрики
```python
class ProcessStatistics:
    template_total         # Всего строк в шаблоне
    template_processed     # Обработано после фильтра ПАО
    template_errors        # Ошибок валидации
    template_matched       # Успешно сматчено
    
    provodki_total         # Всего строк в проводках
    provodki_processed     # Обработано без развёрнутых
    provodki_expanded      # Обработано развёрнутых
    provodki_errors        # Ошибок валидации
    provodki_matched       # Успешно сматчено
```

## Производительность

### Временная сложность
- Чтение данных: O(n)
- Создание ключей: O(n)
- Поиск совпадений: O(1) благодаря словарям
- Запись результатов: O(m), где m - количество совпадений

### Пространственная сложность
- Словари ключей: O(n)
- Результаты для записи: O(m)
- Общая: O(n + m)

### Узкие места
1. Чтение Excel файла - зависит от openpyxl
2. Запись Excel файла - зависит от openpyxl
3. Валидация данных - минимальна (regex проверки)

## Возможные улучшения

### 1. Использование pandas для фильтрации
```python
import pandas as pd
df = pd.read_excel(input_file, sheet_name="Шаблон пров и атриб", header=6)
df_filtered = df[df["Признак плана счетов (ПАО/КИБ)"] == "ПАО"]
```
**Преимущество:** Быстрее для больших файлов  
**Недостаток:** Дополнительная зависимость

### 2. Параллельная обработка
```python
from multiprocessing import Pool
# Обработка разных этапов в параллельных процессах
```
**Преимущество:** Ускорение на многоядерных системах  
**Недостаток:** Усложнение кода, проблемы с памятью

### 3. Прогресс-бар
```python
from tqdm import tqdm
for row_num in tqdm(range(data_start_row, max_row + 1)):
    # Обработка
```
**Преимущество:** Визуализация прогресса  
**Недостаток:** Дополнительная зависимость

### 4. Конфигурационный файл
```yaml
# config.yaml
sheets:
  template:
    name: "Шаблон пров и атриб"
    header_row: 7
    filter_column: "Признак плана счетов (ПАО/КИБ)"
    filter_value: "ПАО"
```
**Преимущество:** Гибкость настройки  
**Недостаток:** Усложнение для простых случаев

## Тестирование

### Юнит-тесты
```python
import unittest

class TestNormalization(unittest.TestCase):
    def test_remove_dots(self):
        result = normalize_template_value("70.606", 1, "Test", "Sheet", logger)
        self.assertEqual(result, "70606")
    
    def test_invalid_chars(self):
        result = normalize_template_value("70606.ABC", 1, "Test", "Sheet", logger)
        self.assertIsNone(result)
```

### Интеграционные тесты
```python
def test_full_workflow():
    # 1. Создать тестовый Excel файл
    # 2. Запустить скрипт
    # 3. Проверить результаты
    # 4. Проверить статистику
```

### Тестовые данные
Создать минимальный Excel файл с:
- 10 строк в шаблоне
- 20 строк в проводках
- Известными совпадениями
- Заведомо невалидными данными

## Отладка

### Включение детального логирования
```python
import logging
logging.basicConfig(level=logging.DEBUG)
```

### Проверка промежуточных результатов
```python
print(f"DEBUG: template_keys = {template_keys}")
print(f"DEBUG: provodki_keys = {provodki_keys}")
```

### Профилирование
```python
import cProfile
cProfile.run('main()', 'profile_stats')

import pstats
p = pstats.Stats('profile_stats')
p.sort_stats('cumulative').print_stats(10)
```

## Обработка ошибок

### Типы исключений
```python
try:
    # Операции с файлом
except FileNotFoundError:
    # Файл не найден
except PermissionError:
    # Нет прав доступа
except KeyError:
    # Столбец не найден
except Exception as e:
    # Общая ошибка
```

### Graceful degradation
Скрипт продолжает работу даже при ошибках валидации отдельных строк, логируя их для последующего анализа.

## Безопасность

### Валидация входных данных
- Проверка существования файла
- Проверка структуры листов
- Проверка типов данных

### Защита от переполнения
- Ограничение на размер файлов (неявно через openpyxl)
- Валидация номеров строк в развёрнутых проводках

## Совместимость

### Python версии
- Минимум: Python 3.7 (f-strings, type hints)
- Рекомендуется: Python 3.9+

### Excel версии
- Excel 2007+ (.xlsx формат)
- LibreOffice Calc (поддержка .xlsx)

### Операционные системы
- Windows ✅
- macOS ✅
- Linux ✅
