"""
Скрипт для проверки проблемы с мэппингом
Детально проверяет почему не записываются ссылки для конкретных строк
"""

import sys
import openpyxl
from openpyxl.utils import get_column_letter
from collections import defaultdict


# Проблемные строки
DEBUG_TEMPLATE_ROWS = [2335, 2857]
DEBUG_PROVODKI_ROWS = [2194, 2203]


def get_column_index_by_name(ws, column_name, header_row):
    """Получить индекс столбца по его имени"""
    for idx, cell in enumerate(ws[header_row], start=1):
        if cell.value and str(cell.value).strip() == column_name:
            return idx
    return None


def check_template_columns(wb):
    """Проверить индексы столбцов в шаблоне"""
    print("\n" + "=" * 80)
    print("ПРОВЕРКА СТОЛБЦОВ ШАБЛОНА")
    print("=" * 80)
    
    ws = wb["Шаблон пров и атриб"]
    header_row = 7
    
    # Проверяем все нужные столбцы
    columns_to_check = [
        "Признак плана счетов (ПАО/КИБ)",
        "Номер счета по Дт",
        "Символ ОФР Дт",
        "Номер счета по Кт",
        "Символ ОФР Кт",
        "№ Показа",
        "Дт Мэпинг Робот",
        "Кт Мэпинг Робот"
    ]
    
    results = {}
    
    for col_name in columns_to_check:
        idx = get_column_index_by_name(ws, col_name, header_row)
        results[col_name] = idx
        
        if idx is None:
            print(f"❌ ОШИБКА: Столбец '{col_name}' НЕ НАЙДЕН!")
        else:
            letter = get_column_letter(idx)
            print(f"✅ '{col_name}': столбец {letter} (индекс {idx})")
    
    # Проверяем конкретные строки
    print(f"\n--- Проверка данных в проблемных строках ---")
    
    col_dt_account = results["Номер счета по Дт"]
    col_kt_account = results["Номер счета по Кт"]
    
    if col_dt_account and col_kt_account:
        for row_num in DEBUG_TEMPLATE_ROWS:
            filter_val = ws.cell(row=row_num, column=results["Признак плана счетов (ПАО/КИБ)"]).value
            
            print(f"\nСтрока {row_num}:")
            print(f"  Фильтр ПАО: '{filter_val}'")
            
            if str(filter_val).strip() == "ПАО":
                dt_col_letter = get_column_letter(col_dt_account)
                kt_col_letter = get_column_letter(col_kt_account)
                
                print(f"  Дт столбец: {dt_col_letter} (индекс {col_dt_account})")
                print(f"  Кт столбец: {kt_col_letter} (индекс {col_kt_account})")
                print(f"  Созданные в словаре:")
                print(f"    'dt_col': '{dt_col_letter}'")
                print(f"    'kt_col': '{kt_col_letter}'")
    
    return results


def check_provodki_columns(wb):
    """Проверить индексы столбцов в проводках"""
    print("\n" + "=" * 80)
    print("ПРОВЕРКА СТОЛБЦОВ ПРОВОДОК")
    print("=" * 80)
    
    ws = wb["Проводки по СП анализ"]
    header_row = 1
    
    # Проверяем все нужные столбцы
    columns_to_check = [
        "Дт",
        "Кт",
        "Номера строк развёрнутых проводок",
        "Дт Мэпинг Робот",
        "Кт Мэпинг Робот",
        "СП Робот"
    ]
    
    results = {}
    
    for col_name in columns_to_check:
        idx = get_column_index_by_name(ws, col_name, header_row)
        results[col_name] = idx
        
        if idx is None:
            print(f"❌ ОШИБКА: Столбец '{col_name}' НЕ НАЙДЕН!")
        else:
            letter = get_column_letter(idx)
            print(f"✅ '{col_name}': столбец {letter} (индекс {idx})")
    
    # Проверяем конкретные строки
    print(f"\n--- Проверка данных в проблемных строках ---")
    
    col_dt = results["Дт"]
    col_kt = results["Кт"]
    col_expanded = results["Номера строк развёрнутых проводок"]
    
    if col_dt and col_kt:
        for row_num in DEBUG_PROVODKI_ROWS:
            expanded_val = ws.cell(row=row_num, column=col_expanded).value if col_expanded else None
            
            print(f"\nСтрока {row_num}:")
            print(f"  Развёрнутые: '{expanded_val}'")
            
            if not expanded_val or not str(expanded_val).strip():
                dt_col_letter = get_column_letter(col_dt)
                kt_col_letter = get_column_letter(col_kt)
                
                print(f"  Дт столбец: {dt_col_letter} (индекс {col_dt})")
                print(f"  Кт столбец: {kt_col_letter} (индекс {col_kt})")
                print(f"  Созданные в словаре:")
                print(f"    'dt_col': '{dt_col_letter}'")
                print(f"    'kt_col': '{kt_col_letter}'")
            else:
                print(f"  ⚠️ ПРОПУСКАЕТСЯ: есть развёрнутые проводки")
    
    return results


def simulate_reference_creation():
    """Симулировать создание ссылок"""
    print("\n" + "=" * 80)
    print("СИМУЛЯЦИЯ СОЗДАНИЯ ССЫЛОК")
    print("=" * 80)
    
    # Примеры с проблемными строками
    examples = [
        {
            'scenario': 'Проводки → Шаблон (строка 2194 проводок)',
            'template_rows': [
                {'row': 2335, 'dt_col': 'B', 'kt_col': 'D', 'pokaza': 'СП-123'},
                {'row': 2857, 'dt_col': 'B', 'kt_col': 'D', 'pokaza': 'СП-123'}
            ]
        }
    ]
    
    for example in examples:
        print(f"\n--- {example['scenario']} ---")
        
        dt_refs = []
        kt_refs = []
        pokaza_values = []
        
        for template_info in example['template_rows']:
            # Проверяем наличие столбцов
            if not template_info.get('dt_col') or not template_info.get('kt_col'):
                print(f"  ❌ ОШИБКА: Строка {template_info['row']} - отсутствует dt_col или kt_col!")
                print(f"     dt_col='{template_info.get('dt_col')}', kt_col='{template_info.get('kt_col')}'")
                continue
            
            # Создаём ссылки
            dt_ref = f"='Шаблон пров и атриб'!{template_info['dt_col']}{template_info['row']}"
            kt_ref = f"='Шаблон пров и атриб'!{template_info['kt_col']}{template_info['row']}"
            
            print(f"  Строка {template_info['row']}:")
            print(f"    dt_ref: {dt_ref}")
            print(f"    kt_ref: {kt_ref}")
            
            dt_refs.append(dt_ref)
            kt_refs.append(kt_ref)
            
            # Pokaza
            pokaza = str(template_info['pokaza']).strip() if template_info.get('pokaza') else ""
            if pokaza and pokaza not in pokaza_values:
                pokaza_values.append(pokaza)
                print(f"    pokaza: '{pokaza}' (добавлено)")
        
        # Объединяем ссылки
        dt_value = ", ".join(dt_refs)
        kt_value = ", ".join(kt_refs)
        sp_value = ", ".join(pokaza_values)
        
        print(f"\n  Результаты:")
        print(f"    dt_refs список: {dt_refs}")
        print(f"    kt_refs список: {kt_refs}")
        print(f"    dt_value: '{dt_value}' [длина={len(dt_value)}]")
        print(f"    kt_value: '{kt_value}' [длина={len(kt_value)}]")
        print(f"    sp_value: '{sp_value}' [длина={len(sp_value)}]")
        
        # Проверяем условия записи
        print(f"\n  Проверка условий записи:")
        print(f"    'dt_value' in result: True")
        print(f"    result['dt_value']: '{dt_value}' → bool={bool(dt_value)}")
        
        if dt_value:
            print(f"    ✅ Дт Мэпинг БУДЕТ записан")
        else:
            print(f"    ❌ Дт Мэпинг НЕ БУДЕТ записан (пустая строка!)")
        
        if kt_value:
            print(f"    ✅ Кт Мэпинг БУДЕТ записан")
        else:
            print(f"    ❌ Кт Мэпинг НЕ БУДЕТ записан (пустая строка!)")
        
        if sp_value:
            print(f"    ✅ СП Робот БУДЕТ записан")
        else:
            print(f"    ❌ СП Робот НЕ БУДЕТ записан (пустая строка!)")


def check_actual_cells(wb):
    """Проверить что записано в ячейках мэппинга"""
    print("\n" + "=" * 80)
    print("ПРОВЕРКА ТЕКУЩЕГО СОДЕРЖИМОГО ЯЧЕЕК МЭППИНГА")
    print("=" * 80)
    
    # Проверка шаблона
    print("\n--- Шаблон пров и атриб ---")
    ws_template = wb["Шаблон пров и атриб"]
    header_row = 7
    
    col_dt_mapping = get_column_index_by_name(ws_template, "Дт Мэпинг Робот", header_row)
    col_kt_mapping = get_column_index_by_name(ws_template, "Кт Мэпинг Робот", header_row)
    
    for row_num in DEBUG_TEMPLATE_ROWS:
        dt_value = ws_template.cell(row=row_num, column=col_dt_mapping).value if col_dt_mapping else None
        kt_value = ws_template.cell(row=row_num, column=col_kt_mapping).value if col_kt_mapping else None
        
        print(f"\nСтрока {row_num}:")
        print(f"  Дт Мэпинг Робот: {repr(dt_value)}")
        print(f"  Кт Мэпинг Робот: {repr(kt_value)}")
        
        if dt_value is None and kt_value is None:
            print(f"  ❌ Оба столбца пустые!")
        elif dt_value is None or kt_value is None:
            print(f"  ⚠️ Один из столбцов пустой!")
    
    # Проверка проводок
    print("\n--- Проводки по СП анализ ---")
    ws_provodki = wb["Проводки по СП анализ"]
    header_row = 1
    
    col_dt_mapping = get_column_index_by_name(ws_provodki, "Дт Мэпинг Робот", header_row)
    col_kt_mapping = get_column_index_by_name(ws_provodki, "Кт Мэпинг Робот", header_row)
    col_sp_robot = get_column_index_by_name(ws_provodki, "СП Робот", header_row)
    
    for row_num in DEBUG_PROVODKI_ROWS:
        dt_value = ws_provodki.cell(row=row_num, column=col_dt_mapping).value if col_dt_mapping else None
        kt_value = ws_provodki.cell(row=row_num, column=col_kt_mapping).value if col_kt_mapping else None
        sp_value = ws_provodki.cell(row=row_num, column=col_sp_robot).value if col_sp_robot else None
        
        print(f"\nСтрока {row_num}:")
        print(f"  Дт Мэпинг Робот: {repr(dt_value)}")
        print(f"  Кт Мэпинг Робот: {repr(kt_value)}")
        print(f"  СП Робот: {repr(sp_value)}")
        
        if sp_value and (not dt_value or not kt_value):
            print(f"  🔴 ПРОБЛЕМА: СП Робот записан, но Дт/Кт Мэпинг НЕ записаны!")
            print(f"      Это подтверждает что ключи совпадают, но ссылки не записываются!")


def main():
    """Главная функция"""
    print("=" * 80)
    print("ДИАГНОСТИКА ПРОБЛЕМЫ С МЭППИНГОМ")
    print("=" * 80)
    print("\nЭтот скрипт проверит:")
    print("1. Правильность индексов столбцов")
    print("2. Создание букв столбцов через get_column_letter")
    print("3. Симуляцию создания ссылок")
    print("4. Текущее содержимое ячеек")
    
    # Получение файла
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    else:
        input_file = input("Введите путь к Excel-файлу: ").strip()
    
    print(f"\nОткрываем файл: {input_file}")
    
    # Открываем с data_only=True
    wb = openpyxl.load_workbook(input_file, data_only=True)
    
    # Проверки
    check_template_columns(wb)
    check_provodki_columns(wb)
    simulate_reference_creation()
    check_actual_cells(wb)
    
    wb.close()
    
    print("\n" + "=" * 80)
    print("ДИАГНОСТИКА ЗАВЕРШЕНА")
    print("=" * 80)
    print("\nВЫВОДЫ:")
    print("- Если все столбцы найдены и индексы правильные")
    print("- Если симуляция показывает что ссылки создаются")
    print("- НО в файле ячейки пустые")
    print("→ Проблема в логике условия записи в write_results_to_excel")
    print("\nПроверьте строки 632-637 в new.py:")
    print("  if 'dt_value' in result and result['dt_value']:")
    print("Возможно dt_value является пустой строкой ''!")
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
