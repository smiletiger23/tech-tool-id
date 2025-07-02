import os
import openpyxl
from db_manager import FixtureDBManager
import shutil

class ExcelClassifierImporter:
    def __init__(self, db_manager_instance):
        self.db_manager = db_manager_instance
        self.sheet_configs = {
            "Категории": {
                "handler": self.db_manager.add_category,
                "columns": ["CategoryCode", "CategoryName"],
                "required_cols": ["CategoryCode", "CategoryName"]
            },
            "Серии": {
                "handler": self.db_manager.add_series_description,
                "columns": ["CategoryCode", "SeriesCode", "SeriesName"],
                "required_cols": ["CategoryCode", "SeriesCode", "SeriesName"]
            },
            "Изделия": {
                "handler": self.db_manager.add_item_number_description,
                "columns": ["CategoryCode", "SeriesCode", "ItemNumberCode", "ItemNumberName"],
                "required_cols": ["CategoryCode", "SeriesCode", "ItemNumberCode", "ItemNumberName"]
            },
            "Операции": {
                "handler": self.db_manager.add_operation_description,
                "columns": ["OperationCode", "OperationName"],
                "required_cols": ["OperationCode", "OperationName"]
            },
            # Листы "Оснастки" и "Версии_Сборок" удалены
        }

    def import_from_excel(self, excel_file_path):
        if not os.path.exists(excel_file_path):
            print(f"Ошибка: Файл '{excel_file_path}' не найден.")
            return False

        try:
            workbook = openpyxl.load_workbook(excel_file_path)
        except Exception as e:
            print(f"Ошибка при открытии Excel-файла '{excel_file_path}': {e}")
            return False

        print(f"\n--- Начинаем импорт данных из '{excel_file_path}' ---")

        for sheet_name, config in self.sheet_configs.items():
            if sheet_name in workbook.sheetnames:
                print(f"\nОбработка листа: '{sheet_name}'")
                sheet = workbook[sheet_name]
                header = [cell.value for cell in sheet[1]]

                missing_cols = [col for col in config["required_cols"] if col not in header]
                if missing_cols:
                    print(f"  Внимание: На листе '{sheet_name}' отсутствуют необходимые колонки: {', '.join(missing_cols)}. Пропуск листа.")
                    continue

                imported_count = 0
                skipped_count = 0
                for row_index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
                    row_data = {}
                    for col_idx, cell_value in enumerate(row):
                        if col_idx < len(header):
                            # Убедимся, что значение - строка, прежде чем применять strip()
                            row_data[header[col_idx]] = str(cell_value).strip() if cell_value is not None else ''

                    args = []
                    is_valid_row = True
                    for col_name in config["columns"]:
                        value = row_data.get(col_name)
                        # Изменение здесь: проверка на пустую строку или None после strip()
                        if value is None or value == '':
                            print(f"  Строка {row_index}: Пропуск из-за отсутствия данных в колонке '{col_name}'.")
                            is_valid_row = False
                            break
                        args.append(value) # value уже очищено от пробелов и не является None

                    if not is_valid_row:
                        skipped_count += 1
                        continue

                    # Если все данные валидны, вызываем обработчик
                    if config["handler"](*args):
                        imported_count += 1
                    else:
                        skipped_count += 1

                print(f"  Импортировано записей: {imported_count}")
                print(f"  Пропущено записей (дубликаты, ошибки): {skipped_count}")
            else:
                print(f"  Лист '{sheet_name}' не найден в файле. Пропуск.")

        print("\n--- Импорт данных завершен ---")
        return True

if __name__ == "__main__":
    excel_file = "classifier_data.xlsx" # Новое имя файла для минимальной версии

    def create_sample_excel_minimal(file_name):
        workbook = openpyxl.Workbook()

        # Категории
        sheet = workbook.active
        sheet.title = "Категории"
        sheet.append(["CategoryCode", "CategoryName"])
        sheet.append(["CS", "Коммутаторы"])
        sheet.append(["NEO", "Неорос"])
        sheet.append(["ROU", "Маршрутизаторы"])

        # Серии
        sheet = workbook.create_sheet("Серии")
        sheet.append(["CategoryCode", "SeriesCode", "SeriesName"])
        sheet.append(["CS", "X", "Расширенная серия коммутаторов"])
        sheet.append(["CS", "A", "Стандартная серия коммутаторов"])
        sheet.append(["ROU", "X", "Расширенная серия маршрутизаторов"])

        # Изделия
        sheet = workbook.create_sheet("Изделия")
        sheet.append(["CategoryCode", "SeriesCode", "ItemNumberCode", "ItemNumberName"])
        sheet.append(["CS", "X", "01", "CS2124"])
        sheet.append(["NEO", "A", "02", "NEO-FPGA-B"])
        sheet.append(["CS", "X", "10", "CS21XX (Старая серия)"])
        sheet.append(["ROU", "X", "03", "Router-X-Pro"])

        # Операции
        sheet = workbook.create_sheet("Операции")
        sheet.append(["OperationCode", "OperationName"])
        sheet.append(["F", "Фрезерование"])
        sheet.append(["D", "Сверление"])
        sheet.append(["M", "Монтаж"])

        workbook.save(file_name)
        print(f"Пример Excel-файла '{file_name}' создан.")

    # Создаем тестовый Excel-файл
    #create_sample_excel_minimal(excel_file) # <- Убедитесь, что эта строка закомментирована, если вы используете свой реальный файл classifier_data.xlsx

    db_name = "my_fixtures.db"
    base_dir = "fixture_database_root_app" # Убедитесь, что это имя каталога соответствует вашему проекту

    # Удаление существующей базы данных и папки для чистого запуска
    db_file_path = os.path.join(base_dir, db_name)
    if os.path.exists(db_file_path):
        print(f"Удаление существующей базы данных: {db_file_path}")
        os.remove(db_file_path)
    # Если вы хотите удалять всю папку, используйте это (но будьте осторожны):
    # if os.path.exists(base_dir):
    #     print(f"Удаление существующей корневой папки: {base_dir}")
    #     shutil.rmtree(base_dir)
    os.makedirs(base_dir, exist_ok=True) # Создаем папку, если она не существует

    db_manager = FixtureDBManager(db_name, base_db_dir=base_dir)

    importer = ExcelClassifierImporter(db_manager)

    importer.import_from_excel(excel_file)

    print("\n--- Проверка импортированных категорий ---")
    print(db_manager.get_categories())
    print("\n--- Проверка импортированных серий ---")
    print(db_manager.get_series_descriptions())
    print("\n--- Проверка импортированных изделий ---")
    print(db_manager.get_item_number_descriptions())
    print("\n--- Проверка импортированных операций ---")
    print(db_manager.get_operation_descriptions())


    # Пробуем добавить тестовую оснастку, чтобы убедиться, что пути генерируются корректно
    print("\n--- Тестирование добавления оснасток ---")
    # Добавьте здесь реальные коды из вашего Excel, если они отличаются от тестовых
    db_manager.add_fixture_id("CS.X01.F12.010101-V0Z") # Пример: Category=CS, Series=X, Item=01, Operation=F, Fixture=12, AssemblyVersion=V0, IntermediateVersion=Z
    db_manager.add_fixture_id("ROU.X03.M34.010101-X1")  # Пример: Category=ROU, Series=X, Item=03, Operation=M, Fixture=34, AssemblyVersion=X1, IntermediateVersion=None

    print("\n--- Все идентификаторы после теста ---")
    all_ids_after_test = db_manager.get_fixture_ids_with_descriptions()
    for item in all_ids_after_test:
        print(f"ID: {item['id']}, Category: {item['CategoryName']}, Series: {item['SeriesName']}, Item: {item['ItemNumberName']}, Operation: {item['OperationName']}, FixtureNumber: {item['FixtureNumber']}, AssemblyVersion: {item['AssemblyVersionCode']}, IntermediateVersion: {item['IntermediateVersion']}, Base Path: {item['BasePath']}")

    db_manager.close()
    # os.remove(excel_file) # Опционально: удалить тестовый файл Excel