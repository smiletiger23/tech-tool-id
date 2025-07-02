import sqlite3
import os


class FixtureDBManager:
    def __init__(self, db_name="my_fixtures_app.db", base_db_dir="."):
        self.db_name = db_name
        self.base_db_dir = base_db_dir
        self.conn = None
        self.cursor = None
        self.db_path = os.path.join(self.base_db_dir, self.db_name)

        # Убедитесь, что директория существует
        try:
            os.makedirs(self.base_db_dir, exist_ok=True)
        except OSError as e:
            print(f"Ошибка: Не удалось создать директорию '{self.base_db_dir}': {e}")
            return  # Выходим, если директория не может быть создана

        try:
            # Проверяем, существует ли файл БД, чтобы вывести соответствующее сообщение
            db_exists = os.path.exists(self.db_path)

            self.conn = sqlite3.connect(self.db_path)
            self.conn.row_factory = sqlite3.Row  # Позволяет получать данные как словари
            self.cursor = self.conn.cursor()

            if not db_exists:
                print(f"База данных '{self.db_path}' не найдена. Создаем новую.")
            else:
                print(f"Подключение к базе данных '{self.db_path}' успешно установлено.")

            # Включаем поддержку внешних ключей, это очень важно для целостности БД
            self.cursor.execute("PRAGMA foreign_keys = ON;")
            self.conn.commit()

            self.create_tables()  # Убедитесь, что это всегда вызывается при подключении

        except sqlite3.Error as e:
            print(f"Ошибка подключения к базе данных '{self.db_path}': {e}")
            self.conn = None  # Сбросить conn и cursor, чтобы предотвратить дальнейшие ошибки
            self.cursor = None
        except Exception as e:
            print(f"Неизвестная ошибка в __init__ db_manager: {e}")
            self.conn = None
            self.cursor = None

    def create_tables(self):
        # Эта функция теперь должна быть уверена, что self.cursor не None
        if self.cursor is None:
            print("Предупреждение: Не удалось создать таблицы, курсор БД равен None (проблема с подключением).")
            return  # Выходим, если курсор не инициализирован

        try:
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS Categories (
                    CategoryCode TEXT PRIMARY KEY,
                    CategoryName TEXT
                )
            """)
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS Series (
                    SeriesCode TEXT NOT NULL,
                    SeriesName TEXT,
                    CategoryCode TEXT NOT NULL,
                    PRIMARY KEY (SeriesCode, CategoryCode),
                    FOREIGN KEY (CategoryCode) REFERENCES Categories(CategoryCode)
                )
            """)
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS ItemNumbers (
                    ItemNumberCode TEXT NOT NULL,
                    ItemNumberName TEXT,
                    CategoryCode TEXT NOT NULL,
                    SeriesCode TEXT NOT NULL,
                    PRIMARY KEY (ItemNumberCode, CategoryCode, SeriesCode),
                    FOREIGN KEY (CategoryCode) REFERENCES Categories(CategoryCode),
                    FOREIGN KEY (SeriesCode, CategoryCode) REFERENCES Series(SeriesCode, CategoryCode)
                )
            """)
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS Operations (
                    OperationCode TEXT PRIMARY KEY,
                    OperationName TEXT
                )
            """)
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS FixtureIDs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    FullIDString TEXT UNIQUE,
                    Category TEXT,
                    Series TEXT,
                    ItemNumber TEXT,
                    Operation TEXT,
                    FixtureNumber TEXT,
                    UniqueParts TEXT,
                    PartInAssembly TEXT,
                    PartQuantity TEXT,
                    AssemblyVersionCode TEXT,
                    IntermediateVersion TEXT,
                    BasePath TEXT,
                    FOREIGN KEY (Category) REFERENCES Categories(CategoryCode),
                    FOREIGN KEY (Series, Category) REFERENCES Series(SeriesCode, CategoryCode),
                    FOREIGN KEY (ItemNumber, Category, Series) REFERENCES ItemNumbers(ItemNumberCode, CategoryCode, SeriesCode),
                    FOREIGN KEY (Operation) REFERENCES Operations(OperationCode)
                )
            """)
            self.conn.commit()
            print("Все таблицы успешно созданы/проверены.")
        except sqlite3.Error as e:
            print(f"Ошибка при создании/проверке таблиц: {e}")
        except Exception as e:
            print(f"Неизвестная ошибка в create_tables db_manager: {e}")

    def add_category(self, code, name):
        if self.conn is None:
            print("Ошибка: Соединение с БД не установлено.")
            return False
        try:
            self.cursor.execute("INSERT INTO Categories (CategoryCode, CategoryName) VALUES (?, ?)", (code, name))
            self.conn.commit()
            print(f"Категория '{code}' ('{name}') успешно добавлена.")
            return True
        except sqlite3.IntegrityError:
            print(f"Категория с кодом '{code}' уже существует.")
            return False
        except sqlite3.Error as e:
            print(f"Ошибка при добавлении категории '{code}': {e}")
            return False

    def add_series_description(self, category_code, series_code, series_name):
        if self.conn is None:
            print("Ошибка: Соединение с БД не установлено.")
            return False
        try:
            self.cursor.execute("INSERT INTO Series (CategoryCode, SeriesCode, SeriesName) VALUES (?, ?, ?)",
                                (category_code, series_code, series_name))
            self.conn.commit()
            print(f"Серия '{series_code}' ('{series_name}') для категории '{category_code}' успешно добавлена.")
            return True
        except sqlite3.IntegrityError as e:
            print(
                f"Серия с кодом '{series_code}' для категории '{category_code}' уже существует или нарушено ограничение уникальности: {e}.")
            return False
        except sqlite3.Error as e:
            print(f"Ошибка при добавлении серии '{series_code}' для категории '{category_code}': {e}")
            return False

    def add_item_number_description(self, category_code, series_code, item_number_code, item_number_name):
        if self.conn is None:
            print("Ошибка: Соединение с БД не установлено.")
            return False
        try:
            self.cursor.execute(
                "INSERT INTO ItemNumbers (CategoryCode, SeriesCode, ItemNumberCode, ItemNumberName) VALUES (?, ?, ?, ?)",
                (category_code, series_code, item_number_code, item_number_name))
            self.conn.commit()
            print(
                f"Изделие '{item_number_code}' ('{item_number_name}') для категории '{category_code}', серии '{series_code}' успешно добавлено.")
            return True
        except sqlite3.IntegrityError as e:
            print(
                f"Изделие с кодом '{item_number_code}' для категории '{category_code}', серии '{series_code}' уже существует или нарушено ограничение уникальности: {e}.")
            return False
        except sqlite3.Error as e:
            print(f"Ошибка при добавлении описания изделия: {e}")
            return False

    def add_operation_description(self, code, name):
        if self.conn is None:
            print("Ошибка: Соединение с БД не установлено.")
            return False
        try:
            self.cursor.execute("INSERT INTO Operations (OperationCode, OperationName) VALUES (?, ?)", (code, name))
            self.conn.commit()
            print(f"Операция '{code}' ('{name}') успешно добавлена.")
            return True
        except sqlite3.IntegrityError:
            print(f"Операция с кодом '{code}' уже существует.")
            return False
        except sqlite3.Error as e:
            print(f"Ошибка при добавлении операции '{code}': {e}")
            return False

    def add_fixture_id(self, full_id_string):
        if self.conn is None:
            print("Ошибка: Соединение с БД не установлено.")
            return None

        # Используем новый метод для парсинга строки ID
        parsed_data = self.parse_id_string(full_id_string)
        if parsed_data is None:
            print(f"Не удалось распарсить полный ID '{full_id_string}'. Оснастка не добавлена.")
            return None

        try:
            self.cursor.execute("""
                INSERT INTO FixtureIDs (
                    FullIDString, Category, Series, ItemNumber, Operation, FixtureNumber,
                    UniqueParts, PartInAssembly, PartQuantity, AssemblyVersionCode,
                    IntermediateVersion, BasePath
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                full_id_string,
                parsed_data['Category'],
                parsed_data['Series'],
                parsed_data['ItemNumber'],
                parsed_data['Operation'],
                parsed_data['FixtureNumber'],
                parsed_data['UniqueParts'],
                parsed_data['PartInAssembly'],
                parsed_data['PartQuantity'],
                parsed_data['AssemblyVersionCode'],
                parsed_data['IntermediateVersion'],
                parsed_data['BasePath']
            ))
            self.conn.commit()
            fixture_id = self.cursor.lastrowid
            print(f"Оснастка '{full_id_string}' успешно добавлена с ID: {fixture_id}")
            return fixture_id
        except sqlite3.IntegrityError as e:
            print(f"Оснастка с полным ID '{full_id_string}' уже существует или нарушено ограничение уникальности: {e}")
            return None
        except sqlite3.Error as e:
            print(f"Ошибка при добавлении оснастки '{full_id_string}': {e}")
            return None
        except Exception as e:
            print(f"Непредвиденная ошибка при добавлении оснастки '{full_id_string}': {e}")
            return None

    def get_categories(self):
        if self.conn is None: return []
        self.cursor.execute("SELECT CategoryCode, CategoryName FROM Categories")
        return [dict(row) for row in self.cursor.fetchall()]

    def get_series_descriptions(self):
        if self.conn is None: return []
        self.cursor.execute("SELECT CategoryCode, SeriesCode, SeriesName FROM Series")
        return [dict(row) for row in self.cursor.fetchall()]

    def get_series_by_category(self, category_code):
        if self.conn is None: return []
        self.cursor.execute("SELECT SeriesCode, SeriesName FROM Series WHERE CategoryCode = ?", (category_code,))
        return [dict(row) for row in self.cursor.fetchall()]

    def get_item_number_descriptions(self):
        if self.conn is None: return []
        self.cursor.execute("SELECT CategoryCode, SeriesCode, ItemNumberCode, ItemNumberName FROM ItemNumbers")
        return [dict(row) for row in self.cursor.fetchall()]

    def get_items_by_category_and_series(self, category_code, series_code):
        if self.conn is None: return []
        self.cursor.execute(
            "SELECT ItemNumberCode, ItemNumberName FROM ItemNumbers WHERE CategoryCode = ? AND SeriesCode = ?",
            (category_code, series_code))
        return [dict(row) for row in self.cursor.fetchall()]

    def get_operation_descriptions(self):
        if self.conn is None: return []
        self.cursor.execute("SELECT OperationCode, OperationName FROM Operations")
        return [dict(row) for row in self.cursor.fetchall()]

    def get_fixture_ids_with_descriptions(self, item_number_code=None):
        if self.conn is None: return []

        query = """
            SELECT
                f.id,
                f.FullIDString,
                c.CategoryName,
                s.SeriesName,
                i.ItemNumberName,
                o.OperationName,
                f.FixtureNumber,
                IFNULL(f.UniqueParts, '') AS UniqueParts,         -- Изменено: IFNULL для отображения пустых строк вместо None
                IFNULL(f.PartInAssembly, '') AS PartInAssembly,  -- Изменено
                IFNULL(f.PartQuantity, '') AS PartQuantity,      -- Изменено
                IFNULL(f.AssemblyVersionCode, '') AS AssemblyVersionCode, -- Изменено
                IFNULL(f.IntermediateVersion, '') AS IntermediateVersion, -- Изменено
                IFNULL(f.BasePath, '') AS BasePath               -- Изменено
            FROM
                FixtureIDs f
            LEFT JOIN Categories c ON f.Category = c.CategoryCode
            LEFT JOIN Series s ON f.Series = s.SeriesCode AND f.Category = s.CategoryCode
            LEFT JOIN ItemNumbers i ON f.ItemNumber = i.ItemNumberCode AND f.Category = i.CategoryCode AND f.Series = i.SeriesCode
            LEFT JOIN Operations o ON f.Operation = o.OperationCode
        """
        params = []

        if item_number_code and item_number_code != "Выберите изделие":
            query += " WHERE f.ItemNumber = ?"
            params.append(item_number_code)

        query += " ORDER BY f.id DESC"  # Changed to DESC for newest first

        try:
            self.cursor.execute(query, tuple(params))
            return [dict(row) for row in self.cursor.fetchall()]
        except sqlite3.Error as e:
            print(f"Ошибка при получении оснасток: {e}")
            return []

    def get_existing_fixture_numbers(self, category_code, series_code, item_number_code, operation_code):
        if self.conn is None: return []

        query = """
            SELECT DISTINCT FixtureNumber
            FROM FixtureIDs
            WHERE Category = ? AND Series = ? AND ItemNumber = ? AND Operation = ?
            ORDER BY FixtureNumber
        """
        params = (category_code, series_code, item_number_code, operation_code)

        try:
            self.cursor.execute(query, params)
            return [row[0] for row in self.cursor.fetchall()]
        except sqlite3.Error as e:
            print(f"Ошибка при получении существующих номеров оснасток: {e}")
            return []

    def get_next_fixture_number(self, category_code, series_code, item_number_code, operation_code):
        if self.conn is None: return "01"  # Возвращаем дефолт, если нет подключения

        query = """
            SELECT MAX(CAST(FixtureNumber AS INTEGER))
            FROM FixtureIDs
            WHERE Category = ? AND Series = ? AND ItemNumber = ? AND Operation = ?
        """
        params = (category_code, series_code, item_number_code, operation_code)

        try:
            self.cursor.execute(query, params)
            max_num = self.cursor.fetchone()[0]
            if max_num is None:
                return "01"  # Если нет существующих, начинаем с '01'
            else:
                next_num = max_num + 1
                return f"{next_num:02d}"  # Форматируем как "01", "02", ..., "10"
        except sqlite3.Error as e:
            print(f"Ошибка при получении следующего номера оснастки: {e}")
            return "01"  # В случае ошибки, возвращаем дефолт
        except ValueError:
            print(
                f"Предупреждение: FixtureNumber содержит нечисловые значения для {category_code}.{series_code}.{item_number_code}.{operation_code}. Начинаем с '01'.")
            return "01"  # На случай, если FixtureNumber не числовой

    def parse_id_string(self, full_id_string):
        """
        Парсит полную строку ID оснастки на составляющие части.
        Ожидаемый формат: Category.Series_ItemNumber.Operation_FixtureNumber.UniqueParts-AssemblyVersion-IntermediateVersion
        Пример: CS.X01.F12.010101-V0Z

        Возвращает словарь с распарсенными значениями или None в случае ошибки.
        """
        try:
            parts_arr = full_id_string.split('.')
            if len(parts_arr) < 4:
                print(f"Ошибка парсинга FullIDString: Недостаточно частей в '{full_id_string}'")
                return None

            category = parts_arr[0]

            # Парсинг Series и ItemNumber из второй части (например "X01")
            series_item = parts_arr[1]
            if len(series_item) >= 2:
                series = series_item[0]
                item_number = series_item[1:]
            else:
                print(f"Ошибка парсинга Series/ItemNumber из '{series_item}' в '{full_id_string}'")
                return None

            # Парсинг Operation и FixtureNumber из третьей части (например "F12")
            operation_fixture = parts_arr[2]
            if len(operation_fixture) >= 2:
                operation = operation_fixture[0]
                fixture_number = operation_fixture[1:]
            else:
                print(f"Ошибка парсинга Operation/FixtureNumber из '{operation_fixture}' в '{full_id_string}'")
                return None

            # Парсинг UniqueParts, AssemblyVersionCode, IntermediateVersion из четвертой части (например "010101-V0Z")
            version_parts = parts_arr[3].split('-')
            unique_parts = version_parts[0] if len(version_parts) > 0 else ""
            assembly_version_code = version_parts[1] if len(version_parts) > 1 else ""
            intermediate_version = version_parts[2] if len(
                version_parts) > 2 else ""  # Если W отсутствует, будет пустая строка

            # PartInAssembly, PartQuantity, BasePath - если их нет в FullIDString, они будут пустыми строками
            part_in_assembly = ""
            part_quantity = ""
            base_path = ""

            return {
                'Category': category,
                'Series': series,
                'ItemNumber': item_number,
                'Operation': operation,
                'FixtureNumber': fixture_number,
                'UniqueParts': unique_parts,
                'PartInAssembly': part_in_assembly,
                'PartQuantity': part_quantity,
                'AssemblyVersionCode': assembly_version_code,
                'IntermediateVersion': intermediate_version,
                'BasePath': base_path
            }
        except Exception as e:
            print(f"Непредвиденная ошибка при парсинге ID '{full_id_string}': {e}")
            return None

    def _to_base36(self, number):
        """Converts an integer to a base36 string."""
        alphabet = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        if number < 0:
            return '-' + self._to_base36(-number)
        res = ''
        while number > 0:
            number, rem = divmod(number, 36)
            res = alphabet[rem] + res
        return res or '0'

    def _from_base36(self, base36_string):
        """Converts a base36 string to an integer."""
        alphabet = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        base36_string = base36_string.upper()
        result = 0
        for char in base36_string:
            result = result * 36 + alphabet.index(char)
        return result

    def close(self):
        if self.conn:
            self.conn.close()
            print("Соединение с базой данных закрыто.")
