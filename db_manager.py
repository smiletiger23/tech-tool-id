import sqlite3
import os
import shutil  # Added for shutil.rmtree
import re  # Added for re.match in parse_id_string


class FixtureDBManager:
    def __init__(self, db_name="my_fixtures_app.db", base_db_dir="."):
        self.db_name = db_name
        self.base_db_dir = base_db_dir
        self.conn = None
        self.cursor = None
        self.db_path = os.path.join(self.base_db_dir, self.db_name)

        try:
            os.makedirs(self.base_db_dir, exist_ok=True)
        except OSError as e:
            print(f"Ошибка: Не удалось создать директорию '{self.base_db_dir}': {e}")
            return

        try:
            db_exists = os.path.exists(self.db_path)
            self.conn = sqlite3.connect(self.db_path)
            self.conn.row_factory = sqlite3.Row
            self.cursor = self.conn.cursor()

            if not db_exists:
                print(f"База данных '{self.db_path}' не найдена. Создаем новую.")
            else:
                print(f"Подключение к базе данных '{self.db_path}' успешно установлено.")
            self.create_tables()

        except sqlite3.Error as e:
            print(f"Ошибка подключения к базе данных или создания таблиц: {e}")
            if self.conn:
                self.conn.close()
            self.conn = None
            self.cursor = None

    def create_tables(self):
        try:
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS Categories (
                    CategoryCode TEXT PRIMARY KEY NOT NULL UNIQUE,
                    CategoryName TEXT NOT NULL
                )
            """)
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS Series (
                    SeriesCode TEXT NOT NULL,
                    SeriesName TEXT NOT NULL,
                    CategoryCode TEXT NOT NULL,
                    PRIMARY KEY (CategoryCode, SeriesCode),
                    FOREIGN KEY (CategoryCode) REFERENCES Categories(CategoryCode)
                )
            """)
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS ItemNumbers (
                    ItemNumberCode TEXT NOT NULL,
                    ItemNumberName TEXT NOT NULL,
                    CategoryCode TEXT NOT NULL,
                    SeriesCode TEXT NOT NULL,
                    PRIMARY KEY (CategoryCode, SeriesCode, ItemNumberCode),
                    FOREIGN KEY (CategoryCode, SeriesCode) REFERENCES Series(CategoryCode, SeriesCode)
                )
            """)
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS Operations (
                    OperationCode TEXT PRIMARY KEY NOT NULL UNIQUE,
                    OperationName TEXT NOT NULL
                )
            """)
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS FixtureIDs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Category TEXT NOT NULL,
                    Series TEXT NOT NULL,
                    ItemNumber TEXT NOT NULL,
                    Operation TEXT NOT NULL,
                    FixtureNumber TEXT NOT NULL,
                    UniqueParts TEXT NOT NULL,
                    PartInAssembly TEXT NOT NULL,
                    PartQuantity TEXT NOT NULL,
                    AssemblyVersionCode TEXT NOT NULL,
                    IntermediateVersion TEXT, -- Может быть пустым
                    BasePath TEXT NOT NULL UNIQUE,
                    FullIDString TEXT NOT NULL UNIQUE,
                    FOREIGN KEY (Category) REFERENCES Categories(CategoryCode),
                    FOREIGN KEY (Category, Series) REFERENCES Series(CategoryCode, SeriesCode),
                    FOREIGN KEY (Category, Series, ItemNumber) REFERENCES ItemNumbers(CategoryCode, SeriesCode, ItemNumberCode),
                    FOREIGN KEY (Operation) REFERENCES Operations(OperationCode)
                )
            """)
            self.conn.commit()
            print("Все таблицы успешно созданы/проверены.")
        except sqlite3.Error as e:
            print(f"Ошибка при создании таблиц: {e}")

    def add_category(self, code, name):
        try:
            self.cursor.execute("INSERT OR IGNORE INTO Categories (CategoryCode, CategoryName) VALUES (?, ?)",
                                (code, name))
            self.conn.commit()
        except sqlite3.Error as e:
            print(f"Ошибка при добавлении категории {code}: {e}")

    def get_categories(self):
        try:
            self.cursor.execute("SELECT * FROM Categories ORDER BY CategoryCode")
            return [dict(row) for row in self.cursor.fetchall()]
        except sqlite3.Error as e:
            print(f"Ошибка при получении категорий: {e}")
            return []

    def add_series_description(self, category_code, series_code, series_name):
        try:
            self.cursor.execute(
                "INSERT OR IGNORE INTO Series (CategoryCode, SeriesCode, SeriesName) VALUES (?, ?, ?)",
                (category_code, series_code, series_name)
            )
            self.conn.commit()
        except sqlite3.Error as e:
            print(f"Ошибка при добавлении серии {series_code}: {e}")

    def get_series_by_category(self, category_code):
        try:
            self.cursor.execute(
                "SELECT SeriesCode, SeriesName FROM Series WHERE CategoryCode = ? ORDER BY SeriesCode",
                (category_code,)
            )
            return [dict(row) for row in self.cursor.fetchall()]
        except sqlite3.Error as e:
            print(f"Ошибка при получении серий для категории {category_code}: {e}")
            return []

    def get_series_descriptions(self):
        try:
            self.cursor.execute("SELECT * FROM Series ORDER BY CategoryCode, SeriesCode")
            return [dict(row) for row in self.cursor.fetchall()]
        except sqlite3.Error as e:
            print(f"Ошибка при получении всех серий: {e}")
            return []

    def add_item_number_description(self, category_code, series_code, item_number_code, item_number_name):
        try:
            self.cursor.execute(
                "INSERT OR IGNORE INTO ItemNumbers (CategoryCode, SeriesCode, ItemNumberCode, ItemNumberName) VALUES (?, ?, ?, ?)",
                (category_code, series_code, item_number_code, item_number_name)
            )
            self.conn.commit()
        except sqlite3.Error as e:
            print(f"Ошибка при добавлении изделия {item_number_code}: {e}")

    def get_items_by_category_and_series(self, category_code, series_code):
        try:
            self.cursor.execute(
                "SELECT ItemNumberCode, ItemNumberName FROM ItemNumbers WHERE CategoryCode = ? AND SeriesCode = ? ORDER BY ItemNumberCode",
                (category_code, series_code)
            )
            return [dict(row) for row in self.cursor.fetchall()]
        except sqlite3.Error as e:
            print(f"Ошибка при получении изделий для категории {category_code} и серии {series_code}: {e}")
            return []

    def get_item_number_descriptions(self):
        try:
            self.cursor.execute("SELECT * FROM ItemNumbers ORDER BY CategoryCode, SeriesCode, ItemNumberCode")
            return [dict(row) for row in self.cursor.fetchall()]
        except sqlite3.Error as e:
            print(f"Ошибка при получении всех изделий: {e}")
            return []

    def add_operation_description(self, operation_code, operation_name):
        try:
            self.cursor.execute("INSERT OR IGNORE INTO Operations (OperationCode, OperationName) VALUES (?, ?)",
                                (operation_code, operation_name))
            self.conn.commit()
        except sqlite3.Error as e:
            print(f"Ошибка при добавлении операции {operation_code}: {e}")

    def get_operation_descriptions(self):
        try:
            self.cursor.execute("SELECT * FROM Operations ORDER BY OperationCode")
            return [dict(row) for row in self.cursor.fetchall()]
        except sqlite3.Error as e:
            print(f"Ошибка при получении операций: {e}")
            return []

    def add_fixture_id(self, full_id_string):
        parsed_id = self.parse_id_string(full_id_string)
        if not parsed_id:
            print(f"Не удалось распарсить ID строку: {full_id_string}")
            return None

        self.cursor.execute("SELECT id FROM FixtureIDs WHERE FullIDString = ?", (full_id_string,))
        if self.cursor.fetchone():
            print(f"Оснастка с FullIDString '{full_id_string}' уже существует. Добавление отменено.")
            return None

        base_path = os.path.join(
            self.base_db_dir,
            parsed_id['Category'],
            f"{parsed_id['Series']}{parsed_id['ItemNumber']}",
            f"{parsed_id['Operation']}{parsed_id['FixtureNumber']}"
        )

        try:
            os.makedirs(base_path, exist_ok=True)
        except OSError as e:
            print(f"Ошибка при создании директории '{base_path}': {e}")
            return None

        try:
            self.cursor.execute("""
                INSERT INTO FixtureIDs (
                    Category, Series, ItemNumber, Operation, FixtureNumber,
                    UniqueParts, PartInAssembly, PartQuantity, AssemblyVersionCode, IntermediateVersion,
                    BasePath, FullIDString
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                parsed_id['Category'],
                parsed_id['Series'],
                parsed_id['ItemNumber'],
                parsed_id['Operation'],
                parsed_id['FixtureNumber'],
                parsed_id['UniqueParts'],
                parsed_id['PartInAssembly'],
                parsed_id['PartQuantity'],
                parsed_id['AssemblyVersionCode'],
                parsed_id['IntermediateVersion'],
                base_path,
                full_id_string
            ))
            self.conn.commit()
            return self.cursor.lastrowid
        except sqlite3.IntegrityError as e:
            print(f"Ошибка целостности данных при добавлении оснастки '{full_id_string}': {e}")
            return None
        except sqlite3.Error as e:
            print(f"Ошибка при добавлении оснастки '{full_id_string}': {e}")
            return None

    def get_existing_fixture_numbers(self, category_code, series_code, item_number_code, operation_code):
        try:
            self.cursor.execute(
                """
                SELECT FixtureNumber FROM FixtureIDs
                WHERE Category = ? AND Series = ? AND ItemNumber = ? AND Operation = ?
                ORDER BY FixtureNumber
                """,
                (category_code, series_code, item_number_code, operation_code)
            )
            return [row['FixtureNumber'] for row in self.cursor.fetchall()]
        except sqlite3.Error as e:
            print(
                f"Ошибка при получении существующих TT для {category_code}.{series_code}{item_number_code}.{operation_code}: {e}")
            return []

    def get_next_fixture_number(self, category_code, series_code, item_number_code, operation_code):
        existing_tts = self.get_existing_fixture_numbers(
            category_code, series_code, item_number_code, operation_code
        )
        if not existing_tts:
            return "01"

        numeric_tts = sorted([self._from_base36(tt) for tt in existing_tts])
        next_num = 1
        for num in numeric_tts:
            if num == next_num:
                next_num += 1
            elif num > next_num:
                break

        next_tt = self._to_base36(next_num).zfill(2)
        return next_tt

    def get_fixture_id_by_id(self, fixture_db_id):
        try:
            self.cursor.execute("SELECT * FROM FixtureIDs WHERE id = ?", (fixture_db_id,))
            result = self.cursor.fetchone()
            return dict(result) if result else None
        except sqlite3.Error as e:
            print(f"Ошибка при получении оснастки по ID {fixture_db_id}: {e}")
            return None

    def get_fixture_ids_with_descriptions(self, category_code=None, series_code=None, item_number_code=None,
                                          operation_code=None):
        query = """
            SELECT
                f.id,
                f.Category, c.CategoryName,
                f.Series, s.SeriesName,
                f.ItemNumber, i.ItemNumberName,
                f.Operation, o.OperationName,
                f.FixtureNumber,
                f.UniqueParts, f.PartInAssembly, f.PartQuantity,
                f.AssemblyVersionCode, f.IntermediateVersion,
                f.BasePath,
                f.FullIDString
            FROM
                FixtureIDs f
            LEFT JOIN Categories c ON f.Category = c.CategoryCode
            LEFT JOIN Series s ON f.Category = s.CategoryCode AND f.Series = s.SeriesCode
            LEFT JOIN ItemNumbers i ON f.Category = i.CategoryCode AND f.Series = i.SeriesCode AND f.ItemNumber = i.ItemNumberCode
            LEFT JOIN Operations o ON f.Operation = o.OperationCode
            WHERE 1=1
        """
        params = []

        if category_code:
            query += " AND f.Category = ?"
            params.append(category_code)
        if series_code:
            query += " AND f.Series = ?"
            params.append(series_code)
        if item_number_code:
            query += " AND f.ItemNumber = ?"
            params.append(item_number_code)
        if operation_code:
            query += " AND f.Operation = ?"
            params.append(operation_code)

        query += " ORDER BY f.Category, f.Series, f.ItemNumber, f.Operation, f.FixtureNumber"

        try:
            self.cursor.execute(query, tuple(params))
            return [dict(row) for row in self.cursor.fetchall()]
        except sqlite3.Error as e:
            print(f"Ошибка при получении всех оснасток с описаниями: {e}")
            return []

    def delete_fixture_id(self, fixture_db_id, delete_files=False):
        fixture_data = self.get_fixture_id_by_id(fixture_db_id)
        if not fixture_data:
            print(f"Оснастка с ID {fixture_db_id} не найдена.")
            return False

        if delete_files:
            base_path = fixture_data.get("BasePath")
            if base_path and os.path.exists(base_path):
                try:
                    shutil.rmtree(base_path)
                    print(f"Папка оснастки '{base_path}' успешно удалена.")
                except OSError as e:
                    print(f"Ошибка при удалении папки оснастки '{base_path}': {e}")
                    return False
            else:
                print(f"Путь к папке оснастки '{base_path}' не найден или не существует.")

        try:
            self.cursor.execute("DELETE FROM FixtureIDs WHERE id = ?", (fixture_db_id,))
            self.conn.commit()
            print(f"Оснастка с ID {fixture_db_id} удалена из базы данных.")
            return True
        except sqlite3.Error as e:
            print(f"Ошибка при удалении оснастки с ID {fixture_db_id} из БД: {e}")
            return False

    def _to_base36(self, number):
        """Converts an integer to a base36 string, excluding I, J, L, O."""
        filtered_chars = [c for c in '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ' if c not in 'IJLO']
        base = len(filtered_chars)

        if number < 0:
            raise ValueError("Cannot convert negative numbers to custom base36.")

        if number == 0:
            return '0'

        res = ''
        while number > 0:
            number, rem = divmod(number, base)
            res = filtered_chars[rem] + res
        return res

    def _from_base36(self, base36_string):
        """Converts a custom base36 string (excluding I, J, L, O) to an integer."""
        filtered_chars = [c for c in '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ' if c not in 'IJLO']
        base = len(filtered_chars)

        reverse_map = {char: i for i, char in enumerate(filtered_chars)}

        base36_string = base36_string.upper().strip()
        result = 0
        power = 0
        for char in reversed(base36_string):
            if char not in reverse_map:
                raise ValueError(f"Invalid character '{char}' for custom base36 encoding.")
            value = reverse_map[char]
            result += value * (base ** power)
            power += 1
        return result

    def parse_id_string(self, full_id_string):
        """
        Парсит полную строку ID оснастки на составляющие.
        Пример формата: KKK.SNN.DTT.AABBCC-VVW
        KKK (Category): 2-3 буквы
        S (Series): 1 символ
        NN (ItemNumber): 2 символа
        D (Operation): 1 символ
        TT (FixtureNumber): 2 символа
        AA (UniqueParts): 2 символа
        BB (PartInAssembly): 2 символа
        CC (PartQuantity): 2 символа
        VV (AssemblyVersionCode): 2 символа
        W (IntermediateVersion): 0 или 1 символ (опционально)
        """
        # Updated regex to match the parsing logic and handle optional 'W'
        match = re.match(
            r"^([A-Z]{2,3})\.([0-9A-Z]{1})([0-9A-Z]{2})\.([0-9A-Z]{1})([0-9A-Z]{2})\.([0-9A-Z]{2})([0-9A-Z]{2})([0-9A-Z]{2})-([0-9A-Z]{2})([0-9A-Z]{0,1})$",
            full_id_string)
        if not match:
            print(f"Ошибка парсинга ID строки: '{full_id_string}'")
            return None

        try:
            return {
                'Category': match.group(1),
                'Series': match.group(2),
                'ItemNumber': match.group(3),
                'Operation': match.group(4),
                'FixtureNumber': match.group(5),
                'UniqueParts': match.group(6),
                'PartInAssembly': match.group(7),
                'PartQuantity': match.group(8),
                'AssemblyVersionCode': match.group(9),
                'IntermediateVersion': match.group(10) if match.group(10) else None  # Store as None if empty
            }
        except Exception as e:
            print(f"Непредвиденная ошибка при парсинге ID '{full_id_string}': {e}")
            return None

    def close(self):
        if self.conn:
            self.conn.close()
            print("Соединение с базой данных закрыто.")