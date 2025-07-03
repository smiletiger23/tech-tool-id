import sqlite3
import os
import shutil
import re


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
                    BasePath TEXT NOT NULL,
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
        """Добавляет или обновляет категорию. Возвращает 'added', 'updated' или 'skipped'."""
        try:
            self.cursor.execute("SELECT CategoryName FROM Categories WHERE CategoryCode = ?", (code,))
            existing_name = self.cursor.fetchone()

            if existing_name:
                if existing_name['CategoryName'] != name:
                    self.cursor.execute("UPDATE Categories SET CategoryName = ? WHERE CategoryCode = ?", (name, code))
                    self.conn.commit()
                    return 'updated'
                else:
                    return 'skipped'
            else:
                self.cursor.execute("INSERT INTO Categories (CategoryCode, CategoryName) VALUES (?, ?)", (code, name))
                self.conn.commit()
                return 'added'
        except sqlite3.Error as e:
            print(f"Ошибка при обработке категории {code}: {e}")
            return 'error'

    def get_categories(self):
        try:
            self.cursor.execute("SELECT * FROM Categories ORDER BY CategoryCode")
            return [dict(row) for row in self.cursor.fetchall()]
        except sqlite3.Error as e:
            print(f"Ошибка при получении категорий: {e}")
            return []

    def add_series_description(self, category_code, series_code, series_name):
        """Добавляет или обновляет серию. Возвращает 'added', 'updated' или 'skipped'."""
        try:
            self.cursor.execute(
                "SELECT SeriesName FROM Series WHERE CategoryCode = ? AND SeriesCode = ?",
                (category_code, series_code)
            )
            existing_name = self.cursor.fetchone()

            if existing_name:
                if existing_name['SeriesName'] != series_name:
                    self.cursor.execute(
                        "UPDATE Series SET SeriesName = ? WHERE CategoryCode = ? AND SeriesCode = ?",
                        (series_name, category_code, series_code)
                    )
                    self.conn.commit()
                    return 'updated'
                else:
                    return 'skipped'
            else:
                self.cursor.execute(
                    "INSERT INTO Series (CategoryCode, SeriesCode, SeriesName) VALUES (?, ?, ?)",
                    (category_code, series_code, series_name)
                )
                self.conn.commit()
                return 'added'
        except sqlite3.Error as e:
            print(f"Ошибка при обработке серии {series_code} для категории {category_code}: {e}")
            return 'error'

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
        """Добавляет или обновляет изделие. Возвращает 'added', 'updated' или 'skipped'."""
        try:
            self.cursor.execute(
                "SELECT ItemNumberName FROM ItemNumbers WHERE CategoryCode = ? AND SeriesCode = ? AND ItemNumberCode = ?",
                (category_code, series_code, item_number_code)
            )
            existing_name = self.cursor.fetchone()

            if existing_name:
                if existing_name['ItemNumberName'] != item_number_name:
                    self.cursor.execute(
                        "UPDATE ItemNumbers SET ItemNumberName = ? WHERE CategoryCode = ? AND SeriesCode = ? AND ItemNumberCode = ?",
                        (item_number_name, category_code, series_code, item_number_code)
                    )
                    self.conn.commit()
                    return 'updated'
                else:
                    return 'skipped'
            else:
                self.cursor.execute(
                    "INSERT INTO ItemNumbers (CategoryCode, SeriesCode, ItemNumberCode, ItemNumberName) VALUES (?, ?, ?, ?)",
                    (category_code, series_code, item_number_code, item_number_name)
                )
                self.conn.commit()
                return 'added'
        except sqlite3.Error as e:
            print(
                f"Ошибка при обработке изделия {item_number_code} для категории {category_code} и серии {series_code}: {e}")
            return 'error'

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
        """Добавляет или обновляет операцию. Возвращает 'added', 'updated' или 'skipped'."""
        try:
            self.cursor.execute("SELECT OperationName FROM Operations WHERE OperationCode = ?", (operation_code,))
            existing_name = self.cursor.fetchone()

            if existing_name:
                if existing_name['OperationName'] != operation_name:
                    self.cursor.execute("UPDATE Operations SET OperationName = ? WHERE OperationCode = ?",
                                        (operation_name, operation_code))
                    self.conn.commit()
                    return 'updated'
                else:
                    return 'skipped'
            else:
                self.cursor.execute("INSERT INTO Operations (OperationCode, OperationName) VALUES (?, ?)",
                                    (operation_code, operation_name))
                self.conn.commit()
                return 'added'
        except sqlite3.Error as e:
            print(f"Ошибка при обработке операции {operation_code}: {e}")
            return 'error'

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

        # Проверяем, существует ли уже оснастка с таким же FullIDString
        self.cursor.execute("SELECT id FROM FixtureIDs WHERE FullIDString = ?", (full_id_string,))
        if self.cursor.fetchone():
            print(f"Оснастка с FullIDString '{full_id_string}' уже существует. Добавление отменено.")
            return None

        # Формируем имя папки для версии сборки (KKK.SNN.DTT.AA0000-VVW)
        folder_version_name = (
            f"{parsed_id['Category']}."
            f"{parsed_id['Series']}{parsed_id['ItemNumber']}."
            f"{parsed_id['Operation']}{parsed_id['FixtureNumber']}."
            f"{parsed_id['UniqueParts']}0000-"  # BB и CC заменены на 0000
            f"{parsed_id['AssemblyVersionCode']}"
            f"{parsed_id['IntermediateVersion'] if parsed_id['IntermediateVersion'] else ''}"
        )

        # Формируем полный путь к папке, используя новое имя папки версии сборки
        base_path = os.path.join(
            self.base_db_dir,
            parsed_id['Category'],
            f"{parsed_id['Category']}.{parsed_id['Series']}{parsed_id['ItemNumber']}",
            f"{parsed_id['Category']}.{parsed_id['Series']}{parsed_id['ItemNumber']}.{parsed_id['Operation']}{parsed_id['FixtureNumber']}",
            folder_version_name  # Используем новое имя папки
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
                base_path,  # Сохраняем путь к папке сборки
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
                f.AssemblyVersionCode,
                IFNULL(f.IntermediateVersion, '') AS IntermediateVersion,
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

    def get_latest_fixture_for_assembly(self, category, series, item_number, operation, fixture_number, unique_parts):
        """
        Retrieves the fixture with the latest version (VVW) for a given assembly base.
        The base is defined by KKK.SNN.DTT.AA.
        """
        query = """
            SELECT
                AssemblyVersionCode,
                IntermediateVersion
            FROM
                FixtureIDs
            WHERE
                Category = ? AND Series = ? AND ItemNumber = ? AND Operation = ? AND FixtureNumber = ? AND UniqueParts = ?
            ORDER BY
                AssemblyVersionCode DESC, IntermediateVersion DESC
        """
        params = (category, series, item_number, operation, fixture_number, unique_parts)

        try:
            self.cursor.execute(query, params)
            # Fetch all results to find the truly "latest" based on custom logic
            all_versions = self.cursor.fetchall()

            latest_fixture_data = None
            latest_version_components = None

            for row in all_versions:
                vv = row['AssemblyVersionCode']
                w = row['IntermediateVersion'] if row['IntermediateVersion'] else ''
                current_version_string = f"{vv}{w}"
                current_version_components = self._parse_version_components(current_version_string)

                if current_version_components['is_special_x']:
                    # Special 'X' versions are not part of the normal ordering for "latest" in this context
                    # If we encounter an 'X' version, we don't consider it for finding the *numeric* latest.
                    # If the user tries to create a numeric version after an 'X' version, it should be allowed.
                    continue

                if latest_version_components is None:
                    latest_version_components = current_version_components
                    latest_fixture_data = row
                else:
                    # Compare using the custom logic
                    if self.is_version_newer(
                            f"{latest_version_components['vv_code']}{latest_version_components['w_code'] if latest_version_components['w_code'] is not None else ''}",
                            current_version_string
                    ):
                        latest_version_components = current_version_components
                        latest_fixture_data = row

            return dict(latest_fixture_data) if latest_fixture_data else None

        except sqlite3.Error as e:
            print(f"Ошибка при получении последней версии для сборки: {e}")
            return None

    def _parse_version_components(self, version_string):
        """
        Parses a version string (VV or VVW) into its components for comparison.
        Returns a dictionary: {'vv_code': str, 'w_code': str or None, 'major_int': int, 'minor_int': int, 'w_int': int or None, 'is_special_x': bool}
        """
        vv_code = version_string[:2]
        w_code = version_string[2:] if len(version_string) > 2 else None

        is_special_x = False
        major_int = 0
        minor_int = 0
        w_int = None

        if vv_code.startswith('X'):
            is_special_x = True
            # For 'X' versions, we don't assign numeric major/minor for comparison purposes
            # We can still convert the second char of VV and W for internal consistency if needed,
            # but for ordering, they are treated as non-comparable in the standard sequence.
        else:
            try:
                major_int = int(vv_code[0])
                minor_int = int(vv_code[1])
            except ValueError:
                # Should be caught by validate_vv_input in GUI, but for robustness
                print(f"Warning: Non-numeric VV code '{vv_code}' encountered during parsing.")
                is_special_x = True  # Treat as special if parsing fails unexpectedly

        if w_code:
            # Convert char to integer (A=1, B=2, ..., Z=26)
            if 'A' <= w_code <= 'Z':
                w_int = ord(w_code) - ord('A') + 1
            else:
                # If W is not a letter, treat as invalid for comparison or special
                w_int = 0  # Or handle as error

        return {
            'vv_code': vv_code,
            'w_code': w_code,
            'major_int': major_int,
            'minor_int': minor_int,
            'w_int': w_int,
            'is_special_x': is_special_x
        }

    def is_version_newer(self, old_version_string, new_version_string):
        """
        Compares two version strings (VV or VVW) based on custom rules.
        Returns True if new_version_string is strictly newer than old_version_string.
        """
        old_comp = self._parse_version_components(old_version_string)
        new_comp = self._parse_version_components(new_version_string)

        # Rule: X versions are not checked for order against normal versions.
        # If new is 'X' type, it's considered valid (True) as per user's "without check" rule.
        if new_comp['is_special_x']:
            return True
        # If old is 'X' type and new is not, new is NOT strictly newer in the numeric sequence.
        # This means a normal version cannot be "newer" than a special 'X' version in the strict sense.
        if old_comp['is_special_x'] and not new_comp['is_special_x']:
            return False
        # If both are 'X' type, their relative order is not defined by this function.
        # For simplicity, if both are 'X', we consider them not strictly newer than each other.
        if old_comp['is_special_x'] and new_comp['is_special_x']:
            return False  # Or True if user implies any X is newer than another X, but that's not specified.

        # Compare major version (first digit of VV)
        if new_comp['major_int'] > old_comp['major_int']:
            return True
        if new_comp['major_int'] < old_comp['major_int']:
            return False

        # If major versions are equal, compare minor version (second digit of VV)
        if new_comp['minor_int'] > old_comp['minor_int']:
            return True
        if new_comp['minor_int'] < old_comp['minor_int']:
            return False

        # If major and minor versions are equal, compare intermediate version (W)
        # Rule: 01A newer than 01, 01B newer than 01A
        if new_comp['w_int'] is not None and old_comp['w_int'] is None:
            return True  # e.g., 01A is newer than 01
        if new_comp['w_int'] is None and old_comp['w_int'] is not None:
            return False  # e.g., 01 is not newer than 01A

        # If both have W or both don't have W, compare W values
        if new_comp['w_int'] is not None and old_comp['w_int'] is not None:
            return new_comp['w_int'] > old_comp['w_int']  # e.g., 01B > 01A

        # If both are exactly the same (VV and W), not strictly newer
        return False

    def delete_fixture_id(self, fixture_db_id, delete_files=False):
        fixture_data = self.get_fixture_id_by_id(fixture_db_id)
        if not fixture_data:
            print(f"Оснастка с ID {fixture_db_id} не найдена.")
            return False

        if delete_files:
            base_path = fixture_data.get("BasePath")
            if base_path and os.path.exists(base_path):
                try:
                    # Проверяем, есть ли другие записи в БД, использующие этот же BasePath
                    self.cursor.execute("SELECT COUNT(*) FROM FixtureIDs WHERE BasePath = ?", (base_path,))
                    count_referencing_path = self.cursor.fetchone()[0]

                    if count_referencing_path == 1:  # Если это единственная запись, ссылающаяся на этот путь
                        shutil.rmtree(base_path)
                        print(f"Папка оснастки '{base_path}' успешно удалена.")
                    else:
                        print(
                            f"Папка оснастки '{base_path}' не будет удалена, так как на нее ссылаются другие записи ({count_referencing_path} шт.).")
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
                'IntermediateVersion': match.group(10) if match.group(10) else None
            }
        except Exception as e:
            print(f"Непредвиденная ошибка при парсинге ID '{full_id_string}': {e}")
            return None

    def close(self):
        if self.conn:
            self.conn.close()
            print("Соединение с базой данных закрыто.")
