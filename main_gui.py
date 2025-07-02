import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import shutil
from db_manager import FixtureDBManager
from excel_importer import ExcelClassifierImporter
import re

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class FixtureApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Управление Оснастками (Tech-Tool-ID)")
        self.geometry("900x700")

        self.db_manager = FixtureDBManager(db_name="my_fixtures_app.db", base_db_dir="fixture_database_root_app")

        # --- Переменные для Combobox'ов ---
        self.category_code_var = ctk.StringVar(value="Выберите категорию")
        self.series_code_var = ctk.StringVar(value="Выберите серию")
        self.item_number_code_var = ctk.StringVar(value="Выберите изделие")
        self.operation_code_var = ctk.StringVar(value="Выберите операцию")

        self.fixture_number_code_var = ctk.StringVar(value="Будет сгенерирован")  # TT
        self.unique_parts_aa_var = ctk.StringVar(value="01")  # AA
        self.part_in_assembly_bb_var = ctk.StringVar(value="01")  # BB
        self.part_quantity_cc_var = ctk.StringVar(value="01")  # CC
        self.assembly_version_vv_var = ctk.StringVar(value="V0")  # VV
        self.intermediate_version_w_var = ctk.StringVar(value="")  # W (опционально)

        # --- Создание элементов интерфейса ---
        self.create_frame = ctk.CTkFrame(self)
        self.create_frame.pack(pady=10, padx=10, fill="x", expand=False)
        self.create_frame.grid_columnconfigure(0, weight=1)
        self.create_frame.grid_columnconfigure(1, weight=3)
        self.create_frame.grid_columnconfigure(2, weight=1)
        self.create_frame.grid_columnconfigure(3, weight=3)

        # Категория
        ctk.CTkLabel(self.create_frame, text="Категория (KKK):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.category_combobox = ctk.CTkComboBox(self.create_frame, variable=self.category_code_var, values=[],
                                                 command=self.on_category_selected, state="readonly")
        self.category_combobox.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        # Серия
        ctk.CTkLabel(self.create_frame, text="Серия (S):").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.series_combobox = ctk.CTkComboBox(self.create_frame, variable=self.series_code_var, values=[],
                                               command=self.on_series_selected, state="readonly")
        self.series_combobox.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        # Изделие
        ctk.CTkLabel(self.create_frame, text="Изделие (NN):").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.item_number_combobox = ctk.CTkComboBox(self.create_frame, variable=self.item_number_code_var, values=[],
                                                    command=self.on_item_number_selected, state="readonly")
        self.item_number_combobox.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

        # Операция
        ctk.CTkLabel(self.create_frame, text="Операция (D):").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.operation_combobox = ctk.CTkComboBox(self.create_frame, variable=self.operation_code_var, values=[],
                                                  command=self.on_operation_selected, state="readonly")
        self.operation_combobox.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

        # Номер оснастки (TT) - теперь Combobox
        ctk.CTkLabel(self.create_frame, text="Номер оснастки (TT):").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.fixture_number_combobox = ctk.CTkComboBox(self.create_frame, variable=self.fixture_number_code_var,
                                                       values=[],
                                                       command=self.on_fixture_number_selected, state="readonly")
        self.fixture_number_combobox.grid(row=0, column=3, padx=5, pady=5, sticky="ew")

        # Уникальные части (AA)
        ctk.CTkLabel(self.create_frame, text="Уникальных деталей (AA):").grid(row=1, column=2, padx=5, pady=5,
                                                                              sticky="w")
        self.unique_parts_entry = ctk.CTkEntry(self.create_frame, textvariable=self.unique_parts_aa_var, width=50)
        self.unique_parts_entry.grid(row=1, column=3, padx=5, pady=5, sticky="ew")
        self.unique_parts_entry.bind("<FocusOut>", self.validate_aa_bb_input)
        self.unique_parts_entry.bind("<Return>", self.validate_aa_bb_input)

        # Деталь в сборке (BB)
        ctk.CTkLabel(self.create_frame, text="Деталь в сборке (BB):").grid(row=2, column=2, padx=5, pady=5, sticky="w")
        self.part_in_assembly_entry = ctk.CTkEntry(self.create_frame, textvariable=self.part_in_assembly_bb_var,
                                                   width=50)
        self.part_in_assembly_entry.grid(row=2, column=3, padx=5, pady=5, sticky="ew")
        self.part_in_assembly_entry.bind("<FocusOut>", self.validate_aa_bb_input)
        self.part_in_assembly_entry.bind("<Return>", self.validate_aa_bb_input)

        # Количество (CC)
        ctk.CTkLabel(self.create_frame, text="Количество (CC):").grid(row=3, column=2, padx=5, pady=5, sticky="w")
        self.part_quantity_entry = ctk.CTkEntry(self.create_frame, textvariable=self.part_quantity_cc_var, width=50)
        self.part_quantity_entry.grid(row=3, column=3, padx=5, pady=5, sticky="ew")
        self.part_quantity_entry.bind("<FocusOut>", self.validate_cc_input)
        self.part_quantity_entry.bind("<Return>", self.validate_cc_input)

        # Версия сборки (VV)
        ctk.CTkLabel(self.create_frame, text="Версия сборки (VV):").grid(row=4, column=0, padx=5, pady=5, sticky="w")
        self.assembly_version_entry = ctk.CTkEntry(self.create_frame, textvariable=self.assembly_version_vv_var,
                                                   width=50)
        self.assembly_version_entry.grid(row=4, column=1, padx=5, pady=5, sticky="ew")
        self.assembly_version_entry.bind("<FocusOut>", self.validate_vv_input)
        self.assembly_version_entry.bind("<Return>", self.validate_vv_input)

        # Промежуточная версия (W)
        ctk.CTkLabel(self.create_frame, text="Пром. версия (W):").grid(row=4, column=2, padx=5, pady=5, sticky="w")
        self.intermediate_version_entry = ctk.CTkEntry(self.create_frame, textvariable=self.intermediate_version_w_var,
                                                       width=50)
        self.intermediate_version_entry.grid(row=4, column=3, padx=5, pady=5, sticky="ew")
        self.intermediate_version_entry.bind("<FocusOut>", self.validate_w_input)
        self.intermediate_version_entry.bind("<Return>", self.validate_w_input)

        self.add_button = ctk.CTkButton(self.create_frame, text="Создать оснастку", command=self.create_fixture_command)
        self.add_button.grid(row=5, column=0, columnspan=4, pady=10, padx=10)

        # Фрейм для сообщений
        self.status_frame = ctk.CTkFrame(self)
        self.status_frame.pack(pady=5, padx=10, fill="x", expand=False)
        self.status_label = ctk.CTkLabel(self.status_frame, text="Статус: Готов", wraplength=850)
        self.status_label.pack(pady=5, padx=10, fill="x", expand=True)

        # Фрейм для копирования файлов
        self.file_copy_frame = ctk.CTkFrame(self)
        self.file_copy_frame.pack(pady=10, padx=10, fill="x", expand=False)

        self.file_label = ctk.CTkLabel(self.file_copy_frame, text="Файлы для копирования:")
        self.file_label.pack(pady=(10, 0), padx=10, anchor="w")

        self.selected_files_path = ctk.CTkEntry(self.file_copy_frame, placeholder_text="Выбранные файлы...",
                                                state="readonly")
        self.selected_files_path.pack(pady=(0, 5), padx=10, fill="x", expand=True)

        self.select_files_button = ctk.CTkButton(self.file_copy_frame, text="Выбрать файлы", command=self.select_files)
        self.select_files_button.pack(pady=(0, 5), padx=10)

        self.copy_files_button = ctk.CTkButton(self.file_copy_frame, text="Копировать файлы в папку выбранной оснастки",
                                               command=self.copy_files)
        self.copy_files_button.pack(pady=(0, 10), padx=10)

        # Фрейм для списка оснасток
        self.list_frame = ctk.CTkFrame(self)
        self.list_frame.pack(pady=10, padx=10, fill="both", expand=True)

        self.list_label = ctk.CTkLabel(self.list_frame, text="Список существующих оснасток (для выбранного изделия):")
        self.list_label.pack(pady=(10, 0), padx=10, anchor="w")

        self.fixture_list_text = ctk.CTkTextbox(self.list_frame, width=860, height=200, state="disabled")
        self.fixture_list_text.pack(pady=(0, 10), padx=10, fill="both", expand=True)

        # Кнопки для управления списком
        self.list_buttons_frame = ctk.CTkFrame(self)
        self.list_buttons_frame.pack(pady=(0, 10), padx=10, fill="x", expand=False)
        self.list_buttons_frame.grid_columnconfigure((0, 1, 2), weight=1)

        # The refresh button will now always refresh based on current item selection
        self.refresh_button = ctk.CTkButton(self.list_buttons_frame, text="Обновить список",
                                            command=lambda: self.load_fixtures_to_list(
                                                self._get_code_from_display_text(self.item_number_code_var.get())))
        self.refresh_button.grid(row=0, column=0, padx=(10, 5), pady=5, sticky="w")

        self.open_folder_button = ctk.CTkButton(self.list_buttons_frame, text="Открыть папку оснастки",
                                                command=self.open_selected_folder)
        self.open_folder_button.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        self.delete_button = ctk.CTkButton(self.list_buttons_frame, text="Удалить выбранную оснастку", fg_color="red",
                                           hover_color="#8B0000", command=self.delete_selected_fixture)
        self.delete_button.grid(row=0, column=2, padx=(5, 10), pady=5, sticky="e")

        self.selected_fixture_id_in_list = None

        self.load_all_combobox_data()
        # Initially, load fixtures based on the default selected item (if any)
        self.load_fixtures_to_list(self._get_code_from_display_text(self.item_number_code_var.get()))

    def set_status(self, message, is_error=False):
        self.status_label.configure(text=f"Статус: {message}", text_color="red" if is_error else "green")
        self.update_idletasks()

    def _get_code_from_display_text(self, display_text):
        print(f"DEBUG: _get_code_from_display_text input: '{display_text}'")
        if not display_text:
            return ""
        # Check if the text contains a description in parentheses
        match = re.match(r"^(.*?)\s*\(.*\)$", display_text)
        if match:
            code = match.group(1).strip()
            print(f"DEBUG: _get_code_from_display_text output (parsed): '{code}'")
            return code
        print(f"DEBUG: _get_code_from_display_text output (raw): '{display_text}'")
        return display_text.strip()

    def _format_display_text(self, code, name):
        return f"{code} ({name})" if name else code

    def load_all_combobox_data(self):
        print("DEBUG: Загрузка всех данных для Combobox'ов...")
        self.categories_data = self.db_manager.get_categories()
        category_options = [self._format_display_text(c['CategoryCode'], c['CategoryName']) for c in
                            self.categories_data]
        self.category_combobox.configure(values=category_options)
        if category_options:
            self.category_code_var.set(category_options[0])
            self.on_category_selected(self.category_code_var.get())
        else:
            self.category_code_var.set("Нет данных")
            self.category_combobox.configure(values=["Нет данных"], state="disabled")
            print("DEBUG: Нет доступных категорий.")

        self.operations_data = self.db_manager.get_operation_descriptions()
        operation_options = [self._format_display_text(o['OperationCode'], o['OperationName']) for o in
                             self.operations_data]
        self.operation_combobox.configure(values=operation_options)
        if operation_options:
            self.operation_code_var.set(operation_options[0])
            self.on_operation_selected(self.operation_code_var.get())
        else:
            self.operation_code_var.set("Нет данных")
            self.operation_combobox.configure(values=["Нет данных"], state="disabled")
            print("DEBUG: Нет доступных операций.")
        print("DEBUG: Загрузка всех данных для Combobox'ов завершена.")

    def on_category_selected(self, event=None):
        category_display_text = self.category_code_var.get()
        print(f"DEBUG: on_category_selected вызван с '{category_display_text}'")

        category_code = self._get_code_from_display_text(category_display_text)
        print(f"DEBUG: _get_code_from_display_text input: '{category_display_text}'")
        print(f"DEBUG: _get_code_from_display_text output (parsed): '{category_code}'")

        if "Выберите" in category_code or "Нет данных" in category_code:
            self.series_combobox.configure(values=[])
            self.series_code_var.set("Выберите серию")
            self.series_combobox.configure(state="disabled")

            self.item_number_combobox.configure(values=[])
            self.item_number_code_var.set("Выберите изделие")
            self.item_number_combobox.configure(state="disabled")

            self.fixture_number_combobox.configure(values=[])
            self.fixture_number_code_var.set("Заполните поля выше")
            self.fixture_number_combobox.configure(state="disabled")
            print("DEBUG: Категория не выбрана, сброс зависимых полей.")
            return

        self.series_combobox.configure(state="readonly")  # Включаем Combobox Серии

        # Передаем декодированный category_code
        self.series_data = self.db_manager.get_series_by_category(category_code)

        # Добавим print для проверки данных серии
        print(f"DEBUG: Полученные данные серий для категории '{category_code}': {self.series_data}")

        if not self.series_data:
            self.series_combobox.configure(values=[])
            self.series_code_var.set("Нет данных")
            self.series_combobox.configure(state="disabled")
            print(f"DEBUG: Нет серий для категории '{category_code}'.")
        else:
            options = [f"{s['SeriesCode']} ({s['SeriesName']})" for s in self.series_data]
            self.series_combobox.configure(values=options)
            self.series_code_var.set(options[0])  # Устанавливаем первый элемент по умолчанию

        # Вызываем on_series_selected, чтобы обновить зависимые поля (Изделие, Оснастка)
        # Это важно, так как при изменении категории, серия также меняется.
        self.on_series_selected(self.series_code_var.get())

    def on_series_selected(self, event=None):
        category_code = self._get_code_from_display_text(self.category_code_var.get())
        series_display_text = self.series_code_var.get()
        print(f"DEBUG: on_series_selected вызван с '{series_display_text}'")

        series_code = self._get_code_from_display_text(series_display_text)
        print(f"DEBUG: _get_code_from_display_text input: '{series_display_text}'")
        print(f"DEBUG: _get_code_from_display_text output (parsed): '{series_code}'")

        if "Выберите" in series_code or "Нет данных" in series_code or \
                "Выберите" in category_code or "Нет данных" in category_code:
            self.item_number_combobox.configure(values=[])
            self.item_number_code_var.set("Выберите изделие")
            self.item_number_combobox.configure(state="disabled")

            self.fixture_number_combobox.configure(values=[])
            self.fixture_number_code_var.set("Заполните поля выше")
            self.fixture_number_combobox.configure(state="disabled")
            print("DEBUG: Серия или Категория не выбраны, сброс зависимых полей.")
            return

        self.item_number_combobox.configure(state="readonly")  # Включаем Combobox Изделия

        # Передаем декодированные category_code и series_code
        self.items_data = self.db_manager.get_items_by_category_and_series(category_code, series_code)

        # Добавим print для проверки данных изделий
        print(
            f"DEBUG: Полученные данные изделий для категории '{category_code}', серии '{series_code}': {self.items_data}")

        if not self.items_data:
            self.item_number_combobox.configure(values=[])
            self.item_number_code_var.set("Нет данных")
            self.item_number_combobox.configure(state="disabled")
            print(f"DEBUG: Нет изделий для категории '{category_code}', серии '{series_code}'.")
        else:
            options = [f"{i['ItemNumberCode']} ({i['ItemNumberName']})" for i in self.items_data]
            self.item_number_combobox.configure(values=options)
            self.item_number_code_var.set(options[0])  # Устанавливаем первый элемент по умолчанию

        # Вызываем update_fixture_number_combobox, чтобы обновить поле Оснастки
        self.update_fixture_number_combobox()

    def on_item_number_selected(self, selected_item_display):
        print(f"DEBUG: on_item_number_selected вызван с '{selected_item_display}'")
        # Обновляем Combobox TT только после выбора всех 4х компонентов
        self.update_fixture_number_combobox()

        # <-- NEW: Filter fixture list by selected item -->
        item_code = self._get_code_from_display_text(selected_item_display)
        self.load_fixtures_to_list(item_code)

    def on_operation_selected(self, selected_operation_display):
        print(f"DEBUG: on_operation_selected вызван с '{selected_operation_display}'")
        # Обновляем Combobox TT только после выбора всех 4х компонентов
        self.update_fixture_number_combobox()

    def update_fixture_number_combobox(self):
        """Обновляет выпадающий список номеров оснасток (TT) и устанавливает выбор по умолчанию."""
        print("DEBUG: update_fixture_number_combobox вызван.")
        category_code = self._get_code_from_display_text(self.category_code_var.get())
        series_code = self._get_code_from_display_text(self.series_code_var.get())
        item_number_code = self._get_code_from_display_text(self.item_number_code_var.get())
        operation_code = self._get_code_from_display_text(self.operation_code_var.get())

        # Проверка, что все необходимые поля выбраны и не содержат "Выберите..." или "Нет данных"
        if "Выберите" in category_code or "Нет данных" in category_code or \
                "Выберите" in series_code or "Нет данных" in series_code or \
                "Выберите" in item_number_code or "Нет данных" in item_number_code or \
                "Выберите" in operation_code or "Нет данных" in operation_code:
            self.fixture_number_combobox.configure(values=[])
            self.fixture_number_code_var.set("Заполните поля выше")
            self.fixture_number_combobox.configure(state="disabled")
            print("DEBUG: Не все поля выбраны для обновления TT.")
            return

        self.fixture_number_combobox.configure(state="readonly")  # Включаем Combobox TT

        existing_tts = self.db_manager.get_existing_fixture_numbers(
            category_code, series_code, item_number_code, operation_code
        )
        print(
            f"DEBUG: Существующие TT для {category_code}.{series_code}{item_number_code}.{operation_code}: {existing_tts}")

        options = ["<Создать новый TT>"] + existing_tts
        self.fixture_number_combobox.configure(values=options)

        # Устанавливаем выбор по умолчанию: либо первый из существующих, либо "Создать новый TT"
        if existing_tts:
            self.fixture_number_code_var.set(existing_tts[0])  # По умолчанию выбираем первый существующий TT
        else:
            self.fixture_number_code_var.set("<Создать новый TT>")  # Если нет существующих, предлагаем создать новый
            print("DEBUG: Нет существующих TT, выбран '<Создать новый TT>'.")

        # Сразу вызываем обработчик, чтобы TT сгенерировался, если выбран "<Создать новый TT>"
        # Или чтобы переменная обновилась, если выбран существующий TT
        self.on_fixture_number_selected(self.fixture_number_code_var.get())
        print("DEBUG: update_fixture_number_combobox завершен.")

    def on_fixture_number_selected(self, selected_tt_value):
        """Обработчик выбора TT из Combobox."""
        print(f"DEBUG: on_fixture_number_selected вызван с '{selected_tt_value}'")
        if selected_tt_value == "<Создать новый TT>":
            self.generate_next_fixture_number()  # Генерируем новый TT
        elif selected_tt_value in ["Заполните поля выше", "Ошибка TT", "Нет данных", ""]:
            # Не делаем ничего, если это промежуточное или ошибочное состояние
            print(f"DEBUG: Промежуточное/ошибочное состояние TT: '{selected_tt_value}'")
            pass
        else:
            # Пользователь выбрал существующий TT
            self.fixture_number_code_var.set(selected_tt_value)
            print(f"DEBUG: Выбран существующий TT: '{selected_tt_value}'")

    def generate_next_fixture_number(self):
        """Генерирует следующий TT на основе выбранных параметров."""
        print("DEBUG: generate_next_fixture_number вызван.")
        category_code = self._get_code_from_display_text(self.category_code_var.get())
        series_code = self._get_code_from_display_text(self.series_code_var.get())
        item_number_code = self._get_code_from_display_text(self.item_number_code_var.get())
        operation_code = self._get_code_from_display_text(self.operation_code_var.get())

        if "Выберите" in category_code or "Нет данных" in category_code or \
                "Выберите" in series_code or "Нет данных" in series_code or \
                "Выберите" in item_number_code or "Нет данных" in item_number_code or \
                "Выберите" in operation_code or "Нет данных" in operation_code:
            self.fixture_number_code_var.set("Заполните поля выше")
            print("DEBUG: Не все поля выбраны для генерации TT. Установлено 'Заполните поля выше'.")
            return

        next_tt = self.db_manager.get_next_fixture_number(
            category_code, series_code, item_number_code, operation_code
        )
        if next_tt:
            self.fixture_number_code_var.set(next_tt)
            print(f"DEBUG: Сгенерирован следующий TT: '{next_tt}'")
        else:
            self.fixture_number_code_var.set("Ошибка TT")
            self.set_status("Не удалось сгенерировать номер оснастки (TT). Проверьте консоль.", is_error=True)
            print("DEBUG: Ошибка при генерации TT.")

    def validate_base36_input(self, entry_var, length, forbidden_chars="IJLO", can_be_empty=False):
        value = entry_var.get().strip().upper()

        if can_be_empty and not value:
            return True

        if not value and not can_be_empty:
            messagebox.showerror("Ошибка ввода", f"Поле не может быть пустым. Длина должна быть {length} символа(ов).")
            entry_var.set("0" * length)
            return False

        if len(value) != length:
            messagebox.showerror("Ошибка ввода", f"Значение '{value}' должно быть длиной {length} символа(ов).")
            entry_var.set("0" * length)
            return False

        valid_chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        for char in value:
            if char not in valid_chars or char in forbidden_chars:
                messagebox.showerror("Ошибка ввода",
                                     f"Значение '{value}' содержит недопустимые символы. Допустимы 0-9, A-Z (кроме I, J, L, O).")
                entry_var.set("0" * length)
                return False
        return True

    def validate_aa_bb_input(self, event=None):
        if not self.validate_base36_input(self.unique_parts_aa_var, 2): return False
        if not self.validate_base36_input(self.part_in_assembly_bb_var, 2): return False

        aa_str = self.unique_parts_aa_var.get()
        bb_str = self.part_in_assembly_bb_var.get()

        try:
            aa_int = self.db_manager._from_base36(aa_str)
            bb_int = self.db_manager._from_base36(bb_str)

            if bb_int > aa_int:
                messagebox.showerror("Ошибка ввода",
                                     f"Номер детали в сборке (BB={bb_str}) не может быть больше количества уникальных деталей (AA={aa_str}).")
                self.part_in_assembly_bb_var.set(self.unique_parts_aa_var.get())
                return False
        except ValueError:
            messagebox.showerror("Ошибка ввода",
                                 "Неверный формат для AA или BB. Используйте буквенно-цифровые символы.")
            self.unique_parts_aa_var.set("01")
            self.part_in_assembly_bb_var.set("01")
            return False
        return True

    def validate_cc_input(self, event=None):
        return self.validate_base36_input(self.part_quantity_cc_var, 2)

    def validate_vv_input(self, event=None):
        return self.validate_base36_input(self.assembly_version_vv_var, 2)

    def validate_w_input(self, event=None):
        # We need to explicitly check if it's empty OR a single valid character.
        value = self.intermediate_version_w_var.get().strip().upper()

        if not value:  # Empty string is allowed
            return True

        if len(value) != 1:  # Must be exactly 1 character if not empty
            messagebox.showerror("Ошибка ввода",
                                 f"Промежуточная версия (W) должна быть 1 символом или пустой. Значение '{value}' имеет длину {len(value)}.")
            self.intermediate_version_w_var.set("")  # Reset to empty
            return False

        valid_chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        forbidden_chars = "IJLO"  # Still exclude these as per _to_base36
        for char in value:
            if char not in valid_chars or char in forbidden_chars:
                messagebox.showerror("Ошибка ввода",
                                     f"Промежуточная версия (W) содержит недопустимые символы. Допустимы 0-9, A-Z (кроме I, J, L, O).")
                self.intermediate_version_w_var.set("")  # Reset to empty
                return False
        return True

    def create_fixture_command(self):
        category = self._get_code_from_display_text(self.category_code_var.get())
        series = self._get_code_from_display_text(self.series_code_var.get())
        item_number = self._get_code_from_display_text(self.item_number_code_var.get())
        operation = self._get_code_from_display_text(self.operation_code_var.get())

        fixture_number_display = self.fixture_number_code_var.get()

        if "Выберите" in category or "Нет данных" in category or \
                "Выберите" in series or "Нет данных" in series or \
                "Выберите" in item_number or "Нет данных" in item_number or \
                "Выберите" in operation or "Нет данных" in operation:
            self.set_status("Пожалуйста, заполните все выпадающие списки.", is_error=True)
            messagebox.showerror("Ошибка", "Пожалуйста, заполните все выпадающие списки.")
            return

        # Если TT Combobox находится в состоянии "Заполните поля выше" или "Ошибка TT"
        if fixture_number_display in ["Заполните поля выше", "Ошибка TT"]:
            self.set_status("Не удалось определить номер оснастки (TT). Проверьте выбранные параметры.", is_error=True)
            messagebox.showerror("Ошибка", "Не удалось определить номер оснастки (TT). Проверьте выбранные параметры.")
            return

        fixture_number = fixture_number_display  # Используем текущее значение из Combobox

        unique_parts = self.unique_parts_aa_var.get().upper().zfill(2)
        part_in_assembly = self.part_in_assembly_bb_var.get().upper().zfill(2)
        part_quantity = self.part_quantity_cc_var.get().upper().zfill(2)
        assembly_version = self.assembly_version_vv_var.get().upper().zfill(2)
        intermediate_version = self.intermediate_version_w_var.get().upper()

        full_id_string = (
            f"{category}.{series}{item_number}.{operation}{fixture_number}."
            f"{unique_parts}{part_in_assembly}{part_quantity}-"
            f"{assembly_version}{intermediate_version}"
        )

        print(f"DEBUG: Сформированная строка ID: '{full_id_string}'")  # <-- ADD THIS DEBUG PRINT

        parsed_id_for_path = self.db_manager.parse_id_string(full_id_string)
        if not parsed_id_for_path:
            self.set_status(f"Ошибка парсинга ID '{full_id_string}'.", is_error=True)
            messagebox.showerror("Ошибка", f"Ошибка парсинга ID '{full_id_string}'.")
            return

        if not self.validate_aa_bb_input() or \
                not self.validate_cc_input() or \
                not self.validate_vv_input() or \
                not self.validate_w_input():
            self.set_status("Ошибка в формате числовых полей. Исправьте.", is_error=True)
            return

        full_id_string = (
            f"{category}.{series}{item_number}.{operation}{fixture_number}."
            f"{unique_parts}{part_in_assembly}{part_quantity}-"
            f"{assembly_version}{intermediate_version}"
        )

        self.set_status(f"Добавление оснастки '{full_id_string}'...")
        # Передаем весь parsed_id_data в db_manager для генерации пути
        parsed_id_for_path = self.db_manager.parse_id_string(full_id_string)
        if not parsed_id_for_path:
            self.set_status(f"Ошибка парсинга ID '{full_id_string}'.", is_error=True)
            messagebox.showerror("Ошибка", f"Ошибка парсинга ID '{full_id_string}'.")
            return

        fixture_id = self.db_manager.add_fixture_id(full_id_string)
        if fixture_id:
            self.set_status(f"Оснастка '{full_id_string}' успешно добавлена. ID в БД: {fixture_id}")
            # Сброс полей к значениям по умолчанию для следующей оснастки
            self.unique_parts_aa_var.set("01")
            self.part_in_assembly_bb_var.set("01")
            self.part_quantity_cc_var.set("01")
            self.assembly_version_vv_var.set("V0")
            self.intermediate_version_w_var.set("")
            # Refresh list for currently selected item
            self.load_fixtures_to_list(self._get_code_from_display_text(self.item_number_code_var.get()))
            self.update_fixture_number_combobox()  # Обновить TT список после добавления
        else:
            self.set_status(f"Не удалось добавить оснастку '{full_id_string}'. Проверьте консоль для деталей.",
                            is_error=True)

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Выберите файлы для копирования",
            filetypes=(("Все файлы", "*.*"), ("PDF файлы", "*.pdf"), ("Изображения", "*.png *.jpg *.jpeg"))
        )
        if files:
            self.selected_files = files
            self.selected_files_path.configure(state="normal")
            self.selected_files_path.delete(0, ctk.END)
            self.selected_files_path.insert(0, "; ".join(files))
            self.selected_files_path.configure(state="readonly")
            self.set_status(f"Выбрано файлов: {len(files)}")
        else:
            self.selected_files = []
            self.selected_files_path.configure(state="normal")
            self.selected_files_path.delete(0, ctk.END)
            self.selected_files_path.configure(state="readonly")
            self.set_status("Выбор файлов отменен.")

    def copy_files(self):
        if not hasattr(self, 'selected_files') or not self.selected_files:
            self.set_status("Сначала выберите файлы для копирования.", is_error=True)
            return

        if self.selected_fixture_id_in_list is None:
            self.set_status("Пожалуйста, выберите оснастку из списка для копирования файлов.", is_error=True)
            return

        selected_fixture_data = self.db_manager.get_fixture_id_by_id(self.selected_fixture_id_in_list)
        if not selected_fixture_data:
            self.set_status("Выбранная оснастка не найдена в базе данных.", is_error=True)
            return

        destination_path = selected_fixture_data.get("BasePath")
        if not destination_path or not os.path.isdir(destination_path):
            self.set_status(f"Не удалось найти или создать папку для оснастки ID {self.selected_fixture_id_in_list}.",
                            is_error=True)
            return

        # Получаем компоненты ID для формирования нового имени файла
        category = selected_fixture_data.get('Category')
        series = selected_fixture_data.get('Series')
        item_number = selected_fixture_data.get('ItemNumber')
        operation = selected_fixture_data.get('Operation')
        fixture_number = selected_fixture_data.get('FixtureNumber')  # Это TT
        unique_parts = selected_fixture_data.get('UniqueParts')  # AA
        part_in_assembly = selected_fixture_data.get('PartInAssembly')  # BB
        part_quantity = selected_fixture_data.get('PartQuantity')  # CC

        # --- ИЗМЕНЕНИЯ ЗДЕСЬ (для исправления "None" в имени файла) ---
        assembly_version = selected_fixture_data.get('AssemblyVersionCode')
        intermediate_version = selected_fixture_data.get('IntermediateVersion')

        # Обрабатываем случай, если значения None или пустые строки, заменяя их на пустые строки или "00"
        # Используем "V0" по умолчанию для VV, так как это начальное значение в GUI
        assembly_version = assembly_version if assembly_version is not None and assembly_version != '' else "V0"
        intermediate_version = intermediate_version if intermediate_version is not None else ""
        # --- КОНЕЦ ИЗМЕНЕНИЙ ---

        # Формируем базовое имя файла по формату KKK.SNN.DTT.AABBCC-VVW
        file_base_name = (
            f"{category}.{series}{item_number}.{operation}{fixture_number}."
            f"{unique_parts}{part_in_assembly}{part_quantity}-"
            f"{assembly_version}{intermediate_version}"
        )

        try:
            copied_count = 0
            for file_path in self.selected_files:
                original_file_name = os.path.basename(file_path)

                # Получаем расширение файла
                _, file_extension = os.path.splitext(original_file_name)

                # Формируем новое имя файла
                new_file_name = f"{file_base_name}{file_extension}"
                dest_file_path = os.path.join(destination_path, new_file_name)

                if os.path.exists(dest_file_path):
                    overwrite = messagebox.askyesno(
                        "Файл существует",
                        f"Файл '{new_file_name}' (оригинал: '{original_file_name}') уже существует в целевой папке.\nПерезаписать его?"
                    )
                    if not overwrite:
                        self.set_status(
                            f"Копирование файла '{original_file_name}' отменено (существует, не перезаписано).")
                        continue

                shutil.copy2(file_path, dest_file_path)
                copied_count += 1
                self.set_status(f"Успешно скопировано и переименовано: '{original_file_name}' -> '{new_file_name}'")

            self.set_status(f"Завершено. Скопировано {copied_count} файлов в '{destination_path}'.")
            self.selected_files = []  # Очищаем список выбранных файлов после копирования
            self.selected_files_path.configure(state="normal")
            self.selected_files_path.delete(0, ctk.END)
            self.selected_files_path.configure(state="readonly")
        except Exception as e:
            self.set_status(f"Ошибка при копировании файлов: {e}", is_error=True)

    def load_fixtures_to_list(self, item_number_code=None):  # <-- NEW: Added item_number_code parameter
        self.fixture_list_text.configure(state="normal")
        self.fixture_list_text.delete("1.0", ctk.END)

        for tag in self.fixture_list_text.tag_names():
            if tag.startswith("fixture_"):
                self.fixture_list_text.tag_config(tag, background="", foreground="")
        self.selected_fixture_id_in_list = None

        # Call db_manager with the item_number_code
        fixtures = self.db_manager.get_fixture_ids_with_descriptions(item_number_code)

        # Update the label to reflect the filtering
        if item_number_code and item_number_code not in ["Нет данных", "Выберите изделие"]:
            self.list_label.configure(text=f"Список оснасток для изделия '{item_number_code}':")
        else:
            self.list_label.configure(
                text="Список существующих оснасток (выберите изделие):")  # Or "Все оснастки" if you want to show all initially

        if not fixtures:
            # Check if an item was actually selected and has no fixtures
            if item_number_code and item_number_code not in ["Нет данных", "Выберите изделие"]:
                self.fixture_list_text.insert(ctk.END,
                                              f"Оснасток для изделия '{item_number_code}' не существует в БД.\n")
            else:
                self.fixture_list_text.insert(ctk.END,
                                              "Оснасток в базе данных пока нет (выберите изделие, чтобы увидеть список).\n")
        else:
            self.fixture_list_text.insert(ctk.END,
                                          "ID | Категория | Серия | Изделие | Операция | Оснастка | AA | BB | CC | Версия | Пром. Версия | Путь\n")
            self.fixture_list_text.insert(ctk.END, "-" * 120 + "\n")
            for fixture in fixtures:
                category_display = fixture['CategoryName'] if fixture['CategoryName'] else fixture['CategoryCode']
                series_display = fixture['SeriesName'] if fixture['SeriesName'] else fixture['SeriesCode']
                item_display = fixture['ItemNumberName'] if fixture['ItemNumberName'] else fixture['ItemNumberCode']
                operation_display = fixture['OperationName'] if fixture['OperationName'] else fixture['OperationCode']

                # Safely get AssemblyVersionCode and IntermediateVersion, handling None
                fixture_assembly_version = fixture.get('AssemblyVersionCode', '')
                if fixture_assembly_version is None:  # Replace None with empty string for display
                    fixture_assembly_version = ''

                fixture_intermediate_version = fixture.get('IntermediateVersion', '')
                if fixture_intermediate_version is None:  # Replace None with empty string for display
                    fixture_intermediate_version = ''

                line = (
                    f"{fixture['id']:<3} | "
                    f"{category_display:<9} | "
                    f"{series_display:<5} | "
                    f"{item_display:<7} | "
                    f"{operation_display:<8} | "
                    f"{fixture['FixtureNumber']!r:<8} | "
                    f"{fixture['UniqueParts']!r:<2} | "
                    f"{fixture['PartInAssembly']!r:<2} | "
                    f"{fixture['PartQuantity']!r:<2} | "
                    f"{fixture_assembly_version!r:<6} | "  # Use the safely handled version
                    f"{fixture_intermediate_version:<13} | "  # Use the safely handled version
                    f"{fixture['BasePath']}\n"
                )
                self.fixture_list_text.insert(ctk.END, line)
                self.fixture_list_text.tag_add(f"fixture_{fixture['id']}",
                                               f"{self.fixture_list_text.index(ctk.END)}-2l",
                                               f"{self.fixture_list_text.index(ctk.END)}-1l")
                self.fixture_list_text.tag_bind(f"fixture_{fixture['id']}", "<Button-1>",
                                                lambda e, fid=fixture['id']: self.on_fixture_select(fid))

        self.fixture_list_text.configure(state="disabled")

    def on_fixture_select(self, fixture_id):
        for tag in self.fixture_list_text.tag_names():
            if tag.startswith("fixture_"):
                self.fixture_list_text.tag_config(tag, background="", foreground="")

        self.fixture_list_text.tag_config(f"fixture_{fixture_id}", background="lightgray", foreground="black")
        self.selected_fixture_id_in_list = fixture_id
        self.set_status(f"Выбрана оснастка ID: {fixture_id}")

    def open_selected_folder(self):
        if self.selected_fixture_id_in_list is None:
            self.set_status("Пожалуйста, выберите оснастку из списка, чтобы открыть её папку.", is_error=True)
            return

        fixture_data = self.db_manager.get_fixture_id_by_id(self.selected_fixture_id_in_list)
        if fixture_data and fixture_data.get("BasePath"):
            folder_path = fixture_data["BasePath"]
            if os.path.exists(folder_path):
                try:
                    os.startfile(folder_path)
                    self.set_status(f"Папка '{folder_path}' открыта.")
                except Exception as e:
                    self.set_status(f"Не удалось открыть папку: {e}", is_error=True)
            else:
                self.set_status(f"Папка не существует: {folder_path}", is_error=True)
        else:
            self.set_status("Не удалось получить путь к папке для выбранной оснастки.", is_error=True)

    def delete_selected_fixture(self):
        if self.selected_fixture_id_in_list is None:
            self.set_status("Пожалуйста, выберите оснастку из списка для удаления.", is_error=True)
            return

        confirm = messagebox.askyesno(
            "Подтверждение удаления",
            f"Вы уверены, что хотите удалить оснастку с ID {self.selected_fixture_id_in_list} И ЕЁ ПАПКУ С ФАЙЛАМИ?\nЭто действие необратимо."
        )

        if confirm:
            self.set_status(f"Удаление оснастки ID: {self.selected_fixture_id_in_list}...")
            if self.db_manager.delete_fixture_id(self.selected_fixture_id_in_list, delete_files=True):
                self.set_status(f"Оснастка ID {self.selected_fixture_id_in_list} успешно удалена.")
                # Refresh list for currently selected item
                self.load_fixtures_to_list(self._get_code_from_display_text(self.item_number_code_var.get()))
                self.update_fixture_number_combobox()  # Обновить TT список после удаления
            else:
                self.set_status(f"Не удалось удалить оснастку ID {self.selected_fixture_id_in_list}.", is_error=True)
        else:
            self.set_status("Удаление отменено.")

    def on_closing(self):
        self.db_manager.close()
        self.destroy()


if __name__ == "__main__":
    excel_file = "classifier_data.xlsx"

    db_name = "my_fixtures_app.db"
    base_db_dir = "fixture_database_root_app"

    db_path = os.path.join(base_db_dir, db_name)

    os.makedirs(base_db_dir, exist_ok=True)

    if not os.path.exists(db_path):
        print(f"База данных '{db_path}' не найдена. Создаем новую и импортируем данные из Excel.")
        temp_db_manager = FixtureDBManager(db_name=db_name, base_db_dir=base_db_dir)
        temp_importer = ExcelClassifierImporter(temp_db_manager)

        try:
            temp_importer.import_from_excel(excel_file)
            print("Импорт завершен.")
        except Exception as e:
            print(f"Ошибка при импорте из Excel: {e}")
            print(f"Убедитесь, что файл '{excel_file}' существует и корректен.")
        finally:
            temp_db_manager.close()
    else:
        print(f"База данных '{db_path}' уже существует. Используем существующие данные.")

    app = FixtureApp()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()