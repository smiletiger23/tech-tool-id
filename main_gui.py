import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
from db_manager import FixtureDBManager
import os
import re
import shutil
import excel_importer


class FixtureApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Управление Оснастками (Tech-Tool-ID)")
        self.geometry("1000x800")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=1)

        # 1. Инициализация переменных (до создания виджетов, чтобы они были доступны)
        self.category_code_var = ctk.StringVar(value="Выберите категорию")
        self.series_code_var = ctk.StringVar(value="Выберите серию")
        self.item_number_code_var = ctk.StringVar(value="Выберите изделие")
        self.operation_code_var = ctk.StringVar(value="Выберите операцию")
        self.fixture_number_code_var = ctk.StringVar(value="Заполните поля выше")

        self.unique_parts_aa_var = ctk.StringVar(value="01")
        self.part_in_assembly_bb_var = ctk.StringVar(value="01")
        self.part_quantity_cc_var = ctk.StringVar(value="01")
        self.assembly_version_vv_var = ctk.StringVar(value="V0")
        self.intermediate_version_w_var = ctk.StringVar(value="")

        self.status_var = ctk.StringVar(value="Статус: Готово")

        self.categories_data = []
        self.series_data = []
        self.items_data = []
        self.operations_data = []

        # Variables for file operations
        self.selected_files_to_copy = []
        self.selected_fixture_id_in_list = None
        self.fixture_id_line_map = {}  # Maps line number in textbox to fixture ID

        # 2. Создание виджетов
        self._create_widgets()

        # 3. Инициализация менеджера базы данных
        self.db_manager = FixtureDBManager(db_name="my_fixtures_app.db", base_db_dir="fixture_database_root_app")

        # 4. Проверяем, пуста ли база данных (или не содержит категорий)
        # Если пуста, выполняем первоначальный импорт из Excel
        if not self.db_manager.get_categories():
            print("DEBUG: База данных пуста или не содержит категорий. Выполняем первоначальный импорт из Excel.")
            excel_file = "classifier_data.xlsx"
            importer = excel_importer.ExcelClassifierImporter(self.db_manager)
            try:
                success, added_counts, updated_counts, skipped_counts, missing_from_excel_data = importer.import_from_excel(
                    excel_file)
                if success:
                    print("DEBUG: Первоначальный импорт завершен успешно.")
                    self.set_status("Первоначальный импорт данных завершен.", is_error=False)
                else:
                    print("DEBUG: Ошибка при первоначальном импорте данных.")
                    self.set_status("Ошибка при первоначальном импорте данных. Проверьте консоль.", is_error=True)
            except Exception as e:
                print(f"DEBUG: Критическая ошибка при первоначальном импорте: {e}")
                self.set_status(f"Критическая ошибка при первоначальном импорте: {e}", is_error=True)

        # 5. Загрузка данных для Combobox'ов и обновление списка оснасток
        self.load_all_combobox_data()
        self._refresh_fixture_list_with_current_selection()

    def _create_widgets(self):
        # Фрейм для элементов управления
        control_frame = ctk.CTkFrame(self)
        control_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        control_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        # Метки и Combobox'ы
        labels = ["Категория (KKK):", "Серия (S):", "Изделие (NN):", "Операция (D):", "Номер оснастки (TT):"]
        combobox_vars = [self.category_code_var, self.series_code_var, self.item_number_code_var,
                         self.operation_code_var, self.fixture_number_code_var]
        combobox_commands = [self.on_category_selected, self.on_series_selected, self.on_item_number_selected,
                             self.on_operation_selected, self.on_fixture_number_selected]
        combobox_options = [[], [], [], [], []]

        for i, label_text in enumerate(labels):
            label = ctk.CTkLabel(control_frame, text=label_text)
            label.grid(row=i, column=0, padx=5, pady=5, sticky="w")

            combobox = ctk.CTkComboBox(control_frame, variable=combobox_vars[i],
                                       values=combobox_options[i],
                                       command=combobox_commands[i],
                                       state="readonly")
            combobox.grid(row=i, column=1, padx=5, pady=5, sticky="ew")

            # Сохраняем ссылки на комбобоксы для дальнейшего доступа
            if label_text == "Категория (KKK):":
                self.category_combobox = combobox
            elif label_text == "Серия (S):":
                self.series_combobox = combobox
            elif label_text == "Изделие (NN):":
                self.item_number_combobox = combobox
            elif label_text == "Операция (D):":
                self.operation_combobox = combobox
            elif label_text == "Номер оснастки (TT):":
                self.fixture_number_combobox = combobox

        # Дополнительные поля
        additional_labels = ["Уникальных деталей (AA):", "Деталь в сборке (BB):", "Количество (CC):",
                             "Версия сборки (VV):", "Пром. версия (W):"]
        additional_vars = [self.unique_parts_aa_var, self.part_in_assembly_bb_var, self.part_quantity_cc_var,
                           self.assembly_version_vv_var, self.intermediate_version_w_var]

        # Валидация для AA, BB, CC, VV, W
        self.unique_parts_aa_var.trace_add("write", lambda name, index, mode: self.validate_aa_bb_input())
        self.part_in_assembly_bb_var.trace_add("write", lambda name, index, mode: self.validate_aa_bb_input())
        self.part_quantity_cc_var.trace_add("write", lambda name, index, mode: self.validate_aa_bb_input())
        self.assembly_version_vv_var.trace_add("write", lambda name, index, mode: self.validate_vv_input())
        self.intermediate_version_w_var.trace_add("write", lambda name, index, mode: self.validate_w_input())

        for i, label_text in enumerate(additional_labels):
            label = ctk.CTkLabel(control_frame, text=label_text)
            label.grid(row=i, column=2, padx=5, pady=5, sticky="w")
            entry = ctk.CTkEntry(control_frame, textvariable=additional_vars[i])
            entry.grid(row=i, column=3, padx=5, pady=5, sticky="ew")

        # Кнопка создания оснастки
        create_button = ctk.CTkButton(control_frame, text="Создать оснастку", command=self.create_fixture_command)
        create_button.grid(row=len(labels), column=0, columnspan=4, padx=5, pady=10, sticky="ew")

        # Кнопка импорта из Excel
        import_excel_button = ctk.CTkButton(control_frame, text="Импортировать из Excel (обновить)",
                                            command=self.import_excel_data_command)
        import_excel_button.grid(row=len(labels) + 1, column=0, columnspan=4, padx=5, pady=10, sticky="ew")

        # Статусная строка
        self.status_label = ctk.CTkLabel(control_frame, textvariable=self.status_var, wraplength=400)
        self.status_label.grid(row=len(labels) + 2, column=0, columnspan=4, padx=5, pady=5, sticky="ew")

        # Фрейм для списка оснасток
        list_frame = ctk.CTkFrame(self)
        list_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(1, weight=1)

        # Инициализация self.list_label
        self.list_label = ctk.CTkLabel(list_frame, text="Список оснасток для изделия '00':")
        self.list_label.grid(row=0, column=0, padx=5, pady=5, sticky="nw")

        self.fixture_list_textbox = ctk.CTkTextbox(list_frame, wrap="none")
        self.fixture_list_textbox.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")
        self.fixture_list_textbox.configure(state="normal")
        self.fixture_list_textbox.bind("<ButtonRelease-1>", self.on_fixture_list_click)

        file_label = ctk.CTkLabel(list_frame, text="Файлы для копирования:")
        file_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")

        select_files_button = ctk.CTkButton(list_frame, text="Выбрать файлы", command=self.select_files_command)
        select_files_button.grid(row=3, column=0, padx=5, pady=5, sticky="ew")

        copy_files_button = ctk.CTkButton(list_frame, text="Копировать файлы в папку выбранной оснастки",
                                          command=self.copy_files_command)
        copy_files_button.grid(row=4, column=0, padx=5, pady=5, sticky="ew")

        # Кнопка удаления оснастки
        delete_fixture_button = ctk.CTkButton(list_frame, text="Удалить выбранную оснастку",
                                              command=self.delete_fixture_command, fg_color="red")
        delete_fixture_button.grid(row=5, column=0, padx=5, pady=5, sticky="ew")

    def set_status(self, message, is_error=False):
        self.status_var.set(f"Статус: {message}")
        if is_error:
            self.status_label.configure(text_color="red")
        else:
            self.status_label.configure(text_color="green")
        self.update_idletasks()

    def load_all_combobox_data(self):
        print("DEBUG: Загрузка всех данных для Combobox'ов...")
        self.categories_data = self.db_manager.get_categories()
        if self.categories_data:
            options = [f"{c['CategoryCode']} ({c['CategoryName']})" for c in self.categories_data]
            self.category_combobox.configure(values=options)
            self.category_code_var.set(options[0])
            self.on_category_selected(self.category_code_var.get())
        else:
            self.category_combobox.configure(values=[])
            self.category_code_var.set("Нет доступных категорий")
            self.category_combobox.configure(state="disabled")
            self.set_status("Ошибка: Нет доступных категорий в базе данных.", is_error=True)

        self.operations_data = self.db_manager.get_operation_descriptions()
        if self.operations_data:
            options = [f"{o['OperationCode']} ({o['OperationName']})" for o in self.operations_data]
            self.operation_combobox.configure(values=options)
            self.operation_code_var.set(options[0])
        else:
            self.operation_combobox.configure(values=[])
            self.operation_code_var.set("Нет доступных операций")
            self.operation_combobox.configure(state="disabled")
            self.set_status("Ошибка: Нет доступных операций в базе данных.", is_error=True)
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
            self._refresh_fixture_list_with_current_selection()
            return

        self.series_combobox.configure(state="readonly")

        self.series_data = self.db_manager.get_series_by_category(category_code)

        print(f"DEBUG: Полученные данные серий для категории '{category_code}': {self.series_data}")

        if not self.series_data:
            self.series_combobox.configure(values=[])
            self.series_code_var.set("Нет данных")
            self.series_combobox.configure(state="disabled")
            print(f"DEBUG: Нет серий для категории '{category_code}'.")
        else:
            options = [f"{s['SeriesCode']} ({s['SeriesName']})" for s in self.series_data]
            self.series_combobox.configure(values=options)
            self.series_code_var.set(options[0])

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
            self._refresh_fixture_list_with_current_selection()
            return

        self.item_number_combobox.configure(state="readonly")

        self.items_data = self.db_manager.get_items_by_category_and_series(category_code, series_code)

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
            self.item_number_code_var.set(options[0])

        self.update_fixture_number_combobox()
        self._refresh_fixture_list_with_current_selection()

    def on_item_number_selected(self, event=None):
        item_number_display_text = self.item_number_code_var.get()
        print(f"DEBUG: on_item_number_selected вызван с '{item_number_display_text}'")
        item_number_code = self._get_code_from_display_text(item_number_display_text)
        print(f"DEBUG: _get_code_from_display_text input: '{item_number_display_text}'")
        print(f"DEBUG: _get_code_from_display_text output (parsed): '{item_number_code}'")
        self.update_fixture_number_combobox()
        self._refresh_fixture_list_with_current_selection()

    def on_operation_selected(self, event=None):
        operation_display_text = self.operation_code_var.get()
        print(f"DEBUG: on_operation_selected вызван с '{operation_display_text}'")
        operation_code = self._get_code_from_display_text(operation_display_text)
        print(f"DEBUG: _get_code_from_display_text input: '{operation_display_text}'")
        print(f"DEBUG: _get_code_from_display_text output (parsed): '{operation_code}'")
        self.update_fixture_number_combobox()
        self._refresh_fixture_list_with_current_selection()

    def on_fixture_number_selected(self, event=None):
        fixture_number_display_text = self.fixture_number_code_var.get()
        print(f"DEBUG: on_fixture_number_selected вызван с '{fixture_number_display_text}'")

        if fixture_number_display_text == "<Создать новый TT>":
            self.generate_next_fixture_number()
            print("DEBUG: Выбран '<Создать новый TT>'.")
        else:
            self.fixture_number_code_var.set(fixture_number_display_text)
            print(f"DEBUG: Выбран существующий TT: '{fixture_number_display_text}'")

    def update_fixture_number_combobox(self):
        print("DEBUG: update_fixture_number_combobox вызван.")
        category_code = self._get_code_from_display_text(self.category_code_var.get())
        series_code = self._get_code_from_display_text(self.series_code_var.get())
        item_number_code = self._get_code_from_display_text(self.item_number_code_var.get())
        operation_code = self._get_code_from_display_text(self.operation_code_var.get())

        if "Выберите" in category_code or "Нет данных" in category_code or \
                "Выберите" in series_code or "Нет данных" in series_code or \
                "Выберите" in item_number_code or "Нет данных" in item_number_code or \
                "Выберите" in operation_code or "Нет данных" in operation_code:
            self.fixture_number_combobox.configure(values=[])
            self.fixture_number_code_var.set("Заполните поля выше")
            self.fixture_number_combobox.configure(state="disabled")
            print("DEBUG: Не все поля выбраны для обновления TT.")
            return

        self.fixture_number_combobox.configure(state="readonly")

        existing_tts = self.db_manager.get_existing_fixture_numbers(
            category_code, series_code, item_number_code, operation_code
        )
        print(
            f"DEBUG: Существующие TT для {category_code}.{series_code}{item_number_code}.{operation_code}: {existing_tts}")

        options = ["<Создать новый TT>"] + existing_tts
        self.fixture_number_combobox.configure(values=options)

        if existing_tts:
            self.fixture_number_code_var.set(existing_tts[0])
        else:
            self.fixture_number_code_var.set("<Создать новый TT>")
            print("DEBUG: Нет существующих TT, выбран '<Создать новый TT>'.")

        self.on_fixture_number_selected(self.fixture_number_code_var.get())
        print("DEBUG: update_fixture_number_combobox завершен.")

    def generate_next_fixture_number(self):
        print("DEBUG: generate_next_fixture_number вызван.")
        category_code = self._get_code_from_display_text(self.category_code_var.get())
        series_code = self._get_code_from_display_text(self.series_code_var.get())
        item_number_code = self._get_code_from_display_text(self.item_number_code_var.get())
        operation_code = self._get_code_from_display_text(self.operation_code_var.get())

        print(f"DEBUG: _get_code_from_display_text input: '{self.category_code_var.get()}'")
        print(f"DEBUG: _get_code_from_display_text output (parsed): '{category_code}'")
        print(f"DEBUG: _get_code_from_display_text input: '{self.series_code_var.get()}'")
        print(f"DEBUG: _get_code_from_display_text output (parsed): '{series_code}'")
        print(f"DEBUG: _get_code_from_display_text input: '{self.item_number_code_var.get()}'")
        print(f"DEBUG: _get_code_from_display_text output (parsed): '{item_number_code}'")
        print(f"DEBUG: _get_code_from_display_text input: '{self.operation_code_var.get()}'")
        print(f"DEBUG: _get_code_from_display_text output (parsed): '{operation_code}'")

        next_tt = self.db_manager.get_next_fixture_number(
            category_code, series_code, item_number_code, operation_code
        )
        self.fixture_number_code_var.set(next_tt)
        print(f"DEBUG: Сгенерирован следующий TT: '{next_tt}'")

    def create_fixture_command(self):
        category = self._get_code_from_display_text(self.category_code_var.get())
        series = self._get_code_from_display_text(self.series_code_var.get())
        item_number = self._get_code_from_display_text(self.item_number_code_var.get())
        operation = self._get_code_from_display_text(self.operation_code_var.get())
        fixture_number = self.fixture_number_code_var.get()

        unique_parts = self.unique_parts_aa_var.get().upper().zfill(2)
        part_in_assembly = self.part_in_assembly_bb_var.get().upper().zfill(2)
        part_quantity = self.part_quantity_cc_var.get().upper().zfill(2)
        assembly_version = self.assembly_version_vv_var.get().upper()
        intermediate_version = self.intermediate_version_w_var.get().upper()

        if "Выберите" in category or "Нет данных" in category or \
                "Выберите" in series or "Нет данных" in series or \
                "Выберите" in item_number or "Нет данных" in item_number or \
                "Выберите" in operation or "Нет данных" in operation or \
                "Заполните" in fixture_number:
            self.set_status("Пожалуйста, заполните все поля классификатора.", is_error=True)
            messagebox.showerror("Ошибка", "Пожалуйста, заполните все поля классификатора.")
            return

        if not self.validate_aa_bb_input() or \
                not self.validate_vv_input() or \
                not self.validate_w_input():
            self.set_status("Ошибка валидации дополнительных полей.", is_error=True)
            messagebox.showerror("Ошибка", "Ошибка валидации дополнительных польных полей. Проверьте формат.")
            return

        full_id_string = (
            f"{category}.{series}{item_number}.{operation}{fixture_number}."
            f"{unique_parts}{part_in_assembly}{part_quantity}-"
            f"{assembly_version}{intermediate_version}"
        )
        print(f"DEBUG: Сформированная строка ID: '{full_id_string}'")

        parsed_id_for_path = self.db_manager.parse_id_string(full_id_string)
        if not parsed_id_for_path:
            self.set_status(f"Ошибка парсинга ID '{full_id_string}'.", is_error=True)
            messagebox.showerror("Ошибка", f"Ошибка парсинга ID '{full_id_string}'.")
            return

        base_folder_path = os.path.join(
            self.db_manager.base_db_dir,
            parsed_id_for_path['Category'],
            f"{parsed_id_for_path['Category']}.{parsed_id_for_path['Series']}{parsed_id_for_path['ItemNumber']}",
            f"{parsed_id_for_path['Category']}.{parsed_id_for_path['Series']}{parsed_id_for_path['ItemNumber']}.{parsed_id_for_path['Operation']}{parsed_id_for_path['FixtureNumber']}",
            full_id_string
        )

        try:
            os.makedirs(base_folder_path, exist_ok=True)
            print(f"Базовая папка версии оснастки создана: {base_folder_path}")
        except OSError as e:
            self.set_status(f"Ошибка создания папки '{base_folder_path}': {e}", is_error=True)
            messagebox.showerror("Ошибка", f"Ошибка создания папки: {e}")
            return

        fixture_id = self.db_manager.add_fixture_id(full_id_string)
        if fixture_id:
            self.set_status(f"Оснастка {full_id_string} успешно добавлена. ID в БД: {fixture_id}")
            self._refresh_fixture_list_with_current_selection()
        else:
            self.set_status(f"Не удалось добавить оснастку {full_id_string}.", is_error=True)

    def select_files_command(self):
        """Allows user to select multiple files for copying."""
        file_paths = filedialog.askopenfilenames(
            title="Выберите файлы для копирования",
            filetypes=[("Все файлы", "*.*")]
        )
        if file_paths:
            self.selected_files_to_copy = list(file_paths)
            self.set_status(f"Выбрано файлов для копирования: {len(self.selected_files_to_copy)}")
        else:
            self.selected_files_to_copy = []
            self.set_status("Выбор файлов отменен.", is_error=False)

    def copy_files_command(self):
        """Copies selected files to the folder of the selected fixture."""
        if not self.selected_files_to_copy:
            self.set_status("Ошибка: Не выбраны файлы для копирования.", is_error=True)
            messagebox.showerror("Ошибка копирования", "Пожалуйста, выберите файлы для копирования.")
            return

        if self.selected_fixture_id_in_list is None:
            self.set_status("Ошибка: Не выбрана оснастка для копирования.", is_error=True)
            messagebox.showerror("Ошибка копирования", "Пожалуйста, выберите оснастку из списка, кликнув на нее.")
            return

        fixture_data = self.db_manager.get_fixture_id_by_id(self.selected_fixture_id_in_list)
        if not fixture_data or not fixture_data.get('BasePath'):
            self.set_status(
                f"Ошибка: Не удалось получить путь для выбранной оснастки ID {self.selected_fixture_id_in_list}.",
                is_error=True)
            messagebox.showerror("Ошибка копирования", "Не удалось найти путь к папке выбранной оснастки.")
            return

        destination_folder = fixture_data['BasePath']

        copied_count = 0
        failed_copies = []

        for file_path in self.selected_files_to_copy:
            try:
                shutil.copy2(file_path, destination_folder)
                copied_count += 1
            except Exception as e:
                failed_copies.append(f"{os.path.basename(file_path)}: {e}")

        if copied_count == len(self.selected_files_to_copy):
            self.set_status(f"Успешно скопировано {copied_count} файл(ов) в '{destination_folder}'.", is_error=False)
            messagebox.showinfo("Копирование завершено", f"Все {copied_count} файл(ов) успешно скопированы.")
        else:
            self.set_status(
                f"Скопировано {copied_count} из {len(self.selected_files_to_copy)} файл(ов). Ошибки: {len(failed_copies)}.",
                is_error=True)
            messagebox.showwarning("Копирование с ошибками",
                                   f"Некоторые файлы не были скопированы.\nОшибки:\n" + "\n".join(failed_copies))

    def on_fixture_list_click(self, event):
        """Handles click events on the fixture list textbox to select a fixture."""
        index = self.fixture_list_textbox.index(f"@{event.x},{event.y}")
        line = int(index.split('.')[0])

        self.fixture_list_textbox.tag_remove("highlight", "1.0", "end")

        if line > 2 and line in self.fixture_id_line_map:
            self.selected_fixture_id_in_list = self.fixture_id_line_map[line]
            self.fixture_list_textbox.tag_add("highlight", f"{line}.0", f"{line}.end")
            self.fixture_list_textbox.tag_config("highlight", background="lightblue", foreground="black")

            fixture_data = self.db_manager.get_fixture_id_by_id(self.selected_fixture_id_in_list)
            if fixture_data:
                self.set_status(
                    f"Выбрана оснастка для копирования: ID {self.selected_fixture_id_in_list} ({fixture_data.get('FullIDString', 'N/A')}).")
            else:
                self.set_status(f"Выбрана оснастка с ID {self.selected_fixture_id_in_list}.", is_error=False)
        else:
            self.selected_fixture_id_in_list = None
            self.set_status("Выбор оснастки сброшен. Пожалуйста, кликните на строку оснастки.", is_error=False)

    def _refresh_fixture_list_with_current_selection(self):
        """Helper to call load_fixtures_to_list with current combobox selections."""
        category_code = self._get_code_from_display_text(self.category_code_var.get())
        series_code = self._get_code_from_display_text(self.series_code_var.get())
        item_number_code = self._get_code_from_display_text(self.item_number_code_var.get())
        operation_code = self._get_code_from_display_text(self.operation_code_var.get())

        if "Выберите" in category_code or "Нет данных" in category_code:
            category_code = None
        if "Выберите" in series_code or "Нет данных" in series_code:
            series_code = None
        if "Выберите" in item_number_code or "Нет данных" in item_number_code:
            item_number_code = None
        if "Выберите" in operation_code or "Нет данных" in operation_code:
            operation_code = None

        self.load_fixtures_to_list(
            category_code=category_code,
            series_code=series_code,
            item_number_code=item_number_code,
            operation_code=operation_code
        )

    def load_fixtures_to_list(self, category_code=None, series_code=None, item_number_code=None, operation_code=None):
        """Загружает и отображает список оснасток в текстовом поле."""
        self.fixture_list_textbox.configure(state="normal")
        self.fixture_list_textbox.delete("1.0", "end")
        self.fixture_id_line_map = {}

        fixtures = self.db_manager.get_fixture_ids_with_descriptions(
            category_code=category_code,
            series_code=series_code,
            item_number_code=item_number_code,
            operation_code=operation_code
        )

        filter_display_parts = []
        if category_code and category_code not in ["Нет данных", "Выберите категорию"]:
            filter_display_parts.append(f"категории '{self._get_name_from_code(category_code, 'category')}'")
        if series_code and series_code not in ["Нет данных", "Выберите серию"]:
            filter_display_parts.append(f"серии '{self._get_name_from_code(series_code, 'series')}'")
        if item_number_code and item_number_code not in ["Нет данных", "Выберите изделие"]:
            filter_display_parts.append(f"изделия '{self._get_name_from_code(item_number_code, 'item_number')}'")
        if operation_code and operation_code not in ["Нет данных", "Выберите операцию"]:
            filter_display_parts.append(f"операции '{self._get_name_from_code(operation_code, 'operation')}'")

        if filter_display_parts:
            self.list_label.configure(text=f"Список оснасток для {', '.join(filter_display_parts)}:")
        else:
            self.list_label.configure(text="Список существующих оснасток (все):")

        if not fixtures:
            if filter_display_parts:
                self.fixture_list_textbox.insert("end",
                                                 f"Оснасток для {', '.join(filter_display_parts)} не существует в БД.\n")
            else:
                self.fixture_list_textbox.insert("end", "Оснасток в базе данных пока нет.\n")
            self.fixture_list_textbox.configure(state="disabled")
            return

        header = (
            f"{'ID':<4} | {'Категория':<10} | {'Серия':<8} | {'Изделие':<10} | {'Операция':<10} | "
            f"{'Оснастка':<10} | {'AA':<5} | {'BB':<5} | {'CC':<5} | {'Версия':<8} | "
            f"{'Путь':<40}"
        )
        self.fixture_list_textbox.insert("end", header + "\n")
        self.fixture_list_textbox.insert("end", "-" * len(header) + "\n")

        current_line = 3

        for fixture in fixtures:
            category_display = fixture.get('CategoryName', fixture.get('Category', ''))
            series_display = fixture.get('SeriesName', fixture.get('Series', ''))
            item_display = fixture.get('ItemNumberName', fixture.get('ItemNumber', ''))
            operation_display = fixture.get('OperationName', fixture.get('Operation', ''))

            fixture_number = fixture.get('FixtureNumber', '')
            unique_parts = fixture.get('UniqueParts', '')
            part_in_assembly = fixture.get('PartInAssembly', '')
            part_quantity = fixture.get('PartQuantity', '')
            assembly_version_code = fixture.get('AssemblyVersionCode', '')
            intermediate_version = fixture.get('IntermediateVersion', '')
            full_id_string_display = fixture.get('FullIDString', '')

            combined_version_vvw = f"{assembly_version_code}{intermediate_version}"

            line_text = (
                f"{fixture['id']:<4} | "
                f"{category_display:<10} | "
                f"{series_display:<8} | "
                f"{item_display:<10} | "
                f"{operation_display:<10} | "
                f"{fixture_number:<10} | "
                f"{unique_parts:<5} | "
                f"{part_in_assembly:<5} | "
                f"{part_quantity:<5} | "
                f"{combined_version_vvw:<8} | "
                f"{full_id_string_display:<40}"
            )
            self.fixture_list_textbox.insert("end", line_text + "\n")
            self.fixture_id_line_map[current_line] = fixture['id']
            current_line += 1

        self.fixture_list_textbox.configure(state="disabled")

    def _get_code_from_display_text(self, display_text):
        """Извлекает код из строки вида 'CODE (Description)' или возвращает исходную строку."""
        match = re.match(r"([A-Z0-9]+)\s*\(.*\)", display_text)
        if match:
            return match.group(1)
        return display_text

    def _get_name_from_code(self, code, type_of_code):
        """Возвращает имя по коду из соответствующего словаря данных."""
        if type_of_code == 'category':
            for item in self.categories_data:
                if item['CategoryCode'] == code:
                    return item['CategoryName']
        elif type_of_code == 'series':
            for item in self.series_data:
                if item['SeriesCode'] == code:
                    return item['SeriesName']
        elif type_of_code == 'item_number':
            for item in self.items_data:
                if item['ItemNumberCode'] == code:
                    return item['ItemNumberName']
        elif type_of_code == 'operation':
            for item in self.operations_data:
                if item['OperationCode'] == code:
                    return item['OperationName']
        return code

    def validate_aa_bb_input(self, event=None):
        aa_str = self.unique_parts_aa_var.get().upper()
        bb_str = self.part_in_assembly_bb_var.get().upper()
        cc_str = self.part_quantity_cc_var.get().upper()

        is_valid = True

        if aa_str and (len(aa_str) != 2 or not all(c in '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ' for c in aa_str)):
            self.set_status("AA должно быть 2 символа (0-9, A-Z).", is_error=True)
            is_valid = False

        if bb_str and (len(bb_str) != 2 or not all(c in '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ' for c in bb_str)):
            self.set_status("BB должно быть 2 символа (0-9, A-Z).", is_error=True)
            is_valid = False

        if cc_str and (len(cc_str) != 2 or not all(c in '01234456789ABCDEFGHIJKLMNOPQRSTUVWXYZ' for c in cc_str)):
            self.set_status("CC должно быть 2 символа (0-9, A-Z).", is_error=True)
            is_valid = False

        if is_valid:
            self.set_status("Поля AA, BB, CC валидны.", is_error=False)
        return is_valid

    def validate_vv_input(self, event=None):
        vv_str = self.assembly_version_vv_var.get().upper()
        if vv_str and (len(vv_str) != 2 or not all(c in '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ' for c in vv_str)):
            self.set_status("Версия сборки (VV) должна быть 2 символа (0-9, A-Z).", is_error=True)
            return False
        self.set_status("Поле VV валидно.", is_error=False)
        return True

    def validate_w_input(self, event=None):
        value = self.intermediate_version_w_var.get().strip().upper()

        if not value:
            self.set_status("Поле W валидно (пусто).", is_error=False)
            return True

        if len(value) != 1:
            self.set_status(
                f"Промежуточная версия (W) должна быть 1 символом или пустой. Значение '{value}' имеет длину {len(value)}.",
                is_error=True)
            return False

        valid_chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        forbidden_chars = "IJLO"
        for char in value:
            if char not in valid_chars or char in forbidden_chars:
                self.set_status(
                    f"Промежуточная версия (W) содержит недопустимые символы. Допустимы 0-9, A-Z (кроме I, J, L, O).",
                    is_error=True)
                return False

        self.set_status("Поле W валидно.", is_error=False)
        return True

    def import_excel_data_command(self):
        """
        Обрабатывает интеллектуальный импорт данных из Excel в БД по нажатию кнопки.
        Не удаляет существующую БД, а добавляет новые данные и сообщает об отсутствующих.
        """
        self.set_status("Начинается интеллектуальный импорт данных из Excel...", is_error=False)

        excel_file_path = "classifier_data.xlsx"
        importer = excel_importer.ExcelClassifierImporter(self.db_manager)

        try:
            success, added_counts, updated_counts, skipped_counts, missing_from_excel_data = importer.import_from_excel(
                excel_file_path)

            if success:
                report_message = "Импорт данных из Excel завершен успешно.\n\n"

                # Report added items
                added_summary = []
                for sheet_name, count in added_counts.items():
                    if count > 0:
                        added_summary.append(f"  {sheet_name}: {count} новых записей")
                if added_summary:
                    report_message += "Добавлено:\n" + "\n".join(added_summary) + "\n\n"
                else:
                    report_message += "Новых записей не добавлено.\n\n"

                # Report updated items
                updated_summary = []
                for sheet_name, count in updated_counts.items():
                    if count > 0:
                        updated_summary.append(f"  {sheet_name}: {count} обновленных записей")
                if updated_summary:
                    report_message += "Обновлено:\n" + "\n".join(updated_summary) + "\n\n"
                else:
                    report_message += "Записей не обновлено.\n\n"

                # Report missing items
                missing_summary = []
                for sheet_name, items in missing_from_excel_data.items():
                    if items:
                        missing_summary.append(f"  {sheet_name}: {len(items)} записей отсутствуют в Excel:")
                        for item in items:
                            # Format missing item for display (e.g., CategoryCode (CategoryName))
                            if sheet_name == "Категории":
                                missing_summary.append(
                                    f"    {item.get('CategoryCode', 'N/A')} ({item.get('CategoryName', 'N/A')})")
                            elif sheet_name == "Серии":
                                missing_summary.append(
                                    f"    {item.get('CategoryCode', 'N/A')}-{item.get('SeriesCode', 'N/A')} ({item.get('SeriesName', 'N/A')})")
                            elif sheet_name == "Изделия":
                                missing_summary.append(
                                    f"    {item.get('CategoryCode', 'N/A')}-{item.get('SeriesCode', 'N/A')}-{item.get('ItemNumberCode', 'N/A')} ({item.get('ItemNumberName', 'N/A')})")
                            elif sheet_name == "Операции":
                                missing_summary.append(
                                    f"    {item.get('OperationCode', 'N/A')} ({item.get('OperationName', 'N/A')})")
                            else:
                                missing_summary.append(f"    {item}")  # Fallback for unknown sheet type
                if missing_summary:
                    report_message += "Записи в базе данных, отсутствующие в Excel:\n" + "\n".join(missing_summary)
                else:
                    report_message += "Все записи из базы данных найдены в Excel."

                self.set_status("Импорт завершен. Подробности в отчете.", is_error=False)
                messagebox.showinfo("Отчет по импорту Excel", report_message)

                self.load_all_combobox_data()
                self._refresh_fixture_list_with_current_selection()
            else:
                self.set_status("Ошибка при импорте данных из Excel. Проверьте консоль.", is_error=True)
                messagebox.showerror("Ошибка импорта",
                                     "Произошла ошибка при импорте данных из Excel. Проверьте консоль для деталей.")
        except Exception as e:
            self.set_status(f"Критическая ошибка при импорте: {e}", is_error=True)
            messagebox.showerror("Критическая ошибка импорта", f"Произошла критическая ошибка при импорте: {e}")

    def delete_fixture_command(self):
        """Deletes the selected fixture from the database and its associated folder."""
        if self.selected_fixture_id_in_list is None:
            self.set_status("Ошибка: Не выбрана оснастка для удаления.", is_error=True)
            messagebox.showerror("Ошибка удаления", "Пожалуйста, выберите оснастку из списка для удаления.")
            return

        fixture_data = self.db_manager.get_fixture_id_by_id(self.selected_fixture_id_in_list)
        if not fixture_data:
            self.set_status(f"Ошибка: Оснастка с ID {self.selected_fixture_id_in_list} не найдена в БД.", is_error=True)
            messagebox.showerror("Ошибка удаления", "Выбранная оснастка не найдена в базе данных.")
            return

        full_id_string = fixture_data.get('FullIDString', 'N/A')
        base_path = fixture_data.get('BasePath')

        confirm = messagebox.askyesno("Подтверждение удаления",
                                      f"Вы уверены, что хотите удалить оснастку:\nID: {full_id_string}\nПуть: {base_path}\n\nЭто действие необратимо и удалит папку на диске!")

        if confirm:
            if self.db_manager.delete_fixture_id(self.selected_fixture_id_in_list, delete_files=True):
                self.set_status(f"Оснастка ID {self.selected_fixture_id_in_list} ({full_id_string}) успешно удалена.",
                                is_error=False)
                self.selected_fixture_id_in_list = None
                self.fixture_list_textbox.tag_remove("highlight", "1.0", "end")
                self._refresh_fixture_list_with_current_selection()
                self.update_fixture_number_combobox()
            else:
                self.set_status(
                    f"Не удалось удалить оснастку ID {self.selected_fixture_id_in_list}. Проверьте консоль.",
                    is_error=True)
        else:
            self.set_status("Удаление отменено.")

    def on_closing(self):
        self.db_manager.close()
        self.destroy()


if __name__ == "__main__":
    app = FixtureApp()
    app.mainloop()
