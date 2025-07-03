import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk  # Import ttk for Treeview
from db_manager import FixtureDBManager
import os
import re
import shutil
import excel_importer
import subprocess


class FixtureApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Управление Оснастками (Tech-Tool-ID)")
        self.geometry("1200x800")  # Increased width for better Treeview display
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
        self.assembly_version_vv_var = ctk.StringVar(value="01")  # Default VV is 01
        self.intermediate_version_w_var = ctk.StringVar(value="")

        self.status_var = ctk.StringVar(value="Статус: Готово")
        self.hide_non_actual_versions_var = ctk.BooleanVar(value=True)  # Checkbox variable, default True
        self.hide_assembled_parts_var = ctk.BooleanVar(value=False)  # Checkbox variable, default False

        self.categories_data = []
        self.series_data = []
        self.items_data = []
        self.operations_data = []

        # Variables for file operations
        self.selected_files_to_copy = []
        self.selected_fixture_id_in_list = None
        # self.fixture_id_line_map = {} # No longer needed with Treeview

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

        # Checkbox for hiding non-actual versions
        self.hide_non_actual_checkbox = ctk.CTkCheckBox(control_frame,
                                                        text="Скрыть неактуальные версии",
                                                        variable=self.hide_non_actual_versions_var,
                                                        command=self.on_filter_checkbox_toggled)
        self.hide_non_actual_checkbox.grid(row=len(labels) + 3, column=0, columnspan=4, padx=5, pady=5, sticky="w")

        # Checkbox for hiding assembled parts
        self.hide_assembled_parts_checkbox = ctk.CTkCheckBox(control_frame,
                                                             text="Не отображать сборные части (BB!=01)",
                                                             variable=self.hide_assembled_parts_var,
                                                             command=self.on_filter_checkbox_toggled)
        self.hide_assembled_parts_checkbox.grid(row=len(labels) + 4, column=0, columnspan=4, padx=5, pady=5, sticky="w")

        # Фрейм для списка оснасток
        list_frame = ctk.CTkFrame(self)
        list_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(1, weight=1)

        # Инициализация self.list_label
        self.list_label = ctk.CTkLabel(list_frame, text="Список оснасток для изделия '00':")
        self.list_label.grid(row=0, column=0, padx=5, pady=5, sticky="nw")

        # --- Treeview для списка оснасток ---
        columns = ('category', 'series', 'item_number', 'operation', 'fixture_number',
                   'aa', 'bb', 'cc', 'version', 'full_id')

        self.fixture_list_tree = ttk.Treeview(list_frame, columns=columns, show='headings')

        # Define headings
        self.fixture_list_tree.heading('category', text='Категория', anchor=tk.W)
        self.fixture_list_tree.heading('series', text='Серия', anchor=tk.W)
        self.fixture_list_tree.heading('item_number', text='Изделие', anchor=tk.W)
        self.fixture_list_tree.heading('operation', text='Операция', anchor=tk.W)
        self.fixture_list_tree.heading('fixture_number', text='Оснастка', anchor=tk.W)
        self.fixture_list_tree.heading('aa', text='AA', anchor=tk.W)
        self.fixture_list_tree.heading('bb', text='BB', anchor=tk.W)
        self.fixture_list_tree.heading('cc', text='CC', anchor=tk.W)
        self.fixture_list_tree.heading('version', text='Версия', anchor=tk.W)
        self.fixture_list_tree.heading('full_id', text='Полный ID (Путь)', anchor=tk.W)

        # Define column widths (initial values, Treeview will adjust)
        self.fixture_list_tree.column('category', width=100, minwidth=50)
        self.fixture_list_tree.column('series', width=80, minwidth=50)
        self.fixture_list_tree.column('item_number', width=100, minwidth=50)
        self.fixture_list_tree.column('operation', width=100, minwidth=50)
        self.fixture_list_tree.column('fixture_number', width=100, minwidth=50)
        self.fixture_list_tree.column('aa', width=50, minwidth=30)
        self.fixture_list_tree.column('bb', width=50, minwidth=30)
        self.fixture_list_tree.column('cc', width=50, minwidth=30)
        self.fixture_list_tree.column('version', width=80, minwidth=50)
        self.fixture_list_tree.column('full_id', width=300, minwidth=150)  # Increased width for ID

        self.fixture_list_tree.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")

        # Scrollbars for Treeview
        tree_scrollbar_y = ctk.CTkScrollbar(list_frame, command=self.fixture_list_tree.yview)
        tree_scrollbar_y.grid(row=1, column=1, sticky="ns")
        self.fixture_list_tree.configure(yscrollcommand=tree_scrollbar_y.set)

        tree_scrollbar_x = ctk.CTkScrollbar(list_frame, orientation="horizontal", command=self.fixture_list_tree.xview)
        tree_scrollbar_x.grid(row=2, column=0, sticky="ew")  # Placed below the treeview
        self.fixture_list_tree.configure(xscrollcommand=tree_scrollbar_x.set)

        self.fixture_list_tree.bind("<<TreeviewSelect>>", self.on_fixture_list_select)

        # --- Styling for ttk.Treeview to match CustomTkinter theme ---
        style = ttk.Style()

        # Get current CustomTkinter appearance mode
        appearance_mode = ctk.get_appearance_mode()
        if appearance_mode == "Dark":
            bg_color = "#2b2b2b"  # CustomTkinter dark mode background
            fg_color = "#ffffff"  # CustomTkinter text color
            selected_bg = "#3a7ebf"  # CustomTkinter blue
            selected_fg = "#ffffff"
            header_bg = "#343638"  # Slightly darker for header
            separator_color = "#565b5e"  # Border color
        else:  # Light mode
            bg_color = "#ededed"
            fg_color = "#000000"
            selected_bg = "#3a7ebf"
            selected_fg = "#ffffff"
            header_bg = "#d9d9d9"
            separator_color = "#a6a6a6"

        style.theme_use("default")  # Use default theme as a base for styling
        style.configure("Treeview",
                        background=bg_color,
                        foreground=fg_color,
                        fieldbackground=bg_color,
                        bordercolor=separator_color,
                        lightcolor=separator_color,
                        darkcolor=separator_color,
                        borderwidth=1,
                        rowheight=25,
                        font=("", 12))  # Increased font size for rows
        style.map('Treeview',
                  background=[('selected', selected_bg)],
                  foreground=[('selected', selected_fg)])

        style.configure("Treeview.Heading",
                        background=header_bg,
                        foreground=fg_color,
                        font=("", 12, "bold"))  # Increased font size for headings
        style.map("Treeview.Heading",
                  background=[('active', header_bg)])  # Keep header background consistent on hover

        # --- End Treeview ---

        file_label = ctk.CTkLabel(list_frame, text="Файлы для копирования:")
        file_label.grid(row=3, column=0, padx=5, pady=5, sticky="w")  # Adjusted row

        select_files_button = ctk.CTkButton(list_frame, text="Выбрать файлы", command=self.select_files_command)
        select_files_button.grid(row=4, column=0, padx=5, pady=5, sticky="ew")  # Adjusted row

        copy_files_button = ctk.CTkButton(list_frame, text="Копировать файлы в папку выбранной оснастки",
                                          command=self.copy_files_command)
        copy_files_button.grid(row=5, column=0, padx=5, pady=5, sticky="ew")  # Adjusted row

        # Кнопка открытия директории оснастки
        open_folder_button = ctk.CTkButton(list_frame, text="Открыть директорию выбранной оснастки",
                                           command=self.open_fixture_folder_command)
        open_folder_button.grid(row=6, column=0, padx=5, pady=5, sticky="ew")  # Adjusted row

        # Кнопка удаления оснастки
        delete_fixture_button = ctk.CTkButton(list_frame, text="Удалить выбранную оснастку",
                                              command=self.delete_fixture_command, fg_color="red")
        delete_fixture_button.grid(row=7, column=0, padx=5, pady=5, sticky="ew")  # Adjusted row

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
            options = ["Все категории"] + [f"{c['CategoryCode']} ({c['CategoryName']})" for c in self.categories_data]
            self.category_combobox.configure(values=options)
            self.category_code_var.set("Все категории")
            self.category_combobox.set("Все категории")
        else:
            self.category_combobox.configure(values=[])
            self.category_code_var.set("Нет доступных категорий")
            self.category_combobox.configure(state="disabled")
            self.set_status("Ошибка: Нет доступных категорий в базе данных.", is_error=True)

        self.operations_data = self.db_manager.get_operation_descriptions()
        if self.operations_data:
            options = ["Все операции"] + [f"{o['OperationCode']} ({o['OperationName']})" for o in self.operations_data]
            self.operation_combobox.configure(values=options)
            self.operation_code_var.set("Все операции")
            self.operation_combobox.set("Все операции")
        else:
            self.operation_combobox.configure(values=[])
            self.operation_code_var.set("Нет доступных операций")
            self.operation_combobox.configure(state="disabled")
            self.set_status("Ошибка: Нет доступных операций в базе данных.", is_error=True)
        print("DEBUG: Загрузка всех данных для Combobox'ов завершена.")

        # Initialize dependent comboboxes to placeholder states
        self.series_combobox.configure(values=[])
        self.series_code_var.set("Выберите серию")
        self.series_combobox.set("Выберите серию")
        self.series_combobox.configure(state="disabled")

        self.item_number_combobox.configure(values=[])
        self.item_number_code_var.set("Выберите изделие")
        self.item_number_combobox.set("Выберите изделие")
        self.item_number_combobox.configure(state="disabled")

        self.fixture_number_combobox.configure(values=[])
        self.fixture_number_code_var.set("Заполните поля выше")
        self.fixture_number_combobox.set("Заполните поля выше")
        self.fixture_number_combobox.configure(state="disabled")

    def on_category_selected(self, event=None):
        category_display_text = self.category_code_var.get()
        print(f"DEBUG: on_category_selected вызван с '{category_display_text}'")

        category_code = self._get_code_from_display_text(category_display_text)
        print(f"DEBUG: _get_code_from_display_text input: '{category_display_text}'")
        print(f"DEBUG: _get_code_from_display_text output (parsed): '{category_code}'")

        if category_code is None:  # Simplified check for None
            self.series_combobox.configure(values=[])
            self.series_code_var.set("Выберите серию")
            self.series_combobox.set("Выберите серию")
            self.series_combobox.configure(state="disabled")

            self.item_number_combobox.configure(values=[])
            self.item_number_code_var.set("Выберите изделие")
            self.item_number_combobox.set("Выберите изделие")
            self.item_number_combobox.configure(state="disabled")

            self.fixture_number_combobox.configure(values=[])
            self.fixture_number_code_var.set("Заполните поля выше")
            self.fixture_number_combobox.set("Заполните поля выше")
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
            self.series_combobox.set("Нет данных")
            self.series_combobox.configure(state="disabled")
            print(f"DEBUG: Нет серий для категории '{category_code}'.")
        else:
            options = ["Все серии"] + [f"{s['SeriesCode']} ({s['SeriesName']})" for s in self.series_data]
            self.series_combobox.configure(values=options)
            self.series_code_var.set("Все серии")
            self.series_combobox.set("Все серии")

        self.on_series_selected(self.series_code_var.get())  # Call with placeholder to reset dependent fields

    def on_series_selected(self, event=None):
        category_code = self._get_code_from_display_text(self.category_code_var.get())
        series_display_text = self.series_code_var.get()
        print(f"DEBUG: on_series_selected вызван с '{series_display_text}'")

        series_code = self._get_code_from_display_text(series_display_text)
        print(f"DEBUG: _get_code_from_display_text input: '{series_display_text}'")
        print(f"DEBUG: _get_code_from_display_text output (parsed): '{series_code}'")

        # Simplified check for None
        if category_code is None or series_code is None:
            self.item_number_combobox.configure(values=[])
            self.item_number_code_var.set("Выберите изделие")
            self.item_number_combobox.set("Выберите изделие")
            self.item_number_combobox.configure(state="disabled")

            self.fixture_number_combobox.configure(values=[])
            self.fixture_number_code_var.set("Заполните поля выше")
            self.fixture_number_combobox.set("Заполните поля выше")
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
            self.item_number_combobox.set("Нет данных")
            self.item_number_combobox.configure(state="disabled")
            print(f"DEBUG: Нет изделий для категории '{category_code}', серии '{series_code}'.")
        else:
            options = ["Все изделия"] + [f"{i['ItemNumberCode']} ({i['ItemNumberName']})" for i in self.items_data]
            self.item_number_combobox.configure(values=options)
            self.item_number_code_var.set("Все изделия")
            self.item_number_combobox.set("Все изделия")

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

        # Simplified check for None
        if category_code is None or \
                series_code is None or \
                item_number_code is None or \
                operation_code is None:
            self.fixture_number_combobox.configure(values=[])
            self.fixture_number_code_var.set("Заполните поля выше")
            self.fixture_number_combobox.set("Заполните поля выше")
            self.fixture_number_combobox.configure(state="disabled")
            print("DEBUG: Не все поля выбраны для обновления TT.")
            return

        self.fixture_number_combobox.configure(state="readonly")

        existing_tts = self.db_manager.get_existing_fixture_numbers(
            category_code, series_code, item_number_code, operation_code
        )
        # Ensure uniqueness for existing_tts
        existing_tts = sorted(list(set(existing_tts)))

        print(
            f"DEBUG: Существующие TT для {category_code}.{series_code}{item_number_code}.{operation_code}: {existing_tts}")

        options = ["<Создать новый TT>"] + existing_tts
        self.fixture_number_combobox.configure(values=options)

        # If there's only one existing TT (and it's not the placeholder), select it
        if len(existing_tts) == 1 and self.fixture_number_code_var.get() not in [
            "Выберите номер оснастки или создайте новый", "<Создать новый TT>"]:
            self.fixture_number_code_var.set(existing_tts[0])
            self.fixture_number_combobox.set(existing_tts[0])
        elif "<Создать новый TT>" not in self.fixture_number_code_var.get() and \
                self.fixture_number_code_var.get() not in existing_tts:
            self.fixture_number_code_var.set("Выберите номер оснастки или создайте новый")
            self.fixture_number_combobox.set("Выберите номер оснастки или создайте новый")

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

        # Check for placeholder values from "Все..." options
        if category is None or series is None or item_number is None or operation is None or \
                "Заполните" in fixture_number or "Выберите номер оснастки" in fixture_number:
            self.set_status("Пожалуйста, заполните все поля классификатора.", is_error=True)
            messagebox.showerror("Ошибка", "Пожалуйста, заполните все поля классификатора.")
            return

        if not self.validate_aa_bb_input() or \
                not self.validate_vv_input() or \
                not self.validate_w_input():
            self.set_status("Ошибка валидации дополнительных полей.", is_error=True)
            messagebox.showerror("Ошибка", "Ошибка валидации дополнительных польных полей. Проверьте формат.")
            return

        # Validate BB <= AA
        try:
            aa_int = int(unique_parts)
            bb_int = int(part_in_assembly)
            if bb_int > aa_int:
                self.set_status("Ошибка: Деталь в сборке (BB) не может быть больше Уникальных деталей (AA).",
                                is_error=True)
                messagebox.showerror("Ошибка валидации",
                                     "Деталь в сборке (BB) не может быть больше Уникальных деталей (AA).")
                return
        except ValueError:
            self.set_status("Ошибка: AA или BB не являются числами.", is_error=True)
            messagebox.showerror("Ошибка валидации", "AA или BB должны быть числовыми значениями.")
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

        # Version ordering validation:
        # Get the latest existing fixture for this specific assembly and unique part (KKK.SNN.DTT.AA)
        latest_existing_fixture = self.db_manager.get_latest_fixture_for_assembly(
            category, series, item_number, operation, fixture_number, unique_parts
        )

        if latest_existing_fixture:
            latest_vv = latest_existing_fixture['AssemblyVersionCode']
            latest_w = latest_existing_fixture['IntermediateVersion'] if latest_existing_fixture[
                'IntermediateVersion'] else ''
            latest_version_string = f"{latest_vv}{latest_w}"
            current_version_string = f"{assembly_version}{intermediate_version}"

            # If the current version is NOT the same as the latest existing version,
            # AND the current version is NOT strictly newer than the latest existing version,
            # then it's an invalid version attempt (i.e., trying to create an older version).
            if current_version_string != latest_version_string and \
                    not self.db_manager.is_version_newer(latest_version_string, current_version_string):
                self.set_status(
                    f"Ошибка: Новая деталь ({full_id_string}) должна иметь версию '{latest_version_string}' или более новую.",
                    is_error=True)
                messagebox.showerror("Ошибка версионирования",
                                     f"Деталь для оснастки {category}.{series}{item_number}.{operation}{fixture_number}.{unique_parts} должна иметь версию '{latest_version_string}' или более новую. Текущая версия: '{current_version_string}'.")
                return

        # Construct the specific folder name for the assembly version (KKK.SNN.DTT.AA0000-VVW)
        folder_version_name = (
            f"{parsed_id_for_path['Category']}."
            f"{parsed_id_for_path['Series']}{parsed_id_for_path['ItemNumber']}."
            f"{parsed_id_for_path['Operation']}{parsed_id_for_path['FixtureNumber']}."
            f"{parsed_id_for_path['UniqueParts']}0000-"  # BB and CC are replaced with 0000
            f"{parsed_id_for_path['AssemblyVersionCode']}"
            f"{parsed_id_for_path['IntermediateVersion'] if parsed_id_for_path['IntermediateVersion'] else ''}"
        )

        base_folder_path = os.path.join(
            self.db_manager.base_db_dir,
            parsed_id_for_path['Category'],
            f"{parsed_id_for_path['Category']}.{parsed_id_for_path['Series']}{parsed_id_for_path['ItemNumber']}",
            f"{parsed_id_for_path['Category']}.{parsed_id_for_path['Series']}{parsed_id_for_path['ItemNumber']}.{parsed_id_for_path['Operation']}{parsed_id_for_path['FixtureNumber']}",
            folder_version_name  # Use the newly constructed folder name
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

    def open_fixture_folder_command(self):
        """Opens the file explorer to the directory of the selected fixture."""
        if self.selected_fixture_id_in_list is None:
            self.set_status("Ошибка: Не выбрана оснастка для открытия директории.", is_error=True)
            messagebox.showerror("Ошибка", "Пожалуйста, выберите оснастку из списка, кликнув на нее.")
            return

        fixture_data = self.db_manager.get_fixture_id_by_id(self.selected_fixture_id_in_list)
        if not fixture_data or not fixture_data.get('BasePath'):
            self.set_status(
                f"Ошибка: Не удалось получить путь для выбранной оснастки ID {self.selected_fixture_id_in_list}.",
                is_error=True)
            messagebox.showerror("Ошибка", "Не удалось найти путь к папке выбранной оснастки.")
            return

        folder_path = fixture_data['BasePath']

        if not os.path.exists(folder_path):
            self.set_status(f"Ошибка: Директория не существует: '{folder_path}'", is_error=True)
            messagebox.showerror("Ошибка", f"Директория не существует:\n{folder_path}")
            return

        try:
            # For Windows
            if os.name == 'nt':
                os.startfile(folder_path)
            # For macOS
            elif os.uname().sysname == 'Darwin':
                subprocess.Popen(['open', folder_path])
            # For Linux
            else:
                subprocess.Popen(['xdg-open', folder_path])
            self.set_status(f"Открыта директория: '{folder_path}'", is_error=False)
        except Exception as e:
            self.set_status(f"Ошибка при открытии директории '{folder_path}': {e}", is_error=True)
            messagebox.showerror("Ошибка", f"Не удалось открыть директорию:\n{folder_path}\nОшибка: {e}")

    def on_fixture_list_select(self, event):
        """Handles selection events on the Treeview to select a fixture."""
        selected_item_id = self.fixture_list_tree.focus()
        if selected_item_id:
            # Get the actual data from the selected item
            values = self.fixture_list_tree.item(selected_item_id, 'values')
            # The fixture ID is stored as the last value (FullIDString)
            # We need to retrieve the actual database ID which is stored in the item's `iid`
            self.selected_fixture_id_in_list = selected_item_id  # Treeview item ID is its internal ID

            fixture_data = self.db_manager.get_fixture_id_by_id(int(self.selected_fixture_id_in_list))  # Convert to int
            if fixture_data:
                self.set_status(f"Выбрана оснастка для копирования: {fixture_data.get('FullIDString', 'N/A')}.")
            else:
                self.set_status(f"Выбрана оснастка с ID {self.selected_fixture_id_in_list}.", is_error=False)
        else:
            self.selected_fixture_id_in_list = None
            self.set_status("Выбор оснастки сброшен. Пожалуйста, кликните на строку оснастки.", is_error=False)

    def on_filter_checkbox_toggled(self):
        """Called when any filter checkbox is toggled."""
        self._refresh_fixture_list_with_current_selection()

    def _refresh_fixture_list_with_current_selection(self):
        """Helper to call load_fixtures_to_list with current combobox selections."""
        category_code = self._get_code_from_display_text(self.category_code_var.get())
        series_code = self._get_code_from_display_text(self.series_code_var.get())
        item_number_code = self._get_code_from_display_text(self.item_number_code_var.get())
        operation_code = self._get_code_from_display_text(self.operation_code_var.get())

        # If "Все..." is selected, pass None to db_manager to indicate no filter
        if "Все категории" in self.category_code_var.get():
            category_code = None
        if "Все серии" in self.series_code_var.get():
            series_code = None
        if "Все изделия" in self.item_number_code_var.get():
            item_number_code = None
        if "Все операции" in self.operation_code_var.get():
            operation_code = None

        self.load_fixtures_to_list(
            category_code=category_code,
            series_code=series_code,
            item_number_code=item_number_code,
            operation_code=operation_code
        )

    def load_fixtures_to_list(self, category_code=None, series_code=None, item_number_code=None, operation_code=None):
        """Загружает и отображает список оснасток в Treeview, с сортировкой и фильтрацией."""
        # Clear existing items in Treeview
        for item in self.fixture_list_tree.get_children():
            self.fixture_list_tree.delete(item)

        all_fixtures = self.db_manager.get_fixture_ids_with_descriptions(
            category_code=category_code,
            series_code=series_code,
            item_number_code=item_number_code,
            operation_code=operation_code
        )

        # Apply "hide assembled parts" filter
        if self.hide_assembled_parts_var.get():
            all_fixtures = [f for f in all_fixtures if f.get('PartInAssembly') == '01']

        filter_display_parts = []
        if category_code:
            filter_display_parts.append(f"категории '{self._get_name_from_code(category_code, 'category')}'")
        if series_code:
            filter_display_parts.append(f"серии '{self._get_name_from_code(series_code, 'series')}'")
        if item_number_code:
            filter_display_parts.append(f"изделия '{self._get_name_from_code(item_number_code, 'item_number')}'")
        if operation_code:
            filter_display_parts.append(f"операции '{self._get_name_from_code(operation_code, 'operation')}'")

        if filter_display_parts:
            self.list_label.configure(text=f"Список оснасток для {', '.join(filter_display_parts)}:")
        else:
            self.list_label.configure(text="Список существующих оснасток (все):")

        # Add filter status to label
        filter_status_parts = []
        if self.hide_non_actual_versions_var.get():
            filter_status_parts.append("скрыты неактуальные версии")
        if self.hide_assembled_parts_var.get():
            filter_status_parts.append("скрыты сборные части (BB!=01)")

        if filter_status_parts:
            self.list_label.configure(
                text=self.list_label.cget("text") + f" (Фильтры: {', '.join(filter_status_parts)})")

        if not all_fixtures:
            if filter_display_parts:
                self.fixture_list_tree.insert("", "end", values=[
                    f"Оснасток для {', '.join(filter_display_parts)} не существует в БД."], tags=('no_data',))
            else:
                self.fixture_list_tree.insert("", "end", values=["Оснасток в базе данных пока нет."], tags=('no_data',))
            self.fixture_list_tree.tag_configure('no_data', foreground='gray')
            return

        # Group fixtures by their assembly base (KKK.SNN.DTT.AA)
        assembly_groups = {}
        for fixture in all_fixtures:
            assembly_base_key = (
                fixture['Category'],
                fixture['Series'],
                fixture['ItemNumber'],
                fixture['Operation'],
                fixture['FixtureNumber'],
                fixture['UniqueParts']
            )
            if assembly_base_key not in assembly_groups:
                assembly_groups[assembly_base_key] = []
            assembly_groups[assembly_base_key].append(fixture)

        actual_fixtures = []
        non_actual_fixtures = []

        for group_key, fixtures_in_group in assembly_groups.items():
            if not fixtures_in_group:
                continue

            # Define a key for sorting to find the absolute latest version string
            def get_version_string_for_sort(f):
                vv_code = f['AssemblyVersionCode']
                w_code = f['IntermediateVersion'] if f['IntermediateVersion'] else ''
                return f"{vv_code}{w_code}"

            # Find the maximum version string in the group
            # This will correctly identify the "latest" version string (e.g., "02A" > "01Z")
            # based on lexicographical comparison, which aligns with the is_version_newer logic.
            latest_version_in_group_str = ""
            if fixtures_in_group:
                # Initialize with the version string of the first fixture
                latest_version_in_group_str = get_version_string_for_sort(fixtures_in_group[0])
                for i in range(1, len(fixtures_in_group)):
                    current_v_str = get_version_string_for_sort(fixtures_in_group[i])
                    # Use db_manager.is_version_newer to compare, ensuring 'X' versions are handled
                    if self.db_manager.is_version_newer(latest_version_in_group_str, current_v_str):
                        latest_version_in_group_str = current_v_str
                    elif current_v_str == latest_version_in_group_str:
                        # If versions are equal, it's also considered latest, no change to latest_version_in_group_str
                        pass
                    # If current_v_str is older, latest_version_in_group_str remains unchanged

            # Now, iterate through the group and categorize based on this latest version string
            for fixture in fixtures_in_group:
                current_fixture_version_string = get_version_string_for_sort(fixture)
                if current_fixture_version_string == latest_version_in_group_str:
                    actual_fixtures.append(fixture)
                else:
                    non_actual_fixtures.append(fixture)

        # Sort both lists alphabetically by FullIDString for consistent display order
        actual_fixtures.sort(key=lambda x: x['FullIDString'])
        non_actual_fixtures.sort(key=lambda x: x['FullIDString'])

        # Conditionally combine for display based on checkbox state
        if self.hide_non_actual_versions_var.get():
            sorted_fixtures_for_display = actual_fixtures
        else:
            sorted_fixtures_for_display = actual_fixtures + non_actual_fixtures

        for fixture in sorted_fixtures_for_display:
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

            self.fixture_list_tree.insert("", "end", iid=fixture['id'],  # Use database ID as Treeview item ID
                                          values=(
                                              category_display,
                                              series_display,
                                              item_display,
                                              operation_display,
                                              fixture_number,
                                              unique_parts,
                                              part_in_assembly,
                                              part_quantity,
                                              combined_version_vvw,
                                              full_id_string_display
                                          ))

    def _get_code_from_display_text(self, display_text):
        """Извлекает код из строки вида 'CODE (Description)' или возвращает None, если это заглушка 'Все...'."""
        # Handle placeholder/all options explicitly by returning None
        if "Выберите" in display_text or "Нет данных" in display_text or "Все " in display_text or "Заполните" in display_text:
            return None
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

        # Validate BB <= AA
        if aa_str.isdigit() and bb_str.isdigit():
            try:
                aa_int = int(aa_str)
                bb_int = int(bb_str)
                if bb_int > aa_int:
                    self.set_status("BB не может быть больше AA.", is_error=True)
                    is_valid = False
            except ValueError:
                # Should not happen if isdigit() is true, but for safety
                self.set_status("Ошибка конвертации AA или BB в число.", is_error=True)
                is_valid = False

        if is_valid:
            self.set_status("Поля AA, BB, CC валидны.", is_error=False)
        return is_valid

    def validate_vv_input(self, event=None):
        vv_str = self.assembly_version_vv_var.get().upper()

        if not vv_str:
            self.set_status("Версия сборки (VV) не может быть пустой.", is_error=True)
            return False

        if len(vv_str) != 2:
            self.set_status("Версия сборки (VV) должна быть 2 символа.", is_error=True)
            return False

        # Validate VV format: 0-9, A-Z, but if first char is not 'X', then both must be digits
        if vv_str[0] == 'X':
            # 'X' versions can be X0-X9, XA-XZ etc. Second char can be anything valid.
            if not all(c in '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ' for c in vv_str[1]):
                self.set_status("Версия сборки (VV) после 'X' должна быть 0-9, A-Z.", is_error=True)
                return False
        else:
            # If not 'X', both characters must be digits
            if not vv_str.isdigit():
                self.set_status("Версия сборки (VV) должна быть числовой (например, 01, 10), или начинаться с 'X'.",
                                is_error=True)
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

                added_summary = []
                for sheet_name, count in added_counts.items():
                    if count > 0:
                        added_summary.append(f"  {sheet_name}: {count} новых записей")
                if added_summary:
                    report_message += "Добавлено:\n" + "\n".join(added_summary) + "\n\n"
                else:
                    report_message += "Новых записей не добавлено.\n\n"

                updated_summary = []
                for sheet_name, count in updated_counts.items():
                    if count > 0:
                        updated_summary.append(f"  {sheet_name}: {count} обновленных записей")
                if updated_summary:
                    report_message += "Обновлено:\n" + "\n".join(updated_summary) + "\n\n"
                else:
                    report_message += "Записей не обновлено.\n\n"

                missing_summary = []
                for sheet_name, items in missing_from_excel_data.items():
                    if items:
                        missing_summary.append(f"  {sheet_name}: {len(items)} записей отсутствуют в Excel:")
                        for item in items:
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
                                missing_summary.append(f"    {item}")
                if missing_summary:
                    report_message += "Записи в базе данных, отсутствующие в Excel:\n" + "\n".join(missing_summary)
                else:
                    report_message += "Все записи из базы данных найдены в Excel."

                self.set_status("Импорт завершен. Подробности в отчете.", is_error=False)
                messagebox.showinfo("Отчет по импорту Excel", report_message)

                self.load_all_combobox_data()  # Reload data and reset comboboxes to placeholders
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
                # self.fixture_list_textbox.tag_remove("highlight", "1.0", "end") # No longer needed
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
