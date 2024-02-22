import os
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QObject, pyqtSignal
import sys, os, openpyxl, subprocess
from pyvis.network import Network
import random

class View(QMainWindow):
    btn_pressed = pyqtSignal(str) # Сигнал для передачи текста кнопки

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Работа с файлами")
        self.pefix_path = 'Текущий путь: '
        self.pefix_file = 'Текущий файл: '
        self.folder_path = './'
        self.file_name = 'Файл не выбран'
        self.key_pattern = "3323 232 323 2212"
        self.path_full = None
        self.open_list_file()
        self.initUI()
        self.create_tree()

    def initUI(self):
        tab_widget = QTabWidget()
        self.setCentralWidget(tab_widget)

        file_tab = QWidget()
        root_layout = QVBoxLayout(file_tab)
        tree_layout = QHBoxLayout()
        btn_layout = QVBoxLayout()

        self.path_label = QLabel(self.pefix_path + self.folder_path)
        root_layout.addWidget(self.path_label)

        self.file_label = QLabel(self.pefix_file + self.file_name)
        root_layout.addWidget(self.file_label)

        root_layout.addWidget(QLabel("Введите код ключа F14:"))

        self.keycod_value = QLineEdit()
        self.keycod_value.setText(self.key_pattern)
        root_layout.addWidget(self.keycod_value)
        self.keycod_value.editingFinished.connect(self.entry_callback)

        root_layout.addWidget(QLabel())

        root_layout.addLayout(tree_layout)
        tree_layout.addLayout(btn_layout)

        self.treeview = QTreeWidget()
        tree_layout.addWidget(self.treeview)
        self.treeview.itemClicked.connect(self.tree_callback)

        self.btn_path = QPushButton("Выбрать путь")
        btn_layout.addWidget(self.btn_path)
        self.btn_path.pressed.connect(self.draft_path)

        self.btn_create_pattern = QPushButton("Создать шаблон")
        btn_layout.addWidget(self.btn_create_pattern)
        self.btn_create_pattern.pressed.connect(lambda: self.btn_pressed.emit(self.btn_create_pattern.text()))
        self.btn_create_pattern.pressed.connect(self.create_xls)

        self.btn_open_xls = QPushButton("Открыть файл")
        btn_layout.addWidget(self.btn_open_xls)
        self.btn_open_xls.pressed.connect(lambda: self.btn_pressed.emit(self.btn_open_xls.text()))

        self.btn_create_ms = QPushButton("Создать МС")
        btn_layout.addWidget(self.btn_create_ms)
        self.btn_create_ms.pressed.connect(lambda: self.btn_pressed.emit(self.btn_create_ms.text()))      

        self.btn_create_g = QPushButton("Создать G")
        btn_layout.addWidget(self.btn_create_g)
        self.btn_create_g.pressed.connect(lambda: self.btn_pressed.emit(self.btn_create_g.text()))      

        self.btn_exit = QPushButton("Выход")
        btn_layout.addWidget(self.btn_exit)
        self.btn_exit.clicked.connect(self.close)

        setting_tab = QWidget()
        setting_layout = QVBoxLayout(setting_tab)
        setting_layout.addWidget(QLabel("Настройки"))

        tab_widget.addTab(file_tab, "Выбор файла")
        tab_widget.addTab(setting_tab, "Настройки")
        self.check_status()

    def create_xls(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'CrossTable'
        file_name = 'Шаблон МС.xlsx'
        file_path = os.path.join(self.folder_path, file_name)
        columns = ['K1', 'K2', 'K3', 'K4', 'K5']
        rows = {1: ['x', '', '', '', '', ], 2: ['x', '', '', '', '', ], 3: ['x', '', '', '', '', ], 4: ['x', '', '', '', '', ], 5: ['x', '', '', '', '', ]}

        # Запись заголовков столбцов
        for col_num, col_value in enumerate(columns, start=2):
            ws.cell(row=1, column=col_num, value=col_value)

        # Запись данных строк
        for row_num, row_values in rows.items():
            ws.cell(row=row_num+1, column=1, value=row_num)
            for col_num, col_value in enumerate(row_values, start=2):
                ws.cell(row=row_num+1, column=col_num, value=col_value)

        # Сохранение файла
        wb.save(file_path)
        self.open_list_file()
        self.create_tree()
    
    def check_status(self, msg_txt = "Готов к работе"):
        self.statusBar().showMessage(msg_txt)

    def tree_callback(self, item:QTreeWidgetItem, column):
        selected_item = item
        parent = item.parent()
        if parent is None:
            self.file_name = selected_item.text(0)
        else:
            self.file_name = parent.text(0)

        self.file_label.setText(self.pefix_file + self.file_name)
        self.path_full = os.path.join(self.folder_path, self.file_name)
        is_valid_file = self.check_sheet_exists(self.path_full)
        if is_valid_file:
            self.check_status()
        else:
            self.check_status(f'В книге нет листа CrossTable')
            self.path_full = None

    def check_sheet_exists(self, file_path):
        sheet_name = "CrossTable"
        wb = openpyxl.load_workbook(file_path)
        sheet_names = wb.sheetnames
        if sheet_name in sheet_names:
            return True
        else:
            return False

    def entry_callback(self):
        key_code_in = self.keycod_value.text().replace(' ', '')
        key_list = list(key_code_in)
        key_set = set(key_list)
        key_code_out = f'{"".join(key_list[0:4])} {"".join(key_list[4:7])} {"".join(key_list[7:10])} {"".join(key_list[10:])}'
        self.key_pattern = None
        if len(key_code_in) != 14:
            self.keycod_value.setStyleSheet("color: red;")
            self.check_status(msg_txt =f'Длина кода ключа {len(key_code_in)} не равна 14')
        elif not key_set.issubset({'1','2','3','4'}):
            self.keycod_value.setStyleSheet("color: red;")
            self.check_status(msg_txt =f'Код ключа {len(key_code_in)} должен быть из цифр от 1 до 4')
        else:
            self.keycod_value.setStyleSheet("color: black;")
            self.key_pattern = key_code_out
            self.check_status()
        self.keycod_value.clear()
        self.keycod_value.insert(key_code_out)

    def open_list_file(self):
        # Получение списка файлов в папке
        self.file_list = {file:[] for file in os.listdir(self.folder_path) if file.endswith('.xlsm') or file.endswith('.xlsx') and not file.startswith('~$')}
        if len(self.file_list) == 0:
            self.file_list = {'Файлов не обнаружено': ['Листов не обнаружено']}
        else:
            for file_name in self.file_list:            
                # Полный путь к выбранному файлу
                file_path = os.path.join(self.folder_path, file_name)
                # Загрузка файла Excel
                wb = openpyxl.load_workbook(file_path)
                # Получение списка листов в wb
                self.file_list[file_name] = wb.sheetnames
    
    def create_tree(self):
        self.treeview.setHeaderLabels(['Файл'])
        self.treeview.clear()
        for file, sheets in self.file_list.items():
            file_item = QTreeWidgetItem(self.treeview, [file])
            for sheet in sheets:
                sheet_item = QTreeWidgetItem(file_item, [sheet])
                file_item.addChild(sheet_item)
            self.treeview.addTopLevelItem(file_item)
    
    def draft_path(self):
        self.check_status(msg_txt ='Обновляем информацию о пути')
        self.folder_path = QFileDialog.getExistingDirectory(self, 'Выберите папку')
        if self.folder_path == '':
            self.folder_path = './'
        self.path_label.setText(self.pefix_path + self.folder_path)
        self.open_list_file()
        self.create_tree()
        self.file_name = 'Файл не выбран'
        self.file_label.setText(self.pefix_file + self.file_name)
        self.check_status()

    def is_file_cheked(self):
        path = os.path.join(self.folder_path, self.file_name)
        if os.path.exists(path) and self.file_name != 'Файл не выбран':
            self.check_status(msg_txt="Готов к работе")
            return True
        else:
            self.check_status(msg_txt="Файл не выбран")
            return False

class Presenter:
    def __init__(self, model, view):
        self.data = None
        self.model = model
        self.view = view
        view.btn_pressed.connect(self.press_btn) # Подключение сигнала к методу

    def press_btn(self, btn_text): # Исправлено здесь
        match btn_text:
            case "Создать МС":
                if self.view.path_full is not None:
                    if self.view.key_pattern is not None:
                        key_cod = list(self.view.keycod_value.text().replace(' ', ''))
                        path = self.view.path_full
                        self.view.check_status(msg_txt='Создаем шаблон ключа ...')
                        self.model.init_cross(key_cod, path)
                        self.view.check_status(msg_txt='Создаем коды нарезки цилиндров ...')
                        if self.model.create_table(): 
                            self.view.check_status(msg_txt='Создаем коды нарезки ключей ...')    
                            self.model.cut_keys()
                            self.view.check_status(msg_txt='Проверяем коды нарезки ключей ...')
                            if self.model.check_keys():
                                self.view.check_status(msg_txt='Выгружаем данные в файл ...')
                                self.model.upload_xlsx(path)
                                self.view.open_list_file()
                                self.view.create_tree()
                                self.view.check_status(msg_txt='МС успешно создана')
                            else:
                                self.view.check_status(msg_txt=self.model.msg_error)
                        else:
                            self.view.check_status(msg_txt=self.model.msg_error)
                    else:
                        self.view.check_status(msg_txt="Ошибка в коде ключа")
                else:
                    self.view.check_status(msg_txt="Не выбран файл с CrossTable")
            case "Создать шаблон":
                print("Нажата кнопка:", btn_text)
            case "Открыть файл":
                if self.view.is_file_cheked():
                    subprocess.Popen(['start', '', self.view.path_full], shell=True)
                print("Нажата кнопка:", btn_text)
            case "Выбрать путь":
                print("Нажата кнопка:", btn_text)
            case "Создать G":
                if self.view.is_file_cheked():
                    ms_graf = MasterGraf()
                print("Нажата кнопка:", btn_text)
        
class Cross():
    KEY_COD = None
    CYLINDER_UNIQ = None
    CYLINDER_DICT = None
    CYLINDER_TABLE = None
    KEY_TABLE = None
    KEY2CYLINDER = None
    MASK = {
        '1':'ZUDLNH',
        '2':'ZUDLNH',
        '3':'ZUDLNH',
        '4':'ZUDLNH',
        '5':'ZNH',
        '6':'ZNH',
        '7':'ZNH',
        '8':'ZUDRNH',
        '9':'ZUDRNH',
        '10':'ZUDRNH',
        '11':'ZUDRNH',       
        '12':'ZNH',
        '13':'ZNH',
        '14':'ZNH',
        }
    KEY_MASK = None
    KEY_DICT = {
'K9':[35, 47, 48, 49, ],
'K12':[65, 87, 89, ],
'K8':[99, ],
'K33':[106, ],
'K31':[58, 65, 66, 109, ],
'K14':[54, 55, 56, 57, 58, 65, 109, ],
'K32':[58, 65, 67, 68, 69, 74, 109, ],
'K13':[53, 54, 55, 56, 57, 58, 65, 109, ],
'K10':[58, 65, 75, 76, 77, 79, 80, 98, 109, ],
'K11':[30, 31, 32, 58, 65, 79, 80, 96, 98, 109, ],
'K15':[27, 28, 29, 30, 31, 33, 34, 38, 58, 65, 74, 79, 80, 98, 109, ],
'K34':[108, 109, 110, ],
'K23':[6, 15, 17, 18, 58, 60, 65, 109, 110, ],
'K24':[6, 16, 17, 18, 58, 60, 65, 109, 110, ],
'K25':[6, 17, 18, 19, 20, 58, 60, 65, 109, 110, ],
'K26':[6, 17, 18, 19, 20, 21, 58, 60, 65, 109, 110, ],
'K7':[27, 28, 29, 30, 31, 33, 34, 35, 36, 37, 38, 58, 64, 65, 74, 75, 76, 77, 79, 80, 95, 97, 98, 109, 111, ],
'K6':[1, 2, 3, 4, 6, 27, 28, 29, 30, 31, 33, 34, 36, 37, 38, 43, 44, 58, 65, 71, 72, 73, 74, 79, 80, 92, 94, 97, 98, 100, 103, 104, 105, 107, 109, 111, ],
'K4':[1, 2, 3, 4, 22, 24, 25, 27, 28, 29, 30, 31, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 52, 58, 59, 62, 63, 65, 71, 72, 73, 74, 78, 79, 80, 87, 88, 92, 93, 94, 97, 98, 100, 101, 103, 104, 105, 107, 109, 111, ],
'K35':[112, ],
'K18':[6, 10, 17, 18, 58, 60, 65, 109, 110, 113, ],
'K19':[6, 11, 17, 18, 58, 60, 65, 109, 110, 113, ],
'K20':[6, 12, 17, 18, 58, 60, 65, 109, 110, 113, ],
'K21':[6, 13, 17, 18, 58, 60, 65, 109, 110, 113, ],
'K22':[6, 14, 17, 18, 58, 60, 65, 109, 110, 113, ],
'K16':[6, 7, 8, 17, 18, 58, 60, 65, 109, 110, 113, ],
'K17':[6, 8, 9, 17, 18, 58, 60, 65, 109, 110, 113, ],
'K5':[1, 2, 3, 4, 6, 27, 28, 29, 30, 31, 33, 34, 36, 37, 38, 43, 44, 58, 65, 70, 71, 72, 73, 74, 79, 80, 92, 94, 97, 98, 100, 103, 104, 105, 107, 109, 110, 111, 113, ],
'K3':[1, 2, 3, 4, 6, 22, 24, 25, 27, 28, 29, 30, 31, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 52, 58, 59, 61, 62, 63, 65, 71, 72, 73, 74, 78, 79, 80, 87, 88, 92, 93, 94, 97, 98, 100, 101, 103, 104, 105, 107, 109, 110, 111, 113, ],
'K2':[6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 23, 24, 25, 27, 28, 29, 30, 31, 32, 33, 34, 35, 38, 51, 54, 55, 56, 57, 58, 60, 61, 62, 63, 64, 65, 67, 68, 69, 70, 71, 72, 73, 74, 79, 80, 87, 94, 95, 96, 98, 106, 108, 109, 110, 113, ],
'K1':[1, 2, 3, 4, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 47, 48, 49, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 87, 88, 89, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, ],
}
    
    STOP_LIST = {'1N', '4H', '1U', '4U', '1D', '4D', '1L', '4L', '1R', '4R'}

    def __init__(self, key_cod, path):
        super().__init__()
        self.KEY_COD = key_cod
        self.KEY_MASK = self.createmask()
        self.KEY_DICT = self.create_key_dict(path)
        my_keys = self.KEY_DICT
        my_cylinders = self.create_cylinders_dict(my_keys)
        my_cylinders = self.resort_cylinders_dict(my_cylinders)
        my_uniq = self.create_uniq_dict(my_cylinders)
        my_uniq_dic = dict()
        for uniq_value in my_uniq.values():
            for itm in uniq_value:
                my_uniq_dic[itm] = uniq_value[0]
        self.KEY2CYLINDER = my_uniq
        self.CYLINDER_DICT = my_uniq_dic
        self.CYLINDER_UNIQ = sorted(list(set(self.CYLINDER_DICT.values())))
  
    def draft_cross(self, path):
        my_path = path
        wb = openpyxl.load_workbook(my_path)
        sheet = wb['CrossTable']
        cross_dict = dict()
        # Находим диапазон непрерывных данных в первой строке
        data_range = sheet[1]
        start_column = None
        end_column = None
        for cell in data_range:
            if cell.value is not None:
                if start_column is None:
                    start_column = cell.column
                end_column = cell.column

        # Выводим адрес диапазона в терминал
        if start_column is not None and end_column is not None:
            cross_dict['columns'] = (start_column, end_column+1)
            start_cell = sheet.cell(row=1, column=start_column)
            end_cell = sheet.cell(row=1, column=end_column)
            print(f"Диапазон непрерывных данных: {start_cell.coordinate}:{end_cell.coordinate}")
        else:
            print("Нет непрерывных данных в первой строке")

        data_column = sheet['A']
        start_row = None
        end_row = None

        for cell in data_column:
            if cell.value is not None:
                if start_row is None:
                    start_row = cell.row
                end_row = cell.row

        # Выводим адрес диапазона в терминал
        if start_row is not None and end_row is not None:
            cross_dict['rows'] = (start_row, end_row+1)
            start_cell = sheet.cell(row=start_row, column=1)
            end_cell = sheet.cell(row=end_row, column=1)
            print(f"Диапазон непрерывных данных: {start_cell.coordinate}:{end_cell.coordinate}")
        else:
            print("Нет непрерывных данных в первой строке")
        return cross_dict

    def create_key_dict(self, path):
        cross_dict = self.draft_cross(path)
        my_path = path
        wb = openpyxl.load_workbook(my_path)
        sheet = wb['CrossTable']
        key_dict = dict()
        columns = cross_dict['columns']
        rows = cross_dict['rows']
        for i_col in range(columns[0], columns[1]):
            key_num = sheet.cell(row=1, column=i_col).value
            key_dict[key_num] = []
            for i_row in range(rows[0], rows[1]):
                cell_value = sheet.cell(row=i_row, column=i_col).value
                if cell_value is not None:
                    key_dict[key_num].append(sheet.cell(row=i_row, column=1).value)
        
        return key_dict   
    
    def createmask(self):
        key_mask = []
        for i, cod in enumerate(self.KEY_COD):
            stroka = []
            for k in range(len(self.MASK[str(i+1)])):
                pin_num = str(cod) + list(self.MASK[str(i+1)])[k]
                is_stop = set([pin_num]).issubset(self.STOP_LIST) 
                if not is_stop:
                    sub_str = str(i+1) + list(self.MASK[str(i+1)])[k]
                    stroka.append(sub_str)
            key_mask.append(stroka)

        return key_mask
    
    def create_cylinders_dict(self, keys_dict):
        cylinders_dict = {}
        for key, cylinders in keys_dict.items():
            for cylinder in cylinders:
                if cylinder not in cylinders_dict:
                    cylinders_dict[cylinder] = []
                cylinders_dict[cylinder].append(key)
        return cylinders_dict

    def resort_cylinders_dict(self, cylinders_dict):
        for idx, itm in cylinders_dict.items():
            cylinders_dict[idx] = itm[::-1]
        return cylinders_dict

    def create_uniq_dict(self, keys_dict):
        uniq_dict = {}
        for key, cylinders in keys_dict.items():
            if str(cylinders) not in uniq_dict:
                uniq_dict[str(cylinders)] = []
            uniq_dict[str(cylinders)].append(key)
        return uniq_dict

class Model():
    def __init__(self):
        super().__init__()
        self.for_check_keys = None
        self.checked = None
        self.key_table = None
        self.cylinder_dict = None
        self.cod_table = None
        self.sizes_of_groups = None
        self.value = None
        self.possible = True
        self.number_options = None
        self.input = None
        self.msg_error = None
    
    def init_cross(self, key_cod, path):
        cross = Cross(key_cod, path)
        self.key_cod = cross.KEY_COD
        self.input = cross.KEY_MASK
        self.key_dict = cross.KEY_DICT
        self.key_list = list(cross.KEY_DICT.values())
        self.cylinder_unique = cross.CYLINDER_UNIQ
        self.cylinder_uniq_dict = cross.CYLINDER_DICT
        self.key2cylinder = cross.KEY2CYLINDER

    def dict_revers(self):
        self.cylinder_dict_revers = dict()
        for idx, value in self.cylinder_dict.items():
            self.cylinder_dict_revers[value] = idx
        
        self.cylinder_uniq_dict_revers = dict()
        for idx, value in self.cylinder_uniq_dict.items():
            if value in self.cylinder_uniq_dict_revers:
                self.cylinder_uniq_dict_revers[value].append(idx)
            else:
                self.cylinder_uniq_dict_revers[value] = [idx]    
        
        self.key_dict_revers = dict()
        for idx, value in self.key_dict.items():
            self.key_dict_revers[str(value)] = idx
    
    def create_cylinder_dic(self):
        self.dict_revers()
        self.cylinder_cut = dict()
        for idx, cod in enumerate(self.cod_table):
            itm_str = '#;'.join(cod)+'#'
            id = self.cylinder_uniq_dict_revers[self.cylinder_dict_revers[idx]]
            id_str = ', '.join([str(i) for i in id])
            self.cylinder_cut[id_str] = itm_str.split(';')
        return self.cylinder_cut
    
    def create_key_dic(self):
        self.key_cut = dict()
        for key in self.key_list:
            cods_key = ''
            id = self.key_dict_revers[str(key)]
            for j in range(len(self.sizes_of_groups)):
                cod_key = set()
                for x in range(len(key)):
                    cod_key.add(self.cod_table[self.cylinder_dict[self.cylinder_uniq_dict[key[x]]]][j])
                cods_key = cods_key  + (''.join(cod_key))+'#;'
            key_item = cods_key.split('#;')
            key_item.remove('')
            self.key_cut[id] = (key_item)
        return self.key_cut
    
    def check_keys(self):
        # @title Тест ключей
        self.possible = True
        self.checked = dict()
        is_err = False
        for key_i, key in self.key_dict.items(): # key_i номер в списке ключей 
            is_cyl_err = list()
            cylinder_list = self.cylinder_unique.copy() # Список цилиндров key_list
            cylinder_lost = list(set(cylinder_list)-set(key)) # рзность множеств

            for cyl_lost_i in cylinder_list: # Перебираем оставшиеся цилиндры
                if cyl_lost_i in cylinder_lost:
                    set_cyl = set(self.cod_table[self.cylinder_dict[cyl_lost_i]])
                    set_key = set(self.for_check_keys[key_i])
                    sub_set = set_cyl.issubset(set_key)
                    if sub_set: # если ключ открывает цилиндр
                        is_cyl_err.append(cyl_lost_i)

            if len(is_cyl_err) == 0:
                self.checked[key_i] = ['Passed']
            else:
                is_err = True
                self.checked[key_i] = is_cyl_err
        if is_err:
            self.possible = False
            self.msg_error = f'Ключи открывают НЕ свои цилиндры!!!'
            # self.show_msg(title='О ужас!', msg_txt=msg_txt)
        return self.possible

    
    def cut_keys(self):
        self.key_table = [] #создаем позиции нарезки ключей
        self.for_check_keys = {}
        for key, cylinders in self.key_dict.items(): #перебираем ключи из списка ключей
            cods_key = '' #создаем позиции нарезки ключа в виде строки
            check_key= list()
            for y in range(len(self.sizes_of_groups)): #записываем число вариантов на каждой позиции
                cod_key = list() #множество для формирования списка только уникальных позиций нарезки
                for x in range(len(cylinders)): #номера цилиндров в коде ключа
                    cod = self.cod_table[self.cylinder_dict[self.cylinder_uniq_dict[cylinders[x]]]][y]
                    cod_key.append(cod)  
                    check_key.append(cod)
                cods_key += (''.join(set(cod_key)))+'#;' #строка нарезки ключа с разделителями
            key_item = cods_key.split('#;') #преобразуем строку нарезки кода ключа в список
            key_item.remove('') #удаляем из списка нарезки кода ключа лишние символы
            self.key_table.append(key_item) #добавляем нарезку кода ключа в список ключей
            self.for_check_keys[key] = set(check_key)
    
    def increment_value(self, value, sizes_of_groups):
        for i in range(len(value)):#reversed()
            if (value[i] + 1) % sizes_of_groups[i] != 0 and sizes_of_groups[i] !=1:
                value[i] += 1
                return
            value[i] = 0
        pass

    def value_to_target(self, value, target):
        result = [None] * len(value)
        for i, v in enumerate(value):
            result[i] = target[i][v]
        return result

    def create_table(self):
        # 
        self.possible =True
        sizes_of_groups = [len(x) for x in self.input]
        self.sizes_of_groups = sizes_of_groups.copy()
        self.number_options = sum([n for n in sizes_of_groups if n!=1])
        value = [0] * len(self.input)
        self.cod_table = []
        # self.cylinder_unique = sorted(list(set(itm for sublist in self.key_list for itm in sublist)))
        if len(self.cylinder_unique) > self.number_options:
            self.possible = False
            self.msg_error = f'Количество цилиндров {len(self.cylinder_unique)} \nпревышает количество допустимых комбинаций \nкода ключа {self.number_options}'
            # self.show_msg(title='О ужас', msg_txt=msg_txt)
            return self.possible
        self.cylinder_dict = dict(zip(self.cylinder_unique, range(len(self.cylinder_unique))))
        j = len(self.cylinder_unique)
        self.increment_value(value, sizes_of_groups)
        self.down_grate(value, sizes_of_groups)
        for i in range(j):#VALUES
            self.cod_table.append(self.value_to_target(value, self.input))
            self.increment_value(value, sizes_of_groups)
            self.down_grate(value, sizes_of_groups)

        # print('Печатаем код ключа\n', list(self.key_cod))       
        return self.possible
    
    def down_grate(self, value, sizes_of_groups):
        for i in range(len(value)):
            if value[i]+1 == sizes_of_groups[i]:
                sizes_of_groups[i] = 1
            else:
                break

    def upload_sheet(self, my_sheet, my_dict):
        # Запись значений в файл XLSX
        for row, (cylinder, keys) in enumerate(my_dict.items(), start=1):
            my_sheet.cell(row=row, column=1, value=cylinder)
            my_sheet.cell(row=row, column=2, value=', '.join([str(i) for i in keys]))

    def upload_xlsx(self, path):
        wb = openpyxl.load_workbook(path)
        # sheet = wb.active
        # wb.remove_sheet(sheet)
        sheet_list = {'Cylinders':self.create_cylinder_dic(), 
                      'Keys':self.create_key_dic(),
                      'CylinderCross':self.key2cylinder,
                      'KeyCross': self.key_dict,
                      'CheckKey': self.checked} # , 'Check keys'
        for title, itm_dic in sheet_list.items():
            sheet = wb.create_sheet(title=title)
            sheet = wb[title]
            self.upload_sheet(sheet, itm_dic)
        self.msg_error = f'Подготовка файлов завершена.\n'
        # self.show_msg(title='Успех!', msg_txt=msg_txt)

        # Сохранение файла XLSX
        wb.save(path)
    

    # def show_msg(self, title='Информационное сообщение', msg_txt='Программа работает в обычном режиме'):
    #     app1 = QApplication(sys.argv)
    #     dialog = QDialog()
    #     dialog.setWindowTitle(title)
    #     layout = QVBoxLayout(dialog)
    #     layout.addWidget(QLabel(msg_txt))
    #     quit_button = QPushButton("Quit")
    #     quit_button.clicked.connect(dialog.reject)
    #     layout.addWidget(quit_button)
    #     dialog.setModal(True)
    #     dialog.exec_()
    #     sys.exit(app1.exec_())

class MasterGraf():
    def __init__(self):
        super().__init__()
        self.dic_sorted = None # количество цилиндров - ключи
        self.dic_cross = None # ключ - цилиндры
        self.dic_g = None # иерархия ключей
        my_path = './Шаблон МС.xlsx'
        my_sheet = 'KeyCross'
        self.draft_cross(my_path, my_sheet)
        self.search_graf()
        self.create_graf()
    
    def generate_hex_color(self):
        color = '#{:06x}'.format(random.randint(0, 0xFFFFFF))
        return color
    
    def draft_cross(self, my_path, my_sheet):
        wb = openpyxl.load_workbook(my_path)
        sheet = wb[my_sheet]
        self.dic_cross = dict()
        # Находим диапазон непрерывных данных в первой строке
        row_range = len(sheet['A'])
        for i_row in range(1, row_range+1):
            idx = sheet.cell(row=i_row, column=1).value
            val = sheet.cell(row=i_row, column=2).value
            val = set(val.replace(' ','').split(','))
            self.dic_cross[idx] = val
        
        dic_count = dict()
        for idx, value in self.dic_cross.items():
            cnt = len(value)
            if cnt in dic_count:
                dic_count[cnt].add(idx)
                next
            else:
                dic_count[cnt] = {idx}
        self.dic_sorted = dict(sorted(dic_count.items(), key=lambda x: x[0], reverse=True))
        return 
    
    def search_graf(self):
        dic = {}
        len_mem = None
        end_while = True
        while end_while:
            if len(dic) == 0:
                len_mem = 0
                idx = max(self.dic_sorted)
                first_key = list(self.dic_sorted[idx])
                for i in first_key:
                    dic = self.incr(dic, i)
                    len_dic = len(dic)
            else:
                for my_list in list(dic.values()):
                    for i in my_list:
                        dic = self.incr(dic, i)
                        len_dic = len(dic)
            if len_dic > len_mem:
                len_mem = len_dic
            else:
                end_while = False
                self.dic_g = dic
        return
    
    def incr(self, dic, m_key):
        if m_key not in dic:
            for keys in self.dic_sorted.values():
                    for key in list(keys):
                        if m_key == key:
                            continue
                        my_set = self.dic_cross[m_key]
                        sub_set = self.dic_cross[key]
                        if sub_set.issubset(my_set):
                            self.dic_cross[m_key] = my_set - sub_set
                            if m_key not in dic:
                                dic[m_key] = [key] 
                            else:
                                dic[m_key].append(key)
        return dic

    def add_nod(self, *args):
        for arg in args:
            if arg not in list(self.G.nodes):
                self.G.add_node(arg, color='#ff0000', size=10) if 'K' in arg else self.G.add_node(arg, size=5)
                    
    def create_graf(self):
        self.G = Network()
        for idx, values in self.dic_g.items():
            self.add_nod(idx)
            for value in values:
                self.add_nod(value)
                self.G.add_edge(idx, value)
                for cyl in self.dic_cross[value]:
                    self.add_nod(cyl)
                    self.G.add_edge(value, cyl) 
            for value in self.dic_cross[idx]:
                self.add_nod(value)
                self.G.add_edge(idx, value)
        for my_nods in self.dic_sorted.values():
            for my_nod in my_nods:
                self.add_nod(my_nod)
        self.G.show('graph.html', notebook=False) # save visualization in 'graph.html'
        return
       
def main():
    app = QApplication(sys.argv)
    model = Model()
    view = View()
    presenter = Presenter(model, view)
    view.show()

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()