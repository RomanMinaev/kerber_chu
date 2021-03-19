import os
import getpass
import tkinter as tk
from tkinter import filedialog, messagebox
from imps import SpeFiles
from imps import DatFiles

root = tk.Tk()
root.withdraw()

try:
    print('Создание папки на Рабочем Столе')
    os.mkdir(f'C:/Users/{getpass.getuser()}/Desktop/Kerber Importer Output')
except FileExistsError:
    print('Папка уже создана')
except OSError:
    pass

null_less_than_zero_ask = messagebox.askyesno('Обнуление', 'Обнулить отрицательные значения интенсивности?')

while True:
    if null_less_than_zero_ask:
        null_check = True
    else:
        null_check = False

    input_directory = filedialog.askdirectory(initialdir='/', title='Выберите папку с логами')  # C:/pyroot/Kerber Importer/test
    input_filename = input_directory.split('/')[-1]

    class Filenames:
        """This class is for sorting files in a directory to pick only .spe and .DAT files"""
        def __init__(self, directory):
            self.directory = directory
            self.walker = os.walk(directory, topdown=False)
            self.filename_list = None
            for folder, dirs, actual_name in self.walker:
                self.filename_list = actual_name
            for i in range(len(self.filename_list)):
                try:
                    chk_if_ends = self.filename_list[i].endswith(('.spe', '.DAT'))
                    if not chk_if_ends:
                        self.filename_list.pop(i)
                except IndexError:
                    break

        def get_lst(self):
            """Gets both .spe and .dat files as a list of file names"""
            return self.filename_list

        def get_lst_spe(self):
            """Gets .spe files as a list of file names"""
            not_an_spe_files = []
            for i in range(len(self.filename_list)):
                if not self.filename_list[i].endswith('.spe'):
                    not_an_spe_files.append(self.filename_list[i])
            spe_filename_list = [item for item in self.filename_list if item not in not_an_spe_files]
            return spe_filename_list

        def get_lst_dat(self):
            """Gets .dat files as a list of file names"""
            not_a_dat_files = []
            for i in range(len(self.filename_list)):
                if not self.filename_list[i].endswith('.DAT'):
                    not_a_dat_files.append(self.filename_list[i])
            dat_filename_list = [item for item in self.filename_list if item not in not_a_dat_files]
            return dat_filename_list


    walker_result = Filenames(input_directory)

    if walker_result.get_lst_spe():
        print('There are some .spe files')
        solve = SpeFiles(walker_result.get_lst_spe(), input_directory, null_check, input_filename)

    if walker_result.get_lst_dat():
        print('There are some .DAT files')
        solve = DatFiles(walker_result.get_lst_dat(), input_directory, null_check, input_filename)

# input('Для закрытия нажмите Enter')
