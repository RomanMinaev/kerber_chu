import xlsxwriter
import uuid
import itertools
import io
import getpass
from string import ascii_uppercase


class DatFiles:
	"""Class for importing .DAT file data to Excel book"""
	dat_line_numbers = list(range(15, 2016))

	def __init__(self, file_lst, directory, null_check, filename):
		self.file_lst = file_lst
		self.directory = directory
		self.null_check = null_check
		self.filename = filename
		print('Создание книги Excel...')
		workbook = xlsxwriter.Workbook(
			f'C:/Users/{getpass.getuser()}/Desktop/Kerber Importer Output/DAT_{filename}_{str(uuid.uuid4())[0:4]}.xlsx')
		worksheet1 = workbook.add_worksheet('Без вычета фона')
		worksheet2 = workbook.add_worksheet('С вычетом фона')
		worksheets = [worksheet1, worksheet2]
		for worksheet in worksheets:
			worksheet.set_column('A:A', 19)
			worksheet.set_column('B:B', 5)
			worksheet.write('A1', 'Коэфф подвижности')
			worksheet.write('B1', 'Фон')
		print('Книга Excel создана.')

		def iter_excel_columns():  # Connects 2 characters to form 'AA' format
			for char1, char2 in itertools.product(first_char, ascii_uppercase):
				yield char1 + char2

		excel_charlist = []  # Excel charlist generator
		first_char = [''] + list(ascii_uppercase)
		for s in iter_excel_columns():
			excel_charlist.append(s)
			if s == 'dd':  # Charlist is limited to DD. Can change it.
				break
		excel_charlist.pop(0)  #
		excel_charlist.pop(0)  # 'A', 'B' chars are not needed.

		# Opens file, takes data, closes file
		# This part is for first file only (background signal file)
		dat_file = io.open(f'{self.directory}/{self.file_lst[0]}', encoding='utf-8', errors='ignore')
		print(f'Обработка файла {self.file_lst[0]}')
		read_result = dat_file.readlines()
		mob_time_raw = read_result.pop(1)
		mob_time = mob_time_raw.split()
		mob_time = float(mob_time.pop(-1))  # Gets mob_time value as float

		def write_lines(line_numbers):  # Takes strings needed from file
			searched_lines = []
			for line in line_numbers:
				searched_lines.append(read_result[line])
			return searched_lines

		raw_coords = write_lines(self.dat_line_numbers)

		def grab_coords(raw_coords_list):  # Лист координат из листа строк .DAT
			fixed_coord_list = []
			for i in range(len(raw_coords_list)):
				raw_coord = raw_coords_list[i]
				listed_raw_coord = raw_coord.split()
				ok_coord = int(listed_raw_coord.pop(-1))
				fixed_coord_list.append(ok_coord)
			return fixed_coord_list

		good_cords = grab_coords(raw_coords)  # Лист координат

		coord_number_list = []  # Takes drift time from .DAT
		for i in range(len(raw_coords)):
			raw_coord = raw_coords[i]
			listed_raw_coord = raw_coord.split()
			coord_number = int(listed_raw_coord.pop(0))
			coord_number_list.append(coord_number)

		def drift_time_calc(mob_coeff):  # Mobility coords list
			drift_time_list_local = []
			for i in range(len(coord_number_list)):
				coord = mob_coeff / coord_number_list[i]
				drift_time_list_local.append(coord)
			return drift_time_list_local

		drift_time_list = drift_time_calc(mob_time)

		if null_check:  # Проверка нужно ли обнулять значения
			print('Обнуление значений интенсивности')
			for i in range(len(good_cords)):
				if good_cords[i] < 0:
					good_cords[i] = 0

		background_coords = good_cords  # Saving background coords

		for z in range(len(good_cords)):
			for worksheet in worksheets:
				worksheet.write('A' + str(z + 2), drift_time_list[z])
				worksheet.write('B' + str(z + 2), background_coords[z])

		self.file_lst.pop(0)  # Pop background signal file
		# !!!
		# !!!
		for file_number in range(len(self.file_lst)):
			print(f'Обработка файла {self.file_lst[file_number]}')
			for worksheet in worksheets:
				worksheet.set_column(f'{excel_charlist[file_number]}:{excel_charlist[file_number]}', 5)

			dat_file = io.open(f'{self.directory}/{self.file_lst[file_number]}', encoding='utf-8', errors='ignore')
			read_result = dat_file.readlines()
			mob_time_raw = read_result.pop(1)
			mob_time = mob_time_raw.split()
			mob_time.pop(-1)
			raw_coords = write_lines(self.dat_line_numbers)
			good_cords = grab_coords(raw_coords)

			def background_excluder(bg_excluded_coords_pol, background_coords_pol, coords_pol):  # Вычитатель фона
				for coord in range(len(coords_pol)):
					bg_excluded_coords_pol.append(coords_pol[coord] - background_coords_pol[coord])
				return bg_excluded_coords_pol

			bg_excluded_coords = background_excluder([], background_coords, good_cords)

			if null_check:
				print('Обнуление значений интенсивности')
				for i in range(len(good_cords)):
					if good_cords[i] < 0:
						good_cords[i] = 0
					if bg_excluded_coords[i] < 0:
						bg_excluded_coords[i] = 0

			for z in range(len(good_cords)):
				worksheet1.write(excel_charlist[file_number] + str(z + 2), good_cords[z])
				worksheet2.write(excel_charlist[file_number] + str(z + 2), bg_excluded_coords[z])

			worksheet1.write(excel_charlist[file_number] + str(1), f'Cнимок {file_number}')
			worksheet2.write(excel_charlist[file_number] + str(1), f'Cнимок {file_number}')

		print('Закрытие книги Excel')
		workbook.close()


class SpeFiles:
	"""Class for importing .spe file data to Excel book"""
	first_line_numbers = list(range(12, 2013))  # Numbers of strings with coords
	second_line_numbers = list(range(2014, 4015))

	def __init__(self, file_lst, directory, null_check, filename):
		self.file_lst = file_lst
		self.directory = directory
		self.null_check = null_check
		self.filename = filename
		print('Создание книги Excel...')
		workbook = xlsxwriter.Workbook(
			f'C:/Users/{getpass.getuser()}/Desktop/Kerber Importer Output/SPE_{filename}_{str(uuid.uuid4())[0:4]}.xlsx')
		worksheet1 = workbook.add_worksheet('P_полярность')
		worksheet2 = workbook.add_worksheet('N_полярность')
		worksheet3 = workbook.add_worksheet('P_вычет фона')
		worksheet4 = workbook.add_worksheet('N_вычет фона')
		worksheets = [worksheet1, worksheet2, worksheet3, worksheet4]
		for worksheet in worksheets:
			worksheet.set_column('A:A', 19)
			worksheet.set_column('B:B', 5)
			worksheet.write('A1', 'Коэфф подвижности')
			worksheet.write('B1', 'Фон')
		print('Книга Excel создана.')

		def iter_excel_columns():  # Connects 2 characters to form 'AA' format
			for char1, char2 in itertools.product(first_char, ascii_uppercase):
				yield char1 + char2

		excel_charlist = []  # Excel charlist generator
		first_char = [''] + list(ascii_uppercase)
		for s in iter_excel_columns():
			excel_charlist.append(s)
			if s == 'dd':  # Charlist is limited to DD. Can change it.
				break
		excel_charlist.pop(0)  #
		excel_charlist.pop(0)  # 'A', 'B' chars are not needed.

		# Opens file, takes data, closes file
		# This part is for first file only (background signal file)
		spe_file = io.open(f'{self.directory}/{self.file_lst[0]}', encoding='utf-8', errors='ignore')
		read_result = spe_file.readlines()

		def write_lines(line_numbers):  # Takes strings needed from file
			searched_lines = []
			for line in line_numbers:
				searched_lines.append(read_result[line])
			return searched_lines

		pol_one = write_lines(self.first_line_numbers)
		pol_two = write_lines(self.second_line_numbers)
		spe_file.close()
		pol_one_props = pol_one.pop(0)  # Cuts out property line
		pol_two_props = pol_two.pop(0)

		def str_to_int(coord_list):
			for coord in range(len(coord_list)):
				coord_list[coord] = int(coord_list[coord])
			return coord_list

		pol_one_int = str_to_int(pol_one)  # Coords
		pol_two_int = str_to_int(pol_two)

		def mob_coeff_grabber(props):  # Grabs mob_time float value from pol_one_props
			raw_str = props.pop(7)
			cut = raw_str.split(':')
			return float(cut.pop(1))

		if 'delay_p' in pol_one_props:  # Renames variables depend on polarity
			print(f'Обработка файла {self.file_lst[0]}, p-n')
			props_p = pol_one_props.split(',')
			props_n = pol_two_props.split(',')
			coords_p = pol_one_int
			coords_n = pol_two_int
			mob_coeff_n = mob_coeff_grabber(props_n)
			mob_coeff_p = mob_coeff_grabber(props_p)

		else:
			print(f'Обработка файла {self.file_lst[0]}, n-p')
			props_p = pol_one_props.split(',')
			props_n = pol_two_props.split(',')
			coords_n = pol_one_int
			coords_p = pol_two_int
			mob_coeff_n = mob_coeff_grabber(props_n)
			mob_coeff_p = mob_coeff_grabber(props_p)

		coord_number_list_p = list(range(800, len(coords_p) + 800))
		coord_number_list_n = list(range(600, len(coords_n) + 600))

		def drift_time_calc(mob_coeff, pol):  # mob_coeff coord list
			drift_time_list = []
			if pol == 'p':
				for number in coord_number_list_p:
					coord = mob_coeff / (number * 0.025)
					drift_time_list.append(coord)
				return drift_time_list
			if pol == 'n':
				for number in coord_number_list_n:
					coord = mob_coeff / (number * 0.025)
					drift_time_list.append(coord)
				return drift_time_list

		drift_time_list_p = drift_time_calc(mob_coeff_p, 'p')  # Drift time calc (x axys)
		drift_time_list_n = drift_time_calc(mob_coeff_n, 'n')

		if null_check:  # Nullify negative numbers
			print('Обнуление значений интенсивности')
			for i in range(len(coords_n)):
				if coords_n[i] < 0:
					coords_n[i] = 0
			for i in range(len(coords_p)):
				if coords_p[i] < 0:
					coords_p[i] = 0

		background_coords_p = coords_p  # Saving background coords
		background_coords_n = coords_n

		for z in range(len(coords_p)):  # Writing coords in Excel book
			worksheet1.write('A' + str(z + 2), drift_time_list_p[z])
			worksheet2.write('A' + str(z + 2), drift_time_list_n[z])
			worksheet3.write('A' + str(z + 2), drift_time_list_p[z])
			worksheet4.write('A' + str(z + 2), drift_time_list_n[z])

			worksheet1.write('B' + str(z + 2), background_coords_p[z])
			worksheet2.write('B' + str(z + 2), background_coords_n[z])
			worksheet3.write('B' + str(z + 2), background_coords_p[z])
			worksheet4.write('B' + str(z + 2), background_coords_n[z])

	# Now everything besides background signal:
		self.file_lst.pop(0)

		def background_excluder(bg_excluded_coords_pol, background_coords_pol, coords_pol):  # Вычитатель фона
			for coord in range(len(coords_pol)):
				bg_excluded_coords_pol.append(coords_pol[coord] - background_coords_pol[coord])
			return bg_excluded_coords_pol

		for file_number in range(len(self.file_lst)):
			for worksheet in worksheets:
				worksheet.set_column(f'{excel_charlist[file_number]}:{excel_charlist[file_number]}', 5)

			spe_file = io.open(f'{self.directory}/{self.file_lst[file_number]}', encoding='utf-8', errors='ignore')
			read_result = spe_file.readlines()
			pol_one = write_lines(self.first_line_numbers)
			pol_two = write_lines(self.second_line_numbers)

			spe_file.close()

			pol_one_props = pol_one.pop(0)  # Cuts out property line
			pol_two.pop(0)
			pol_one_int = str_to_int(pol_one)  # Coords
			pol_two_int = str_to_int(pol_two)

			if 'delay_p' in pol_one_props:  # Renames variables depend on polarity
				print(f'Обработка файла {self.file_lst[file_number]}, p-n')
				coords_p = pol_one_int
				coords_n = pol_two_int

			else:
				print(f'Обработка файла {self.file_lst[file_number]}, n-p')
				coords_n = pol_one_int
				coords_p = pol_two_int

			coord_number_list_p = list(range(800, len(coords_p) + 800))
			coord_number_list_n = list(range(600, len(coords_n) + 600))

			bg_excluded_coords_p = background_excluder([], background_coords_p, coords_p)
			bg_excluded_coords_n = background_excluder([], background_coords_n, coords_n)

			if null_check:
				print('Обнуление значений интенсивности')
				for i in range(len(coords_n)):
					if coords_n[i] < 0:
						coords_n[i] = 0
					if bg_excluded_coords_n[i] < 0:
						bg_excluded_coords_n[i] = 0
				for i in range(len(coords_p)):
					if coords_p[i] < 0:
						coords_p[i] = 0
					if bg_excluded_coords_p[i] < 0:
						bg_excluded_coords_p[i] = 0

			for z in range(len(coords_p)):  # Запись координат в оба листа
				worksheet1.write(excel_charlist[file_number] + str(z + 2), coords_p[z])
				worksheet2.write(excel_charlist[file_number] + str(z + 2), coords_n[z])
				worksheet3.write(excel_charlist[file_number] + str(z + 2), bg_excluded_coords_p[z])
				worksheet4.write(excel_charlist[file_number] + str(z + 2), bg_excluded_coords_n[z])

			for worksheet in worksheets:
				worksheet.write(excel_charlist[file_number] + str(1), f'Cнимок {file_number}')

		print('Закрытие книги Excel')
		workbook.close()
