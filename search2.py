import os
import openpyxl
from openpyxl import Workbook
import pickle
import csv


BASA_FILE_NAME = "BASA"
DIRECTORY_AVITO_PARSER = 'архив ави парсер'
DIRECTORY_REPORT_ARCH = 'REPORTS/архив'
DIRECTORY_NEW_REPORT = 'REPORTS/new'

def main():
	basa = init()
	response = None
	while response != '0':
		print('''
			1 - Проверить файл
			3 - Поиск по номеру
			5 - Валидация имен
			6 - Добавить в feedback
			10 - Проверка категорий
			0 - Выйти''')
		response = input("Введите: ")

		if response == '1':
			res1(basa)
		if response == '2':
			res3(basa)
		if response == '5':
			import name
			name.main()

# def res3(basa):
# 	L=[]
# 	for k in basa

def res1(basa):
	file_name = input('Имя файла: ') + '.xlsx'
	if chek_file(file_name):
		file = read_exel_in_dict(file_name)
		print('В файле ', len(file), " номеров.")
		clean_dt = {}
		defolt = {}
		for k in file:
			fl = True
			for kb in basa:
				if kb == k:
					fl = False
					break
			if fl:
				clean_dt[k] = file[k]
			else:
				defolt[k] = file[k]
		if clean_dt:
			print('Найдено ', len(defolt), ' совподений из базы.')
			request = 'Записать '+ str(len(clean_dt)) + ' номеров в базу? '
			if input(request) == 'y':
				add_atribute(clean_dt)
				basa.update(clean_dt)
				save_pickle(BASA_FILE_NAME, basa)
				move_to(file_name, DIRECTORY_AVITO_PARSER)
	else:
		print('File ', file_name,' not found.')

def res2():
	dt_file = read_exel_in_dict('avito_92052.xlsx')
	dt_coord = {}
	dt_sity = {}
	for k in dt_file:
		if dt_file[k].get('Широта') != None:
			if is_valid_coord((dt_file[k]['Широта'], dt_file[k]['Долгота'])):
				dt_coord[k] = dt_file[k]
			else:
				dt_sity[k] = dt_file[k]
	dt_sity = validate_sity(dt_sity)
	print(len(dt_coord), len(dt_sity))
	dt_sity.update(dt_coord)
	writin_new_exele('12345.xlsx', dt_sity)

def move_to(file_name, to_directory, from_directory=None):
	if from_directory:
		os.rename(from_directory+'/'+file_name, to_directory +'/' + file_name)
		print('Move from: ',  from_directory+'/'+file_name)
		print('Move to: ', from_directory+'/'+file_name)
	else:
		os.rename(file_name, to_directory +'/' + file_name)
		print('Move from: ',  file_name)
		print('Move to: ', to_directory+'/'+file_name)

def save_pickle(file_name, file):
	with open(file_name, 'wb') as D:
		pickle.dump(file, D)
	print('Pickle dump ', os.getcwd()+'\\'+ file_name)

def add_atribute(dt):
	for n in dt:
		dt[n]['Отправлено'] = 0
		dt[n]['Доставлено'] = 0
		dt[n]['Дата отправки'] = []
		dt[n]['Текст сообщения'] = []

def validate_sity(dt):
	SITY_VAL = ["Нижний Новгород", "Дзержинск", "Бор", "Кстово" ,
				 "Павлово" , "Богородск" , "Городец", "Балахна", 
				 "Семенов", "Заволжье", "Чкаловск", "Володарск", "Ворсма",
				  "Ковернино", "Большое Козино", "Лукино", "Линда"]
	resp = {}
	for k in dt:
		for s in SITY_VAL:
			if s.lower() == dt[k]['Город'].lower():
				resp[k] = dt[k]
				break
	return resp

def writin_new_exele(path, rec_dt):
	
	def get_keys(dt):
		l_key = ['Телефон']
		for i in dt:
			for i2 in dt[i]:
				l_key.append(i2)
			break
		

		return l_key

	def dict_in_rec_list(rec_dt, rec_list):
		l_key = rec_list[0]
		for k in rec_dt:
			l_val = [k]
			for i in l_key[1:]:
				try:
					l_val.append(rec_dt[k][i])
				except:
					l_val.append('')
			rec_list.append(l_val)

		return rec_list

	def save_exele(path, list_val):
		wb = Workbook()
		ws = wb.active
		for row in list_val:
			ws.append(row)
		wb.save(path)
		print('complite: ', path)

	print('recording: ', path)
	l_key = get_keys(rec_dt)
	l_val_main = [l_key]
	l_val_main = dict_in_rec_list(rec_dt, l_val_main)
	save_exele(path, l_val_main)

def search_bulding_person(s_dt):
	main_w = ['ремонт', 'строительство', 'отделочники',
			 'cтроители', 'стройка', 'отделка', 'плиточник']
	dub_w = ['ремонт', 'отдел', 'строит', 'клей', 'штукату', 'каркас', 'плит']
	obj_w = ['офис', 'дом', 'кварт', 'работ', 'бригад', 'услуг', 'капит', 'плит',
			 'обо', 'гипс', 'внутр', 'помещ', 'ванн', 'туал', 'ключ', 'коттедж', 'унив']

	t_dt = {}
	for n in s_dt:
		title = s_dt[n].get('Заголовок')
		if title:
			title = title.strip()
			title = title.lower()
			fm = False
			for w in main_w:
				if w == title:
					t_dt[n] = s_dt[n]
					fm = True
					break
			if fm:
				continue
			fd = 0
			for w in dub_w:
				if w in title:
					fd +=1
			if fd > 1:
				t_dt[n] = s_dt[n]
				continue
			if fd == 1:
				fo = False
				for ow in obj_w:
					if ow in title:
						t_dt[n] = s_dt[n]
						fo = True
						break
	return t_dt

def sort_dic_for_rec(dt):
	map_dt = {}
	l_dict = []
	for k in dt:
		map_dt[k] = False
	for k in dt:
		if map_dt[k]:
			continue
		map_dt[k] = True
		prov = {}
		prov[k] = dt[k]
		for k2 in dt:
			if map_dt[k2]:
				continue
			# if prov[k]['Категории'] == dt[k2]['Категории']:
			if prov[k]['Город'] == dt[k2]['Город']:
				prov[k2] = dt[k2]
				map_dt[k2] = True
		l_dict.append(prov)
	l_dict.sort(key=len)
	# l_dict.reverse()
	return l_dict

def is_valid_coord(coord):
	noth_west = (56.794862261400546, 42.286376953125)
	noth_east = (56.96893619436121, 44.57153320312501)
	south_east = (56.105746831832064, 44.5330810546875)
	south_west = (55.54417295022067, 43.4619140625)
	if coord[0] < noth_west[0] and coord[1] > noth_west[1]:
		if coord[0] < noth_east[0] and coord[1] < noth_east[1]:
			if coord[0] > south_west[0] and coord[1] > south_west[1]:
				if coord[0] > south_east[0] and coord[1] < south_east[1]:
					return True
	return False

def init():
	def greate_report(csv_list):		
		defolt = []
		for row in range(1, len(csv_list)):
			list_data = csv_list[row][0].split(';')
			if basa.get(list_data[1]):
				if list_data[4]:
					basa[list_data[1]]['Отправлено'] += int(list_data[4])
					basa[list_data[1]]['Доставлено'] += int(list_data[5])
					if basa[list_data[1]]['Дата отправки']:
						basa[list_data[1]]['Дата отправки'].append(list_data[3][1:-1])
						basa[list_data[1]]['Текст сообщения'].append(list_data[7][1:-1])
					else:
						basa[list_data[1]]['Дата отправки'] = [list_data[3][1:-1]]
						basa[list_data[1]]['Текст сообщения'] = [list_data[7][1:-1]]
			else:
				defolt.append(list_data)
		print('Complite')
		if defolt:
			print('Not found ', len(defolt), ' numbers.')






	if chek_file(BASA_FILE_NAME):
		with open(BASA_FILE_NAME, 'rb') as D:
			basa = pickle.load(D)
		print('База загружена. ', len(basa), ' номеров.')
	else:
		if input('База не найдена, создать?: ') == 'y':
			print("Создается новая база")
			basa = {}
			path_list_xl = get_path(DIRECTORY_AVITO_PARSER)
			for p in path_list_xl:
				dt = read_exel_in_dict(p)
				basa.update(dt)
			add_atribute(basa)
			path_report = get_path(DIRECTORY_REPORT_ARCH)
			for p in path_report:
				csv_list = read_csv(p)
				
				# Там есть два файла которие отличаются от все остальных
				# для них своя реализация
				if len(csv_list[2][0].split(';'))<2:
					defolt = []
					i = 0
					for row in range(1, int(len(csv_list[1:])/4)+1):
						list_data = csv_list[row+i][0].split(';')
						if basa.get(list_data[1]):
							if list_data[4]:
								basa[list_data[1]]['Отправлено'] += int(list_data[4])
								basa[list_data[1]]['Доставлено'] += int(list_data[5])
								if basa[list_data[1]]['Дата отправки']:
									basa[list_data[1]]['Дата отправки'].append(list_data[3][1:-1])
									basa[list_data[1]]['Текст сообщения'].append(list_data[7][1:] + csv_list[row+i+1][0] + csv_list[row+i+2][0] + csv_list[row+i+3][0])
								else:
									basa[list_data[1]]['Дата отправки'] = [list_data[3][1:-1]]
									basa[list_data[1]]['Текст сообщения'] = [list_data[7][1:] + csv_list[row+i+1][0] + csv_list[row+i+2][0] + csv_list[row+i+3][0]]
						else:
							defolt.append(list_data)
						i+= 3
					print('Complite')
					if defolt:
						print('Not found ', len(defolt), ' numbers.')
				else:
					greate_report(csv_list)
			save_pickle(BASA_FILE_NAME, basa)
			print('Создалась новая база. ', len(basa), ' номеров.')

	# проверяем новые отчеты
	path_report = os.listdir(DIRECTORY_NEW_REPORT)
	if path_report:
		print(path_report)
		if input('Выполнить?: ') == 'y':
			for p in path_report:
				csv_list = read_csv('REPORTS/new/'+p)
				greate_report(csv_list)
				move_to(p, DIRECTORY_REPORT_ARCH, form_directory=DIRECTORY_NEW_REPORT)				

	return basa

def read_csv(path):
	print('Load report: ', path)
	csv_list = []
	with open(path, newline = '', encoding= 'utf-8') as f:
		reader = csv.reader(f)
		for row in reader:
			csv_list.append(row)
	return csv_list

def save_exele(path, list_val):
	wb = Workbook()
	ws = wb.active
	for row in list_val:
		ws.append(row)
	wb.save(path)

def read_exel_in_dict(path):
	print("load: ", path)
	CHAR = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j' , 'k', 'l', 'm', 'n',
			'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', 'aa', 'ab',
			'ac', 'ad', 'ae', 'af', 'ag', 'ah', 'ai', 'aj', 'ak', 'al', 'am', 'an',
			'ao', 'ap', 'aq', 'ar', 'as']
	exel = openpyxl.load_workbook(path)
	sheet = exel.active
	cols = sheet.max_column
	rows = sheet.max_row
	l_key = []	
	for s in range(0, cols):
		k = CHAR[s] + '1'
		l_key.append(sheet[k].value)
	dt = {}	
	for r in range(2, rows +1):
		phone = sheet[CHAR[0] + str(r)].value 
		dt[phone] = {}
		for i in range(1, len(l_key)):
			dt[phone][l_key[i]] = sheet[CHAR[i] + str(r)].value
	return dt

def get_path(directory = None, type_fl = None):
	if type_fl:
		pint('raliase func')
	else:
		L = os.listdir(directory)
		path_list = []
		for e in L:
			path_list.append(directory+'/'+e)
		return path_list

def chek_file(path):
	l = path.split("/")
	if len(l)>1:
		dir_ = "/".join(l[:-1])
	else:
		dir_ = None
	list_dir = os.listdir(dir_)
	flag = False
	for f in list_dir:
		if f == l[-1]:
			flag = True
			break
	return flag


if __name__ == '__main__':
	main()