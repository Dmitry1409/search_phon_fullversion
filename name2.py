from openpyxl import Workbook
import os, search2, openpyxl


def main():
	file_name = input('Имя файла?: ') +'.xlsx'
	if search2.chek_file(file_name):
		dt = search2.read_exel_in_dict(file_name)
		print('В файле ', file_name, len(dt), ' номеров.')
		list_key = list(dt.keys())
		del_list = []
		for n in list_key:
			if n[2] != '9':
				del_list.append(n)
		if del_list:
			print(del_list)
			print('Найдено ', len(del_list), ' городских номеров.')
			if input('Удалить: ') == 'y':
				for n in del_list:
					del dt[n]
		del_list = []
		for k in dt:
			if 'натяжн' in dt[k]['Заголовок'].lower():
				del_list.append(k)
				del_list.append(dt[k]['Заголовок'])
		if del_list:
			print(del_list)
			print('Найдено ', int(len(del_list)/2), ' конкурентов')
			if input('Удалить: ') == 'y':
				i = 1
				for data in del_list:
					if i % 2 != 0:
						del dt[data]
					i+=1
		del_list = []
		feed_dt = search2.init_feedback()
		for k in dt:
			if k in feed_dt['black']:
				del_list.append(k)
				del_list.append(feed_dt['black'][k])
		
		if del_list:
			print('Найдено ', int(len(del_list)/2), ' номеров из черного списка')
			if input('Удалить: ') == 'y':
				i = 1
				for data in del_list:
					if i % 2 != 0:
						del dt[data]
					i+=1
		
		dt = testing_word(dt)

		rec_dt = {}
		for k in dt:
			rec_dt[k] = {}
			rec_dt[k]['Продавец'] = dt[k]['Продавец']

		resp = int(input('Количество файлов?: '))
		count = len(rec_dt) // resp
		l_key = list(rec_dt.keys())
		rec_list = []
		for i in range(1, resp + 1):
			if i == 1:
				rec_list.append(l_key[:count])
				continue
			if i == resp:
				rec_list.append(l_key[(i * count) - count :])
				break
			rec_list.append(l_key[(i * count) - count : i * count])
		
		for i, r_lict in enumerate(rec_list, 1):
			d = {}
			for k in r_lict:
				d[k] = rec_dt[k]
			search2.writin_new_exele(str(i)+'.xlsx', d, writin_keys=False)
		

	else:
		print('Not found', file_name)


def testing_word(dt):
	
	def read_txt(p):
		f = open(p)
		l = []
		l = f.read().splitlines()
		f.close()
		l2 = []
		for n in l:
			l2.append(n.title())
		return l2

	def searcher(client_n):
		client_n = client_n.title()
		fl = False
		for tn in test_name:
			if client_n == tn:
				dt[k]['Продавец'] = client_n
				fl = True
				break
		return fl

	test_name = read_txt('name.txt')

	for k in dt:
		seller = dt[k]['Продавец']
		cont_name = dt[k]['Контактное лицо']
		if type(seller) == int:
			seller = str(seller)
		if type(cont_name) == int:
			cont_name = str(cont_name)
		fl = False

		if seller is None:
			if cont_name is None:
				continue
			if ' ' in cont_name:
				l_name = cont_name.split(' ')
				for name in l_name:
					fl = searcher(name)
					if fl:
						break
				if fl:
					continue
				else:
					dt[k]['Продавец'] = None
					dt[k]['Контактное лицо'] = cont_name
			
			else:
				fl = searcher(cont_name)
	
				if fl:
					continue
				else:
					dt[k]['Продавец'] = None
					dt[k]['Контактное лицо'] = cont_name
		else:
			if ' ' in seller:
				l_name = seller.split(' ')
				for ln in l_name:
					fl = searcher(ln)
					if fl:
						break
				if fl:
					continue

				if cont_name != None:
					if ' ' in cont_name:
						l_name = cont_name.split(' ')
						for ln in l_name:
							fl = searcher(ln)
							if fl:
								break
						if fl:
							continue
						else:
							dt[k]['Продавец'] = None
							dt[k]['Контактное лицо'] = seller + ' ' + cont_name
					else:
						fl = searcher(cont_name)
						if fl: 
							continue
						else:
							dt[k]['Продавец'] = None
							dt[k]['Контактное лицо'] = seller + ' ' + cont_name
				else:
					dt[k]['Продавец'] = None
					dt[k]['Контактное лицо'] = seller
			
			else:
				fl = searcher(seller)
				if fl:
					continue
				else:
					if cont_name != None:
						if ' ' in cont_name:
							l_name = cont_name.split(' ')
							for ln in l_name:
								fl = searcher(ln)
								if fl:
									break
							if fl:
								continue
							else:
								dt[k]['Продавец'] = None
								dt[k]['Контактное лицо'] = seller + ' ' + cont_name

						else:
							fl = searcher(cont_name)
							if fl:
								continue
							else:
								dt[k]['Продавец'] = None
								dt[k]['Контактное лицо'] = seller + ' ' + cont_name
					else:
						dt[k]['Продавец'] = None
						dt[k]['Контактное лицо'] = seller
	return dt
