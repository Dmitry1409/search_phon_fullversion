import openpyxl
from openpyxl import Workbook


def strip_str(n):
	n = n.strip()
	n = n.strip('.')
	n = n.title()
	return n
def read_txt(f):
	f = open(f)
	l = []
	l = f.read().splitlines()
	f.close()
	return l
def read_xl(f):
	exel = openpyxl.load_workbook(f)
	sheet = exel.active
	list_num = []
	for row in sheet['a']:
		list_num.append(row.value)
	nam = []
	for row in sheet['b']:
		nam.append(row.value)
	nam2 = []
	for row in sheet['c']:
		nam2.append(row.value)
	return list_num, nam, nam2

def del_8(list_num, nam, nam2):
	lnum = []
	name = []
	name2 = []
	i = 0
	dl = []
	for n in list_num:
		if n[2] != '9':
			dl.append(n)
			dl.append(nam[i])
			dl.append(nam2[i])
		else:
			lnum.append(n)
			name.append(nam[i])
			name2.append(nam2[i])
		i +=1
	print('Удалено : ', round(len(dl)/3),  dl)
	return lnum, name, name2


def str_stroke(name, name2):
	i = 0
	for n in name:
		if n is None:
			i+=1
			continue
		n = str(n)
		n = strip_str(n)
		name[i] = n
		i +=1
	i = 0
	for n in name2:
		if n is None:
			i += 1
			continue
		n = str(n)
		n = strip_str(n)
		name2[i] = n
		i+=1
	return name, name2

def gr_rec_list(list_num, nam, nam2):
	i =0
	l_writ = []
	l_writ2 = []
	for n in list_num:
		l_writ = [list_num[i], nam[i], nam2[i]]
		l_writ2.append(l_writ)
		i +=1
	return l_writ2

def writ_xl(rec_l):
	wb = Workbook()
	ws = wb.active
	for row in rec_l:
		ws.append(row)
	res = input('Имя нового файла? : ')
	res = res + '.xlsx'
	wb.save(res)
	print('Comlete')

def testing_word(name, name2, test_name):
	i = 0
	for n in name:
		if n is None:
			if name2[i] is None:
				i+=1
				continue
			if ' ' in name2[i]:
				s = name2[i].split(' ')
				fl = False
				for sn in s:
					sn = sn.title()
					for tw in test_name:
						if tw == sn:
							name[i] = sn
							js = ' '.join(s)
							name2[i] = None
							fl = True
							break
					if fl:
						i+=1
						break
				if fl == False:
					i+=1

			else:
				fl = False
				for tw in test_name:
					if name2[i] == tw:
						name[i] = name2[i]
						name2[i] = None
						i+=1
						fl = True
						break
				if fl == False:
					i+=1
		else:
			if ' ' in n:
				s = n.split(' ')
				fl = False
				for sn in s:
					sn = sn.title()
					for tw in test_name:
						if sn == tw:
							fl = True
							name[i] = sn
							name2[i] = None
							i += 1
							break
					if fl:
						break
				if fl == False:
					if name2[i] != None:
						if ' 'in name2[i]:
							s = name2[i].split(' ')
							fl = False
							for sn in s:
								sn = sn.title()
								for tw in test_name:
									if tw == sn:
										fl = True
										name[i] = sn										
										name2[i] = None
										break
								if fl:
									i +=1
									break
							if fl == False:
								name2[i] = name2[i] + ' '+ name[i]
								name[i] = None
								i+=1
						else:
							fl = False
							for tw in test_name:
								if tw == name2[i]:
									name[i] = name2[i]
									name2[i] = None
									fl = True
									i+=1
									break
							if fl == False:
								name2[i] = name2[i] + ' ' + name[i]
								name[i] = None
								i+=1
					else:
						name[i] = None
						name2[i] = n
						i+=1 
			else:
				fl = False
				for tw in test_name:
					if n == tw:
						fl  = True
						name2[i] = None
						i+=1
						break
				if fl == False:
					if name2[i] == None:
						name2[i] = n
						name[i] = None
						i+=1
					else:
						if ' ' in name2[i]:
							s = name2[i].split(' ')
							fl = False
							for sn in s:
								sn = sn.title()
								for tw in test_name:
									if sn == tw:
										name[i] = sn
										name2[i] = None
										fl = True
										i+=1
										break
								if fl:
									break
							if fl == False:
								name2[i] = name2[i] +' '+ name[i]
								name[i] = None
								i+=1
						else:
							fl = False
							for tw in test_name:
								if name2[i] == tw:
									name[i] = tw
									name2[i] = None
									fl = True
									i+=1
									break
							if fl == False:
								name2[i] = n + ' ' + name2[i]
								name[i] = None
								i+=1
	return name, name2
					
def main():
	nexel = input('Введите .xlsx: ')
	nexel = nexel + '.xlsx'
	num, name, name2 = read_xl(nexel)
	num, name, name2 = del_8(num, name, name2)
	name, name2 = str_stroke(name, name2)
	test_name = read_txt('name.txt')
	name, name2 = testing_word(name, name2, test_name)
	rec_l = gr_rec_list(num, name, name2)
	writ_xl(rec_l)



