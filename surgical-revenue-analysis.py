import numpy as np
import pandas as pd
import re
import os
import sqlalchemy
import requests
import xlsxwriter
import telebot
import getpass
import platform
import time
import threading
from dateutil import parser
from datetime import datetime, timedelta
from fast_bitrix24 import Bitrix
from more_itertools.recipes import unique
from sqlalchemy import create_engine
from dotenv import dotenv_values
from mysql.connector import Error
import plotly
import plotly.graph_objs as go
import plotly.express as px
from plotly.subplots import make_subplots
import dash
import dash_bootstrap_components as dbc
from dash import dash_table
from dash import dcc
from dash import html
from dash.dependencies import Input, Output

# Паттерны
re_1 = r'[^0-9,.;/]' # Регулярное выражение для отсева букв и знаков
pd.set_option('display.max_columns', None)
pd.set_option('mode.chained_assignment', None)
lock = threading.Lock() # Локер процессов
# Создание коннекторов данных
datename = datetime.now().strftime('%d.%m %H.%M.%S') # Время создания файла
bot = telebot.TeleBot(os.getenv("TELEGRAM_TOKEN"))
engine = create_engine(os.getenv("DB_URL"))
kategory_list = [10003976, 10003987, 10003391, 10003988, 10003366, 10003986, 10003403, 10007481, 10002913, 10005920, 10005921, 10005922, 10003982, 10003436, 10003981, 10003469, 10003980, 10003979, 10003487, 10005932, 10003507, 10003977, 10003519, 10003811, 10006088, 10005928, 10003841, 10003842, 10003856, 10006089]
start_time = time.time()
def Input_lag():
    global start_date, end_date
    lock.acquire()
    while True:
        try:
            start_date = input("Введите дату начала в формате dd.mm.yy:\n")
            start_date = pd.to_datetime(datetime.strptime(start_date, '%d.%m.%y'))
            print("Начальная дата: " + str(start_date)); break
        except: print("Введённый параметр не соответствует допустимому формату даты - попробуйте написать дату по другому")
    while True:
        try:
            end_date = input("Теперь введите конечную дату (обратите внимание, что по умолчанию дате присвоится время 00:00:00, т.е. если вы хотите посчитать, например, до 10.10.24 ВКЛЮЧИТЕЛЬНО, то следует написать 11.10.24):\n")
            if end_date == "": end_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0); print(end_date)
            else: end_date = pd.to_datetime(datetime.strptime(end_date, '%d.%m.%y'))
            if end_date < start_date: print("Конечная дата меньше начальной, так быть не должно, попробуйте снова"); continue
            print("Конечная дата: " + str(end_date)); break
        except: print("Введённый параметр не соответствует допустимому формату даты - попробуйте написать дату по другому")
    print("Все опциональные параментры введены пользователем. Ожидаем дозагрузки данных и начала расчёта...")
    lock.release()
def GetSQL():
	global start_time
	# Запрос наименований услуг
	lightquery_WSCHEMA = "SELECT schid, schname, parentschid FROM WSCHEMA"  # Описание услуг из прайс-листа
	WSCHEMA = pd.read_sql(lightquery_WSCHEMA, engine) # dict_WSCHEMA
	WSCHEMA_filter = WSCHEMA[(WSCHEMA['parentschid'].isin(list(filter(None, kategory_list))))]
	WSCHEMA_list = list(filter(None, WSCHEMA_filter['schid'].tolist()))
	dict_WSCHEMA_kat =  dict(WSCHEMA[['schid', 'schname']].values) # Словарь соотвествий ID услуги и наименования
	dict_WSCHEMA =  dict(WSCHEMA_filter[['schid', 'schname']].values) # Словарь соотвествий ID услуги и наименования
	dict_WSCHEMA_ktegory =  dict(WSCHEMA[['schid', 'parentschid']].values) # Словарь соотвествий ID услуги и наименования
	WSCHEMA_filter["Наименование услуги"] = WSCHEMA_filter['schid'].replace(dict_WSCHEMA)
	# Запрос действующих врачей хирургов
	lightquery_BI_DOCTORS = "SELECT dcode, dname, depnum, lockdate from BI_DOCTORS" # Доктора WHERE lockdate IS NULL - без уволенных
	BI_DOCTORS = pd.read_sql(lightquery_BI_DOCTORS, engine)  # dict_BI_DOCTORS_name, dict_BI_DOCTORS_dep
	doctor_list = BI_DOCTORS[(BI_DOCTORS['depnum'] == 120363) & (BI_DOCTORS['lockdate'].isnull())] # Создаём список действующих врачей хирургов
	dict_BI_DOCTORS_name = dict(BI_DOCTORS[['dcode', 'dname']].values) # Словарь соответствий ID сотрудника и полное имя
	doctor_list = list(filter(None, doctor_list['dname'].tolist())) # Конвертируем в список
	# Запрос кодов лечений по подходящим параметрам
	lightquery_TREAT = "SELECT orderno, treatdate, treatcode, dcode, pcode FROM TREAT WHERE (depnum = '120363') AND (kateg = 1)" # Лечения (отчёт по приёмам), AND (kateg = 1) - только наличный расчёт
	TREAT = pd.read_sql(lightquery_TREAT, engine) # Отчёт по приёмам хирургов
	lock.acquire()
	TREAT = TREAT[(TREAT['treatdate'] >= start_date) & (TREAT['treatdate'] < end_date)]
	dict_TREAT_pacient = dict(TREAT[['orderno', 'pcode']].values)
	dict_TREAT_doctor = dict(TREAT[['orderno', 'dcode']].values)
	dict_TREAT_data = dict(TREAT[['orderno', 'treatdate']].values)
	dict_TREAT_code = dict(TREAT[['orderno', 'treatcode']].values)
	TREAT_list = list(filter(None, TREAT['orderno'].tolist())) # Конвертируем в список
	# Запрос сделок с фильтрам по кодам лечений
	lightquery_ORDERDET = "SELECT orderno, schcode, schcount, schprice FROM ORDERDET" # Сделки с суммами по категориям
	ORDERDET = pd.read_sql(lightquery_ORDERDET, engine) # dict_ORDERDET
	ORDERDET = ORDERDET[(ORDERDET['orderno'].isin(TREAT_list)) & (ORDERDET['schcode'].isin(WSCHEMA_list))]
	print(f"Все базы загружены в буфер за :{(time.time() - start_time):.2f}"); start_time = time.time()
	# Преобразование таблицы
	ORDERDET['Стоимость для пациента'] = ORDERDET['schprice'] * ORDERDET['schcount'] # Расчитать стоимость сделки (стоимость * количество)
	ORDERDET['Стоимость для пациента, с вычтенными бонусами'] = ORDERDET['schprice'] * ORDERDET['schcount'] # Расчитать стоимость сделки (стоимость * количество)
	ORDERDET['Категория работы'] = ORDERDET['schcode'].replace(dict_WSCHEMA_ktegory) # Присвоить сделке категорию работы
	ORDERDET['Категория работы'] = ORDERDET['Категория работы'].replace(dict_WSCHEMA_kat) # Присвоить сделке наименование работы
	ORDERDET['Выполненная работа'] = ORDERDET['schcode'].replace(dict_WSCHEMA) # Присвоить сделке наименование работы
	ORDERDET['Пациент'] = ORDERDET['orderno'].replace(dict_TREAT_pacient) # Присвоить сделке имя пациента
	ORDERDET['Доктор'] = ORDERDET['orderno'].replace(dict_TREAT_doctor) # Присвоить сделке доктора
	ORDERDET['Доктор'] = ORDERDET['Доктор'].replace(dict_BI_DOCTORS_name) # Присвоить коду доктора имя доктора
	ORDERDET['Дата операции'] = ORDERDET['orderno'].replace(dict_TREAT_data) # Присвоить сделке пациента
	ORDERDET['Код лечения'] = ORDERDET['orderno'].replace(dict_TREAT_code) # Присвоить сделке пациента
	ORDERDET = ORDERDET.sort_values(by=['Дата операции'], ascending=False) # Сортировка по дате назначения
	ORDERDET["Полная дата"] = ORDERDET["Дата операции"].dt.to_period("D") # Конвертация даты в обычные даты без времени
	ORDERDET["Месяц, год"] = ORDERDET["Дата операции"].dt.to_period("M") # Конвертация даты в месяцы
	ORDERDET["Квартал, год"] = ORDERDET["Дата операции"].dt.to_period("Q") # Конвертация даты в кварталы
	ORDERDET["Год"] = ORDERDET["Дата операции"].dt.to_period("Y") # Конвертация даты в годы
	Make_base(ORDERDET, WSCHEMA_filter)
def Make_base(base, base_operations):
	DataList = pd.DataFrame(pd.date_range(min(start_date, end_date), max(start_date, end_date)).tolist()); DataList.columns =  ['Дата'] # Сборка списка дат
	Data_list_D = DataList['Дата'].dt.to_period("D"); Data_list_M = DataList['Дата'].dt.to_period("M"); Data_list_Q = DataList['Дата'].dt.to_period("Q"); Data_list_Y = DataList['Дата'].dt.to_period("Y") # Форматирование списков дат в дни, месяцы, кварталы и годы
	tables_money = []; tables_services = []; tables_pacients = []; tables_main = []; tables_treats = [] # Объявление циклов для хранения массивов цикла for
	###################################################################################################################################################################################################
	# Результаты работы по периодам
	for Data_list, Period in zip([Data_list_D, Data_list_M, Data_list_Q, Data_list_Y], ['Полная дата', 'Месяц, год', 'Квартал, год', 'Год']):
		Money = pd.DataFrame(columns=['Доктор']); Services = pd.DataFrame(columns=['Доктор']); Pacients = pd.DataFrame(columns=['Доктор']); Treats = pd.DataFrame(columns=['Доктор']); Main_list = pd.DataFrame(columns=['Доктор'])
	# Выручка хир. части ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		for date in Data_list.unique().tolist():
			collect = base[(base[Period] == date)]
			collect_cost = collect.groupby('Доктор').agg({'Стоимость для пациента': ['sum']}).reset_index(); collect_cost.columns = ['Доктор', date] # Группировка с суммами лечений по докторам
			collect_cost = pd.concat([collect_cost, pd.DataFrame({'Доктор': 'Вся выручка хир. части:', date: collect_cost[date].sum()}, index=[0])], axis=0) # Добавление итогов по отделению
			collect_cost = pd.DataFrame({'Доктор': list(filter(None, base['Доктор'].unique().tolist()))}).merge(collect_cost, on='Доктор', how='outer') # Объединяем таблицу со списком операторов (для одинаковой структуры данных)
			collect_cost = pd.concat([collect_cost, pd.DataFrame({'Доктор': np.nan, date: np.nan}, index=[0])], axis=0) # Добавление промежуточной строки между блоками отчёта
			Money = Money.merge(collect_cost, on='Доктор', how='outer')
		collect_cost = base.groupby('Доктор').agg({'Стоимость для пациента': ['sum']}).reset_index(); collect_cost.columns = ['Доктор', "За весь период"]
		collect_cost = pd.concat([collect_cost, pd.DataFrame({'Доктор': 'Вся выручка хир. части:', 'За весь период': collect_cost["За весь период"].sum()}, index=[0])], axis=0) # Добавление итогов по отделению
		Money = Money.merge(collect_cost, on='Доктор', how='outer')
		Money = Money.sort_values(by=['За весь период'], ascending=False) # Сортировка по сумме выручки
	# Оказанные услуги --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		for date in Data_list.unique().tolist():
			collect = base[(base[Period] == date)]
			collect_count = collect.groupby('Доктор').agg(unique_count=('Пациент', 'nunique')).reset_index(); collect_count.columns = ['Доктор', date] # Группировка с количеством уникальных пациентов по докторам
			collect_count = pd.concat([collect_count, pd.DataFrame({'Доктор': 'Все уникальные клиенты для доктора:', date: collect_count[date].sum()}, index=[0])], axis=0) # Добавление итогов по отделению
			collect_count = pd.DataFrame({'Доктор': list(filter(None, base['Доктор'].unique().tolist()))}).merge(collect_count, on='Доктор', how='outer') # Объединяем таблицу со списком операторов (для одинаковой структуры данных)
			Services = Services.merge(collect_count, on='Доктор', how='outer')
		collect_count = base.groupby('Доктор').agg(unique_count=('Пациент', 'nunique')).reset_index(); collect_count.columns = ['Доктор', "За весь период"]
		collect_count = pd.concat([collect_count, pd.DataFrame({'Доктор': 'Все услуги, оказанные доктором:', 'За весь период': collect_count["За весь период"].sum()}, index=[0])], axis=0) # Добавление итогов по отделению
		Services = Services.merge(collect_count, on='Доктор', how='outer')
		Services = Services.sort_values(by=['За весь период'], ascending=False) # Сортировка по количеству клиентов
		Services = pd.concat([Services, pd.DataFrame({'Доктор': np.nan}, index=[0])], axis=0) # Добавление промежуточной строки между блоками отчёта
	# Количество лечений ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		for date in Data_list.unique().tolist():
			collect = base[(base[Period] == date)]
			collect_count = collect.groupby('Доктор').agg(unique_count=('Код лечения', 'nunique')).reset_index(); collect_count.columns = ['Доктор', date] # Группировка с количеством уникальных лечений по докторам
			collect_count = pd.concat([collect_count, pd.DataFrame({'Доктор': 'Все уникальные клиенты для доктора:', date: collect_count[date].sum()}, index=[0])], axis=0) # Добавление итогов по отделению
			collect_count = pd.DataFrame({'Доктор': list(filter(None, base['Доктор'].unique().tolist()))}).merge(collect_count, on='Доктор', how='outer') # Объединяем таблицу со списком операторов (для одинаковой структуры данных)
			Treats = Treats.merge(collect_count, on='Доктор', how='outer')
		collect_count = base.groupby('Доктор').agg(unique_count=('Код лечения', 'nunique')).reset_index(); collect_count.columns = ['Доктор', "За весь период"]
		collect_count = pd.concat([collect_count, pd.DataFrame({'Доктор': 'Все лечения пациентов для доктора:', 'За весь период': collect_count["За весь период"].sum()}, index=[0])], axis=0) # Добавление итогов по отделению
		Treats = Treats.merge(collect_count, on='Доктор', how='outer')
		Treats = Treats.sort_values(by=['За весь период'], ascending=False) # Сортировка по количеству лечений
		Treats = pd.concat([Treats, pd.DataFrame({'Доктор': np.nan}, index=[0])], axis=0) # Добавление промежуточной строки между блоками отчёта
	# Количество уникальных пациентов -----------------------------------------------------------------------------------------------------------------------------------------------------------------
		for date in Data_list.unique().tolist():
			collect = base[(base[Period] == date)]
			collect_count = collect.groupby('Доктор').agg(unique_count=('Пациент', 'nunique')).reset_index(); collect_count.columns = ['Доктор', date] # Группировка с количеством уникальных пациентов по докторам
			collect_count = pd.concat([collect_count, pd.DataFrame({'Доктор': 'Все уникальные клиенты для доктора:', date: collect_count[date].sum()}, index=[0])], axis=0) # Добавление итогов по отделению
			collect_count = pd.DataFrame({'Доктор': list(filter(None, base['Доктор'].unique().tolist()))}).merge(collect_count, on='Доктор', how='outer') # Объединяем таблицу со списком операторов (для одинаковой структуры данных)
			Pacients = Pacients.merge(collect_count, on='Доктор', how='outer')
		collect_count = base.groupby('Доктор').agg(unique_count=('Пациент', 'nunique')).reset_index(); collect_count.columns = ['Доктор', "За весь период"]
		collect_count = pd.concat([collect_count, pd.DataFrame({'Доктор': 'Все уникальные клиенты для доктора:', 'За весь период': collect_count["За весь период"].sum()}, index=[0])], axis=0) # Добавление итогов по отделению
		Pacients = Pacients.merge(collect_count, on='Доктор', how='outer')
		Pacients = Pacients.sort_values(by=['За весь период'], ascending=False) # Сортировка по количеству клиентов
		Pacients = pd.concat([Pacients, pd.DataFrame({'Доктор': np.nan}, index=[0])], axis=0) # Добавление промежуточной строки между блоками отчёта
	# Пополнение списков массивов ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Main_list = pd.concat([Money, Services, Treats, Pacients], axis=0)
		tables_money.append(Money)
		tables_services.append(Services)
		tables_treats.append(Treats)
		tables_pacients.append(Pacients)
		tables_main.append(Main_list)
	###################################################################################################################################################################################################
	# По списку услуг
	Top_doс_work = pd.DataFrame(columns=['Доктор']); Top_doс_kateg = pd.DataFrame(columns=['Доктор'])  # Создание массивов для оценки топов работ
	Base_Doc = tables_money[1]
	Doctor_list_Top = Base_Doc['Доктор'].unique().tolist() # Сортированный список ТОП докторов
	Doctor_list_Top = Doctor_list_Top[1:]
	for doc in Doctor_list_Top:
		collect_doc = base[(base['Доктор'] == doc)]
		Top_doс_doc = pd.DataFrame(columns=['Выполненная работа'])
		try:
			for date in Data_list_M.unique().tolist():
				collect =  collect_doc[(collect_doc['Месяц, год'] == date)]
				doс_work = collect.groupby('Выполненная работа').agg({'Стоимость для пациента': ['sum']}).reset_index(); doс_work.columns = ['Выполненная работа', date] # Группировка с количеством уникальных пациентов по докторам
				doс_work = pd.DataFrame({'Выполненная работа': list(filter(None, collect_doc['Выполненная работа'].unique().tolist()))}).merge(doс_work, on='Выполненная работа', how='outer') # Объединяем таблицу со списком операторов (для одинаковой структуры данных)
				Top_doс_doc = Top_doс_doc.merge(doс_work, on='Выполненная работа', how='outer')
			doс_work = collect_doc.groupby('Выполненная работа').agg({'Стоимость для пациента': ['sum']}).reset_index(); doс_work.columns = ['Выполненная работа', "За весь период"]
			Top_doс_doc = Top_doс_doc.merge(doс_work, on='Выполненная работа', how='outer')
			Top_doс_doc = Top_doс_doc.sort_values(by=['За весь период'], ascending=False).head(10)  # Сортировка работ по сумме выручки и обрезка для топ 10
			Top_doс_doc['Доктор'] = doc # Добавляет имя сотрудника
			Top_doс_work = pd.concat([Top_doс_work, Top_doс_doc], axis=0)
			Top_doс_work = pd.concat([Top_doс_work, pd.DataFrame({'Выполненная работа': np.nan}, index=[0])], axis=0) # Добавление промежуточной строки между блоками отчёта
		except: pass
	###################################################################################################################################################################################################
	# По списку категорий
	for doc in Doctor_list_Top:
		collect_doc = base[(base['Доктор'] == doc)]
		Top_doс_doc = pd.DataFrame(columns=['Категория работы'])
		try:
			for date in Data_list_M.unique().tolist():
				collect = collect_doc[(collect_doc['Месяц, год'] == date)]
				doс_work = collect.groupby('Категория работы').agg({'Стоимость для пациента': ['sum']}).reset_index();
				doс_work.columns = ['Категория работы', date] # Группировка с количеством уникальных пациентов по докторам
				doс_work = pd.DataFrame({'Категория работы': list(filter(None, collect_doc['Категория работы'].unique().tolist()))}).merge(doс_work, on='Категория работы', how='outer') # Объединяем таблицу со списком операторов (для одинаковой структуры данных)
				Top_doс_doc = Top_doс_doc.merge(doс_work, on='Категория работы', how='outer')
			doс_work = collect_doc.groupby('Категория работы').agg({'Стоимость для пациента': ['sum']}).reset_index(); doс_work.columns = ['Категория работы', "За весь период"]
			Top_doс_doc = Top_doс_doc.merge(doс_work, on='Категория работы', how='outer')
			Top_doс_doc = Top_doс_doc.sort_values(by=['За весь период'], ascending=False).head(5) # Сортировка категорий по сумме выручки и обрезка для топ 5
			Top_doс_doc['Доктор'] = doc # Добавляет имя сотрудника
			Top_doс_kateg = pd.concat([Top_doс_kateg, Top_doс_doc], axis=0)
			Top_doс_kateg = pd.concat([Top_doс_kateg, pd.DataFrame({'Категория работы': np.nan}, index=[0])], axis=0) # Добавление промежуточной строки между блоками отчёта
		except: pass
	excel_writer(tables_main, Top_doс_work, Top_doс_kateg, base)

def excel_writer(tables_main, Top_doс_work, Top_doс_kateg, ORDERDET):
	with pd.ExcelWriter("Результаты хирургической части от " + datename + '.xlsx', engine="xlsxwriter") as writer:
		tables_main[0].to_excel(writer, sheet_name='ПХ по дням', index=False, freeze_panes=(1, 0))
		tables_main[1].to_excel(writer, sheet_name='ПХ по месяцам', index=False, freeze_panes=(1, 0))
		tables_main[2].to_excel(writer, sheet_name='ПХ по кварталам', index=False, freeze_panes=(1, 0))
		tables_main[3].to_excel(writer, sheet_name='ПХ по годам', index=False, freeze_panes=(1, 0))
		Top_doс_work.to_excel(writer, sheet_name='Топ10 услуг', index=False, freeze_panes=(1, 0))
		Top_doс_kateg.to_excel(writer, sheet_name='Топ5 категорий', index=False, freeze_panes=(1, 0))
		ORDERDET.to_excel(writer, sheet_name='Все оказанные услуги', index=False, freeze_panes=(1, 0))







Input_lag()
GetSQL()





