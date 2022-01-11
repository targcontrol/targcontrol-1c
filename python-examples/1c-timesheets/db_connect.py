"""
Запросы в 1С ЗУП. Для подключения к базе данных Microsoft sql нужно ввести ip, название базы, пользователя
и пароль.
Название таблиц также могут отличаться.
"""


import pyodbc
import binascii
from targcontrol_query import *
from export_logs import *
from formit_1c_events import events_formited
from execute_samples import timesheet_load_sample, calendarevents_load_sample


"""Подключение к базе данных 1С"""
try:
    conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=server;DATABASE=db;UID=sa;PWD=password')
    cursor = conn.cursor()
except:
    print('Error: unable to connect to database')


"""Получить список всех имеющихся календарных событий в 1С"""
def get_1c_events():
    events1c = {}
    cursor.execute("select * from _Reference45")
    row = cursor.fetchall()
    for x in row:
        y = str(binascii.b2a_hex(x[0]))
        result_y = '0x' + y[2:-1]
        events1c[result_y] = x[6]

    return events1c


"""Получение номера таблицы"""
def get_id_table():
    cursor.execute(f"select _IDRRef from _Document427 where _Fld16109 = '{settings.year_1c}-{settings.month}-01 00:00:00'")
    row = cursor.fetchall()
    # берем последний документ (последний месяц)
    for x in row[-1]:
        y = str(binascii.b2a_hex(x))
        result = '0x' + y[2:-1]

    return result


"""Запрос на таблицу 2"""
def query_table(id_table):
    result = []
    cursor.execute(f"select * from _Document427_VT16125 where _Document427_IDRRef = {id_table}")
    row = cursor.fetchone()
    y = str(binascii.b2a_hex(row[4]))
    for x in row[5:35]:
        result.append(x)
    result.append(y)
    return result


"""Получение табелей, ФИО и табельного номера сотрудника"""
def tabels_query_1c(id_table):
    # общий список
    result = []

    cursor.execute(f"SELECT *FROM [_Document427_VT16125 ] INNER JOIN [_Reference221]"
                   f" ON _Document427_VT16125._Fld16127RRef = _Reference221._IDRRef where _Document427_IDRRef = {id_table}")
    query = cursor.fetchall()

    for row in query:
        # словарь для хранения данных сотрудника
        employee_dict_1c = {'Employee': []}

        # цикл получения ФИО и табельного номера сотрудника, занесение в словарь
        for y in row:
            # проверяем, является ли переменная str
            if isinstance(y, str):
                employee_dict_1c.get('Employee').append(y)

        # цикл получения данных (от начала до конца табеля). Позже нужно добавить проверку месяца
        for x in row[0:35]:
            employee_dict_1c.get('Employee').append(x)

        result.append(employee_dict_1c)

    return result


"""Загрузка часов в базу данных 1С"""
def loaded_time(tables1c, tables_targcontrol, id_table):
    n = 0
    for employee in tables_targcontrol:
        for employee1c in tables1c:
            hours = []
            if employee.get('Employee name') == employee1c.get('Employee')[1]:
                for x in employee.get('date_work').values():
                    hours.append(x)
                y = str(binascii.b2a_hex(employee1c.get('Employee')[7]))
                # id сотрудника, для которого будем грузить табель
                id1c = '0x' + y[2:-1]
                export_logs(f'Табель сформирован для сотрудника {employee1c.get("Employee")[1]} за ', f'{settings.month}-{settings.year}')

                # если 31 день в месяце
                if len(hours) == 31:
                    timesheet_load_sample(cursor, conn, id_table, id1c, hours)
                    cursor.execute(f"UPDATE _Document427_VT16125  "
                        f"SET _Fld16156 = {hours[28]}, _Fld16157 = {hours[29]}, "
                        f"_Fld16158 = {hours[30]}"
                        f" where _Document427_IDRRef = {id_table} and _Fld16127RRef = {id1c}")
                    conn.commit()
                    print(employee.get('Employee name'))

                    # если 30 дней в месяце
                if len(hours) == 30:
                    try:
                        timesheet_load_sample(cursor, conn, id_table, id1c, hours)
                        cursor.execute(f"UPDATE _Document427_VT16125  "
                        f"SET _Fld16156 = {hours[28]}, _Fld16157 = {hours[29]}, _Fld16158 = 0"
                        f" where _Document427_IDRRef = {id_table} and _Fld16127RRef = {id1c}")
                        conn.commit()
                        print(employee.get('Employee name'))
                    except:
                        print('Error time loaded')
                        continue

                # если 28 дней в месяце
                if len(hours) == 28:
                    timesheet_load_sample(cursor, conn, id_table, id1c, hours)
                    print(employee.get('Employee name'))
                    conn.commit()

                # если 29 дней в месяце
                if len(hours) == 29:
                    timesheet_load_sample(cursor, conn, id_table, id1c, hours)
                    cursor.execute(f"UPDATE _Document427_VT16125  "
                        f"SET _Fld16156 = {hours[28]}, _Fld16157 = 0, _Fld16158 = 0"
                        f" where _Document427_IDRRef = {id_table} and _Fld16127RRef = {id1c}")
                    conn.commit()
                    print(employee.get('Employee name'))
                else:
                    pass
            else:
                continue


"""Загрузка в БД 1С календарных событий"""
def loaded_calendar_events(tables1c, tables_targcontrol, id_table, events1c):
    n = 0
    for employee in tables_targcontrol:
        for employee1c in tables1c:
            pre_events = []
            if employee.get('Employee name') == employee1c.get('Employee')[1]:
                for calendar_event in employee.get('calendar_events').values():
                    pre_events.append(calendar_event)
                y = str(binascii.b2a_hex(employee1c.get('Employee')[7]))
                # id сотрудника, для которого будем грузить табель
                id1c = '0x' + y[2:-1]

                events = events_formited(pre_events, events1c)

                export_logs(f'Табель сформирован для сотрудника {employee1c.get("Employee")[1]} за ', f'{settings.month}-{settings.year}')
                # если 31 день в месяце
                if len(events) == 31:
                    calendarevents_load_sample(cursor, conn, id_table, id1c, events)
                    cursor.execute(f"UPDATE _Document427_VT16125  "
                        f"SET _Fld16187RRef = {events[28]}, _Fld16188RRef = {events[29]}, "
                        f"_Fld16189RRef = {events[30]}"
                        f" where _Document427_IDRRef = {id_table} and _Fld16127RRef = {id1c}")
                    print(employee.get('Employee name'))
                    conn.commit()

                # если 30 дней в месяце
                if len(events) == 30:
                    try:
                        calendarevents_load_sample(cursor, conn, id_table, id1c, events)
                        cursor.execute(f"UPDATE _Document427_VT16125  "
                                       f"SET _Fld16187RRef = {events[28]}, _Fld16188RRef = {events[29]}, "
                                       f"_Fld16189RRef = 0"
                                       f" where _Document427_IDRRef = {id_table} and _Fld16127RRef = {id1c}")
                        print(employee.get('Employee name'))
                        conn.commit()
                    except:
                        print('Error events loaded')
                        continue

                # если 28 дней в месяце
                if len(events) == 28:
                    calendarevents_load_sample(cursor, conn, id_table, id1c, events)
                    print(employee.get('Employee name'))
                    conn.commit()

                # если 29 дней в месяце
                if len(events) == 29:
                    calendarevents_load_sample(cursor, conn, id_table, id1c, events)
                    cursor.execute(f"UPDATE _Document427_VT16125  "
                                   f"SET _Fld16187RRef = {events[28]}, _Fld16188RRef = 0, "
                                   f"_Fld16189RRef = 0"
                                   f" where _Document427_IDRRef = {id_table} and _Fld16127RRef = {id1c}")
                    print(employee.get('Employee name'))
                    conn.commit()
                else:
                    pass
            else:
                continue
