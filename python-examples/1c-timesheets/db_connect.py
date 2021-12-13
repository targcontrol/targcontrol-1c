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


"""Подключение к базе данных 1С"""
try:
    conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=ip;DATABASE=database;UID=user;PWD=password')
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
                    cursor.execute(f"UPDATE _Document427_VT16125  "
                        f"SET _Fld16128 = {hours[0]}, _Fld16129 = {hours[1]}, "
                        f"_Fld16130 = {hours[2]}, _Fld16131 = {hours[3]}, _Fld16132 = {hours[4]},"
                        f" _Fld16133 = {hours[5]},"
                        f" _Fld16134 = {hours[6]}, _Fld16135 = {hours[7]}, _Fld16136 = {hours[8]}, "
                        f"_Fld16137 = {hours[9]}, _Fld16138 = {hours[10]}, _Fld16139 = {hours[11]}, "
                        f"_Fld16140 = {hours[12]}, _Fld16141 = {hours[13]}, _Fld16142 = {hours[14]},"
                        f" _Fld16143 = {hours[15]}, _Fld16144 = {hours[16]}, _Fld16145 = {hours[17]}, _Fld16146 = {hours[18]},"
                        f" _Fld16147 = {hours[19]}, "
                        f"_Fld16148 = {hours[20]}, _Fld16149 = {hours[21]}, _Fld16150 = {hours[22]}, _Fld16151 = {hours[23]}, "
                        f"_Fld16152 = {hours[24]}, _Fld16153 = {hours[25]}, _Fld16154 = {hours[26]}, _Fld16155 = {hours[27]},"
                        f" _Fld16156 = {hours[28]}, _Fld16157 = {hours[29]}, "
                        f"_Fld16158 = {hours[30]}"
                        f" where _Document427_IDRRef = {id_table} and _Fld16127RRef = {id1c}")
                    conn.commit()
                    print(employee.get('Employee name'))

                    # если 30 дней в месяце
                if len(hours) == 30:
                    try:
                        cursor.execute(f"UPDATE _Document427_VT16125  "
                        f"SET _Fld16128 = {hours[0]}, _Fld16129 = {hours[1]}, "
                        f"_Fld16130 = {hours[2]}, _Fld16131 = {hours[3]}, _Fld16132 = {hours[4]},"
                        f" _Fld16133 = {hours[5]},"
                        f" _Fld16134 = {hours[6]}, _Fld16135 = {hours[7]}, _Fld16136 = {hours[8]}, "
                        f"_Fld16137 = {hours[9]}, _Fld16138 = {hours[10]}, _Fld16139 = {hours[11]}, "
                        f"_Fld16140 = {hours[12]}, _Fld16141 = {hours[13]}, _Fld16142 = {hours[14]},"
                        f" _Fld16143 = {hours[15]}, _Fld16144 = {hours[16]}, _Fld16145 = {hours[17]}, _Fld16146 = {hours[18]},"
                        f" _Fld16147 = {hours[19]}, "
                        f"_Fld16148 = {hours[20]}, _Fld16149 = {hours[21]}, _Fld16150 = {hours[22]}, _Fld16151 = {hours[23]}, "
                        f"_Fld16152 = {hours[24]}, _Fld16153 = {hours[25]}, _Fld16154 = {hours[26]}, _Fld16155 = {hours[27]},"
                        f" _Fld16156 = {hours[28]}, _Fld16157 = {hours[29]}, _Fld16158 = 0"
                        f" where _Document427_IDRRef = {id_table} and _Fld16127RRef = {id1c}")
                        conn.commit()
                        print(employee.get('Employee name'))
                    except:
                        print('Error time loaded')
                        continue

                # если 28 дней в месяце
                if len(hours) == 28:
                    cursor.execute(f"UPDATE _Document427_VT16125  "
                        f"SET _Fld16128 = {hours[0]}, _Fld16129 = {hours[1]}, "
                        f"_Fld16130 = {hours[2]}, _Fld16131 = {hours[3]}, _Fld16132 = {hours[4]},"
                        f" _Fld16133 = {hours[5]},"
                        f" _Fld16134 = {hours[6]}, _Fld16135 = {hours[7]}, _Fld16136 = {hours[8]}, "
                        f"_Fld16137 = {hours[9]}, _Fld16138 = {hours[10]}, _Fld16139 = {hours[11]}, "
                        f"_Fld16140 = {hours[12]}, _Fld16141 = {hours[13]}, _Fld16142 = {hours[14]},"
                        f" _Fld16143 = {hours[15]}, _Fld16144 = {hours[16]}, _Fld16145 = {hours[17]}, _Fld16146 = {hours[18]},"
                        f" _Fld16147 = {hours[19]}, "
                        f"_Fld16148 = {hours[20]}, _Fld16149 = {hours[21]}, _Fld16150 = {hours[22]}, _Fld16151 = {hours[23]}, "
                        f"_Fld16152 = {hours[24]}, _Fld16153 = {hours[25]}, _Fld16154 = {hours[26]}, _Fld16155 = {hours[27]},"
                        f" _Fld16156 = 0, _Fld16157 = 0, _Fld16158 = 0"
                        f" where _Document427_IDRRef = {id_table} and _Fld16127RRef = {id1c}")
                    conn.commit()
                    print(employee.get('Employee name'))

                # если 29 дней в месяце
                if len(hours) == 29:
                    cursor.execute(f"UPDATE _Document427_VT16125  "
                        f"SET _Fld16128 = {hours[0]}, _Fld16129 = {hours[1]}, "
                        f"_Fld16130 = {hours[2]}, _Fld16131 = {hours[3]}, _Fld16132 = {hours[4]},"
                        f" _Fld16133 = {hours[5]},"
                        f" _Fld16134 = {hours[6]}, _Fld16135 = {hours[7]}, _Fld16136 = {hours[8]}, "
                        f"_Fld16137 = {hours[9]}, _Fld16138 = {hours[10]}, _Fld16139 = {hours[11]}, "
                        f"_Fld16140 = {hours[12]}, _Fld16141 = {hours[13]}, _Fld16142 = {hours[14]},"
                        f" _Fld16143 = {hours[15]}, _Fld16144 = {hours[16]}, _Fld16145 = {hours[17]}, _Fld16146 = {hours[18]},"
                        f" _Fld16147 = {hours[19]}, "
                        f"_Fld16148 = {hours[20]}, _Fld16149 = {hours[21]}, _Fld16150 = {hours[22]}, _Fld16151 = {hours[23]}, "
                        f"_Fld16152 = {hours[24]}, _Fld16153 = {hours[25]}, _Fld16154 = {hours[26]}, _Fld16155 = {hours[27]},"
                        f" _Fld16156 = {hours[28]}, _Fld16157 = 0, _Fld16158 = 0"
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
                    cursor.execute(f"UPDATE _Document427_VT16125  "
                        f"SET _Fld16159RRef = {events[0]}, _Fld16160RRef = {events[1]}, "
                        f"_Fld16161RRef = {events[2]}, _Fld16162RRef = {events[3]}, _Fld16163RRef = {events[4]},"
                        f" _Fld16164RRef = {events[5]},"
                        f" _Fld16165RRef = {events[6]}, _Fld16166RRef = {events[7]}, _Fld16167RRef = {events[8]}, "
                        f"_Fld16168RRef = {events[9]}, _Fld16169RRef = {events[10]}, _Fld16170RRef = {events[11]}, "
                        f"_Fld16171RRef = {events[12]}, _Fld16172RRef = {events[13]}, _Fld16173RRef = {events[14]},"
                        f" _Fld16174RRef = {events[15]}, _Fld16175RRef = {events[16]}, _Fld16176RRef = {events[17]}, _Fld16177RRef = {events[18]},"
                        f" _Fld16178RRef = {events[19]}, "
                        f"_Fld16179RRef = {events[20]}, _Fld16180RRef = {events[21]}, _Fld16181RRef = {events[22]}, _Fld16182RRef = {events[23]}, "
                        f"_Fld16183RRef = {events[24]}, _Fld16184RRef = {events[25]}, _Fld16185RRef = {events[26]}, _Fld16186RRef = {events[27]},"
                        f" _Fld16187RRef = {events[28]}, _Fld16188RRef = {events[29]}, "
                        f"_Fld16189RRef = {events[30]}"
                        f" where _Document427_IDRRef = {id_table} and _Fld16127RRef = {id1c}")
                    print(employee.get('Employee name'))
                    conn.commit()

                # если 30 дней в месяце
                if len(events) == 30:
                    try:
                        cursor.execute(f"UPDATE _Document427_VT16125  "
                                       f"SET _Fld16159RRef = {events[0]}, _Fld16160RRef = {events[1]}, "
                                       f"_Fld16161RRef = {events[2]}, _Fld16162RRef = {events[3]}, _Fld16163RRef = {events[4]},"
                                       f" _Fld16164RRef = {events[5]},"
                                       f" _Fld16165RRef = {events[6]}, _Fld16166RRef = {events[7]}, _Fld16167RRef = {events[8]}, "
                                       f"_Fld16168RRef = {events[9]}, _Fld16169RRef = {events[10]}, _Fld16170RRef = {events[11]}, "
                                       f"_Fld16171RRef = {events[12]}, _Fld16172RRef = {events[13]}, _Fld16173RRef = {events[14]},"
                                       f" _Fld16174RRef = {events[15]}, _Fld16175RRef = {events[16]}, _Fld16176RRef = {events[17]}, _Fld16177RRef = {events[18]},"
                                       f" _Fld16178RRef = {events[19]}, "
                                       f"_Fld16179RRef = {events[20]}, _Fld16180RRef = {events[21]}, _Fld16181RRef = {events[22]}, _Fld16182RRef = {events[23]}, "
                                       f"_Fld16183RRef = {events[24]}, _Fld16184RRef = {events[25]}, _Fld16185RRef = {events[26]}, _Fld16186RRef = {events[27]},"
                                       f" _Fld16187RRef = {events[28]}, _Fld16188RRef = {events[29]}, "
                                       f"_Fld16189RRef = 0"
                                       f" where _Document427_IDRRef = {id_table} and _Fld16127RRef = {id1c}")
                        print(employee.get('Employee name'))
                        conn.commit()
                    except:
                        print('Error events loaded')
                        continue

                # если 28 дней в месяце
                if len(events) == 28:
                    cursor.execute(f"UPDATE _Document427_VT16125  "
                                   f"SET _Fld16159RRef = {events[0]}, _Fld16160RRef = {events[1]}, "
                                   f"_Fld16161RRef = {events[2]}, _Fld16162RRef = {events[3]}, _Fld16163RRef = {events[4]},"
                                   f" _Fld16164RRef = {events[5]},"
                                   f" _Fld16165RRef = {events[6]}, _Fld16166RRef = {events[7]}, _Fld16167RRef = {events[8]}, "
                                   f"_Fld16168RRef = {events[9]}, _Fld16169RRef = {events[10]}, _Fld16170RRef = {events[11]}, "
                                   f"_Fld16171RRef = {events[12]}, _Fld16172RRef = {events[13]}, _Fld16173RRef = {events[14]},"
                                   f" _Fld16174RRef = {events[15]}, _Fld16175RRef = {events[16]}, _Fld16176RRef = {events[17]}, _Fld16177RRef = {events[18]},"
                                   f" _Fld16178RRef = {events[19]}, "
                                   f"_Fld16179RRef = {events[20]}, _Fld16180RRef = {events[21]}, _Fld16181RRef = {events[22]}, _Fld16182RRef = {events[23]}, "
                                   f"_Fld16183RRef = {events[24]}, _Fld16184RRef = {events[25]}, _Fld16185RRef = {events[26]}, _Fld16186RRef = {events[27]},"
                                   f" _Fld16187RRef = 0, _Fld16188RRef = 0, "
                                   f"_Fld16189RRef = 0"
                                   f" where _Document427_IDRRef = {id_table} and _Fld16127RRef = {id1c}")
                    print(employee.get('Employee name'))

                    conn.commit()

                # если 29 дней в месяце
                if len(events) == 29:
                    cursor.execute(f"UPDATE _Document427_VT16125  "
                                   f"SET _Fld16159RRef = {events[0]}, _Fld16160RRef = {events[1]}, "
                                   f"_Fld16161RRef = {events[2]}, _Fld16162RRef = {events[3]}, _Fld16163RRef = {events[4]},"
                                   f" _Fld16164RRef = {events[5]},"
                                   f" _Fld16165RRef = {events[6]}, _Fld16166RRef = {events[7]}, _Fld16167RRef = {events[8]}, "
                                   f"_Fld16168RRef = {events[9]}, _Fld16169RRef = {events[10]}, _Fld16170RRef = {events[11]}, "
                                   f"_Fld16171RRef = {events[12]}, _Fld16172RRef = {events[13]}, _Fld16173RRef = {events[14]},"
                                   f" _Fld16174RRef = {events[15]}, _Fld16175RRef = {events[16]}, _Fld16176RRef = {events[17]}, _Fld16177RRef = {events[18]},"
                                   f" _Fld16178RRef = {events[19]}, "
                                   f"_Fld16179RRef = {events[20]}, _Fld16180RRef = {events[21]}, _Fld16181RRef = {events[22]}, _Fld16182RRef = {events[23]}, "
                                   f"_Fld16183RRef = {events[24]}, _Fld16184RRef = {events[25]}, _Fld16185RRef = {events[26]}, _Fld16186RRef = {events[27]},"
                                   f" _Fld16187RRef = {events[28]}, _Fld16188RRef = 0, "
                                   f"_Fld16189RRef = 0"
                                   f" where _Document427_IDRRef = {id_table} and _Fld16127RRef = {id1c}")
                    print(employee.get('Employee name'))
                    conn.commit()
                else:
                    pass
            else:
                continue
