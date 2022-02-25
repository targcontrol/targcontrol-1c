"""
Requests to get data from TARGControl. Integration with 1С ЗУП by default
Rename 'dev.tarcontrol.com' to 'cloud.targcontrol.com'
Descriptions others modules in russian
 """


import requests
import json
import settings
from settings import *

# unsorted employees
employees_global = []


"""Get calendar events from TARGControl"""
def get_calendar_events():
    try:
        response = requests.get('https://dev.targcontrol.com/external/api/employee-schedules/calendar/types',
                                headers=headers).json()
    except:
        print('Error: get calendar events')

    return response


"""Get table numbers"""
def get_employee_timesheet_number():
    personal_numbers = []
    try:
        response = requests.get('https://dev.targcontrol.com/external/api/employees/query', headers=headers).json()
        for employee in response:
            # remove fired
            if employee.get('fired', True):
                continue
            employees_global.append(employee)
            # get table numbers
        for employee in employees_global:
            for x in employee.get('profileFieldValues'):
                personal_numbers.append(x['value'])
    except:
        print('Error: get employees personal numbers')
        pass

    return personal_numbers


""""Get employees and ids, which have personnel numbers"""
def get_employee_id(personal_numbers):
    employee_ids = []
    employees = []
    try:
        for employee in employees_global:
            for x in employee.get('profileFieldValues'):
                if x['value'] in personal_numbers:
                    employee_ids.append(employee['id'])
                    employees.append(employee)
    except:
        print('Error: get employees ids in an TARGControl Cloud')
        pass

    return employee_ids, employees


"""Create a timesheet list"""
def create_time_sheets(employee_ids, month: str, year: str, day: str, calendar_events):
    tables = []
    tables_empl_ids = []
    data = {"month": "{}-{}-{}".format(year, month, day)}
    try:
        response = requests.post('https://dev.targcontrol.com/external/api/time-sheets', headers=headers,
                                 data=json.dumps(data)).json()
        for table in response:
            tabel_dict = {'id': table.get('employeeId'), 'day': [], 'work_time': [], 'calendar_events':[]}

            days = table.get('days')

            for day in days:
                work_time = day.get('accountedTimeMillis')
                work_time = (work_time/(1000*60*60))%24
                tabel_dict.get('day').append(day.get('dayDate'))
                tabel_dict.get('work_time').append(work_time)

                # day off and without a visit
                if day.get('dayType') == 'DAY_OFF' and len(day.get('calendarEventIds')) == 0 and day.get('accountedTimeMillis') == 0:
                    tabel_dict.get('calendar_events').append('В')
                # day off and with a visit
                if day.get('dayType') == 'DAY_OFF' and len(day.get('calendarEventIds')) == 0 and day.get('accountedTimeMillis') > 0:
                    tabel_dict.get('calendar_events').append('Я')
                # work day and with a visit
                if day.get('dayType') == 'WORK_DAY' and len(day.get('calendarEventIds')) == 0 and day.get('accountedTimeMillis') != 0:
                    tabel_dict.get('calendar_events').append('Я')
                # work day and without a visit
                if day.get('dayType') == 'WORK_DAY' and len(day.get('calendarEventIds')) == 0 and day.get('accountedTimeMillis') == 0:
                    tabel_dict.get('calendar_events').append('ПР')

                # calendar events
                for calendar_event in calendar_events:
                    try:
                        for idcalendar_in_date in day.get('calendarEventIds'):
                            if calendar_event.get('id') == idcalendar_in_date:
                                tabel_dict.get('calendar_events').append(calendar_event.get('abbreviation'))
                    except:
                        tabel_dict.get('calendar_events').append('0')
                        print('Error calendar event add')
                        continue

            tables_empl_ids.append(table.get('employeeId'))
            tables.append(tabel_dict)
    except:
        print('Error: create time sheets for employees')
        pass

    return tables, tables_empl_ids


"""Make dict with employees and timesheets"""
def formit_tabels(employees, tables, tables_empl_ids):
    result = []
    try:
        for employee in employees:
            employee_id = employee.get('id')
            if employee_id in tables_empl_ids:
                employee_name = (employee.get('name', ).get('lastName', '') or '') + ' ' + \
                            (employee.get('name').get('firstName', '') or '') + ' ' + \
                            (employee.get('name').get('middleName', '') or '')

                # personal number acquisition cycle
                for x in employee.get('profileFieldValues'):
                    tabel_number = x['value']

                date_work = {}
                calendar_event = {}
                employee_dict = {'Employee name': employee_name,  'id': employee_id, 'tabel': tabel_number,
                             'date_work': date_work, 'calendar_events': calendar_event}

                for day in tables:
                    if day.get('id') in employee_id:
                        for date, timework in sorted(zip(day.get('day'), day.get('work_time'))):
                            date_work[date] = timework
                        for date, calendar in sorted(zip(day.get('day'), day.get('calendar_events'))):
                            calendar_event[date] = calendar
                    else:
                        continue
                result.append(employee_dict)
    except:
        print('Error')
        pass

    return result
