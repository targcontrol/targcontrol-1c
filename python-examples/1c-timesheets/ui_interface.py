"""Интерфейс для ввода данных"""


from tkinter import *
from tkinter import messagebox
from db_connect import *
import settings
from targcontrol_query import *
from settings import month, year
import webbrowser

"""Интерфейс для загрузки табеля в 1С"""
def click():
    # сохраняем введенное значение в месяц
    settings.month = message_month.get()
    # сохраняем введенное значение в год
    settings.year = message_year.get()

# функции с модуля targcontrol_query
    calendar_events = get_calendar_events()
    personal_numbers = get_employee_timesheet_number()
    employee_ids, employees = get_employee_id(personal_numbers)
    tables, tables_empl_ids = create_time_sheets(employee_ids, settings.month, settings.year, '17',calendar_events)
    targcontrol_tabels = formit_tabels(employees, tables, tables_empl_ids)

# функции с модуля db_connect
    events1c = get_1c_events()
    id_table = get_id_table()
    query_table(id_table)
    tables1c = tabels_query_1c(id_table)
    loaded_calendar_events(tables1c, targcontrol_tabels, id_table, events1c)
    loaded_time(tables1c, targcontrol_tabels, id_table)

# всплывающая форма после загрузки табеля
    messagebox.showwarning(title='Построение табеля', message='Табель построен')


"""Функция вызова ссылки в браузере"""
def callback(event):
    webbrowser.open_new(r"https://targcontrol.com/knowledge-base/work-with-timesheet/")

if __name__ == '__main__':
    """Tkinter форма"""
    window = Tk()
    window.title('Импорт табелей в 1С')
    # подсказка в форме
    label1 = Label(text="Для импорта графиков в 1С не забудьте закрыть ваши табели в TARGControl Cloud", fg="#eee", bg="blue")
    label1.pack()
    # ссылка
    label2 = Label(text="Подробнее о работе с табелями:", fg="#eee", bg="blue")
    label2.pack()
    # сделать ссылку активной
    label3 = Label(text="https://targcontrol.com/knowledge-base/work-with-timesheet/", fg="blue")
    label3.pack()
    label3.bind("<Button-1>", callback)
    # получить из формы месяц
    message_month = StringVar()
    message_entry = Entry(textvariable=message_month)
    message_entry.place(relx=.5, rely=.4, anchor="c")
    # получить год
    message_year = StringVar()
    message_entry_year = Entry(textvariable=message_year)
    message_entry_year.place(relx=.5, rely=.6, anchor="c")
    # отображение в поле ввода текущего месяца
    message_entry.insert(0, f"{datetime.now().month}")
    # отображение в поле ввода текущего года
    message_entry_year.insert(0, f"{datetime.now().year}")
    # размер окна
    window.geometry('600x200')
    button = Button(window, command=click, text=f'Импортировать табель за выбранный месяц', padx="15", pady="6", font="15")
    button.pack(side=BOTTOM)
    # отображение icon
    window.iconbitmap('images/icon.ico')
    window.mainloop()