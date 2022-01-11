"""Конфигурация date и headers"""


from datetime import datetime

headers = {
    'accept': 'application/json',
    'X-API-Key': '',
    'Content-Type': 'application/json',
}

"""Получаем дату и месяц. Месяц будет использоваться как "-1" чтобы использовать табель прошлого месяца"""
# текущий год, будет исользоватья по умолчанию
year = datetime.now().year
# в 1с год вида 4021
year_1c = year + 2000
# месяц, будет использоватья по умолчанию
month = datetime.now().month
