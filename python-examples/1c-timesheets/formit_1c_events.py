"""
Формируем список ids календарных событий для записи в 1С ЗУП
Используется в loaded_calendar_events
"""


def events_formited(pre_events, events1c):
    # list ids calendar events
    events = []

    for event in pre_events:
        if event in events1c.values():
            for event1c_id, event1c_name in events1c.items():
                if event == event1c_name:
                    # добавляем в список ids календарных событий
                    events.append(event1c_id)
                else:
                    continue
        else:
            # если нет совпадений, пишем "Я"
            # можно вставить свое значение
            events.append('0x80E500155D01940311E8E8A7332ED753')

    return events
