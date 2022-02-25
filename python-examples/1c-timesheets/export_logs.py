def export_logs(event: str, datetime):
    with open('logs.txt', 'a', encoding="UTF-8") as outlogs:
        outlogs.write('{} {}\n'.format(event, datetime))
