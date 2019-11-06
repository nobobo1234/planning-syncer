from datetime import datetime, timedelta

start = datetime.today() - timedelta(days=datetime.today().weekday())
print(start)
