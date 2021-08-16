
import xlwings
import pyodbc

i = 2
s = xlwings.books.active.sheets.active

CS = "DRIVER={SQL Server};SERVER=HIIWINBL18;DATABASE=SNDBase91;"
conn = pyodbc.connect(CS)
cursor = conn.cursor()

dates = list()
while s.range(i, 2).value:
    part = s.range(i, 2).value.replace("-", "_", 1)
    cursor.execute("SELECT ArcDateTime, PartName FROM PIPArchive WHERE PartName=? AND TransType='SN102'", part)
    for d, p in cursor.fetchall():
        dates.append(str(d.date()))

    i += 1

conn.close()

print("\n".join(sorted(set(dates))))
