import xlwings
import pyodbc
import os

import parsers

from re import compile as regex
from collections import defaultdict
from types import SimpleNamespace
from datetime import datetime

def main():
    # regexes
    CNF = regex("\d{10}")
    NOT_CNF = regex("\d{7}")

    # parse active xl sheet
    cohv = parsers.parse_sheet()

    # parse out dict structure
    cnf = defaultdict(int)
    not_cnf = defaultdict(list)
    for x in cohv:
        if CNF.match(x.order):
            cnf[x.matl] += x.qty
        elif NOT_CNF.match(x.order):
            not_cnf[x.matl].append((x.wbs, x.qty))
        else:
            print("Order not matched:", x.order)
        

    reader = SnReader()
    confirmations = list()
    for part, qty_confirmed in cnf.items():
        qty_burned = reader.get_part_burned_qty(part)

        if qty_confirmed < qty_burned:
            qty_to_confirm = qty_burned - qty_confirmed
            for wbs, qty in sorted(not_cnf[part]):
                if qty < qty_to_confirm:
                    confirmations.append((part, wbs, qty))
                    qty_to_confirm -= qty
                else:  # open quantity is equal or greater than what needs to be confirmed
                    confirmations.append((part, wbs, qty_to_confirm))
                    break

        try:
            del not_cnf[part]
        except KeyError:
            pass

    processed_lines = parsers.get_cnf_file_rows([x[0] for x in confirmations], processed_only=True)

    index = SimpleNamespace(matl=0, qty=4, wbs=2, plant=11)
    # matl = SimpleNamespace(matl=6, qty=8, loc=10, wbs=7, plant=11)

    templates = dict()
    for line in processed_lines:
        part = line[index.matl]

        templates[part] = line

    # create output file
    ready_file = os.path.join(
        os.path.dirname(os.path.realpath(__file__)), 
        "Production_{}.ready".format(datetime.now().strftime("%Y%m%d%H%M%S"))
    )

    with open(ready_file, 'w') as cnf_file:
        for part, wbs, qty in confirmations:
            line = templates[part]
            line[index.wbs] = wbs
            line[index.qty] = str(int(qty))

            cnf_file.write("\t".join(line))


class SnReader:

    def __init__(self):
        self.simtrans_cutoff = get_simtrans_cutoff()

        CS = "DRIVER={SQL Server};SERVER=HIIWINBL18;DATABASE=SNDBase91;"
        self.conn = pyodbc.connect(CS)

    def get_part_burned_qty(self, part_name):
        cursor = self.conn.cursor()
        cursor.execute("""
            SELECT
                WONumber, QtyInProcess
            FROM PIPArchive
            WHERE
                PartName = ?
            AND
                TransType = 'SN102'
            AND ArcDateTime < ?
        """, part_name.replace("-", "_", 1), self.simtrans_cutoff)

        total = 0
        for wo, qty in cursor.fetchall():
            if wo not in ('REMAKES', 'EXTRAS'):
                total += qty

        return total


def get_simtrans_cutoff():
    # sim trans files are generated every 4 hours on the half hours
    # 00:30, 04:30, 08:30, 12:30, 16:30, 20:30 

    now = datetime.now()

    last_run = now.replace(hour=now.hour - now.hour % 4)

    return "{0:%m}/{0:%d}/{0:%Y} {0:%I}:30:00 {0:%p}".format(last_run)



if __name__ == "__main__":
    main()
