
import xlwings
import pyodbc

from types import SimpleNamespace
from collections import defaultdict, namedtuple
from re import compile as regex
from tqdm import tqdm

import time

CNF_PATTERN = regex("\d{10}")
OPEN_PATTERN = regex("\d{7}")


def main():
    sndb = SnReader()

    parsed_xl = read_xl()

    print_later = []
    needs_cnf = []

    outfile = open('cnf_these.csv', 'w')
    outfile.write("Part,Qty\n")

    with tqdm(total=len(parsed_xl)) as progress:
        for part, records in parsed_xl.items():
            progress.set_description(part)
            progress.update()

            cnf_qty = qty_if_regex(CNF_PATTERN, records)
            open_qty = qty_if_regex(OPEN_PATTERN, records)
            if part.count('-') > 1 and cnf_qty > 0:
                continue

            burned_qty = sndb.get_part_burned_qty(part)

            if cnf_qty < burned_qty:
                print_later.append(
                    "{}: {} / {}".format(part, int(cnf_qty), burned_qty))
                outfile.write("{},{}\n".format(
                    part, int(burned_qty - cnf_qty)))

    outfile.close()

    if print_later:
        print("\n".join(print_later))


class SnReader:

    def __init__(self):
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
        """, part_name.replace("-", "_", 1))

        total = 0
        for wo, qty in cursor.fetchall():
            if wo not in ('REMAKES', 'EXTRAS'):
                total += qty

        return total


def read_xl():
    sht = xlwings.books.active.sheets.active

    # parse header
    col_map = SimpleNamespace()
    max_col = 0
    for i, col in enumerate(sht.range("A1").expand('right').value):
        if col == "Order":
            col_map.order = i
        elif col == "Material Number":
            col_map.part = i
        elif col == "Order quantity (GMEIN)":
            col_map.qty = i
        elif col == "WBS Element":
            col_map.wbs = i
        elif col == "Occurrence":
            col_map.shipment = i
        elif col == "Plant":
            col_map.plant = i
        elif col == "Material description":
            col_map.desc = i

        max_col = i

    data = defaultdict(list)
    Record = namedtuple('Record', ['order', 'qty', 'wbs', 'shipment', 'plant'])
    items = sht.range((2, 1), (2, max_col+1)).expand('down').value

    for row in tqdm(items):
        data[row[col_map.part]].append(Record(
            row[col_map.order],
            row[col_map.qty],
            row[col_map.wbs],
            row[col_map.shipment],
            row[col_map.plant],
        ))

    return data


def qty_if_regex(order_regex, records):
    total = 0

    for r in records:
        if order_regex.match(r.order):
            total += r.qty

    return total


if __name__ == "__main__":
    main()
