import xlwings
import os
import re

import parsers

from collections import defaultdict
from types import SimpleNamespace
from datetime import datetime

def main():
    # regexes
    CNF = re.compile("\d{10}")
    NOT_CNF = re.compile("\d{7}")

    # parse active xl sheet
    cohv = parsers.parse_sheet()

    # parse out dict structure
    cnf = defaultdict(lambda: defaultdict(int))
    not_cnf = defaultdict(lambda: defaultdict(int))
    for x in cohv:
        if CNF.match(x.order):
            collection = cnf
        elif NOT_CNF.match(x.order):
            collection = not_cnf
        else:
            print("Order not matched:", x.order)
        
        collection[x.matl][x.wbs] =+ x.qty


    processed_lines = parsers.get_cnf_file_rows(cnf.keys(), processed_only=True)

    index = SimpleNamespace(matl=0, qty=4, wbs=2, plant=11)
    # matl = SimpleNamespace(matl=6, qty=8, loc=10, wbs=7, plant=11)

    failures = list()
    for line in processed_lines:
        part = line[index.matl]
        wbs = line[index.wbs]
        qty = int(line[index.qty])

        try:
            cnf_qty = cnf[part][wbs]
            if cnf_qty >= qty:
                cnf[part][wbs] -= qty
            else:
                qty -= cnf[part][wbs]
                cnf[part][wbs] = 0

                raise KeyError

        except KeyError:
            try:
                while qty > 0:
                    # get last key-value pair
                    order_wbs, order_qty = not_cnf[part].popitem()

                    if order_qty <= qty:
                        line[index.qty] = str(int(order_qty))
                        qty -= order_qty
                    else:
                        not_cnf[part][order_wbs] = order_qty - qty
                        line[index.qty] = str(int(qty))
                
                    failures.append(line)
            except KeyError:
                print("Part not found in open orders:", part)

    # create output file
    ready_file = os.path.join(
        os.path.dirname(os.path.realpath(__file__)), 
        "Production_{}.ready".format(datetime.now().strftime("%Y%m%d%H%M%S"))
    )
    with open(ready_file, 'w') as fail_file:
        for line in sorted(failures):
            fail_file.write("\t".join(line))


if __name__ == "__main__":
    main()
