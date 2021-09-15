
import xlwings

from collections import defaultdict
from re import compile as regex

import parsers
import sndb

# regexes
CNF = regex("\d{10}")
NOT_CNF = regex("\d{7}")


def main():
    cohv = parsers.parse_sheet()

    active = defaultdict(list)
    not_found = list()

    with sndb.get_sndb_conn() as db:
        cursor = db.cursor()

        for x in cohv:
            if CNF.match(x.order):
                continue

            has_result = False

            cursor.execute("SELECT ProgramName FROM PIP WHERE PartName=?", x.part.replace("-", "_", 1))
            for row in cursor.fetchall():
                has_result = True
                prog = row[0]

                active[prog].append(x.part)

            if not has_result:
                not_found.append(x.part)

    if not_found:
        print("\nNot active:")
        for x in not_found:
            print(" - {}".format(x))


    if active:
        print("Needs updated:")
        for k, v in active.items():
            print(k)
            for x in v:
                print(" - {}".format(x))


if __name__ == "__main__":
    main()
