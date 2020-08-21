import xlwings
import os

from types import SimpleNamespace

aliases = SimpleNamespace()
aliases.matl = ("Material", "Material Number")
aliases.qty = ("Qty in unit of entry", "Unrestricted")
aliases.loc = ("Storage Location")
aliases.wbs = ("WBS Element", "Special stock number")
aliases.plant = ("Plant")
aliases.order = ("Order")


def parse_header(row):
    header = SimpleNamespace()

    for i, item in enumerate(row):
        # same as
        # `
        #   if item in aliases.matl:
        #       header.matl = i
        # `
        # just less code to maintain
        for k, v in aliases.__dict__.items():
            if item in v:
                setattr(header, k, i)

    return header


def parse_row(row, header):
    res = SimpleNamespace()

    for k, v in header.__dict__.items():
        # same as res.matl = row[header.matl]
        setattr(res, k, row[v])

    return res


def parse_sheet(sheet=xlwings.books.active.sheets.active):
    header = parse_header(sheet.range("A1").expand('right').value)

    max_column = max(header.__dict__.values())
    last_row = sheet.range((2, header.matl + 1)).end('down').row
    for row in sheet.range((2, 1), (last_row, max_column + 1)).value:
        yield parse_row(row, header)


def get_cnf_file_rows(parts):

    part = SimpleNamespace(matl=0, qty=4, wbs=2, plant=11)
    matl = SimpleNamespace(matl=6, qty=8, loc=10, wbs=7, plant=11)

    def files():
        dirs = [
            r"\\hssieng\SNData\SimTrans\SAP Data Files\Processed",
            r"\\hssieng\SNData\SimTrans\SAP Data Files\deleted files",
            r"\\hssieng\SNData\SimTrans\SAP Data Files\old deleted files",
        ]

        # get data from SAP Data Files
        for d in dirs:
            for f in os.listdir(d):
                yield os.path.join(d, f)

    def fileWorker(f):
        result = list()
        with open(f, "r") as prod_file:
            for line in prod_file.readlines():
                result.append(line.upper().split("\t"))
        return result

    prod_data = list()
    with tqdm.tqdm(desc="Fetching Data") as progress:
        for result in Pool().imap(fileWorker, files()):
            progress.update(1)
            for res in result:
                if res[part.matl] in parts:
                    prod_data[res[part.matl]] = x

    return prod_data
