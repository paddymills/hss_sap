import xlwings
import pyodbc

import numpy as np
import os
import sys
import datetime

from multiprocessing import Pool
import tqdm

from operator import itemgetter
from types import SimpleNamespace
from collections import defaultdict
from itertools import product

index = SimpleNamespace()

# if P1_DUMP is not null,
# all plant 1 locations will be set to to dumping ground
P1_DUMP = "T19"
P1_LOCS = ["{}{}".format(*x) for x in product("HIJKLMNT", range(1, 20))]

stock_locs = {
    "50/50W-0006": "T19",
    "50/50W-0008": "T19",
    "50/50W-0010": "T19",
    "50/50W-0012": "T19",
    "50/50W-0014": "T19",
    "50/50W-0100": "T19",
}


def main():
    sheet = xlwings.books.active.sheets.active
    # find order of Part, Qty, WBS, Order#, Plant
    if sheet.range("B1").value == "Material":  # COGI w/o changes
        indexOrder = (1, 6, 8, 9, 3)
    elif sheet.range("A1").value == "Material":  # COGI w/ status column removed
        indexOrder = (0, 5, 7, 8, 2)
    # COGI w/ 4 columns inserted at beginning
    elif sheet.range("E1").value == "Material":
        indexOrder = (4, 1, 11, 12, 6)
    elif sheet.range("B1").value == "Material Number":  # COHV
        if sheet.range("C1").value == "Material description":  # COHV CNF
            indexOrder = (1, 6, 3, 0, 8)
        else:
            indexOrder = (1, 2, 3, 0, 7)
    else:
        print("header format not matched")
        exit()
        # indexOrder = (0,2,3,0)

    index.PART, index.QTY, index.WBS, index.ORDER, index.PLANT = indexOrder

    skip_pip = (len(sys.argv) > 1 and sys.argv[1] == 'nopip')

    end = sheet.range((2, index.PART+1)).expand().last_cell
    data = sheet.range((2, 1), end).options(ndim=2).value
    data = findActivePrograms(data, trim=skip_pip)

    refinedData = [x for x in data if x[index.QTY]]
    # dumpXlWinShuttles(refinedData)
    createCnfFile(refinedData)


def findActivePrograms(data, trim=False):
    conn = pyodbc.connect(
        "DRIVER={SQL Server};SERVER=HIIWINBL18;DATABASE=SNDBase91;UID=SNUser;PWD=BestNest1445")
    cursor = conn.cursor()

    new_data = []
    for x in data:
        part = x[index.PART].replace("-", "_", 1)
        cursor.execute("""
            SELECT PIP.PartName, PIP.ProgramName, Program.MachineName
            FROM PIP
            INNER JOIN Program
                ON PIP.ProgramName=Program.ProgramName
            WHERE PIP.PartName=?
        """, part)

        occurences = 0
        for a in cursor.fetchall():
            print(a)
            occurences += 1

        if occurences < 1:
            new_data.append(x)

    conn.close()

    return new_data


def dumpXlWinShuttles(data):
    def uniqueNthItems(ls): return np.unique(
        np.take(ls, [index.ORDER], axis=1))
    data_OrderPartQty = itemgetter(index.ORDER, index.PART, index.QTY)
    WinShuttle_DataFiles = os.path.join(os.path.expanduser(
        "~\\Documents"), "WinShuttle", "TRANSACTION", "Data")

    wb = xlwings.books.open(os.path.join(
        WinShuttle_DataFiles, "CO02 Add Material Components.xlsx"))
    wb.sheets[0].range("A2").value = [data_OrderPartQty(x) for x in data]
    wb.save()
    wb.close()

    wb = xlwings.books.open(os.path.join(
        WinShuttle_DataFiles, "CO02 Add Material Line.xlsx"))
    wb.sheets[0].range("A2").value = [[x] for x in uniqueNthItems(data)]
    wb.save()
    wb.close()


def createCnfFile(data):
    def uniqueNthItems(ls): return np.unique(np.take(ls, [index.PART], axis=1))
    data_PartQtyWbsPlant = itemgetter(
        index.PART, index.QTY, index.WBS, index.PLANT)

    def location_handler(loc, plant, mm):
        if not loc:
            return 'RAW'
        elif plant == 'HS02':
            return 'RAW'

        if plant == 'HS01' and P1_DUMP and loc in P1_LOCS:
            loc = P1_DUMP
        elif mm in stock_locs.keys():
            loc = stock_locs[mm]

        return loc

    dirs = [
        r"\\hssieng\SNData\SimTrans\SAP Data Files\Processed",
        r"\\hssieng\SNData\SimTrans\SAP Data Files\deleted files",
        # r"\\hssieng\SNData\SimTrans\SAP Data Files\old deleted files",
    ]

    cnf_parts = list()
    faro_map = dict()
    for part in tqdm.tqdm(uniqueNthItems(data), desc="Adding FARO names"):
        job, mark = part.split("-", 1)
        faro_name = "{}-{}".format(job[-4:], mark)
        cnf_parts += [part, faro_name]
        faro_map[faro_name] = part

    prod_data = dict()

    # get data from SAP Data Files
    files = list()
    for d in dirs:
        for f in os.listdir(d):
            files.append(os.path.join(d, f))

    with tqdm.tqdm(desc="Fetching Data", total=len(files)) as progress:
        for result in Pool().imap(fileWorker, files):
            progress.update(1)
            for x in result:
                trials = [
                    x[0],
                    x[0][:x[0].find("-", 5)],
                    x[0][:x[0].find("-", 9)],
                ]
                try:
                    trials.append("{0}-{3}".format(*x[0].split("-")))
                except IndexError:
                    pass

                for t in trials:
                    if t in cnf_parts:
                        if t in faro_map.keys():
                            t = faro_map[t]
                        prod_data[t] = x

    # group items by part and wbs
    dataSubtotal = defaultdict(lambda: defaultdict(int))
    dataPlants = dict()
    for x in data:
        part, qty, wbs, plant = data_PartQtyWbsPlant(x)
        dataSubtotal[part][wbs] += int(qty)
        dataPlants[part] = plant

    # create output file
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    thisFilesDir = os.path.dirname(os.path.realpath(__file__))
    with open(os.path.join(thisFilesDir, "Production_{}.ready".format(timestamp)), "w") as cnf_file:
        for part, val in dataSubtotal.items():
            if part not in prod_data.keys():
                print(part, "not found &&", findPartCompletionDate(part))
                continue

            d = prod_data[part]
            area_ea = float(d[8]) / float(d[4])
            d[10] = location_handler(d[10], d[11], d[6])
            for wbs, qty in val.items():
                d[0] = part
                d[2] = wbs
                d[4] = str(int(qty))
                d[8] = str(round(area_ea * qty, 3))
                d[11] = dataPlants[part]
                cnf_file.write("\t".join(d))


def fileWorker(f):
    result = list()
    with open(f, "r") as prod_file:
        for line in prod_file.readlines():
            result.append(line.upper().split("\t"))
    return result


def findPartCompletionDate(part):
    conn = pyodbc.connect(
        "DRIVER={SQL Server};SERVER=HIIWINBL18;DATABASE=SNDBase91;UID=SNUser;PWD=BestNest1445")
    cursor = conn.cursor()
    cursor.execute("""
            SELECT ArcDateTime
            FROM PIPArchive
            WHERE PartName=?
            AND TransType='SN102'
        """, part.replace('-', '_', 1))

    response = [x[0].isoformat() for x in cursor.fetchall()]

    return response


def test():
    conn = pyodbc.connect(
        "DRIVER={SQL Server};SERVER=HIIWINBL18;DATABASE=SNDBase91;UID=SNUser;PWD=BestNest1445")
    cursor = conn.cursor()

    parts = [
        ("32290", "823D71490", "4500205649"),
        ("32291", "823D71490", "4500205649"),
        ("32321", "821D09260", "4500200448"),
        ("32324", "821D09260", "4500200448"),
        ("32323", "821D09260", "4500200448"),
        ("32320", "821D09260", "4500200448"),
        ("32292", "823D71490", "4500205649"),
        ("32322", "821D09260", "4500200448"),
        ("32329", "821D09260", "4500200448"),
        ("32333", "823D67540", "4500200448"),
    ]

    for x in parts:
        cursor.execute("""
            UPDATE StockHistory
            SET HeatNumber=?, BinNumber=?
            WHERE ProgramName=?
        """, x[1], x[2], x[0])

        cursor.execute("""
            UPDATE StockArchive
            SET HeatNumber=?, BinNumber=?
            WHERE ProgramName=?
        """, x[1], x[2], x[0])

    conn.commit()
    conn.close()


if __name__ == '__main__':
    main()
    # test()
