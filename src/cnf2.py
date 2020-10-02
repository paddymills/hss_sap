import parsers
import xlwings

from collections import defaultdict

TR_COLS = 17
TR_ROWS = 25


def main():
    determine_tr()


def determine_tr():
    cogi = defaultdict(list)
    mb52 = defaultdict(list)

    for wb in xlwings.books:
        for s in wb.sheets:
            data = parsers.parse_sheet(s)

            ##############################################
            #                    COGI                    #
            ##############################################
            if s.range("A1").value == 'Processing Status':
                for row in data:
                    cogi[row.matl].append(row)

            ##############################################
            #                    MB52                    #
            ##############################################
            elif s.range("A1").value == 'Material Number':
                for row in data:
                    mb52[row.matl].append(row)

    tr = list()
    for k in cogi.keys():
        for res in determine_tr_for_parts(cogi[k], mb52[k]):
            if res:
                tr.append(res)

    # ^^^ rewritten above ^^^
    # for part_key, part_items in cogi.items():
    #     for inv_key, inv_items in mb52.items():
    #         if part_key == inv_key:
    #             for res in determine_tr_for_parts(part_items, inv_items):
    #                 if res:
    #                     tr.append(res)

    # insert blank row every TR_ROWS
    index = TR_ROWS
    while index < len(tr):
        tr.insert(index, [None] * TR_COLS)
        index += TR_ROWS + 1

    wb = xlwings.books.add()
    s = wb.sheets[0]
    s.range("A1").value = tr
    s.autofit('c')
    s.range("B:B").column_width = 0.1
    s.range("G:M").column_width = 0.1
    s.range("P:P").color = (0, 0, 0)
    wb.save(r"C:\temp\migo_tr.xlsx")  # masks save error on close


def determine_tr_for_parts(parts, inventory):
    keep_wbs = [p.wbs for p in parts]
    can_move = [(i.wbs, i.qty) for i in inventory if i.wbs not in keep_wbs]

    tr = list()  # (from, to, qty)
    for p in parts:
        for i in inventory:
            if i.wbs == p.wbs:
                p.qty -= i.qty

        while p.qty > 0 and can_move:
            wbs, qty = can_move.pop()
            if qty > p.qty:
                can_move.append((wbs, qty - p.qty))
                qty = p.qty
            tr.append(tr_format(p, wbs, qty))
            p.qty -= qty

    return tr


def tr_format(part, from_wbs, qty):

    row = [None] * TR_COLS
    row[0] = part.matl
    row[2] = part.plant
    row[3] = "PROD"
    row[4] = qty
    row[5] = "EA"
    row[13] = "PROD"
    row[14] = part.wbs
    row[16] = from_wbs

    return row


if __name__ == "__main__":
    main()
