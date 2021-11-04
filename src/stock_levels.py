
import sndb
from collections import defaultdict

header_wid = 30

def main():
    levels = {
        "50/50W-0004": defaultdict(float),
        "50/50W-0006": defaultdict(float),
        "50/50W-0008": defaultdict(float),
        "50/50W-0010": defaultdict(float),
        "50/50W-0012": defaultdict(float),
        "50/50W-0014": defaultdict(float),
        "50/50W-0010": defaultdict(float),
    }

    with sndb.get_sndb_conn() as db:
        cursor = db.cursor()
        for key in levels:
            cursor.execute("""
                SELECT Location, Area
                FROM Stock
                WHERE PrimeCode=?
            """, key)

            for row in cursor.fetchall():
                levels[key][row.Location] += row.Area

    for key, value in levels.items():
        print("\n{1}\n{0:^{width}}\n{1}".format(key, "=" * header_wid, width=header_wid))
        for loc, area in value.items():
            print("  - {:>3}: {}".format(loc, area))

if __name__ == "__main__":
    main()
