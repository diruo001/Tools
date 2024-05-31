from operator import le
from bs4 import BeautifulSoup
import re


class Table:
    def __init__(self) -> None:
        self.caption = None
        self.headers = None
        self.rows = []

    def shape(self):
        if len(self.rows) == 0:
            return 0, 0
        else:
            return len(self.rows), len(self.rows[0])


def extract_tables(filepath, table_names):
    soup = BeautifulSoup(open(filepath), features="lxml")
    table_list = []
    existed_table_names = []
    for table in soup.find_all("table"):
        current_table = Table()
        current_table.caption = table.caption.contents
        table_name = re.search(r"Table \d", table.caption.contents[0]).group()
        if table_name not in table_names or table_name in existed_table_names:
            continue
        existed_table_names.append(table_name)
        if table.thead is not None:
            current_table.headers = [h.contents for h in table.thead.tr.find_all()]    
        for i, row in enumerate(table.find_all("tr")):
            if table.thead is not None and i == 0:
                continue
            row_cells = [cell.contents for cell in row.find_all("td")]
            current_table.rows.append(row_cells)
        table_list.append(current_table)
    return table_list


if __name__ == "__main__":
    filename = "Lnibt1_La_tables.html"
    table_names = ["Table 1", "Table 4", "Table 5", "Table 6", "Table 8"]
    table_list = extract_tables(filepath=filename, table_names=table_names)
    print(table_list)