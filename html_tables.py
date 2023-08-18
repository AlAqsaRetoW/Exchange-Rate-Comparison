from bs4 import BeautifulSoup
from RPA.Tables import Table

def read_table_from_html(html_table: str) -> Table:
    """Parses and returns the given HTML table as a Table structure.

    :param html_table: Table HTML markup.
    """
    table_rows = []
    soup = BeautifulSoup(html_table, "html.parser")
    # table header row
    for table_row in soup.select('tr'):
        cells = table_row.find_all('th')

        if len(cells) > 0:
            cell_values = [cell.text.strip() for cell in cells]
            table_rows.append(cell_values)

    # table rows
    for table_row in soup.select('tr'):
        cells = table_row.find_all('td')

        if len(cells) > 0:
            cell_values = [cell.text.strip() for cell in cells]
            table_rows.append(cell_values)

    return Table(table_rows)
