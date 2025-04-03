
def to_table(data, document):
    total_columns = len(data.columns)
    table = document.add_table(rows =1, cols = total_columns)
    table.style = "Table Grid"
    header_cells = table.rows[0].cells

    for index, columns in enumerate(data):
        header_cells[index].text = columns

    for _, row in data.iterrows():
        row_cells = table.add_row().cells
        for idx, values in enumerate(row):
            row_cells[idx].text = str(values)