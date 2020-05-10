from openpyxl.chart import Reference, ScatterChart, Series


def create_diagram(quasicycle):
    chart = ScatterChart()
    chart.title = quasicycle.name
    chart.style = 2
    chart.x_axis.title = ''
    chart.y_axis.title = ''
    chart.legend = None
    rows_reference = Reference(quasicycle.sheet, min_col=quasicycle.start_cell_col,
                               min_row=quasicycle.start_cell_row, max_row=quasicycle.start_cell_row + quasicycle.size)
    cols_reference = Reference(quasicycle.sheet, min_col=quasicycle.start_cell_col + 1,
                               min_row=quasicycle.start_cell_row, max_row=quasicycle.start_cell_row + quasicycle.size)
    chart.x_axis.scaling.min = quasicycle.row_min - 10
    chart.y_axis.scaling.min = quasicycle.col_min - 10
    chart.x_axis.scaling.max = quasicycle.row_max + 10
    chart.y_axis.scaling.max = quasicycle.col_max + 10
    series = Series(cols_reference, rows_reference, title_from_data=True)
    chart.layoutTarget = "inner"
    chart.series.append(series)
    return chart


def calculate_derivative(sheet, config):
    current_row = config.start_row
    while sheet.cell(current_row + 1, config.start_col).value is not None:
        sheet.cell(current_row, config.start_col + 1).value = sheet.cell(current_row + 1, config.start_col).value
        current_row += 1
