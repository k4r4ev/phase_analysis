from openpyxl.chart import Reference, ScatterChart, Series
from quasicycle import Quasicycle


def import_source_data(workbook, source_sheet, config):
    col = 1
    while True:
        if source_sheet.cell(row=1, column=col).value == config.source_col:
            break
        col += 1
    workbook.create_sheet(config.output_sheet, len(workbook.sheetnames) + 1)
    output_sheet = workbook[config.output_sheet]
    i = 1
    while source_sheet.cell(row=i + 1, column=col).value is not None:
        output_sheet.cell(i, 1).value = float(source_sheet.cell(i + 1, col).value)
        i += 1


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


def distance(point1, point2):
    square_x = (point2[0] - point1[0]) ** 2
    square_y = (point2[1] - point1[1]) ** 2
    dist = (square_x + square_y) ** 0.5
    return dist


def check_near(points_list, time_check, position, next_num, min_value):
    while time_check != 0:
        try:
            if distance(points_list[position], points_list[position + next_num + 1]) < min_value:
                min_value = distance(points_list[position], points_list[position + next_num + 1])
                next_num += 1
            time_check -= 1
        except IndexError:
            break
    return min_value, next_num


def get_quasicycles(sheet, config):
    quasicycles = []
    points_list = []
    row = 1
    while sheet.cell(row, 2).value is not None:
        points_list.append([sheet.cell(row, 1).value, sheet.cell(row, 2).value])
        row += 1
    position = 0
    q_index = 1
    while position + config.min_size < len(points_list) - 1:
        q_size = config.min_size
        min_value = distance(points_list[position], points_list[q_size + position])
        time_check = 3
        min_value, q_size = check_near(points_list, time_check, position, q_size, min_value)
        if q_size == q_size + time_check:
            while distance(points_list[position], points_list[position + q_size + 1]) < min_value:
                min_value = distance(points_list[position], points_list[position + q_size])
                q_size += 1
        quasicycles.append(Quasicycle(sheet, "Квазицикл " + str(q_index), position + 1, 1, q_size))
        position = position + q_size + 1
        q_index += 1
    return quasicycles
