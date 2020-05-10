class Quasicycle:

    def __init__(self, sheet, name, start_cell_row, start_cell_col, size):
        self.sheet = sheet
        self.name = name
        self.start_cell_row = start_cell_row
        self.start_cell_col = start_cell_col
        self.row_values = []
        self.col_values = []
        self.size = size
        self.row_max = 0
        self.row_min = 0
        self.row_average = 0
        self.col_max = 0
        self.col_min = 0
        self.col_average = 0
        self.square = 0
        self.calculate_parameters()

    def calculate_parameters(self):
        for row_item in range(self.start_cell_row, self.start_cell_row + self.size):
            self.row_values.append(self.sheet.cell(row_item, self.start_cell_col).value)
            self.col_values.append(self.sheet.cell(row_item, self.start_cell_col + 1).value)
        self.row_max = max(self.row_values)
        self.row_min = min(self.row_values)
        self.row_average = (self.row_max + self.row_min) / 2
        self.col_max = max(self.col_values)
        self.col_min = min(self.col_values)
        self.col_average = (self.col_max + self.col_min) / 2
        self.square = (self.row_max - self.row_min) * (self.col_max - self.col_min)
