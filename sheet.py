import math
import pandas
import numbers

from xlsxrender.cell import Cell

class Sheet:

    """ Representation of Excel sheet

    Should to manipulate an Excel sheet in just one of
    Sheet class instance

    Args:
        book (xlsxwriter.Book): instance of xlsxwriter book
        sheetname (str): name of existing sheet in workbook or for create new one
        row: (int): current row with focus in excel sheet, default is 0
        col (int): current column with focus in excel sheet, default is 0


    Attributes:
        COLUMN_SIZE (int): start size for every column in excel sheet
    """

    COLUMN_SIZE = 9

    def __init__(self, book, sheetname, row=0, col=0):
        sheet = book.get_worksheet_by_name(sheetname)

        if sheet is None:
            book.add_worksheet(sheetname)
            sheet = book.get_worksheet_by_name(sheetname)

        self.sheet = sheet
        self.sheet.hide_gridlines(option=2)

        self._columns_size = []
        self.book = book
        self.row = row
        self.col = col

    def __write_table_title(self, irow, icol, title='', merge_until=0, title_bg='', title_fg='#007944'):
        """Write general title for combined cells or specific column

        Args:
            irow (int): row to focus
            icol (int): column to focus
            title (str, optional): Defaults to ''.
            merge_until (int, optional): Numbers of columns to merge for write title. Defaults to 0.
            title_bg (str, optional): Title background color. Defaults transparent.
            title_fg (str, optional): Title foreground. Defaults is primary theme color.
        """
        if merge_until in [0, 1]:
            self.sheet.write(
                irow, icol, title, self.book.add_format({
                    'align': 'center',
                    'font_color': title_fg if title_fg is not None else '#007944',
                    'bold': True,
                    'bg_color': title_bg
                })
            )
        else:
            self.sheet.merge_range(irow, icol, irow, icol + merge_until-1, title, self.book.add_format({
                'align': 'center',
                'font_color': title_fg if title_fg is not None else '#007944',
                'bold': True,
                'bg_color': title_bg
            }))

    def __resize_column(self, column_index, cell_value):
        for x in range(-1, column_index - len(self._columns_size)):
            self._columns_size.append(8)

        current_col_size = self._columns_size[column_index]

        value_size = (
            (len(str(cell_value)) * 20) / current_col_size
        )

        if value_size > current_col_size:
            if value_size > 20:
                self.sheet.set_column(self.row, column_index , 20)
            else:
                self.sheet.set_column(self.row, column_index, value_size)


    def get_totalcol(self, rows, cols, data):
        totalcolumn = [Cell('TOTAL GENERAL', header=True), ]
        for i in range(0, rows):
            sumatory = 0
            for j in cols:
                item = data[j][i+1]
                if item.value not in ['', None]:
                    sumatory += item.value
            
            totalcolumn.append(Cell(sumatory))

        return totalcolumn

    def next_row(self):
        self.row += 1

    def get_depth_y(self, data, depth=0):
        columns = list(data.columns.values)
        for col in columns:
            if isinstance(data[col][0].value, pandas.DataFrame):
                aux_depth = self.get_depth_y(data[col][0].value, depth + 1)
                depth = aux_depth if aux_depth > depth else depth
                return depth
        return depth

    def render(self, irow, icol, data, title=None, combine_title_cols=None, title_bg='#ffffff', title_fg=None):
        depth = self.get_depth_y(data)
        if title is not None:
            merge_until = len(list(data)) if combine_title_cols is None else combine_title_cols
            self.__write_table_title(
                irow, icol,
                title=title,
                merge_until=merge_until,
                title_bg=title_bg,
                title_fg=title_fg
            )
            self.__resize_column(icol, title)
            irow += 1

        columns = list(data.columns)

        col = icol
        irow += depth
        row = 0
        for (index, column) in enumerate(columns):
            row = irow      
            cells = data[column].tolist()

            for cell in cells:
                if cell is None:
                    continue
                if isinstance(cell.value, pandas.DataFrame):
                    self.render(
                        row-depth,
                        col,
                        cell.value,
                        title=cell.title,
                        title_bg=cell.title_bg,
                        title_fg=cell.title_fg
                    )
                    col += len(list(cell.value.columns)) - 1
                    break

                if isinstance(cell.value, numbers.Number):
                    self.sheet.write(
                        row, col,
                        cell.value if not math.isnan(cell.value) else 0,
                        self.book.add_format(cell.style)
                    )
                else:
                    self.sheet.write(
                        row, col,
                        cell.value,
                        self.book.add_format(cell.style)
                    )

                self.__resize_column(col, cell.value)
                row += 1
            col += 1

        self.row = row