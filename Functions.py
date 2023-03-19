from openpyxl.comments import Comment
from openpyxl.styles import PatternFill


def get_col_no(sheet, desc, max_col):
    """
    According to given active_sheet name and searched column description (str), function
    returns column number where that string is.
    Parameters
    ----------
    sheet : workbook.active
        active sheet name
    desc : str
        searched string
    max_col : int
        number of not empty columns
    """
    for s in range(max_col):
        if str(desc) == str(sheet.cell(row=1, column=s+1).value):
            return s+1
    return 0


def set_cell_color(sheet, row, col, color_rgn='n'):
    """
    Change cell background color on 'r'-red, 'g'-green, 'n'-none.
    Parameters
    ----------
    sheet : workbook.active
        active sheet name
    row : int
        cell row position
    col : int
        cell column position
    color_rgn : str, optional
        (OPTIONAL) color_rgn = 'r'-red; 'g'-green; 'n' or empty - NONE
    """
    if color_rgn == 'r':
        fg_color = 'E74C3C'
        pattern = PatternFill(patternType='solid', fgColor=fg_color)
        sheet.cell(row=row, column=col).fill = pattern
    if color_rgn == 'g':
        fg_color = '2ECC71'
        pattern = PatternFill(patternType='solid', fgColor=fg_color)
        sheet.cell(row=row, column=col).fill = pattern
    if color_rgn == 'n':
        pattern = PatternFill(patternType=None)
        sheet.cell(row=row, column=col).fill = pattern


def get_cell_value(sheet, row, col):
    """
    Returns value from cell in active_sheet at specific row & column position.
    Parameters
    ----------
    sheet : workbook.active
        active sheet name
    row : int
        cell row position
    col : int
        cell column position
    """
    return sheet.cell(row=row, column=col).value


def set_cell_value(sheet, row, col, val, typ=0):
    """
    Set value at cell in active_sheet at specific row & column position.
    Parameters
    ----------
    sheet : workbook.active
        active sheet name
    row : int
        cell row position
    col : int
        cell column position
    val : any
        value
    typ : int, optional
        (OPTIONAL) type of data (0 or empty - float/int; 1 - force string; 2 - None)
    """
    if typ == 0:
        sheet.cell(row=row, column=col).value = float(val)
    elif typ == 1:
        sheet.cell(row=row, column=col).value = str(val)
    elif typ == 2:
        sheet.cell(row=row, column=col).value = None
    elif typ == 3:
        sheet.cell(row=row, column=col).value = int(val)


def set_cell_comment(sheet, row, col, commentary, add=False, delete=False):
    """
    Function add 'commentary' to specified cell given as (active_sheet, row & column position)
    If optional parameter add=True then function add commentary to existing one.
    In another way removes old commentary and add a new one.
    Parameters
    ----------
    sheet : workbook.active
        active sheet name
    row : int
        cell row position
    col : int
        cell column position
    commentary : str
        commentary that will be added to specified cell
    add : bool, optional
        (OPTIONAL) add=False or NONE - function swap commentary into new on , add=True - add second commentary
    delete : bool, optional
        (OPTIONAL) delete=False or NONE - function do nothing with existing commentary, delete=True - remove commentary
    """
    if str(sheet.cell(row=row, column=col).comment) == 'None' or add is False:
        comment = Comment(commentary, 'CdA analyzer')
        comment.width = 400
        comment.height = 150
        sheet.cell(row=row, column=col).comment = comment
    else:
        tmp_txt1 = str(str(sheet.cell(row=row, column=col).comment).replace('Comment: ', '')).\
            replace('by CdA analyzer', '')
        comment = Comment(tmp_txt1 + ' ::\n' + commentary, 'CdA analyzer')
        comment.width = 400
        comment.height = 150
        sheet.cell(row=row, column=col).comment = comment
    if delete is True:
        sheet.cell(row=row, column=col).comment = None


