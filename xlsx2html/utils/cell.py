import re
import openpyxl

CELL_LOCATION_RE = re.compile(
    r"""
    ^
        \#
        (?:
            (?P<sheet_name>\w+)[.!]
            |
        )
        (?P<coord>[A-Za-z]+[\d]+)
    $
    """,
    re.VERBOSE,
)


def parse_cell_location(cell_location):
    """
    >>> parse_cell_location("#Sheet1.C1")
    {'sheet_name': 'Sheet1', 'coord': 'C1'}
    >>> parse_cell_location("#Sheet1!C1")
    {'sheet_name': 'Sheet1', 'coord': 'C1'}
    >>> parse_cell_location("#C1")
    {'sheet_name': None, 'coord': 'C1'}
    """

    m = CELL_LOCATION_RE.match(cell_location)
    if not m:
        return None

    return m.groupdict()

def is_formula_range(formula):
    try:
        openpyxl.utils.cell.range_to_tuple(formula)
        return True
    except:
        return False

def get_values_for_range(workbook, ref_string):
    [sheet_name, range_data] = openpyxl.utils.cell.range_to_tuple(ref_string)
    [start_column, start_row, end_column, end_row] = range_data
    sheet = workbook[sheet_name] or workbook.active
    values = []
    for row_index in range(start_row, end_row + 1):
        row = sheet[row_index]
        for cell in row[start_column - 1:end_column]:
            values.append(str(cell.value))
    return values

def dropdown_options_for_cell(workbook, sheet, cell):
    options = None

    # Go find the answer options
    for data_validation in sheet.data_validations.dataValidation:
        if cell.coordinate in data_validation.sqref:
            formula = data_validation.formula1

            # Sometimes, the formula is None. If so, we can't do anything with it. Just skip it.
            if formula is None:
                continue

            defined_name = workbook.defined_names[formula] if formula in workbook.defined_names else None

            if defined_name is not None and defined_name.type == 'RANGE':
                options = get_values_for_range(workbook, defined_name.attr_text)

            elif is_formula_range(formula):
                options = get_values_for_range(workbook, formula)
            else:
                # Strip leading and trailing quotes, if present
                if formula[0] == '"':
                    formula = formula[1:]
                if formula[-1] == '"':
                    formula = formula[:-1]

                options = formula.split(',')
            break

    return options