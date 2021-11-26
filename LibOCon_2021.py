'''
Module to demonstrate SourceForge Methods in the Calc Service	
	
Activate	in create_sheet_example
ClearAll	in clear_contents_v1
ClearFormats	in clear_contents_v2
ClearValues	in clear_contents_v3
CopySheet	in copy_sheet_example
CopySheetFromFile	in copy_from_file_example
CopyToCell	in copy_cells_v1
CopyToRange	in copy_cells_v2
DAvg	in calculate_average
DCount	see Davg
DMax	see Davg
DMin	see Davg
DSum	see Davg
Forms	
GetColumnName	see GetValue
GetFormula	see GetValue
GetValue	in mark_invalid
ImportFromCSVFile	in open_csv_file_v1
ImportFromDatabase	see ImportFromCSVFile
InsertSheet	in create_sheet_example
MoveRange	see CopyToRange
MoveSheet	see InsertSheet 
Offset	in create_random_matrix_v1
RemoveSheet	in remove_sheet_example
RenameSheet	see InsertSheet 
SetArray	in create_random_matrix_v2
SetValue	in create_random_matrix_v1
SetCellStyle	in mark_invalid
SetFormula	see SetValue
SortRange	'''


from scriptforge import CreateScriptService
import random as rnd
import pathlib

# Creates a 6x6 matrix starting at A1
def create_random_matrix_v1(args=None):
    doc = CreateScriptService("Calc")
    for i in range(6):
        for j in range(6):
            target_cell = doc.Offset("A1", i, j)
            r = rnd.random()
            if r < 0.5:
                doc.setValue(target_cell, "EVEN")
            else:
                doc.setValue(target_cell, "ODD")

# Creates a 6x6 matrix starting at A1
# Uses the method setArray to insert values
def create_random_matrix_v2(args=None):
    doc = CreateScriptService("Calc")
    rnd_word = lambda : "EVEN" if rnd.random() < 0.5 else "ODD"
    values = [[rnd_word() for _ in range(6)] for _ in range(6)]
    doc.setArray("A1", values)

# Creates an mxn matrix starting at A1 and asks the desired size
def create_random_matrix_v3(args=None):
    doc = CreateScriptService("Calc")
    bas = CreateScriptService("Basic")
    n_rows = bas.InputBox("Number of rows")
    n_cols = bas.InputBox("Number of columns")
    for i in range(int(n_rows)):
        for j in range(int(n_cols)):
            target_cell = doc.Offset("A1", i, j)
            r = rnd.random()
            if r < 0.5:
                doc.setValue(target_cell, "EVEN")
            else:
                doc.setValue(target_cell, "ODD")

# Creates an mxn matrix starting at A1 and asks the desired size
# Uses the method setArray to insert values
def create_random_matrix_v4(args=None):
    doc = CreateScriptService("Calc")
    bas = CreateScriptService("Basic")
    n_rows = bas.InputBox("Number of rows")
    n_cols = bas.InputBox("Number of columns")
    rnd_word = lambda : "EVEN" if rnd.random() < 0.5 else "ODD"
    values = [[rnd_word() for _ in range(int(n_cols))]
              for _ in range(int(n_rows))]
    doc.setArray("A1", values)

# Clear region starting at A1
def clear_region_a1(args=None):
    doc = CreateScriptService("Calc")
    cur_sheet = XSCRIPTCONTEXT.getDocument().CurrentController.ActiveSheet
    cell = cur_sheet.getCellRangeByName("A1")
    cursor = cur_sheet.createCursorByRange(cell)
    cursor.collapseToCurrentRegion()
    doc.clearAll(cursor.AbsoluteName)

# Creates a matrix of size 10x8 with random integers between -20 and 100
def create_values_for_example_2(args=None):
    clear_region_a1()
    doc = CreateScriptService("Calc")
    data = [[rnd.randint(-20, 100) for _ in range(8)] for _ in range(10)]
    doc.setArray("A1", data)

# Marks cells with negative values as INVALID and apply the 'Bad' cell style
def mark_invalid(args=None):
    doc = CreateScriptService("Calc")
    # Gets address of current selection
    cur_selection = doc.CurrentSelection
    # Gets address of first cell in the selection
    first_cell = doc.Offset(cur_selection, 0, 0, 1, 1)
    for i in range(doc.Height(cur_selection)):
        for j in range(doc.Width(cur_selection)):
            cell = doc.Offset(cur_selection, i, j, 1, 1)
            value = doc.getValue(cell)
            if value < 0:
                doc.setValue(cell, "INVALID")
                doc.setCellStyle(cell, "Bad")

# Example of using ClearAll
def clear_contents_v1(args=None):
    doc = CreateScriptService("Calc")
    doc.clearAll("B2:B7")

# Example of using ClearFormats
def clear_contents_v2(args=None):
    doc = CreateScriptService("Calc")
    doc.clearFormats("D2:D7")

# Example of using ClearValues
def clear_contents_v3(args=None):
    doc = CreateScriptService("Calc")
    doc.clearValues("F2:F7")

# Copying to a single cell
def copy_cells_v1(args=None):
    doc = CreateScriptService("Calc")
    doc.copyToCell("A1:A4", "C1")

# Copying cells into a larger range
def copy_cells_v2(args=None):
    doc = CreateScriptService("Calc")
    doc.copyToRange("A1:A4", "E1:F6")

# Copies range from an open file
def copy_range_from_file(args=None):
    # Reference to current document (destination)
    doc = CreateScriptService("Calc")
    # Reference to the source document
    svc = CreateScriptService("UI")
    source_doc = svc.getDocument("DataSource.ods")
    source_range = source_doc.Range("Sheet1.A1:A5")
    # Pastes the contents into the destination
    doc.copyToCell(source_range, "A1")

# Inserting new sheet
def create_sheet_example(args=None):
    doc = CreateScriptService("Calc")
    doc.insertSheet("TestSheet", 2)
    doc.activate("TestSheet")

# Copying an existing sheet
def copy_sheet_example(args=None):
    doc = CreateScriptService("Calc")
    doc.copySheet("TestSheet", "Copy_TestSheet")

# Removing a sheet
def remove_sheet_example(args=None):
    doc = CreateScriptService("Calc")
    doc.removeSheet("Copy_TestSheet")

# Copies sheet from another file (open or closed)
def copy_from_file_example(args=None):
    doc = CreateScriptService("Calc")
    wb = str(pathlib.Path.home().joinpath("Documents", "DataSource.ods"))
    doc.copySheetFromFile(wb, "Sheet2", "Copy_Sheet2")

# Example using the DAvg method
def calculate_average(args=None):
    doc = CreateScriptService("Calc")
    bas = CreateScriptService("Basic")
    result = doc.DAvg("A1:E1")
    bas.MsgBox("The average is {:.02f}".format(result))

# Open CSV file JobData_v1.csv using default configuration
def open_csv_file_v1(args=None):
    doc = CreateScriptService("Calc")
    csvfile = str(pathlib.Path.home().joinpath("Documents", "JobData_v1.csv"))
    doc.ImportFromCSVFile(csvfile, "A1")

# Open CSV file JobData_v2.csv using default configuration
def open_csv_file_v2(args=None):
    doc = CreateScriptService("Calc")
    csvfile = str(pathlib.Path.home().joinpath("Documents", "JobData_v2.csv"))
    doc.ImportFromCSVFile(csvfile, "A1")

# Open CSV file using custom configuration
def open_csv_file_v3(args=None):
    doc = CreateScriptService("Calc")
    csvfile = str(pathlib.Path.home().joinpath("Documents", "JobData_v2.csv"))
    filter_option = "59,34,UTF-8,1"
    doc.ImportFromCSVFile(csvfile, "A1", filter_option)

g_exportedScripts = (
    create_random_matrix_v1,
    create_random_matrix_v2,
    create_random_matrix_v3,
    create_random_matrix_v4,
    clear_region_a1,
    mark_invalid,
    create_values_for_example_2,
    clear_contents_v1,
    clear_contents_v2,
    clear_contents_v3,
    copy_cells_v1,
    copy_cells_v2,
    copy_range_from_file,
    create_sheet_example,
    copy_sheet_example,
    remove_sheet_example,
    copy_from_file_example,
    calculate_average,
    open_csv_file_v1,
    open_csv_file_v2,
    open_csv_file_v3,
	)
