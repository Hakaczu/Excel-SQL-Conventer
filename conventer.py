from openpyxl import load_workbook

#TODO: wyczyscic kod 
# config
table_name = 'sprzet2020'
create_stmt = False
first_row_as_col_names = True
dest_file_name = 'x.xlsx'
output_file_name = 'create.sql'
#end config

def createTable(table_name, columns):
    create = 'CREATE TABLE ' + table_name + '(id INT AUTO_INCREMENT PRIMARY KEY, '
    col_len = len(columns)
    print(col_len)
    i = 1
    for x in columns:
        if(i == col_len):
            stmt = x + ' VARCHAR(2000)'
        else:
            stmt = x + ' VARCHAR(2000), '
        create += stmt
        i = i + 1
    create += ');\n'
    return create

def getColumnsNames(sheet, first_row_as_col_names):
    columns = []

    if first_row_as_col_names == True:
        for x in range(1, sheet.max_column + 1):
            value = str(sheet.cell(row = 1, column = x).value)
            columns.append(value)
    else:
        for x in range(1, sheet.max_column + 1):
            columns.append(str('Col' + x))
    return columns

def insert(table_name, columns, values):
    insert = 'INSERT INTO ' + table_name + '('
    col_len = len(columns)
    i = 1
    for col in columns:
        if(i==col_len):
            insert += col
        else:
            insert += col + ', '
        i = i + 1
    insert += ') VALUES('
    i = 1
    for val in values:
        if(i==col_len):
            insert += "'"+ val +"'"
        else: 
            insert += "'"+ val +"',"
        i = i + 1
    insert += ');\n'
    return insert


# load workbook
wb = load_workbook(filename = dest_file_name)
# load active sheet
ws = wb.active
max_col = ws.max_column
max_row = ws.max_row
convert = ''

print("MAX COL: " + str(max_col))
print("MAX ROW: " + str(max_row))

columnsNames = getColumnsNames(ws, first_row_as_col_names)

if create_stmt == True:
    create = createTable(table_name, columnsNames)
    convert += create

if first_row_as_col_names == True:
    
    for row in range(2, max_row + 1):
        val = []
        for col in range (1, max_col + 1):
            print(str(ws.cell(row = row, column = col).value))
            val.append(str(ws.cell(row = row, column = col).value))
        ins = insert(table_name, columnsNames, val)
        convert += ins

f = open(output_file_name, "w")
f.write(convert)
f.close()