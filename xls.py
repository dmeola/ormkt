# Standard Library Imports

# Related Third Party Imports
import xlrd
import pymysql


book = xlrd.open_workbook("sales.xlsx")
sh = book.sheet_by_index(0)

insert_values = []
row_values = ''
column_names = {'Ord No','Booked Date','Order Date','Cust PO','Order Type','Hospital PO','Loc No','Site No','Cust No','Ship To Name','Item No','Description','Qty','Total','Dcode'}

#skip first 11 rows because it doesn't contain relevant data
for row_idx in range(11, 13):
    insert_values.append([])					# add a dimension for row
    for col_idx in range(0, sh.ncols):  				# Iterate through columns
        cell_obj = sh.cell(row_idx, col_idx)  			# Get cell object by row, col
        insert_values[row_idx-11].append(cell_obj.value)

#MySQL connector
conn = pymysql.connect(host='localhost', port=3306, user='root', passwd='', db='ormkt_db')

cur = conn.cursor()

SQL_column_names = ', '.join('?' * len(column_names))

#creating tempory table for direct import from xls
createTemp = 'CREATE TEMPORARY TABLE IF NOT EXISTS import(Ord_No INT, \
    Booked_Date DATE, \
    Order_Date DATE, \
    Cust_PO INT, \
    Order_Type VARCHAR(45), \
    Hospital_PO INT, \
    Loc_No INT, \
    Site_No INT, \
    Cust_No INT, \
    Ship_To_Name VARCHAR(45), \
    Item_No INT, \
    Description VARCHAR(45), \
    Qty INT, \
    Total INT, \
    Dcode VARCHAR(45))'
cur.execute(createTemp)

#inserting xls data into temporary table
for counter in range(len(insert_values)):
    params = ['?' for item in insert_values[counter]]
    values = ", ".join([str(val) for val in insert_values[counter]])
    print insert_values[counter]
    statement = "INSERT INTO import (Ord_No,Booked_Date,Order_Date,Cust_PO,Order_Type,Hospital_PO,Loc_No,Site_No,Cust_No,Ship_To_Name,Item_No,Description,Qty,Total,Dcode) VALUES ({0});".format(params)

    #cur.executemany(statement, insert_values[counter])
    cur.executemany(statement, ('493773','13-Mar-15','13-Mar-15','PO982628','ILS Standard Order','test','71209','144755','108222','HAYMARKET MEDICAL CENTER HAYMARKET VA 20169','225444','HOHMANN RETR 6-1/4 15MM WIDE','6','375','D1741'))

statement = 'INSERT INTO orders VALUES (%s);' %SQL_column_names

#insert_values = conn.escape('[{0}]'.format(insert_values))

print('INSERT VALUES')
print(insert_values)


#insertStatement = conn.escape(insertStatement)

print('STATEMENT')
print(statement)

cur.executemany(statement, insert_values)

cur.execute('SELECT * FROM orders')

print(cur.description)

print()

for row in cur:
   print(row)

cur.close()
conn.close()
