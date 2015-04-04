# Standard Library Imports

# Related Third Party Imports
import xlrd
import pymysql


book = xlrd.open_workbook("sales.xlsx")
#print "The number of worksheets is", book.nsheets
#print "Worksheet name(s):", book.sheet_names()
sh = book.sheet_by_index(0)
#print sh.name, sh.nrows, sh.ncols
#print "Cell D30 is", sh.cell_value(rowx=29, colx=3)

insert_values = ''
row_values = ''
column_names = {'Ord No','Booked Date','Order Date','Cust PO','Order Type','Hospital PO','Loc No','Site No','Cust No','Ship To Name','Item No','Description','Qty','Total','Dcode'}

#skip first 11 rows because it doesn't contain relevant data
for row_idx in range(11, 13): #sh.nrows):
    for col_idx in range(0, sh.ncols):  				# Iterate through columns
        cell_obj = sh.cell(row_idx, col_idx)  			# Get cell object by row, col
        row_values += "'{0}',".format(cell_obj.value)
        print (cell_obj.value)
    insert_values += '({0}),'.format(row_values)
    print "================"

    #insert_values+='({0}),'.format(sh.row(rx))

#MySQL connector
conn = pymysql.connect(host='localhost', port=3306, user='root', passwd='', db='ormkt_db')

cur = conn.cursor()

statement = 'CREATE TEMPORARY TABLE IF NOT EXISTS import'
statement = 'INSERT INTO orders (order_id, customer_id, customer_PO, order_type, order_date, booked_date) VALUES (%s, %s, %s, %s, %s, %s)'

insert_values = conn.escape('[{0}]'.format(insert_values))

print('INSERT VALUES')
print(insert_values)

#insertStatement = 'INSERT INTO orders (order_id, customer_id, customer_PO, order_type, order_date, booked_date)	VALUES {0}'.format(insert_values)

statement = 'INSERT INTO orders (order_id, customer_id, customer_PO, order_type, order_date, booked_date) VALUES (%s, %s, %s, %s, %s, %s)'

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
