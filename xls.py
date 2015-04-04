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
        print (cell_obj.value)
    print "================"

#MySQL connector
conn = pymysql.connect(host='localhost', port=3306, user='root', passwd='', db='ormkt_db')

cur = conn.cursor()

SQL_column_names = ', '.join('?' * len(column_names))


statement = 'CREATE TEMPORARY TABLE IF NOT EXISTS import'
statement = 'INSERT INTO orders VALUES (%s);' %SQL_column_names

#insert_values = conn.escape('[{0}]'.format(insert_values))

print('INSERT VALUES')
print(insert_values)

#insertStatement = 'INSERT INTO orders (order_id, customer_id, customer_PO, order_type, order_date, booked_date)	VALUES {0}'.format(insert_values)

#statement = 'INSERT INTO orders (order_id, customer_id, customer_PO, order_type, order_date, booked_date) VALUES (%s, %s, %s, %s, %s, %s)'

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
