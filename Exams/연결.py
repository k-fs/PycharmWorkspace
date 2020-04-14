import pyodbc 
# Some other example server values are
# server = 'localhost\sqlexpress' # for a named instance
# server = 'myserver,port' # to specify an alternate port
server = 'tcp:172.20.200.252,8629' 
cnxn = pyodbc.connect('DSN=TDB;UID=icam;PWD=kfam1801')
cursor = cnxn.cursor()

#Sample select query
cursor.execute("select user_id, username from dba_users")
row = cursor.fetchone()
if row:
    print(row)

cursor.execute("select user_id, username from dba_users")
rows = cursor.fetchall()
for row in rows:
    print(row.user_id, row.username)


#cursor.execute("delete from products where id <> ?", 'pyodbc')
#print('Deleted {} inferior products'.format(cursor.rowcount))
#cnxn.commit()
