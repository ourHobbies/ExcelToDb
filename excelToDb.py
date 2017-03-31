"""testing data upload
"""
import xlrd
import MySQLdb

# pylint:   disable=invalid-name
# pylint: disable=C0103

book = xlrd.open_workbook("pythonInputExcel.xlsx")

sheet = book.sheet_by_name("inputSheet")

database = MySQLdb.Connect(host="localhost", user="root", passwd="Baki@mysql17", db="pythondb")

cursor = database.cursor()

# Create the INSERT INTO sql query
query = """INSERT INTO pythontable (name, age, technology) VALUES (%s, %s, %s)"""

# Create a For loop to iterate through each row in the XLS file,
# starting at row 2 to skip the headers
for r in range(0, sheet.nrows):
    name = sheet.cell(r, 0).value
    age = sheet.cell(r, 1).value
    tech = sheet.cell(r, 2).value

    # Assign values from each row
    values = (name, age, tech)

    # Execute sql Query
    cursor.execute(query, values)

# Close the cursor
cursor.close()

# Commit the transaction
database.commit()

# Close the database connection
database.close()

# Print results
print("")
print("All Done! Bye, for now.")
print("")
columns = str(sheet.ncols)
rows = str(sheet.nrows)
#print("I just imported " %2B columns %2B " columns and " %2B rows %2B " rows to MySQL!")
