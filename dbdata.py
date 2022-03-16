import sqlite3
conn = sqlite3.connect("ALSbook.sqlite")

cursor = conn.cursor()

sql_query = """CREATE TABLE ALSs (
)
"""

cursor.execute(sql_query)