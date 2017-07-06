import sqlite3

try:
    conn = sqlite3.connect("FODAY SAHID.db")
    c = conn.cursor()
except Exception as e:
    print(str(e))
else:
    c.executescript(
            """CREATE TABLE CONTACTS(
                ContactId INTEGER PRIMARY KEY AUTOINCREMENT,
                ContactName text,
                ContactNumber text
                )
            """
            )
    c.execute("INSERT INTO CONTACTS(ContactName, ContactNumber) VALUES('FODAY S.N KAMARA', '+232-88-76-77-95')")
    conn.commit()
    conn.close()
    


