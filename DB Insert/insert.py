from glob import glob
import random
import string
import mysql.connector

msg = ""


def insert():
    try:
        mydb = mysql.connector.connect(
            host="bm1cehzgazdmczql5toa-mysql.services.clever-cloud.com",
            user="ubucfxpk856ntdmg",
            password="s0U5zDcArkKxKVl1xGBh",
            database="bm1cehzgazdmczql5toa")

        mycursor = mydb.cursor()
        sql = "INSERT INTO Users (serial) VALUES (%s)"
        val = []
        for i in range(0, 100):
            x = ""
            while True:
                x = ''.join(random.choice(string.ascii_uppercase +
                            string.ascii_lowercase + string.digits) for _ in range(32))
                mycursor.execute("SELECT * FROM Users WHERE serial=%s", [x, ])
                if(mycursor.fetchall().__len__() == 0):
                    break

            val.append([x, ])

        mycursor.executemany(sql, val)
        mydb.commit()

        msg = "Done, Do You Want to do it again? (y, n): "
    except NameError:
        print(NameError)
        msg = "Error Happened, Do you want to try again? (y, n): "

    inpt = str(input(msg))
    if(inpt == "y"):
        insert()
    elif(inpt == "n"):
        exit()


insert()
