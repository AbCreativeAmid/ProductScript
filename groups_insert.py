import xlrd
import os
import psycopg2
def db_connect():
        connection = None

        try:
            connection = psycopg2.connect(user="",password="",host="",port="",database="")
            
        except (Exception, psycopg2.Error) as error :
            print ("Error while connecting to PostgreSQL", error)

        return connection

def insert_groups():
    """ read the groups from excel and distinct them and insert them to the database
        """
        os.chdir(os.path.dirname(__file__))
        work_book = xlrd.open_workbook("items.xlsx")
        sheet = work_book.sheet_by_name("List of Items")
        groups = []

        for i in range(sheet.nrows):
            row = sheet.row(i)
            item = ()
            if row[3].value != '':
                item = (str(row[3].value),)
                if item not in groups:
                    groups.append(item)
        connection = db_connect()
        if connection:
            cursor       = connection.cursor()
            for g in groups:
                insert_query = """INSERT INTO product_category(name) VALUES(%s)"""
                cursor.execute(insert_query,g)
            connection.commit()
            cursor.close()
            connection.close()
            print(" Groups Successfully inserted")
insert_groups()