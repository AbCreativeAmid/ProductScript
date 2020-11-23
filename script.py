import os
import xlrd
import base64
import time
import psycopg2
import xmlrpc.client
class Write_products():

    def db_connect(self):
        """ connect to postgres database
            
            returns:
            database connection or None
        """
        connection = None

        try:
            connection = psycopg2.connect(user="ubuntu",password="12345",host="127.0.0.1",port="5432",database="script")
            
        except (Exception, psycopg2.Error) as error :
            print ("Error while connecting to PostgreSQL", error)

        return connection      

    def read_excel(self):
        """ read the records and write them into database
            
            returns:
            list of image,id of records
        """
        itemlist = []
        fileExist = False
        try:
            work_book = xlrd.open_workbook(os.path.dirname(__file__)+"/items.xlsx")
            fileExist = True
        except:
            print("done")

        if fileExist:
            sheet = work_book.sheet_by_name("List of Items")
            groups = self.get_groups()
            imgpath = os.path.dirname(__file__)+"/images/"
            if groups:
                for i in range(1,500):
                    row = sheet.row(i)
                    item = {}
                    item["name"]            = "No Name" if row[1].value == '' else str(row[1].value)
                    item["part_no"]         = "No Part No" if  row[1].value == '' else str(row[2].value)
                    item["categ_id"]        = groups["All"] if row[3].value =='' else groups[row[3].value]
                    item["default_code"]    = "No default Code" if  row[1].value == '' else str(row[0].value)
                    item["image_1920"] = ""
                    if row[4].value != '' and row[4].value != 'N/A':
                        if os.path.isfile(imgpath+row[4].value):
                            f = open(imgpath+row[4].value,"rb")
                            item["image_1920"] = base64.b64encode(f.read()).decode('utf-8')
                            
                    itemlist.append(item)
        return itemlist    
        

    def connect_odoo(self):
        """ connect to odoo through xml rpc and update the images of products """
        url ="http://127.0.0.1:8000"
        db = "script"
        username = "admin"
        password = "admin"
        common = None
        try:
            common   = xmlrpc.client.ServerProxy('{}/xmlrpc/2/common'.format(url))
        except Exception as error:
            print(error)
            
        
        if common:
            uid      = common.authenticate(db,username,password,{})
            models = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))
            items = self.read_excel()
            iterates = len(items)//25

            for i in range(iterates+1):
                st = i*25
                end = st+25
                if i == iterates:
                    insert_items = items[st:]
                else:
                    insert_items = items[st:end]
                models.execute_kw(db,uid,password,"product.template","create",[insert_items])
    
    def get_groups(self):
        """ read the categories or groups from the database
            
            returns:
            group list 
        """
        connection = self.db_connect()
        if connection:
            cursor = connection.cursor()
            query = "SELECT name,id FROM product_category"
            cursor.execute(query)
            records = cursor.fetchall()
            return dict(records)
        else:
            print("connection not supplied")
            return None
    
start = time.time()
Write_products = Write_products()
Write_products.connect_odoo()
print(time.time()-start)