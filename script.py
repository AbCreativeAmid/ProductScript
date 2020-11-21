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
            connection = psycopg2.connect(user="",password="",host="",port="",database="")
            
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
            print("file not found! \n please put the script file in excel file directory")

        if fileExist:
            sheet = work_book.sheet_by_name("List of Items")
            groups = self.get_groups()
            connection = self.db_connect()
            imgpath = os.path.dirname(__file__)+"/images/"
            if groups and connection:
                cursor = connection.cursor()
                for i in range(1,sheet.nrows):
                    row = sheet.row(i)
                    item = {}
                    name            = "No Name" if row[1].value == '' else str(row[1].value)
                    part_no         = "No Part No" if  row[1].value == '' else str(row[2].value)
                    category_id     = groups["All"] if row[3].value =='' else groups[row[3].value]
                    default_code    =  "No default Code" if  row[1].value == '' else str(row[0].value)
                    record          = (name,category_id,part_no,default_code)
                
                    query = """INSERT INTO product_template(name,categ_id,part_no,default_code,type,uom_id,uom_po_id,tracking,active)
                                Values(%s,%s,%s,%s,'consu',1,1,'none',TRUE) returning id"""
                    cursor.execute(query,record)
                    id = cursor.fetchone()[0]
                    item["id"] = id
                    item["image_1920"] = ""
                    if row[4].value != '' and row[4].value != 'N/A':
                        if os.path.isfile(imgpath+row[4].value):
                            f = open(imgpath+row[4].value,"rb")
                            item["image_1920"] = base64.b64encode(f.read()).decode('utf-8')
                            
                    itemlist.append(item)

                connection.commit()  
                cursor.close()
                connection.close()
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
        except:
            print("not connected")
        
        if common:
            uid      = common.authenticate(db,username,password,{})
            models = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))
            items = self.read_excel()
            if len(items)>1:
                for item in items:
                    models.execute_kw(db,uid,password,"product.template","write",[[item["id"]],{"image_1920":item["image_1920"]}])

    
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