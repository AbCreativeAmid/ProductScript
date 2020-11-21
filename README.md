# ProductScript
It's an script for bulk insert of products from an excel file.
# Script Structure
This script two files one for making distinct the product category and inserting them in database. Becuase we need once to insert the categories then I get all of them and insert them to the database seperately.In this file we have two functions one for db connection the other one for inserting group.
The other file is for inserting the products. which contains a class by the name of 'Write Products'
in the Class we have four methods.
## 1: db_connect() 
This function is for connection to the postgres database.
## 2: read_excel()
This function is for reading all the records from the excel file and inserting them to the database except of thier images. The images just encoded and puted beside of it's appropriat id in a dictionary.
## 3: get_groups()
This methods is for reading the groups from database and converting them to a dictionary.
## 4: conned_odoo
This function connect to odoo project through the XMLRPC and gets all the products ids with thier encoded image file and then update images through write method of product_template method.

# Odoo Changes
For adding the products we needed to add a field for product.templated. Becuase of this we added a model.py and add the field for the model.

# Dependencies
for Completing this Script we used some of libraries.
## 1: xlrd for reading excel file 
```pip3 install xlrd```
## 2: os for directory management
```import os```
## 3: psycopg2 for database connection
```pip3 install psycopg2```
## 
