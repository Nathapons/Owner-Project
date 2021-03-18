import sqlite3


def connect_sqlite():
    try:
        with sqlite3.connect('Prosatang.sqlite') as conn:
            query = '''create table Products(
                        ProductID INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
                        ProductName varchar(255),
                        ProductCategory varchar(10)
                    )'''
            cursor = conn.execute(query)
            print('Create DB')
    except Exception:
        print('Prosatang.sqlite has been created!!')

def insert_products():
    print('Welcome to Insert Program \n')

    product_name = ''
    product_category = ''
    while product_category == '' and product_category == '':
        product_name = input('ชื่อสินค้า : ')
        product_category = input('ผัก หรือ ผลไม้?: ')

    with sqlite3.connect("Prosatang.sqlite") as con:
        sql_cmd = f"insert into Products(ProductName,ProductCategory) values('{product_name}', '{product_category}');"
        con.execute(sql_cmd)
        print('Insert Data Complete')

def get_products():
    print('Welcome to Get Data Program \n')

    with sqlite3.connect("Prosatang.sqlite") as con:
        sql_cmd = "select * from Products"
        data_table = con.execute(sql_cmd)

        for row in data_table:
            print(row)

def update_products():
    print('Welcome to Update Program \n')
    ask = ''
    while ask.upper() not in ['YES', 'NO']:
        ask = input('คุณต้องการอัพเดทข้อมูลหรือไม่? => Yes Or No: ')
    
    if ask.upper() == 'YES':
        try:
            product_id = input('Update ID no.?: ')
            product_name = input('new productname: ')
            product_category = input('new product category(ผัก หรือผลไม้): ')

            with sqlite3.connect("Prosatang.sqlite") as con:
                sql_cmd = f"UPDATE Products SET ProductName='{product_name}', ProductCategory='{product_category}' WHERE ProductID={product_id};"
                con.execute(sql_cmd)
                print('Update Complete!!')
        except Exception:
            print('Update Error!!')
    else:
        print('End Update Program')

def delete_products():
    product_name = ''
    
    while product_name == '':
        product_name = input('Delete productname?: ')

    try:
        with sqlite3.connect("Prosatang.sqlite") as con:
            sql_cmd = f"delete from Products where ProductName = '{product_name}'"
            con.execute(sql_cmd)
            print('Delete Complete')
    except Exception:
        print(f'Product {product_name} is not in Table')


if __name__ == "__main__":
    connect_sqlite()
    insert_products()
    get_products()
    update_products()
    get_products()
    delete_products()
