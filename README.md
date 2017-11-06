# Connect sql server and Excel by Python
1.关于实现通过python访问sql server中的数据并将数据存储在excel表中代码。
import pypyodbc
import xlwt


class Connect:
    pass


def save_sql_data(driver, server, database, pid, uid):
    # frist connect sql
    con = pypyodbc.connect(DRIVER=driver, SERVER=server, DATABASE=database, UID=pid, PWD=uid)
    cur = con.cursor()
    sql = '''select * from customers_test'''
    cur.execute(sql)
    list1 = cur.fetchall()
    cur.close()
    con.close()
    # second connect excel
    workbook = xlwt.Workbook()
    booksheet = workbook.add_sheet('savadata_sheet')
    params = ['CustomerNo', 'CustomerName', 'Address1', 'Address2', 'City', 'State', 'Zip', 'Contact', 'Phone', 'FedIDNo', 'DateInSystem']
    for index in range(len(params)):
        booksheet.write(0, index, params[index])

    for i in range(len(list1)):
        hang = list1[i]
        for j in range(len(hang)):
            booksheet.write(i+1, j, hang[j])
    workbook.save('practice.xls')

if __name__ == '__main__':
    save_sql_data('{SQL SERVER}', 'localhost', 'jdksql', 'sa', '123456')
