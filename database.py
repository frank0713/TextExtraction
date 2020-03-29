# _*_ coding: UTF-8 _*_
# @version: python 3.7.4
# @author: frank0713


import sqlite3
import index


def create_db(index_name):
    """
    Create a sqlite database.
    :param index_name: A str. The name of database, named by user.
    """

    db_names = index_name + ".db"
    conn = sqlite3.connect(db_names)
    table_name = index_name
    c = conn.cursor()
    c.execute("CREATE TABLE %s (Name TEXT, Extension TEXT, CTime TEXT, \
    MTime TEXT, Path TEXT, Size TEXT, Text TEXT)" % table_name)
    conn.commit()
    conn.close()


def data2db(user_dirs, user_extensions, index_name):
    """
    Push the data into sqlite database.
    :param user_dirs: A list. User's selected directions for scanning. e.g.
    ["C:/text", "D:/Programs"]
    :param user_extensions: A list. User's selected formats for indexing. e.g.
    [".txt", ".xlsx"]
    :param index_name: A str. The name of current index setted by user.
    """

    information = index.indexing(user_dirs=user_dirs, user_extensions=user_extensions,
                                 index_name=index_name)
    db_names = str(index_name) + '.db'
    table_name = str(index_name)
    conn = sqlite3.connect(db_names)
    c = conn.cursor()
    for i in information:
        c.execute("INSERT INTO %s (Name, Extension, CTime, MTime, Path,\
         Size, Text) VALUES ('%s', '%s', '%s', '%s', '%s', '%s', '%s')"
                  % (table_name, i['Name'], i['Extension'], i['CTime'],
                     i['MTime'], i['Path'], i['Size'], i['Text']))
        conn.commit()
    conn.close()
