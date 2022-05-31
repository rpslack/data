import cx_Oracle
import pandas as pd

def Table_GET(tbl_name, db):
    cursor = db.cursor()
    cursor.execute("SELECT * FROM all_col_comments WHERE table_name = '%s'" % tbl_name)
    colname = cursor.description
    col = []

    for i in colname:
        col.append(i[0])

    row = cursor.fetchall()
    df_col = pd.DataFrame(row, columns = col)

    col_dic = dict()
    for i in range(len(df_col)):
        col_dic[df_col.loc[i, 'COLUMN_NAME']] = df_col.loc[i, 'COMMENTS']

    cursor.execute("SELECT * FROM %s" % tbl_name)
    colname = cursor.description
    col = []

    for i in colname:
        col.append(i[0])

    row = cursor.fetchall()
    df = pd.DataFrame(row, columns = col)

    for i in df.columns:
        new_col = "".join([col_dic[i],"(",i,")"])
        df.rename(columns = {i : new_col}, inplace = True)

    return df

def Table_GET_TOP100(tbl_name, db):
    cursor = db.cursor()
    cursor.execute("SELECT * FROM all_col_comments WHERE table_name = '%s'" % tbl_name)
    colname = cursor.description
    col = []

    for i in colname:
        col.append(i[0])

    row = cursor.fetchall()
    df_col = pd.DataFrame(row, columns = col)

    col_dic = dict()
    for i in range(len(df_col)):
        col_dic[df_col.loc[i, 'COLUMN_NAME']] = df_col.loc[i, 'COMMENTS']

    cursor.execute("SELECT * FROM %s WHERE rownum <= 100" % tbl_name)
    colname = cursor.description
    col = []

    for i in colname:
        col.append(i[0])

    row = cursor.fetchall()
    df = pd.DataFrame(row, columns = col)

    for i in df.columns:
        new_col = "".join([col_dic[i],"(",i,")"])
        df.rename(columns = {i : new_col}, inplace = True)

    return df

def Table_Columns_GET(tbl_name, db):
    cursor = db.cursor()
    cursor.execute("SELECT * FROM all_col_comments WHERE table_name = '%s'" % tbl_name)
    colname = cursor.description
    col = []

    for i in colname:
        col.append(i[0])

    row = cursor.fetchall()
    df_col = pd.DataFrame(row, columns = col)
    return df_col

def All_Tables(db):
    cursor = db.cursor()
    cursor.execute("""SELECT * FROM all_tab_comments""")
    colname = cursor.description
    col = []

    for i in colname:
        col.append(i[0])

    row = cursor.fetchall()
    table_list = pd.DataFrame(row, columns = col)
    return table_list
