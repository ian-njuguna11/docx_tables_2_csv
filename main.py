# import pandas library
import pandas as pd
import pymysql

pymysql.install_as_MySQLdb()

import docx
from docx import Document
import docx
import pandas as pd

import mysql.connector



doc = docx.Document("C:\\Users\\user\\Desktop\\Dr. Thiga List of Units offered\DEPARTMENT OF COMMERCE LIST OF UNITS ON OFFER _ MAY - AUG 2021 SEMESTER.docx")


def read_docx_table(document, table_num=1, nheader=1):
    table = document.tables[table_num - 1]
    data = [[cell.text for cell in row.cells] for row in table.rows]
    df = pd.DataFrame(data)
    if nheader == 1:
        df = df.rename(columns=df.iloc[1]).drop(df.index[0]).reset_index(drop=True)
    elif nheader == 2:
        outside_col, inside_col = df.iloc[0], df.iloc[1]
        hier_index = pd.MultiIndex.from_tuples(list(zip(outside_col, inside_col)))
        df = pd.DataFrame(data, columns=hier_index).drop(df.index[[0, 1]]).reset_index(drop=True)
    elif nheader > 2:
        print("more Then two headers not currently supported")
        df = pd.DataFrame()
    return df

# //count table doc iter 6
table_num = 2
nheader = 1
df = read_docx_table(doc, table_num, nheader)
df.to_csv("Addendum_II_Masters.csv")
print(df)
