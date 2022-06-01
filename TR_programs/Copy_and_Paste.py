import xlsxwriter
import os
import numpy as np
from TR_programs import paste1
def copy(book,Sheet,path1,Graph_Tittle):
    print(Sheet)
    ws = book.get_worksheet_by_name(Sheet)
    ws.write('A1', "Delay Time(ps)")
    ws.write('B1',"Kerr Signal[%]")
    ws.write('D1',Graph_Tittle)
    paste1.paste(ws,book,Sheet,path1,Graph_Tittle)
