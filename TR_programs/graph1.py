import xlsxwriter

def add_to_chart(ws,book,Sheet,Graph_Tittle,i,min_x,max_x):

    chart = book.add_chart({'type': 'scatter','subtype': 'smooth'})
    chart.set_title ({'name': Graph_Tittle,'name_font':{'name': 'Calibri', 'size': 14} })
    chart.set_y_axis({'name': 'V[V]','name_font':{'name': 'Calibri', 'size': 12},'label_position': 'low'})
    chart.set_x_axis({'name': 'Delay Time[ps]','name_font':{'name': 'Calibri', 'size': 12},'label_position': 'low','major_gridlines':{'visible':True}})
    chart.set_legend({
     "position": 'none'})
    chart.add_series({
    'categories': [Sheet, 1, 0, i, 0],
    'values':     [Sheet, 1, 1, i, 1],})
    chart.set_plotarea({'border': {'color': 'black', 'width': 1}})
    #chart.set_x_axis({'min': min_x, 'max':max_x})
    chart.set_style(11)
    chart.set_size({'width': 650, 'height': 400})
    ws.insert_chart("D3",chart)


def add_to_chart1(ws,book,Sheet,Graph_Tittle,i,min,max):

    chart = book.add_chart({'type': 'scatter'})
    chart.set_title ({'name': Graph_Tittle,'name_font':{'name': 'Calibri', 'size': 14}})
    chart.set_y_axis({'name': 'V[V]','name_font':{'name': 'Calibri', 'size': 12},'label_position': 'low'})
    chart.set_x_axis({'name': 'Delay Time[ps]','name_font':{'name': 'Calibri', 'size': 12},'label_position': 'low','major_gridlines':{'visible':True}})
    chart.add_series({
    'categories': [Sheet, 1, 0, i, 0],
    'values':     [Sheet, 1, 1, i, 1],})
    chart.set_size({'width': 650, 'height': 400})
    chart.set_plotarea({'border': {'color': 'black', 'width': 1}})
    chart.set_legend({
     "position": 'none'})



    ws.insert_chart("D23", chart)
