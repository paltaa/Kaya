from tkinter import *
from tkinter import filedialog
from tkinter import ttk
import pandas as pd
import math
import os


arreglo_tiendas=['TRAPENSES', 'DOMINICOS','OESTE','VESPUCIO','PARQUE ARAUCO','SUBCENTRO']

##Definicion funciones para uso de eventos/botones
def browse_button():
    filename = filedialog.askdirectory()
    dir.set(str(filename))
    #print(filename)
    #print(dir.get())


def next():
    vt=saleNameEntry.get()+'.xlsx'
    vtSheet=saleSheetEntry.get()
    stockA=stockaNameEntry.get()+'.xlsx'
    #stockA_Sheet=stockaSheetEntry.get()
    stockB=stockbNameEntry.get()+'.xlsx'
    #stockB_Sheet=stockbSheetEntry.get()
    week=semanaEntry.get()
    #tienda=tiendaEntry.get()
    if((dir.get()!='')):

        os.chdir(dir.get())
        print(os.getcwd())

    """
    saleMatrix=GetVentas( week, vt, tienda, vtSheet)
    ##INVENTARIO ACTUAL
    stockAMatrix=GetStock(stockA, tienda)
    ##INVENTARIO ANTERIOR
    stockBMatrix=GetStock(stockB, tienda)
    #calcularRotacion(ventas, stock Anterior, stock Actual)


    rotaDataFrame=calcularRotation(saleMatrix, stockAMatrix, stockBMatrix)
    name=tienda+week+'.csv'
    rotaDataFrame.to_csv(name, index=False)
"""

    writer = pd.ExcelWriter(week+'.xlsx')

    for i in range(len(arreglo_tiendas)):
        saleMatrix=GetFourVentas(vt, week, arreglo_tiendas[i], vtSheet)
        stockAMatrix=GetStock(stockA, arreglo_tiendas[i])
        stockBMatrix=GetStock(stockB, arreglo_tiendas[i])
        rotaDataFrame=calcularRotation_B(saleMatrix, stockAMatrix, stockBMatrix, week)
        print("sheet tienda "+ arreglo_tiendas[i]+' CREADO CON EXITO')
        rotaDataFrame.to_excel(writer, arreglo_tiendas[i], header=False, index=False)

    writer.save()

    print('Arhivo generado con exito')
#Definicion funciones para trabajo de excels.
def GetFourVentas(file, week, tienda, sheetName):
    xl=pd.ExcelFile(file)
    Ventas= (xl.parse(sheetName))
    w , h = 6 , len(Ventas)
    global Data
    Data= [[ '' for x in range(w)] for y in range(h)]
    global ventas
    ventas= [[ '' for x in range(w)] for y in range(h)]
    
    headers = Ventas.dtypes.index
    for i in range(len(Ventas.iloc[0])):
        if (headers[i]=='Código'):
            colCodProd=i
        elif(headers[i]=='Producto'):
            colProd=i
        elif(headers[i]=='Cant. Facturada'):
            colFact=i
        elif(headers[i]=='SEMANA'):
            colSemana=i
        elif(headers[i]=='LOCAL'):
            colTienda=i
    i=0        
    setSku=set()
    for index, rows in Ventas.iterrows():
        
        if ( rows[colTienda]== tienda and(rows[colSemana] == week or rows[colSemana] == week-1 or rows[colSemana] == week -2 or rows[colSemana]== week-3 )):
            ventas[i][0]=rows[colCodProd]
            ventas[i][1]=rows[colProd]
            ventas[i][2]=rows[colFact]
            ventas[i][3]=rows[colSemana]
            ventas[i][4]=rows[colTienda]
            i=i+1
            setSku.add(rows[colCodProd])
    del(ventas[i:len(ventas)])
    #Set de todos los sku sin repetir para rellenar el arreglo con ventas x sku x tienda semanas t, t-1, t-2, t-3
    i=0
    for sku in setSku:
        print(sku)
        Data[i][0]=sku     #sku
      #  Data[i][1]=rows[colProd]     #Producto
        for k in range(len(ventas)):
            if(ventas[k][0]==sku): 
                Data[i][1]=ventas[k][1]
                if(ventas[k][3]==week):
                    Data[i][5]=ventas[k][2]
                elif(ventas[k][3]==week-1):
                    Data[i][4]=ventas[k][2]
                elif(ventas[k][3]==week-2):
                    Data[i][3]=ventas[k][2]
                elif(ventas[k][3]==week-3):
                    Data[i][2]=ventas[k][2]
        print(Data[i])
        i=i+1
    del(Data[i:len(Data)])
    print (pd.DataFrame(Data))
    print("******Matriz de ventas tienda "+ tienda +" cargada con exito**********")
    #print(Matriz)
    return Data


def GetVentas ( semana_B, file , tienda, sheetName):
    xl=pd.ExcelFile(file)
    Ventas= (xl.parse(sheetName))
    w , h = 6 , len(Ventas)
    Data= [[ '' for x in range(w)] for y in range(h)]
    headers = Ventas.dtypes.index
    for i in range(len(Ventas.iloc[0])):
        if (headers[i]=='Código'):
            colCodProd=i
        elif(headers[i]=='Producto'):
            colProd=i
        elif(headers[i]=='Cant. Facturada'):
            colFact=i
        elif(headers[i]=='SEMANA'):
            colSemana=i
        elif(headers[i]=='MES'):
            colMes=i
        elif(headers[i]=='LOCAL'):
            colTienda=i
    i=0
    for index, rows in Ventas.iterrows():
        if(rows[colFact]!= -1  and rows[colSemana] == int(semana_B) and tienda == rows[colTienda]) :
            Data[i][0]=rows[colCodProd]     #sku
            Data[i][1]=rows[colProd]     #Producto
            if (  rows[colFact]=='.'):
                Data[i][2]=0
            else:
                Data[i][2]=rows[colFact]     #Cantidad
            Data[i][3]=rows[colSemana]    #semana
            Data[i][4]=rows[colMes]    #mes
            Data[i][5]=rows[colTienda]    #tienda
            i=i+1

    del(Data[i:len(Data)])
    print("******Matriz de ventas tienda "+ tienda +" cargada con exito**********")
    #print(Matriz)
    return Data

def GetStock(file, tienda):
    xl=pd.ExcelFile(file)
    stock=(xl.parse('stock'))
    w , h = 6 , len(stock)
    Data= [[ '' for x in range(w)] for y in range(h)]
    #buscar columna con disponible
    headers = stock.dtypes.index
    for i in range(len(stock.iloc[0])):
        if (headers[i]=='Cod. Producto'):
            colCodProd=i
        elif(headers[i]=='Producto'):
            colProd=i
        elif(headers[i]=='Disponible'):
            colDisp=i
        elif(headers[i]=='Semana'):
            colSemana=i
        elif(headers[i]=='mes'):
            colMes=i
        elif(headers[i]=='BODEGA nombre'):
            colTienda=i
    contadorDisp=0
    """ Codigo para iterrows
    for index, rows in ventas.iterrows():
    print(rows['Bodega'])
    """
    for index, rows in stock.iterrows():
        if(rows[colDisp]>0 and tienda == rows[colTienda]):
            Data[contadorDisp][0]=rows[colCodProd]
            Data[contadorDisp][1]=rows[colProd]
            Data[contadorDisp][2]=rows[colDisp]
            Data[contadorDisp][3]=rows[colSemana]
            Data[contadorDisp][4]=rows[colMes]
            Data[contadorDisp][5]=rows[colTienda]
            contadorDisp=contadorDisp+1
    #Matriz=pd.DataFrame(Data)
    del(Data[contadorDisp:len(Data)])
    print("******Matriz de Stock "+tienda +" cargada con exito   *********")
    #print(pd.DataFrame(Data))
    return Data
def calcularRotation_B( ventas, stock_a, stock_b, week):
    Data= [[ '' for x in range(9)] for y in range(len(ventas))]
    #Data=[]
    len_a=len(stock_a)
    len_b=len(stock_b)
    Data[0][0]='SKU'
    Data[0][1]='Nombre Producto'
    Data[0][2]='Vendidos '+week-3
    Data[0][3]='Vendidos '+week-2
    Data[0][4]='Vendidos '+week-1
    Data[0][5]='Vendidos '+week
    Data[0][6]='Inventario Periodo Anterior'
    Data[0][7]='Inventario Periodo'
    Data[0][8]='Rotacion Periodo'
    #Data[0][9]='Tienda'
    Data[0][9]='Pedido Actual'
    for i in range(1,len(ventas)):
        Data[i][0]= ventas[i][0]     #SKU
        Data[i][1]= ventas[i][1]     #Nombre
        Data[i][2]= ventas[i][2]     #venta w-3
        Data[i][3]= ventas[i][3]     #venta w-2
        Data[i][4]= ventas[i][4]     #venta w-1 
        Data[i][5]= ventas[i][5]     #venta w
        for j in range(len_a):

            if(ventas[i][0] == stock_a[j][0]):    #SKU iguales
                    Data[i][6]=stock_a[j][2]  #inventario anterior
                    break
        for k in range(len_b):
                if( ventas[i][0] == stock_b[k][0]):   #SKU iguales
                    Data[i][7]=stock_b[k][2] #inventario actual
                    break
        if(type(Data[i][2])==str or  type(Data[i][3]) == str or  type(Data[i][4])==str):
            Data[i][8]='No Hay info'

        else:
            #Data[i][9]= Data[i][2]/((Data[i][3] + Data[i][4])//2)
            Data[i][9]="ingresar formula excel culiao"
            Data[i][8]= math.ceil(Data[i][2]**Data[i][5])
    print('*************DATA FRAME DE ROTACION DE INVENTARIO  ***************')
    dataframe=pd.DataFrame(Data)
    print(dataframe)
    return dataframe


def calcularRotation_A( ventas, stock_a, stock_b):
    Data= [[ '' for x in range(9)] for y in range(len(ventas))]
    #Data=[]
    len_a=len(stock_a)
    len_b=len(stock_b)
    Data[0][0]='SKU'
    Data[0][1]='Nombre Producto'
    Data[0][2]='Vendidos Periodo'
    Data[0][3]='Inventario Periodo Anterior'
    Data[0][4]='Inventario Periodo'
    Data[0][5]='Rotacion Periodo'
    Data[0][6]='Semana Actual'
    Data[0][7]='Tienda'
    Data[0][8]='Pedido Actual'
    for i in range(1,len(ventas)):
        Data[i][0]= ventas[i][0]     #SKU
        Data[i][1]= ventas[i][1]     #Nombre
        Data[i][2]= ventas[i][2]     #venta
        Data[i][6]= ventas[i][3]    #semana
        Data[i][7]= ventas[i][5]    #tienda
        for j in range(len_a):

            if(ventas[i][3]==stock_a[j][3]  and #semanas iguales
            ventas[i][0] == stock_a[j][0]):    #SKU iguales

                    Data[i][4]=stock_a[j][2]  #inventario anterior
                    break
        for k in range(len_b):
                if(ventas[i][3]-1 ==stock_b[k][3]  and   #Semanas iguales
                    ventas[i][0] == stock_b[k][0]):   #SKU iguales
                    Data[i][3]=stock_b[k][2] #inventario actual
                    break
        if(type(Data[i][2])==str or  type(Data[i][3]) == str or  type(Data[i][4])==str):
            Data[i][5]='No Hay info'

        else:
            Data[i][5]= Data[i][2]/((Data[i][3] + Data[i][4])//2)
            Data[i][8]= math.ceil(Data[i][2]**Data[i][5])
    print('*************DATA FRAME DE ROTACION DE INVENTARIO  ***************')
    dataframe=pd.DataFrame(Data)
    print(dataframe)
    return dataframe
root = Tk()
###Global variables
dir=StringVar()
vt=StringVar()
vtSheet=StringVar()
stockA=StringVar()
stockA_Sheet=StringVar()
stockB=StringVar()
stockB_Sheet=StringVar()
week=StringVar()
tienda=IntVar()


buttonBrowse = Button(text="Buscar directorio", command=browse_button)
title = Label(root, text= "Manejo de inventarios KayaUnite")
config=Label(root, text="Kaya more Faya")
browseLabel= Label(root, text="Directorio de trabajo")
browseText= Label(root, text=str(dir.get()))


saleLabel=Label(root, text="Nombre archivo ventas")
saleNameEntry=Entry(root)
saleSheetName=Label(root, text= "Hoja de trabajo ")
saleSheetEntry=Entry(root)


stockaLabel=Label(root, text="Nombre archivo stock inicial")
stockaNameEntry=Entry(root)
stockaSheetName=Label(root,text="Hoja de trabajo tiene que ser stock")
#stockaSheetEntry=Entry(root)

stockbLabel=Label(root, text= "Nombre archivo stock anterior")
stockbNameEntry=Entry(root)
stockbSheetName=Label(root,text="Hoja de trabajo tiene que ser stock")
#stockbSheetEntry=Entry(root)

buttonNext=Button(text="Aplicar configuracion", command=next)
buttonNext.bind('<Button-1>', next)


semanaLabel=Label(root, text="Semana")
semanaEntry=Entry(root)
#tiendaLabel=Label(root, text="Tienda")
#tiendaEntry=Entry(root)


#grid row column
title.grid(row=0)
config.grid(row=1)
browseLabel.grid(row=2, column=0)
browseText.grid(row=2, column= 1)
buttonBrowse.grid(row=2, column=2)
saleLabel.grid(row=3, column=0)
saleNameEntry.grid(row=3, column=1)
saleSheetName.grid(row=3,column=2)
saleSheetEntry.grid(row=3, column=3)

stockaLabel.grid(row=4, column=0)
stockaNameEntry.grid(row=4, column=1)
stockaSheetName.grid(row=4, column=2)
#stockaSheetEntry.grid(row=4, column=3)

stockbLabel.grid(row=5, column=0)
stockbNameEntry.grid(row=5, column=1)
stockbSheetName.grid(row=5, column=2)
#stockbSheetEntry.grid(row=5, column=3)

semanaLabel.grid(row=6, column=0)
semanaEntry.grid(row=6, column=1)
#tiendaLabel.grid(row=6, column=2)
#tiendaEntry.grid(row=6, column=3)
buttonNext.grid(row=7, column=3)

root.mainloop()
