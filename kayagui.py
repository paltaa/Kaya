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
        saleMatrix=GetVentas(week,vt,arreglo_tiendas[i], vtSheet)
        stockAMatrix=GetStock(stockA, arreglo_tiendas[i])
        stockBMatrix=GetStock(stockB, arreglo_tiendas[i])
        rotaDataFrame=calcularRotation(saleMatrix, stockAMatrix, stockBMatrix)
        print("sheet tienda "+ arreglo_tiendas[i]+' CREADO CON EXITO')
        rotaDataFrame.to_excel(writer, arreglo_tiendas[i], header=False, index=False)

    writer.save()

    print('Arhivo generado con exito')
#Definicion funciones para trabajo de excels.

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
    for k in range(len(Ventas)):
        if(Ventas.iloc[k,colFact]!= -1  and Ventas.iloc[k,colSemana] == int(semana_B) and tienda == Ventas.iloc[k,colTienda]) :
            Data[i][0]=Ventas.iloc[k,colCodProd]     #sku
            Data[i][1]=Ventas.iloc[k,colProd]     #Producto
            if (  Ventas.iloc[k,colFact]=='.'):
                Data[i][2]=0
            else:
                Data[i][2]=Ventas.iloc[k,colFact]     #Cantidad
            Data[i][3]=Ventas.iloc[k,colSemana]    #semana
            Data[i][4]=Ventas.iloc[k,colMes]    #mes
            Data[i][5]=Ventas.iloc[k,colTienda]    #tienda
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
    for k in range(len(stock)):
        if(stock.iloc[k,colDisp]>0 and tienda == stock.iloc[k,colTienda]):
            Data[contadorDisp][0]=stock.iloc[k,colCodProd]
            Data[contadorDisp][1]=stock.iloc[k,colProd]
            Data[contadorDisp][2]=stock.iloc[k,colDisp]
            Data[contadorDisp][3]=stock.iloc[k,colSemana]
            Data[contadorDisp][4]=stock.iloc[k,colMes]
            Data[contadorDisp][5]=stock.iloc[k,colTienda]
            contadorDisp=contadorDisp+1
    #Matriz=pd.DataFrame(Data)
    del(Data[contadorDisp:len(Data)])
    print("******Matriz de Stock "+tienda +" cargada con exito   *********")
    #print(pd.DataFrame(Data))
    return Data
def calcularRotation( ventas, stock_a, stock_b):
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
        contada=0
        Data[i][0]= ventas[i][0]     #SKU
        Data[i][1]= ventas[i][1]     #Nombre
        Data[i][2]= ventas[i][2]     #venta
        Data[i][6]= ventas[i][3]    #semana
        Data[i][7]= ventas[i][5]    #tienda
        for j in range(len_a):

            if(ventas[i][3]==stock_a[j][3]  and #semanas iguales
            ventas[i][0] == stock_a[j][0]):    #SKU iguales

                    Data[i][4]=stock_a[j][2]  #inventario anterior
                    
        for k in range(len_b):
                if(ventas[i][3]-1 ==stock_b[k][3]  and   #Semanas iguales
                    ventas[i][0] == stock_b[k][0]):   #SKU iguales
                    Data[i][3]=stock_b[k][2] #inventario actual

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
