# -*- coding: utf-8 -*-
"""
Created on Tue Jun 19 15:30:23 2018

@author: JuanPablo
"""

#import pyomo

import pandas as pd


def getFourVentas(file, week, tienda, sheetName):
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

k=getFourVentas("ventas.xlsx", 15, "DOMINICOS", "vta 18")
