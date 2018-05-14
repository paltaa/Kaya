import pandas as pd
import math
#*** CONFIGURACION****

#*****NOMBRE DE LOS ARCHIVOS*******
fileVentas = 'ventas.xlsx'
#Nombre archivo con los stocks por semana
fileStocks = [
'stock semana 10.xlsx',
'stock semana 11.xlsx',
'stock semana 12.xlsx',
'stock semana 13.xlsx',
'stock semana 14.xlsx',
]
#***********SEMANAS Y TIENDA  Q VAMOS A CALCULAR ROTACION***********
semana_A=13
semana_B=12
tienda='TRAPENSES'



#****INICIO SCRIPT***********
#Fila, columna
def GetVentas ( semana_B, file , tienda):
    xl=pd.ExcelFile(file)
    Ventas= (xl.parse('vta 18'))
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
        if(Ventas.iloc[k,colFact]!= -1  and Ventas.iloc[k,colSemana] ==semana_B and tienda ==Ventas.iloc[k,colTienda]) :
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
    Matriz=pd.DataFrame(Data)
    return Data
#print (GetVentas(fileVentas))
#construir matriz de stock de la siguiente forma:
#sku, nombre, cantidad, semana, mes, tienda
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
    return Data
print("********** CARGANDO STOCK SEMANA 13 ****************")
stock_a=GetStock('stock semana 13.xlsx', tienda)
print("********** CARGANDO VENTAS SEMANA 13 *****************")
ventas=GetVentas(13,'ventas.xlsx', tienda)
print("********** CARGANDO STOCK SEMANA 12 *****************")
stock_b=GetStock('stock semana 12.xlsx', tienda)
print('SEMANA 13 ********************')
print(pd.DataFrame(stock_a))

print('SEMANA 12 ********************')
print(pd.DataFrame(stock_b))

#variables funcion inicio semana, semana final, matriz de ventas, matriz de stock, tienda
def calcularRotation( ventas, stock_a, stock_b):
    Data= [[ '' for x in range(9)] for y in range(len(ventas))]
    #Data=[]
    len_a=len(stock_a)
    len_b=len(stock_b)
    Data[0][0]='SKU'
    Data[0][1]='Nombre Producto'
    Data[0][2]='Vendidos Periodo'
    Data[0][3]='Inventario Periodo'
    Data[0][4]='Inventario Periodo Anterior'
    Data[0][5]='Rotacion Periodo'
    Data[0][6]='Semana Actual'
    Data[0][7]='Tienda'
    Data[0][8]='Pedido Actual'
    for i in range(1,len(ventas)):
        contada=0
        Data[i][0]=ventas[i][0]     #SKU
        Data[i][1]=ventas[i][1]     #Nombre
        Data[i][2]=ventas[i][2]     #venta
        Data[i][6]= ventas[i][3]    #semana
        Data[i][7]= ventas[i][5]    #tienda
        for j in range(len_a):

            if(ventas[i][3]==stock_a[j][3]  and #semanas iguales
            ventas[i][0] == stock_a[j][0]):    #SKU iguales

                    Data[i][3]=stock_a[j][2]  #inventario anterior

        for k in range(len_b):
                if(ventas[i][3]-1 ==stock_b[k][3]  and   #Semanas iguales
                    ventas[i][0] == stock_b[k][0]):   #SKU iguales
                    Data[i][4]=stock_b[k][2] #inventario actual

        if(type(Data[i][2])==str or  type(Data[i][3]) == str or  type(Data[i][4])==str):
            Data[i][5]='No Hay info'

        else:
            #PON LA FORMULA ACA FEO CULIAO
            Data[i][5]= Data[i][2]/((Data[i][3] + Data[i][4])//2)
            Data[i][8]= math.ceil(Data[i][2]**Data[i][5])
    print('*************DATA FRAME DE ROTACION DE INVENTARIO ***************')
    dataframe=pd.DataFrame(Data)
    print(dataframe)
    return dataframe


print("*****CALCULANDO ROTACION DE INVENTARIO ********")
rotacion=calcularRotation(ventas,stock_a,stock_b)

print("******CREANDO EXCEL CON ROTAION DE INVENTARIO **********")
rotacion.to_csv(path_or_buf=tienda+".csv", index=False)
print("*****SCRIPT TERMINADO*********")


#writer = pd.ExcelWriter(('Rotacionqlia.xlsx', engine='xlsxwriter')


#rotacion.to_excel(writer, sheet_name='rotacion')

#writer.save()
