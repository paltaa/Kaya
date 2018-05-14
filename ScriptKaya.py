import pandas as pd
#A, B, J, M , N, O,
#Nombre archivo con las Ventas
fileVentas = 'ventas.xlsx'
#Nombre archivo con los stocks por semana
fileStocks = [
'stock semana 10.xlsx',
'stock semana 11.xlsx',
'stock semana 12.xlsx',
'stock semana 13.xlsx',
'stock semana 14.xlsx',
]
#Fila, columna
def GetVentas (file ):
    xl=pd.ExcelFile(file)
    Ventas= (xl.parse('vta 18'))
    w , h = 6 , len(Ventas)
    Data= [[ '' for x in range(w)] for y in range(h)]
    for i in range(len(Ventas)):
        if(Ventas.iloc[i,9]>0):
            Data[i][0]=Ventas.iloc[i,0]     #sku
            Data[i][1]=Ventas.iloc[i,1]     #Producto
            Data[i][2]=Ventas.iloc[i,9]     #Cantidad
            Data[i][3]=Ventas.iloc[i,12]    #semana
            Data[i][4]=Ventas.iloc[i,13]    #mes
            Data[i][5]=Ventas.iloc[i,14]    #tienda
    Matriz=pd.DataFrame(Data)
    return Data
#print (GetVentas(fileVentas))
#construir matriz de stock de la siguiente forma:
#sku, nombre, cantidad, semana, mes, tienda
def GetStock(file):
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
        if(stock.iloc[k,colDisp]>0):
            Data[contadorDisp][0]=stock.iloc[k,colCodProd]
            Data[contadorDisp][1]=stock.iloc[k,colProd]
            Data[contadorDisp][2]=stock.iloc[k,colDisp]
            Data[contadorDisp][3]=stock.iloc[k,colSemana]
            Data[contadorDisp][4]=stock.iloc[k,colMes]
            Data[contadorDisp][5]=stock.iloc[k,colTienda]
            contadorDisp=contadorDisp+1
    #Matriz=pd.DataFrame(Data)
    return Data

stock_a=GetStock('stock semana 10.xlsx')
ventas=GetVentas('ventas.xlsx')
stock_b=GetStock('stock semana 11.xlsx')
print('len stock_A : '+str(len(stock_a)))
print('len ventas:' + str(len(ventas)))
print('len stock_b : '+ str(len(stock_b)))
print( ventas[0])
print(stock_a[0])
print(stock_b[0])
#variables funcion inicio semana, semana final, matriz de ventas, matriz de stock, tienda
def calcularRotation(inicio, final, ventas, stock_a, stock_b, tienda):
    Data= [[ '' for x in range(6)] for y in range(len(stock_a))]
    #Data=[]
    rows=1
    len_a=len(stock_a)
    len_b=len(stock_b)
    maxlen=max(len_a , len_b)
    print (maxlen)
    for i in range(len(ventas)):
        for j in range(maxlen):
            if(j<len_a):
                if(ventas[i][3]==stock_a[j][3]  and #semanas iguales
                ventas[i][0] == stock_a[j][0] and   #SKU iguales
                stock_a[j][5]==tienda and     #tienda igual stock a
                ventas[i][5]==tienda):         #tienda igual matriz ventas
                    Data[i][0]=ventas[i][0]
                    Data[i][1]=ventas[i][1]
                    Data[i][2]=ventas[i][2]
                    Data[i][3]=stock_a[j][2]
                    print(Data[i])
            if(j<len_b):
                if(ventas[i][3]==stock_b[j][3]  and
                    ventas[i][0] == stock_b[j][0] and
                    stock_b[j][5]==tienda and
                    ventas[i][5]==tienda):
                    Data[i][4]=stock_b[j][2]
                    Data[i][5]='aca va rotacion'

    dataframe=pd.DataFrame(Data)
    print(Data)
    return dataframe
tienda='TRAPENSES'

calcularRotation(0,0,ventas,stock_a,stock_b,tienda)
