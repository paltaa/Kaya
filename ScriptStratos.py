# Import pandas libreria q usaremos
import pandas as pd
######NACHO SOLO PUEDES TOCAR ESTO###################

#VARIABLE DEL CODIGO QUE NOS INTERESA, MODIFICAR SOLO ESTA VARIABLE

Code='11-1011.02'
file = 'maestro.xlsx'


####################################################
#declaracion funciones:

def getJob (Code, jobZones, jobReference ):
	exp = ['','']
	for i in xrange(len(jobZones)):
		if(Code == jobZones.iloc[i,0]):

			Zone = jobZones.iloc[i,1]

	for j in xrange(len(jobReference)):

		if(Zone == jobReference.iloc[j,0]):

			exp[0]=jobReference.iloc[j,2]
			exp[1]=jobReference.iloc[j,4]
	Data=pd.DataFrame(exp)
	return Data

def GetName(Data, Code ):
	for i in xrange(len(Data)):
		name = ['' , '']
		if ( Code == Data.iloc[i,0]):

			name[0] = Data.iloc[i,1]
			name[1] = Data.iloc[i,2]
			Data=pd.DataFrame(name)
			return Data

def excelsior (matrizSkills , skills, Code, scaleReference ):
	#Valor de donde empieza a escribir las filas

	j=1


	for i in xrange(len(skills)):

		if(Code == str(skills.iloc[i,0])):

			#asignar nombre de la skill con su valor pedido, minimo y maximo
			#nombre
			matrizSkills[j][0]=skills.iloc[i,2]
			#valor pedido
			matrizSkills[j][1]=skills.iloc[i,4]

			#VARIABLE DE LA ESCALA EN LA QUE ESTA EL SKILL PARA SACAR MINIMO Y MAXIMO
			Escala = skills.iloc[j,3]

			#recorrer la escala
			for k in xrange(len(scaleReference)):

				if (Escala == scaleReference.iloc[k,0]):
					#dar valor minimo maximo y nombre de la escala
					matrizSkills[j][2] = scaleReference.iloc[k,2] #min
					matrizSkills[j][3] = scaleReference.iloc[k,3] #max
					matrizSkills[j][4] = scaleReference.iloc[k,1] #nombre
			j=j+1


	print (j)

	#print matrizSkills

	Data=pd.DataFrame(matrizSkills)


	return Data
#funcion para sacar las tareas de cada cargo
def getTasks ( matrizTasks, Tasks, Code ):
	j=0
	for i in xrange(len(Tasks)):
		if ( Code == Tasks.iloc[i,0]):

			matrizTasks[j][0]= Tasks.iloc[i,2]

			j=j+1
	Data=pd.DataFrame(matrizTasks)

	return Data

#nombre del archivo excel a cargar
# Cargar el excel como objeto
xl = pd.ExcelFile(file)
# ver nombres de cada spreadsheet
#cargar cada spreadsheet
scaleReference = (xl.parse('Scale Reference'))
skills=(xl.parse('skills'))
workActivity=(xl.parse('work Activity'))
interests=(xl.parse('Interest'))
workValue=(xl.parse('Work Value'))
knowledge=(xl.parse('Knowledge'))
abilities=(xl.parse('Abilities'))
workContext=(xl.parse('Work Context'))
tasks = (xl.parse('Task'))
jobZones =(xl.parse('Job Zones'))
jobReference=(xl.parse('Zone Job Reference'))
Data = (xl.parse(' Data'))

#Matriz de skills con minimo maximo y valor necesario
w, h = 5, 150
matrizSkills = [[ '' for x in range(w)] for y in range(h)]
matrizInterests = [[ '' for x in range(w)] for y in range(h)]
matrizWorkActivity = [['' for x in range(w)] for y in range(h)]
matrizTasks = [['' for x in range(1)] for y in range(h)]
#nombre del cargo
nombre = GetName(Data, Code)

#NOMBRES DE COLUMNAS EN MATRIZ DE SKILLS
matrizWorkActivity[0][0] = 'Nombre Elemento'
matrizWorkActivity[0][1] = 'Valor'
matrizWorkActivity[0][2] = 'Minimo'
matrizWorkActivity[0][3] = 'Maximo'
matrizWorkActivity[0][4] = 'Nombre Escala'

matrizSkills[0][0] = 'Nombre Elemento'
matrizSkills[0][1] = 'Valor Minimo Esperado'
matrizSkills[0][2] = 'Minimo'
matrizSkills[0][3] = 'Maximo'
matrizSkills[0][4] = 'Nombre Escala'


matrizInterests[0][0] = 'Nombre Elemento'
matrizInterests[0][1] = 'Valor'
matrizInterests[0][2] = 'Minimo'
matrizInterests[0][3] = 'Maximo'
matrizInterests[0][4] = 'Nombre Escala'

#obtener todas las matricces en forma data frame para ser exportadas al excel
mJob=getJob (Code, jobZones, jobReference )
mSkills = excelsior( matrizSkills , skills, Code, scaleReference)
mWorkA = excelsior( matrizWorkActivity, workActivity, Code, scaleReference )
mInterests= excelsior( matrizInterests, interests, Code, scaleReference )
mWorkValue= excelsior( matrizInterests, workValue, Code, scaleReference )
mKnowledge=  excelsior( matrizInterests, knowledge, Code, scaleReference )
mAbilities=  excelsior( matrizInterests, abilities, Code, scaleReference )
mWorkContext=  excelsior( matrizInterests, workContext, Code, scaleReference )
mTasks = getTasks(matrizTasks , tasks , Code  )

writer = pd.ExcelWriter((nombre.iloc[0,0])+'.xlsx', engine='xlsxwriter')
mCaracteristicas = nombre

# Escribir todos los dataframes a un excel
mSkills.to_excel(writer, index=False, sheet_name='Skills')
mWorkA.to_excel(writer, index=False, sheet_name='Work Activity')
mInterests.to_excel(writer, index=False, sheet_name='Interests')
mWorkValue.to_excel(writer, index=False, sheet_name='Work Value')
mKnowledge.to_excel(writer, index=False, sheet_name='Knowledge')
mAbilities.to_excel(writer, index=False, sheet_name='Abilities')
mWorkContext.to_excel(writer, index=False, sheet_name='Work Context')
mTasks.to_excel(writer, index=False, sheet_name = 'Tasks')
mJob.to_excel(writer, index=False , sheet_name = 'Job Zone')
mCaracteristicas.to_excel(writer, index=False, sheet_name='Caracteristicas')
# Guardar el excel
writer.save()
#print " joya"
