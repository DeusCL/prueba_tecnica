"""
	@archivo: prueba.py
	@autor: Sebastián Morales Álvarez
	@fecha: 08-01-2023

	@función: A partir de un documento .xlsx genera unas planillas .csv
	con datos de matriculación de personas. Además, genera una copia del
	documento pero con datos normalizados.

	Forma de uso:

		prueba.py <carpeta/archivo excel>

	Ejemplos:

		prueba.py input.xlsx

		prueba.py carpeta/documento excel.xlsx

		prueba.py "C:/Users/Usuardium/Desktop/program/excel input.xlsx"


	Se generarán archivos .csv y un archivo .xlsx dentro de una carpeta llamada "output" en
	la misma ruta donde se encuentre este script.


	Pre-requisitos: librería pandas, y probablemente la librería openpyxl

"""



import sys, os
import pandas as pd


TXT_RUTS_INVALIDOS = "ruts invalidos.txt"
CARPETA_SALIDA = "output"



def main():
	df_documento = obtener_excel()
	if df_documento is None: return


	# Crea la carpeta de salida en caso que no exista
	if not os.path.exists(CARPETA_SALIDA):
		os.mkdir(CARPETA_SALIDA)


	# Identificar cursos y crear DataFrames vacíos para irlos rellenando luego
	lista_cursos = identificar_cursos(df_documento)

	print("\nInstancias de cursos detectadas:\n  - "+"\n  - ".join(lista_cursos)+"\n")

	cursos_df_list = {}

	for curso in lista_cursos:
		cursos_df_list[curso] = pd.DataFrame(columns = [
			'username','password','firstname','lastname','email',
			'course1','role1','institution','profile_field_RUT'
		])

	# Lista con los indices de los ruts invalidos para poder eliminarlos después
	indices_ruts_invalidos = list()


	print("\n\nProcesando datos...")
	for index in range(len(df_documento)):
		row = df_documento.iloc[index]
		marca_temporal = row['Marca temporal']


		# Normalizar Nombres
		lista_nombres = [nombre.capitalize() for nombre in str(row['Nombres']).split(' ')]
		nombres = ' '.join(lista_nombres).strip()
		firstname = lista_nombres[0]
		df_documento.at[index, 'Nombres'] = nombres


		# Validar rut y registrarlo si fuera inválido
		rut = str(row['RUT'])
		if not validar_rut(rut):
			print(f" ** Advertencia: RUT inválido en el registro \"{marca_temporal}\" ({nombres}). RUT: \"{rut}\". **")
			registrar_rut_invalido(dict(row))
			indices_ruts_invalidos.append(index)
			continue


		# Identificar campos vacíos
		for campo in ["Nombres", "Apellidos", "Establecimiento", "Teléfono"]:
			if str(row[campo]).strip() == "" or str(row[campo]) == "nan":
				print(f" ** Advertencia: Campo \"{campo}\" vacío en el registro \"{marca_temporal}\" ({nombres}). **")


		# Normalizar rut
		profile_field_RUT = normalizar_rut(rut)
		df_documento.at[index, 'RUT'] = profile_field_RUT

		username = profile_field_RUT.replace('.','').replace('-','').replace('K','0')
		password = username[:4]

		# Normalizar Apellidos
		lastname = " ".join([apellido for apellido in str(row['Apellidos']).split(' ') if apellido != ''])
		lastname = lastname.title().strip()
		df_documento.at[index, 'Apellidos'] = lastname

		# Normalizar correo electrónico
		email = str(row['Dirección de correo electrónico']).lower().strip()
		df_documento.at[index, 'Dirección de correo electrónico'] = email

		# Normalizar institución
		institution = str(row['Establecimiento']).title().strip()
		df_documento.at[index, 'Establecimiento'] = institution

		# Normalizar teléfono
		norm_telefono = normalizar_telefono(str(row['Teléfono']))
		df_documento.at[index, 'Teléfono'] = int(norm_telefono) if norm_telefono.isnumeric() else norm_telefono

		if len(norm_telefono) < 9:
			print(f" ** Advertencia: Teléfono inválido en el registro \"{marca_temporal}\" ({nombres}). Teléfono: \"{norm_telefono}\". **")


		# Identificar los cursos de esta persona
		cursos = row['¿Cuál o cuáles cursos le interesan?']
		cursos_separados = [curso.strip() for curso in cursos.split(',')]


		# Agregar todos estos datos a la tabla de cada curso de interés de esta persona
		for course1 in cursos_separados:
			curso_df = cursos_df_list[course1]

			# Verificar que esta persona no se encuentre agregada al curso
			if username in curso_df.values:
				#print(f" ** Advertencia: {profile_field_RUT} ({nombres}) ya se encontraba en el curso {course1}")
				continue

			datos = [
				username, password, firstname, lastname,email,
				course1, 5, institution, profile_field_RUT
			]

			# Concatenar dataframes
			cursos_df_list[course1] = pd.concat(
				[curso_df, pd.DataFrame([datos], columns=curso_df.columns)],
				ignore_index=True
			)


	# Exportar planillas

	print("\n\nExportando planillas csv de los cursos...")

	for curso in lista_cursos:
		archivo_planilla = os.path.join(CARPETA_SALIDA, f"{curso}.csv")

		print(f"Exportando {os.path.abspath(archivo_planilla)}")

		curso_df = cursos_df_list[curso]
		curso_df.to_csv(archivo_planilla, index=False, encoding='utf-8-sig')

	print("\n\nExportando archivo excel con los campos normalizados...")

	# Remover los ruts defectuosos
	df_documento = df_documento.drop(indices_ruts_invalidos)

	# Exportar archivo xlsx con los campos normalizados
	archivo_excel = os.path.join(CARPETA_SALIDA, "archivo_normalizado.xlsx")

	try:
		df_documento.to_excel(archivo_excel, index=False)
	except PermissionError as e:
		print(f" **** Error: No se ha podido generar el archivo ****\n Motivo: {e}")
	else:
		print(f"\nEl archivo excel con los campos normalizados se ha guardado en {os.path.abspath(archivo_excel)}")
 
	print("\nFinalizado.\n")


def normalizar_telefono(telefono:str):
	""" Normaliza el número telefónico quitando el +56 y otros caracteres no deseados. """

	# Del telefono conservar sólo los números y el +
	telefono = "".join([char for char in telefono if char.isnumeric() or char=="+"])

	# Remover el +56 si lo tuviera
	if "+56" in telefono:
		telefono = telefono.replace("+56", "")

	# Si tiene 11 caracteres es probable que tenga el 56 sin el +
	elif len(telefono) == 11 and telefono.startswith("56"):
		telefono = telefono[2:]

	return telefono


def registrar_rut_invalido(datos_usuario:dict):
	""" Registra en un archivo TXT un rut que ha sido detectado como inválido. """

	archivo_destino = os.path.join(CARPETA_SALIDA, TXT_RUTS_INVALIDOS)

	if not os.path.exists(archivo_destino):
		archivo_ruts = open(archivo_destino, 'w')
		archivo_ruts.write("*"*80)
		archivo_ruts.write("\n\tRegistro de ruts inválidos\n")
		archivo_ruts.write("*"*80)
		archivo_ruts.write("\n\n")
	else:
		archivo_ruts = open(archivo_destino, 'at')

	archivo_ruts.write(f"\n\n\n\nRut Inválido: {datos_usuario['RUT']}")

	for key in datos_usuario.keys():
		archivo_ruts.write(f"\n  {key} : {datos_usuario[key]}")

	archivo_ruts.close()



def normalizar_rut(rut:str):
	""" Arregla el formato de un rut colocandole puntos y un guión. """

	rut = rut.lower().strip()

	# Guardar todos los números o la K del rut en una lista
	lista_digitos = [x for x in list(filter(lambda x: x.isnumeric() or x == 'k', rut))]

	# Sacar el digito verificador de la lista
	digito_verificador = lista_digitos.pop(len(lista_digitos)-1)

	# Convertir la lista de digitos en un rut numérico
	run = int("".join(lista_digitos))

	# Separar los miles del rut numérico y convertir las comas en puntos
	norm_rut = "{:,}".format(run).replace(',','.')

	return norm_rut + "-" + digito_verificador



def validar_rut(rut:str):
	""" Verifica si el rut es válido o inválido """

	# Guardar todos los números o la K del rut en una lista
	rut = rut.lower().strip()
	list_rut = [x for x in list(filter(lambda x: (x.isnumeric() or x=='k'), rut))]

	if len(list_rut) <= 1:
		return False

	digito_verificador = list_rut.pop(len(list_rut)-1)

	# Calcular digito verificador
	serie = range(2, 8)
	reverse_rut = list_rut[::-1]
	real_digito_verificador = str(11 - sum([int(reverse_rut[i]) * serie[i%6] for i in range(len(reverse_rut))])%11)

	if real_digito_verificador == '11':
		real_digito_verificador = '0'

	if real_digito_verificador == '10':
		real_digito_verificador = 'k'

	return digito_verificador == real_digito_verificador




def identificar_cursos(dataframe:object):
	""" Retorna una lista de todos los cursos del documento que
	están en el campo de los cursos de interés, separados por una coma. 

	IE: cada instancia de curso está escrita perfectamente. """

	nombre_campo_cursos = '¿Cuál o cuáles cursos le interesan?'

	# Obtener los datos que hay en el campo de los cursos
	df_cursos = dataframe.get(nombre_campo_cursos)

	# Lista de listas de cursos separados por la coma
	cursos_separados = [cursos.split(',') for cursos in df_cursos]

	# Juntar todas las listas de cursos_separados en esta única lista
	lista_cursos = list()

	for cursos in cursos_separados:
		cursos_limpios = [curso.strip() for curso in cursos]
		lista_cursos = lista_cursos + cursos_limpios

	return list(set(lista_cursos))




def obtener_excel():
	""" Retorna un objeto DataFrame del archivo excel obtenido desde los
	argumentos. O None si el archivo no fue especificado, encontrado o válido """

	# Validar que exista algun argumento para el archivo del documento fuente
	if len(sys.argv) == 1:
		print("\nDocumento fuente no especificado.")
		print("\nUso:\n  program.py <ruta_documento>")
		print("\nPor ejemplo:\n  program.py mi_carpeta/archivo.xlsx")
		return None

	ruta_documento = " ".join(sys.argv[1:])


	# Comprobar que exista el archivo especificado
	if not os.path.exists(ruta_documento):
		print(f"\n No se ha podido encontrar el documento \"{ruta_documento}\".")
		return None

	# Tratar de leer el archivo
	try:
		print(f"\nLeyendo \"{ruta_documento}\"...")
		df_documento = pd.read_excel(ruta_documento)
	except ValueError:
		print(" **** Error: El archivo debe corresponder a un formato excel. **** ")
		return None
	except ImportError as e:
		print(e)
		return None
	except PermissionError as e:
		print(f" **** Error: No se ha podido abrir el archivo ****\n Motivo: {e}")
		return None

	return df_documento


if __name__ == '__main__':

	main()

