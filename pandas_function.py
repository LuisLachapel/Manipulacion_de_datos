import pandas as pd
from docx import Document

def pandas_manipulation(source,document):    
    document = Document()
    
    document.add_heading(r"Abrir archivos con pandas", level=2)
    data = pd.read_csv(source)
    document.add_paragraph(data.to_string())
    document.add_paragraph(r"""para saltar tabulaciones agrega el parametro delimiter al pd.read asi:
               pd.read_csv(source, delimiter='\t')""")
    
    file_txt1 = pd.read_csv(r'Resources\notebooks\archivo1.txt',skiprows=2)


    document.add_paragraph("""Para archivos con comentarios con el signo #, 
              se debe saltar las filas con el parametro skiprows = n""")
    document.add_paragraph(file_txt1.to_string())
    document.add_paragraph("Dependiendo el caso agregar el parametro: engine='python'")

    #Descripcion de datos

    document.add_heading("Descripcion preliminar de los datos", level=2)

    data_txt = pd.read_csv(r'Resources\notebooks\datos.txt', delimiter='|')

    document.add_paragraph('Para tener una descripcion de los datos utlizar el metodo describe de pandas ')
    document.add_paragraph(data_txt.describe().to_string())

    document.add_paragraph('Si quiero saber los tipos utilizo la propiedad dtypes: ')
    document.add_paragraph(data_txt.dtypes.to_string())

    document.add_paragraph(r"""Si quero obtener los valores unicos de una columna utilizo el metodo unique, e
              specificando la columna, ejemplo: df["nombre"].unique():""")

    document.add_paragraph(str(data_txt['nombre'].unique()))

    document.add_paragraph("Si quiero contarlos utilizo el metodo value_counts: ")
    document.add_paragraph(data_txt['nombre'].value_counts().to_string())

    #Datos faltantes
    document.add_heading("Manejo de datos faltantes", level=2)
    data_faltante = pd.read_csv(r'Resources/notebooks/datos_con_faltantes.csv')

    document.add_paragraph("""Para contar la cantidad de datos faltantes por columna 
              utilizar los metodos .isnull().sum():""")
    document.add_paragraph(data_faltante.isnull().sum().to_string())

    document.add_paragraph("Para eliminar los datos faltantes se utiliza el metodo .dropna()")

    data_drop = data_faltante.dropna()

    data_columns = data_faltante.dropna(axis= 1)
    document.add_paragraph(data_drop.to_string())
    document.add_paragraph("""Si quiero eliminar las columnas con  datos faltantes se utiliza el parametro axis=1
               en el metodo .dropna()""")
    document.add_paragraph(data_columns.to_string())

    document.add_paragraph("""Se puede rellenar los datos faltantes utilizando el metodo fillna()
               y especificar que deseas rellenar""")

    data_fill = data_faltante.copy()

    data_faltante['Nombre'] = data_faltante['Nombre'].fillna("Desconocido")

    document.add_paragraph(data_faltante.to_string())

    document.add_paragraph("Para rellenar datos numericos como el salario se puede rellenar con la media")

    data_faltante["Salario"] = data_faltante["Salario"].fillna(data_faltante["Salario"].mean())
    document.add_paragraph(data_faltante.to_string())

    document.add_paragraph("Para  datos numericos como la edad, es mas util  la mediana:")
    data_faltante["Edad"] = data_faltante["Edad"].fillna(data_faltante["Edad"].median())
    document.add_paragraph(data_faltante.to_string())

    document.add_paragraph("Para rellenar datos como la cuidad, seria util el uso de la moda")
    data_faltante["Ciudad"] = data_faltante["Ciudad"].fillna(data_faltante["Ciudad"].mode()[0])
    document.add_paragraph(data_faltante.to_string())


    document.add_paragraph("""Para especificar los datos en donde se deban eliminar las columnas
                            se utiliza la propiedad subset, ejemplo de subset= [Nombre]:""")
    document.add_paragraph(data_fill.dropna(subset=['Nombre'],inplace= False).to_string())

        #Reestructurar datos
    document.add_heading("Reestructurar datos", level=2)

    document.add_paragraph("Para segmentar y clasificar datos continuos en grupos o intevalos se utiliza pd.cut():")
    data_faltante["Rango_Edad"] = pd.cut(data_faltante["Edad"], bins=[20,25,30,35,40],labels=['20-25','25-30','30-35','35-40'],include_lowest= True)
    document.add_paragraph(data_faltante.to_string())


    document.add_paragraph("Para obtener el salario promedio por edad utilizo groupby() ")
    data_agrupada = data_faltante.groupby('Rango_Edad',observed= True)['Salario'].mean()
    document.add_paragraph(data_agrupada.to_string())

    #Manejo de duplicados
    
    duplicate = pd.read_csv(r'Resources\notebooks\data_duplicada.csv')

    document.add_heading("Manejo de duplicados", level=2)
    document.add_paragraph("""Al usar el metodo .duplicated() me indica por medio de booleans que filas son duplicadas, sin necesidad de que estos sean consecutivos""")

    document.add_paragraph(duplicate.duplicated().to_string())

    document.add_paragraph("""Se puede usar el parametro subset para especificar las columnas y saber si contiene duplicados""")
    
    document.add_paragraph(duplicate.duplicated(subset='Nombre').to_string())

    document.add_paragraph("""Obteniendo las filas de los duplicados con: duplicate[duplicate.duplicated()""")

    duplicated_rows = duplicate[duplicate.duplicated()]
    document.add_paragraph(duplicated_rows.to_string())

    document.add_paragraph("""Para eliminar los duplicados se utiliza el metodo .drop_duplicates()""")
    duplicate_delete = duplicate.drop_duplicates()
    document.add_paragraph(duplicate_delete.to_string())

    document.add_paragraph("Para crear una columna que indique los valores duplicados se hace de este metodo: duplicate['Es duplicado?'] = duplicate.duplicated() ")
    duplicate['Es duplicado?'] = duplicate.duplicated()
    document.add_paragraph(duplicate.to_string())

    document.add_paragraph("Usando .map puedo cambiar los valores de la columna 'Es duplicado' de un booleano a Si 0 No")
    duplicate["Es duplicado?"] = duplicate["Es duplicado?"].map({True: 'Si', False: 'No'})
    document.add_paragraph(duplicate.to_string())
    """ duplicate_summarized = duplicate.groupby('Nombre').agg({
        'Edad': 'first',
        'Salario': 'mean',
        'Fecha_Ingreso': 'first'
    }).reset_index()"""
    

    




   


    document.save(r"Output/documento de practicas.docx")



