import pandas as pd
from docx import Document



def pandas_manipulation():    
    document = Document()
    #Abrir archivos con pandas
    open_files_with_pd(document)
    
    #Descripcion preeliminar
    data_description(document)
    
    #Datos faltantes
    data_missing(document)
    
    #Reestructurar datos
    restructure_data(document)
    
    #Manejo de duplicados y Manejo de columnas 
    duplicate_and_columns_handling(document)

    #Concatenacion
    concatenation_handling(document)
    
    #Formato de largo a ancho e inverso
    format_lenght_width(document)

    #Separar columnas
    sepate_columns(document)

    #Conversion de datos categoricos
    categorical_data_conversion(document)

    #Variables dummy
    dummy_variables(document)

    #Funciones de agregación
    aggregation_function(document)

    #Funciones personalizadas
    custom_functions(document)
    
    #Funciones de grupo
    grouping_functions(document)

    
    #Exportación de datos
    export_data(document)

    #Manejo de json
    handling_json(document)

    #Manejo de xml y html
    handling_xml_html(document)

    document.save(r"Output/documento de practicas.docx")


def open_files_with_pd(document):
    source = r'D:\Cursos\Python\Manipulacion y limpieza\Resources\notebooks\datos.csv'
    
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

def data_description(document):
     
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

def data_missing(document):
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

def restructure_data(document):

    document.add_heading("Reestructurar datos", level=2)
    data_faltante = pd.read_csv(r'Resources/notebooks/datos_con_faltantes.csv')
    document.add_paragraph("Para segmentar y clasificar datos continuos en grupos o intevalos se utiliza pd.cut():")
    data_faltante["Rango_Edad"] = pd.cut(data_faltante["Edad"], bins=[20,25,30,35,40],labels=['20-25','25-30','30-35','35-40'],include_lowest= True)
    document.add_paragraph(data_faltante.to_string())


    document.add_paragraph("Para obtener el salario promedio por edad utilizo groupby() ")
    data_agrupada = data_faltante.groupby('Rango_Edad',observed= True)['Salario'].mean()
    document.add_paragraph(data_agrupada.to_string())

def duplicate_and_columns_handling(document):
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
  #Manejo de columnas
 document.add_paragraph("Para reordenar las columnas de un dataframe se hace de esta forma:duplicate[['ID','Nombre','Edad', 'Salario','Es duplicado?','Fecha_Ingreso']]")
 document.add_paragraph(duplicate[['ID','Nombre','Edad', 'Salario','Es duplicado?','Fecha_Ingreso']].to_string())
    
 document.add_paragraph("Para elegir columnas especificas de un dataframe se puede usa la propiedad .loc de esta manera: duplicate.loc[:,['ID','Nombre']] ")
 duplicate_loc = duplicate.loc[:,["ID",'Nombre']]
 document.add_paragraph(duplicate_loc.to_string())
 document.add_paragraph("El primer parametro de la propiedad .loc se utiliza para especificar el rango de las filas que se selecionaran, ejemplo de filas del 1 al 9 / 1:9")
 document.add_paragraph(duplicate_loc[1:9].to_string())

 document.add_paragraph("Para eliminar una columna, ejemplo la duplicado, se utiliza el metodo ..drop(columns=['Es duplicado?'])")
 duplicate = duplicate.drop(columns=['Es duplicado?'])
 document.add_paragraph(duplicate.to_string())
 document.add_paragraph("Si quiero filtra los salarios superiores a 50,000 se hace de la siguiente manera: duplicate.loc[duplicate['Salario' ]> 50000]")
    
 document.add_paragraph(duplicate.loc[duplicate['Salario' ]> 50000].to_string())

 document.add_paragraph("Para agregar una columna nueva a un dataframe se hace de esta manera:duplicate['Posición'] agregando los valores que tendra la columna: ")
 duplicate['Posición'] = pd.cut(duplicate["Salario"],bins=[40000,60000,65000,75000],labels=['junior','mid','Senior'])
 document.add_paragraph(duplicate.to_string())

 document.add_paragraph("Esta nueva columna se calcula cuanto se le descuenta de afp + ars: ")
 duplicate['AFP + ARS'] = duplicate["Salario"] * (5.91 /100)
 duplicate['Salario_Neto'] = duplicate['Salario'] * 0.9409
 document.add_paragraph(duplicate.to_string())

def concatenation_handling(document):
    document.add_heading("Concatenacion y combinacion", level=2)
    
    data_lima = {
    'Producto': ['A', 'B'],
    'Ventas': [250, 150],
    'Ciudad': ['Lima', 'Lima']
    }

    df_lima = pd.DataFrame(data_lima)
    
    
    data_bogota = {
    'Producto': ['A', 'C'],
    'Ventas': [100, 200],
    'Ciudad': ['Bogotá', 'Bogotá']
}
    df_bogota = pd.DataFrame(data_bogota)
    document.add_paragraph("Para concatenar dos o mas dataframes se utiliza el metodo concat de esta manera: pd.concat([df_lima, df_bogota]), este es el resultado: ")
    df_concat = pd.concat([df_lima, df_bogota])
    document.add_paragraph(df_concat.to_string())
    data_inventario = {
    'Producto': ['A', 'B'],
    'Inventario': [30, 45]
    }

    df_inventario = pd.DataFrame(data_inventario)
    document.add_paragraph("Al concatenar datframes es recomendable reiniciar indices usando la propiedad: reset_index(drop=True)  ")
    df_concat = pd.concat([df_concat.reset_index(drop=True), df_inventario],axis=1)
    document.add_paragraph(df_concat.to_string())

def format_lenght_width(document):
    document.add_heading("Formatear tablas de largo a ancho", level=2)
    data_largo = {
    'Producto': ['A', 'A', 'B', 'B', 'C', 'C'],
    'Mes': ['Enero', 'Febrero', 'Enero', 'Febrero', 'Enero', 'Febrero'],
    'Ventas': [100, 150, 200, 250, 300, 350]
    }

    df = pd.DataFrame(data_largo)
    document.add_paragraph("Para convertir un dataframe de largo a ancho se utiliza el metodo pivot de un dataframe, que te pide tres parametros, index, columns y values")
    document.add_paragraph("Antes de utilizar el metodo pivot:")
    document.add_paragraph(df.to_string())
    document.add_paragraph("Despues de utilizar el metodo pivot:")
    document.add_paragraph(df.pivot(index='Producto', columns='Mes', values='Ventas').to_string())
    document.add_heading("Formatear tablas de ancho a largo", level=2)
    df_width = df.pivot(index='Producto', columns='Mes', values='Ventas')
    document.add_paragraph("Para formatear dataframes, para pasarlos de ancho a largo se utiliza el metodo melt, el cual te pide un dataframe y que llenes los parametros: id_vars,value_vars,var_name, value_name, el resultado es este:")
    df_new_lenght = pd.melt(df_width.reset_index(),id_vars=['Producto'],value_vars=['Enero', 'Febrero'],var_name='Mes', value_name='Ventas')
    document.add_paragraph(df_new_lenght.to_string())

def sepate_columns(document):

    document.add_heading("Separar columnas", level=2)
    data = {
    'Nombre_Completo': ['Juan Perez', 'Maria Gomez', 'Luis Martinez']}
    data_fechas = {
    'Fecha': ['01-01-2024', '15-02-2024', '30-03-2024']
}

    
    df = pd.DataFrame(data)
    df_fechas = pd.DataFrame(data_fechas)
    document.add_paragraph("Para separar los datoscde una tabla de un dataframe en columnas se utiliza el metodo split de la propiedad str: .str.split, donde en el primer parametro se expecifica que elemento se debe separar del texto. ")
    document.add_paragraph("Antes de separar:")
    document.add_paragraph(df.to_string())
    document.add_paragraph(df_fechas.to_string())
    
    document.add_paragraph("Despues de separar:")
    df[['Nombre', 'Apellido']] = df['Nombre_Completo'].str.split(' ', expand=True)
    df_fechas[['dia','mes','año']] = df_fechas['Fecha'].str.split('-', expand=True)
    
    document.add_paragraph(df.to_string())
    document.add_paragraph(df_fechas.to_string())
    document.add_paragraph("Para crear una fecha completa con las / se usa el metodo agg, y el metodo join de esta manera: .agg('/'.join,axis=1) ")
    df_fechas['Fecha_completa'] = df_fechas[['dia','mes','año']].agg('/'.join,axis=1)
    document.add_paragraph("Resultado:")
    document.add_paragraph(df_fechas.to_string())  

def categorical_data_conversion(document):
    document.add_heading("Conversion de datos categoricos", level=2)
    data = {
    'producto': ['Manzana', 'Banana', 'Cereza', 'Durazno', 'Pera'],
    'categoria': ['Fruta', 'Fruta', 'Fruta', 'Fruta', 'Fruta'],
    'calidad': ['Alta', 'Media', 'Baja', 'Alta', 'Media'],  # Columna categórica
    'ventas': [50, 30, 70, 85, 40]
}

    document.add_paragraph("Para la conversion de datos categoricos se puede utilizar el metodo de codificacion ordinal que consiste en darle un valor a las variables categoricas para poder darles un orden  lógico o jerárquico  ")
    df = pd.DataFrame(data)
    document.add_paragraph("Datatframe si en el orden categorico:")
    document.add_paragraph(df.to_string())
    category_map = {'Baja':1, 'Media':2, 'Alta':3}
    df['Categoria ordinal'] = df['calidad'].map(category_map)
    document.add_paragraph("Despues:")
    document.add_paragraph(df.to_string())

    document.add_paragraph("La codificacion one hot permite convertir cada columna categoria a binaria:")

    df_one_hot = pd.get_dummies(df,columns=['calidad'])
    document.add_paragraph(df_one_hot.to_string())

def dummy_variables(document):
    document.add_heading("Variables dummy", level=2)
    data = {
    'vehiculo': ['Auto', 'Camioneta', 'Moto', 'Camion', 'Auto'],
    'color': ['Rojo', 'Azul', 'Negro', 'Blanco', 'Rojo'],
    'precio': [20000, 30000, 15000, 40000, 18000],
    'ventas': [150, 120, 130, 60, 180]
     }

    df = pd.DataFrame(data)
    document.add_paragraph("Las variables dummy son variables categoricas que toman valores binarios, usando el metodo get_dummies se crean columnas binarias con estos valores")
    document.add_paragraph("Antes de la conversión:")
    document.add_paragraph(df.to_string())
    dummy_variables = pd.get_dummies(df,columns=['vehiculo', 'color'])
    document.add_paragraph("Despues:")
    document.add_paragraph(dummy_variables.to_string())
    document.add_paragraph("Usando el parametro drop_first, se elimina la multicolinealidad que consiste en eliminar las columnas que son redundates")
    dummy_variables_drop = pd.get_dummies(df, columns=['vehiculo','color'],drop_first=True)
    document.add_paragraph(dummy_variables_drop.to_string())
    document.add_paragraph("Si quiero que los valores de las columnas dummy se muestren con texto descriptivos, se puede usar el metodo map, pero usando una expresion lambda en vez de booleanos:")
    rename_dummy_values = dummy_variables_drop[['vehiculo_Camion', 'vehiculo_Camioneta','vehiculo_Moto','color_Blanco', 'color_Negro','color_Rojo']].map(lambda x: 'Si' if x else 'No')
    document.add_paragraph(rename_dummy_values.to_string())

def aggregation_function(document):
    document.add_heading("Funciones de agregación",level=2)
    
    data = {
    'producto': ['Manzana', 'Manzana', 'Banana', 'Banana', 'Cereza', 'Cereza', 'Manzana', 'Banana', 'Cereza'],
    'categoría': ['Fruta', 'Fruta', 'Fruta', 'Legumbre', 'Fruta', 'Fruta', 'Fruta', 'Legumbre', 'Fruta'],
    'ventas': [50, 30, 20, 15, 10, 25, 45, 55, 40],
    'precio_unitario': [0.5, 0.5, 0.3, 0.3, 0.8, 0.8, 0.5, 0.3, 0.8],
    'fecha': pd.date_range(start='2024-01-01', periods=9, freq='D')
    }
    document.add_paragraph("Para las funciones de agregacion usando el metodo groupby se debe especificar la columna categorica, los valores los cuales seran agrupados y el calculo que se realizara")
    df = pd.DataFrame(data)
    document.add_paragraph("Datos sin agrupar")
    document.add_paragraph(df.to_string())
    document.add_paragraph("Datos de suma de ventas agrupados por producto, df.groupby('producto')['ventas'].sum():  ")
    product_sales = df.groupby('producto')['ventas'].sum()
    document.add_paragraph(product_sales.to_string())
    document.add_paragraph("Datos del promedio de precio por unidad agrupados por producto, df.groupby('producto')['precio_unitario'].mean() ")
    average_product_price = df.groupby('producto')['precio_unitario'].mean()
    document.add_paragraph(average_product_price.to_string())
    document.add_paragraph("Con el metodo agg se puede agregar multiples funciones, ejemplo:")
    sales_summary = df.groupby('producto')['ventas'].agg(['sum','min','max'])
    document.add_paragraph(sales_summary.to_string())
    document.add_paragraph("Si quiero cambiar los datos agrupados convertidos en una serie a un dataframe se utiliza el metodo .reset_index() ")
    sales_summary = df.groupby('producto')['ventas'].agg(['sum','min','max']).reset_index()
    document.add_paragraph(sales_summary.to_string())
    document.add_paragraph("Se puede agrupar varias columnas con groupby y crear varias funciones con el metodo agg: ")
    category_summary = df.groupby(['categoría','producto']).agg(
        total_sales = ('ventas','sum'),
        average_price = ('precio_unitario','mean')
    )
    document.add_paragraph(category_summary.to_string())


def custom_functions(document):
    document.add_heading("Funciones personalizadas", level=2)

    data = {
    'producto': ['Manzana', 'Manzana', 'Banana', 'Banana', 'Cereza', 'Cereza', 'Manzana', 'Banana', 'Cereza'],
    'categoría': ['Fruta', 'Fruta', 'Fruta', 'Fruta', 'Fruta', 'Fruta', 'Fruta', 'Fruta', 'Fruta'],
    'ventas': [50, 30, 20, 15, 10, 25, 45, 55, 40],
    'precio_unitario': [0.5, 0.5, 0.3, 0.3, 0.8, 0.8, 0.5, 0.3, 0.8],
    'fecha': pd.date_range(start='2024-01-01', periods=9, freq='D')
}

    df = pd.DataFrame(data)

    document.add_paragraph("Usan el metodo apply, se puede integrar funciones personalizadas")
    document.add_paragraph("Dataframe original:")
    document.add_paragraph(df.to_string())


    def total_renevue(row):
        return row['ventas'] * row['precio_unitario']
    
    df['Ingresos totales'] = df.apply(total_renevue,axis=1)
    document.add_paragraph("dataframe con los ingresos totales:")
    document.add_paragraph(df.to_string())

    def classification_sales(ventas):
        classification = {
            'Alto': ventas > 40,
            'Medio': 20 <= ventas <= 40,
            'Bajo': ventas < 20
        }
        for categoria, cantidad in classification.items():
            if cantidad:
                return categoria

    document.add_paragraph("dataframe con la clasificacion de las ventas:")
    df['clasificaciones'] = df['ventas'].apply(classification_sales)
    document.add_paragraph(df.to_string())
    
    def percentage_sales(ventas):
        total = ventas.sum()
        porcentaje = (ventas / total) * 100
        return porcentaje.astype(int)
    
    document.add_paragraph("dataframe con el porcentaje de las ventas:")
    df['Porcentaje'] = df.groupby('producto')['ventas'].transform(percentage_sales)
    document.add_paragraph(df.to_string())

def grouping_functions(document):
    document.add_heading("Funciones de grupo", level=2)
    document.add_paragraph("Estas funciones por grupo esta hecho con los metodos groupby y agg")

    data = {
    'producto': ['Manzana', 'Manzana', 'Banana', 'Banana', 'Cereza', 'Cereza', 'Manzana', 'Banana', 'Cereza'],
    'categoría': ['Fruta', 'Fruta', 'Fruta', 'Fruta', 'Fruta', 'Fruta', 'Fruta', 'Fruta', 'Fruta'],
    'ventas': [50, 30, 20, 15, 10, 25, 45, 55, 40],
    'precio_unitario': [0.5, 0.5, 0.3, 0.3, 0.8, 0.8, 0.5, 0.3, 0.8],
    'fecha': pd.date_range(start='2024-01-01', periods=9, freq='D')
}

    df = pd.DataFrame(data)

    
    document.add_paragraph("Dataframe original:")
    document.add_paragraph(df.to_string())

    grouping_products_by_sales = df.groupby('producto')['ventas'].agg(['sum','mean']).reset_index()
    document.add_paragraph("Agrupación de productos por ventas:")
    document.add_paragraph(grouping_products_by_sales.to_string())

    def range_function(ventas):
        return ventas.max() - ventas.min()

    range_sales = df.groupby('producto')['ventas'].agg(range_function).reset_index()
    document.add_paragraph("Rango de ventas, (hecho con una función personalizada):")
    document.add_paragraph(range_sales.to_string())
    document.add_paragraph("Agrupación de ingreso total por producto:")
    df['ingreso_total'] = df['ventas'] * df['precio_unitario']
    
    grouping_reneuve = df.groupby(['producto','categoría']).agg({
        'ventas': 'sum',
        'ingreso_total': 'sum'
    }).reset_index()

    document.add_paragraph(grouping_reneuve.to_string())

def export_data(document):
    document.add_heading("Exportación de datos", level=2)
    document.add_paragraph("""Para exportar datos se debe primero convertir el conjunto de datos en un dataframe, despues se elige el formato para exportar, de los cuales estan:
    to_csv
    to_excel
    to_parquet
    to_html
    etc.
                           """)
    document.add_paragraph("Al elegir cualquiera de los formatos para exportar se debe ingresar el nombre que tendra el archivo y su extensión de esta manera: to_excel('Output/documento excel.xlsx',index=True) donde la propiedad index indica si deberia estar numerada cada fila del dataframe. ")
    document.add_paragraph("""Recomendaciones para el guardado de archivos:
    En archivos csv, el parametro sep es para separar los delimitadores.
    float_format='%.2f' Para expecificar que solo haya 2 decimales
    El parametro compression es para especificar el formato en el que se comprimirá el archivo
    el parametro compression permite los siguientes procesos compression={'method': 'zip', 'archive_name': 'datos_comprimido.xlsx'}
    el parametro encoding, permite codificar el archivo a multiples formatos: encoding='utf-8'
    Con el parametro na_rep puedes especificar con que rellenar los datos nulos o vacios, ej: na_rep='N/A'        """)

def handling_json(document):
    document.add_heading("Manejo de json", level=2)
    document.add_paragraph("Para pasar cadenas de json al metodo  pd.read_json, es necesario convertir esas cadenas en un formato StringIO de la siguiente forma: ")
    document.add_paragraph("""json_data = '''
[
    {"producto": "A", "ventas": 100},
    {"producto": "B", "ventas": 200},
    {"producto": "C", "ventas": 300}
]
'''

# Envolver la cadena JSON en un objeto StringIO
json_stream = StringIO(json_data)

Es necesario hacer estas importaciones: from io import StringIO """)
    

    json = pd.read_json(r'Resources\notebooks\datos.json',orient='records')
    document.add_paragraph("Diferentes orientaciones que se pueden utilizar a exportar en json:")
    document.add_paragraph("orient='split':")
    format_split = json.to_json(orient='split')
    document.add_paragraph(str(format_split))
    document.add_paragraph("orient='records' uso de la propiedad lines = True (unicamente utilizable con esta orientación):")
    document.add_paragraph(str(json.to_json(orient='records', lines= True)))
    document.add_paragraph("orient='index'")
    document.add_paragraph(str(json.to_json(orient='index')))
    document.add_paragraph("orient='columns:'")
    document.add_paragraph(str(json.to_json(orient='columns')))
    document.add_paragraph("orient='values':")
    document.add_paragraph(str(json.to_json(orient='values')))

def handling_xml_html(document):
    document.add_heading("Manejo de xml y html",level=2)
    document.add_paragraph("Para la lectura de archivos xml se usa el metodo read_xml, si no es un archivo, se usa StringIO de la libreria io al igual que los json  ")
    df = pd.read_xml(r'Resources\notebooks\datos.xml')
    document.add_paragraph(df.to_string())
    
    tables = pd.read_html(r'Resources\notebooks\datos.html')
    document.add_paragraph("Si un documento html posee varias tablas usando tables[n] donde n es el numero de tabla que se busca acceder:")
    document.add_paragraph(tables[0].to_string())


