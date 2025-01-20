import os
from pandas_function import pandas_manipulation
from docx import Document
def main():

    document = Document()
    source = r'D:\Cursos\Python\Manipulacion y limpieza\Resources\notebooks\datos.csv'
    output_doc = r"Output/documento de practicas.docx"
    

    if os.path.exists(output_doc):
        os.remove(output_doc)

    pandas_manipulation(source, document)


if __name__ == "__main__":
    main()