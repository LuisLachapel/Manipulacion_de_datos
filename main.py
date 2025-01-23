import os
from pandas_function import pandas_manipulation
from docx import Document
def main():

    
    output_doc = r"Output/documento de practicas.docx"
    

    if os.path.exists(output_doc):
        os.remove(output_doc)

    pandas_manipulation()


if __name__ == "__main__":
    main()