import pyexcel as pe
from os import listdir

def xlsconverter(input):
    print(input)
    for file in listdir(input):
        pe.save_book_as(file_name=str(input + file), dest_file_name=str(input + file).replace('.xls', '')+'.xlsx')
