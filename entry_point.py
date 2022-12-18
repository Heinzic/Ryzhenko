import pretty_table_generator
import pdf_generator

key = input("Введите \"g\" для генерации таблицы. Введите \"f\" для формирования pdf ")
if key == "g":
    pretty_table_generator.get_pretty_table()
elif key == "f":
    pdf_generator.get_pdf_statistic()