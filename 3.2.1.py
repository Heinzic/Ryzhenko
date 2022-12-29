import csv

file_name = input("Введите название файла: ")
is_first = True
rows_count = 0
new_files = {}
headlines_list = []

with open(file_name, encoding="utf-8-sig") as File:
    reader = csv.reader(File)
    for row in reader:
        if is_first:
            headlines_list = row
            is_first = False
        else:
            if row[-1][:4] in new_files.keys():
                new_files[row[-1][:4]].append(row)
            else:
                new_files[row[-1][:4]] = [headlines_list, row]

for year in new_files:
    with open(f'new_files/new_file_{year}.csv', 'w', newline='', encoding="utf-8-sig") as csvfile:
        writer = csv.writer(csvfile, delimiter=',',
                            quotechar='|', quoting=csv.QUOTE_MINIMAL)
        for row in new_files[year]:
            writer.writerow(row)