# arr = []
# with open('src.csv') as file:
#     x = file.read().split('\n')
#     for i in x[1:]:
#         # arr.append(i.split(',')[15])
#         print(len(i.split(',')))
# print(arr)
import csv
with open('src.csv') as f:
    reader = csv.reader(f)
    for row in reader:
        print(row[14])
        isbns = str(row[14]).zfill(13)
print(isbns)