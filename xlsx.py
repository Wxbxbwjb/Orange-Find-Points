import pandas as a
b = a.read_excel('a.xlsx', c='2', usecols='A:BD')
d = b.columns.tolist()
print(d)
