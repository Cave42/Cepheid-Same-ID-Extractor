import pandas as pd

path = ("H:\OCR Extraction.xlsx")

data = pd.read_excel(r'H:\OCR Extraction.xlsx', header=None)

print("rows:")
print(len(data[0]))

data = data[0].str.findall(r"274S\d+\d+\d+\d+\d+\d+")

data = data.explode()

data = data.drop_duplicates()

print(data)

data.to_excel("SampleIDS.xlsx")
