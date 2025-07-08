# Program to extract CUV worksheet from Excel appraisal workbook


# Import modules and set parameters
import pandas as pd
import numpy as np
import os


# Use predefined path and file for now but will need to allow for user to select/input this info
folder_path = '/Users/dougroehm/Code/repos/mtdor-extract-cuv/TY23-cap-final-appraisals'
file_name = 'U202_TY23_CAP_Final_Appraisal.xlsx'
file_path = folder_path + "/" + file_name


# Create a Pandas DataFrame of the "CUV" worksheet
df = pd.read_excel(file_path, sheet_name='CUV', header=None, usecols="A:D", index_col=None)


# Extract company name and first word to include in worksheet name
company_name = df.iloc[1, 0]
co_first_word = company_name.split()[0]


# Write data to new Excel workbook
df.to_excel("cuv-extract.xlsx", sheet_name=co_first_word + "_CUV", header=False, index=False)
print("Done")