# Import file
# Import modules and set parameters
import pandas as pd
import numpy as np
import os


# Set parameters
pd.set_option('display.max_rows', None)


# Use predefined path and file for now but will need to allow for user to select/input this info
folder_path = '/Users/dougroehm/Code/repos/mtdor-extract-cuv/TY23-cap-final-appraisals/'
file_name = 'U202_TY23_CAP_Final_Appraisal.xlsx'
file_path = folder_path + file_name


# Create a Pandas DataFrame from file
# df = pd.ExcelFile(file_path)


# Create a Pandas DataFrame of the "CUV" worksheet
df = pd.read_excel(file_path, sheet_name='CUV', header=None, usecols="A:D", names=["Col1", "Col2", "Col3", "Col4"])


# Use .notna to remove rows that are blank
df_notna = df[df['Col4'].notna()]


# Extract the rows to be kept
df_to_extract = df_notna[df_notna['Col1'].str.contains("Cost") | 
df_notna['Col1'].str.contains("Capitalization") | 
df_notna['Col1'].str.contains("Stock") |
df_notna['Col1'].str.contains("Correlated Unit Value") |
df_notna['Col1'].str.contains("Montana Allocation Factor") |
df_notna['Col1'].str.contains("Montana Allocated Value") |
df_notna['Col1'].str.contains("Value to be Apportioned")
]

print(df_to_extract)

# I am going to need to pull the data using iloc because there is too much overlap in string names