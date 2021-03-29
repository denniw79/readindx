# -*- coding: utf-8 -*-
"""
Created on Mon Mar 29 16:51:30 2021

@author: Dennis
"""

import pandas as pd
import io
import os
import sys
import glob

def get_script_path():
    return os.path.dirname(os.path.realpath(sys.argv[0]))

liste = []
i= 0
for file in glob.glob(os.path.join(get_script_path() + "/*.txt")):
        i+=1
        with io.open(file, mode="r", encoding="utf-8") as fd:
            text = fd.read()
            liste.append([i, text, 0])
next
df_r = pd.DataFrame(liste, columns={'Satznummer', 'Satz', 'Version'})

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('bundledTexts.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df_r.to_excel(writer, sheet_name='Sheet1', index=False)

# Close the Pandas Excel writer and output the Excel file.
writer.save()
