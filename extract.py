# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""


from docx import Document
import pandas as pd

document = Document("C:/Users/ydecress/Desktop/Cleaning_Small.docx")
 

for p in document.paragraphs:
    print (p.text)
    
    
for table in document.tables:
    for row in table.rows:
        for cell in row.cells:
            print (cell.text)
            
            
all_tables = {}
for table in document.tables:
    table_data = {}
    keys = None
    #i indexes the rows, row is the row object
    for i, row in enumerate(table.rows):
        text = (cell.text for cell in row.cells) #converts all tables cells to text  
        text = tuple(text) #need to convert from row object to a tuple
        if i == 0: # Establish the table name based on first row
            key = text[0]
            table_key = key
            entries = text[1:]
        
        elif i !=0 : # Establishing the contents in the rows afterwards
            j=0 #let j index over the entries in the row tuple
            while j < len(text)-1:
                curr = text[j]
                if len(curr) > 2:
                    key = curr
                    entries = text[j+1]
                    j=j+2
                else:
                    j=j+1
                    print("not longer than 2")
                
            #key = text[0]
            #entries = text[1:]
            #print(key)
            #print(entries)

        # add the keys and entries into a table specific dictionary
        table_data[key] = entries
        
    # add the dictionary to the overall dictionary
    all_tables[table_key]=table_data
    print(table_key)

all_tables

cover_page = all_tables['SmartAssessment Report']

cover_page_df = pd.DataFrame.from_dict(cover_page, '')
cover_page_df




