# -*- coding: utf-8 -*-
"""
This script reads in one word document.
It then reads in all of the tables  in that word document.
For each of these tables, it creates a dataframe.

To do:
    1) Instead of creating new dataframes, have future documents add new 
        rows to the existing dataframes, one row per executive
    2) Automatically detect which table is being interpreted. This can only
        be done once we look at many documents, as the tables might vary
        by document. This could be implemented by checking if the table
        contains certin key words, or by trying to find the header immediately
        prior to the table.
        i) Front Page Info
        ii) ghSMART Competency Ratings
        iii) Key Strengths and Key Risks
    3) Output the dataframe to an excel file. Preferable similar in structure
        to the excel file already created.
    4) Try to detect other relevant information
        i) Table of Contents -- to get an overview of frequency of various 
                data in the overall dataset
        ii) Firm-Specific Mission Text
        iii) Firm-Specific Key Outcomes Ratings/Comments
        iv) Company Specific Questions
        v) Candidate Specific Recommendations
        vi) Career Goals and Motivations
    5) Transform the entire document into one parseable text. Then check
        to see how many times certain words occur in the text using
        dictionaries or arrays of strings. This can be a crude linguistic
        approach i.e. measuring the frequency of words associated with
        overconfidence (overconfident, assertive, optimistic, etc.) versus
        the frequency of words associated with caution (cautious, timid,
        scientific, etc.).

Notes:
    - Merged Cells are read as multiple cells all with the same information
"""

"""
for p in document.paragraphs:
    print (p.text)
    
for table in document.tables:
    for row in table.rows:
        for cell in row.cells:
            print (cell.text)
"""


from docx import Document
import pandas as pd

document_id = 1 #tracks the document index
doc_paths = ["C:/Users/ydecress/Desktop/Cleaning_Small.docx", "C:/Users/ydecress/Desktop/Cleaning_Small_Copy.docx"]
for path in doc_paths:
    #first read in the given document
    document = Document(path)
    all_tables = {} #dictionary of all tables
    table_ind = 1 #tracks the table index
    for table in document.tables:
        table_data = {}
        table_ind += 1
        keys = None
        #i indexes the rows, row is the row object
        for i, row in enumerate(table.rows):
            text = (cell.text for cell in row.cells) #converts all tables cells to text  
            text = tuple(text) #need to convert from row object to a tuple
            if i == 0: # Establish the table name based on first row
                key = text[0]
                table_key = key
                entries = text[1:]
                if table_ind==1: #if this is the first table in the document
                    table_key = "info"
                table_data["table_name"] = table_key
            
            elif i !=0 : # Establishing the contents in the rows afterwards
                j=0 #let j index over the entries in the row tuple
                while j < len(text)-1: #while we have at least one more entry in the row
                    #if len(text) > 2: #the key must be longer than 2 characters
                    key = text[j]
                    if len(key) == 0:
                        j+=1 #if the current cell is empty, skip it
                        continue
                    entries = text[j+1]
                    j=j+2
                    if key==entries:
                        continue  #if the cell if a merged cell, skip it
                    
                    # add the keys and entries into a table specific dictionary
                    table_data[key] = entries
                #this code below saves the entries as an array
                #key = text[0]
                #entries = text[1:]
                #print(key)
                #print(entries)
                
        # add the table to the overall document dictionary
        all_tables[table_key]=table_data
        print(table_key)
    
    #if this was the first document, we need to create the master dataframe
    #each row will be one executive and each column will contain a variable
    #value for that executive. The first column contains the document id, the 
    #second column contains the filename
    if document_id==1:
        pass #create the dataframe
    else:
        pass #add the dataframe
    document_id+=1


    
all_tables


cover_page = all_tables['SmartAssessment Report']
cover_page_df = pd.DataFrame.from_dict(cover_page, 'columns')
cover_page_df




