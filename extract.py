# -*- coding: utf-8 -*-
"""
This script reads in one word document.
It then reads in all of the tables in that word document.
It then creates a dictionary for each table, and and overall dictionary
    connecting the doc_id to these document specific dictionaries per table
It also creates a dataframe with a row per document and with entries
    corresponding to the identified variables in all of the tables

To do:
    - Automatically detect which table is being interpreted. 
        - This can only be done once we look at many documents, as the tables 
            might vary by document. This could be implemented by checking if 
            the table contains certin key words, or by trying to find the 
            header immediately prior to the table.
                - Front Page Info
                - ghSMART Competency Ratings 
        - Only add variables and table dictionary for these tables
        - create seperate table parsers for each type of table being saved
    - Output the dataframe to an excel file. Preferably similar in structure
        to the excel file already created.
    - Try to detect other relevant information
        - Table of Contents -- to get an overview of frequency of various 
                data in the overall dataset
        - Key Strengths and Key Risks
        - Firm-Specific Mission Text
        - Firm-Specific Key Outcomes Ratings/Comments
        - Company Specific Questions
        - Candidate Specific Recommendations
        - Career Goals and Motivations
    - Transform the entire document into one parseable text. Then check
        to see how many times certain words occur in the text using
        dictionaries or arrays of strings. This can be a crude linguistic
        approach i.e. measuring the frequency of words associated with
        overconfidence (overconfident, assertive, optimistic, etc.) versus
        the frequency of words associated with caution (cautious, timid,
        scientific, etc.).

Notes:
    - Merged Cells are read as multiple cells all with the same information
    - info on the docx package: https://python-docx.readthedocs.io/en/latest/index.html
    - info on the pandas package: https://pandas.pydata.org/
"""


from docx import Document
import pandas as pd

doc_paths = ["C:/Users/ydecress/Desktop/Cleaning_Small.docx", "C:/Users/ydecress/Desktop/Cleaning_Small_Copy.docx"]
for doc_id in range(len(doc_paths)):
    #first read in the given document
    print(doc_id)
    path = doc_paths[doc_id]
    document = Document(path)
    dict_tables = {} #dictionary of all tables
    data_tables = pd.DataFrame({'doc_id':[doc_id], 'doc_path':[path]})
    table_id = 0 #tracks the table index
    for table in document.tables:
        table_dict = {}
        table_id += 1
        #i indexes the rows, row is the row object
        for i, row in enumerate(table.rows):
            text = (cell.text for cell in row.cells) #converts all tables cells to text  
            text = tuple(text) #need to convert from row object to a tuple
            if i == 0: # Establish the table name based on first row
                table_key = text[0]
                if table_id==1: #if this is the first table in the document
                    table_key = "info"
                table_name = "table_name_" + str(table_id)
                data_tables[table_name] = table_key
                table_dict[table_name] = table_key
            
            elif i !=0 : # Establishing the contents in the rows afterwards
                j=0 #let j index over the entries in the row tuple
                while j < len(text)-1: #while we have at least one more entry in the row
                    key = text[j]
                    if len(key) == 0: #if the current cell is empty, skip it
                        j+=1
                        del key
                        continue
                    entries = text[j+1]
                    j=j+2 #the next key is after the current key's entry...so two indices away
                    if key==entries:
                        del key, entries
                        continue  #if the cell if a merged cell, skip it
                    table_dict[key] = entries # add the keys and entries into a table specific dictionary
                    data_tables[key] = entries
                    del key, entries
                #this code below saves the entries as an array
                #key = text[0]
                #entries = text[1:]
                #print(key)
                #print(entries)
                
        # add the table to the overall document dictionary...just keep adding columns to the dataframe
        dict_tables[table_key]=table_dict
        del table_key
    
    #if this was the first document, we need to create the master dataframe
    #each row will be one executive and each column will contain a variable
    #value for that executive. The first column contains the document id, the 
    #second column contains the filename
    if doc_id==0: 
        data_master = data_tables
        dict_master = {doc_id:dict_tables}
    else:
        data_master = data_master.append(data_tables)
        dict_master[doc_id] = dict_tables


dict_master
data_master


cover_page = all_tables['SmartAssessment Report']
cover_page_df = pd.DataFrame.from_dict(cover_page, 'columns')
cover_page_df




