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
from parse_document import parse_document
import pandas as pd

document = Document("C:/Users/ydecress/Desktop/Cleaning_Small.docx")
for para in document.paragraphs:
    print(para.text)

doc_paths = ["C:/Users/ydecress/Desktop/Cleaning_Small.docx", "C:/Users/ydecress/Desktop/Cleaning_Small_Copy.docx"]
tables_to_parse = ["info", "competencies"]

for doc_id in range(len(doc_paths)):
    #first read in the given document
    print(doc_id)
    path = doc_paths[doc_id]
    doc_dict, doc_data = parse_document(doc_id, path, tables_to_parse)
    candidate_name=doc_dict['info']['Candidate Name']
    #if this was the first document, we need to create the master dataframe
    #each row will be one executive and each column will contain a variable
    #each column gets a unique doc_id and path to document
    if doc_id==0: 
        data_master = doc_data
        dict_master = {candidate_name:doc_dict}
    else:
        data_master = data_master.append(doc_data)
        dict_master[candidate_name] = doc_dict


dict_master
data_master









