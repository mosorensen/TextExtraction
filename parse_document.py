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

import pandas as pd
from docx import Document
from parse_table import parse_table
from which_table import which_table

tables_to_parse = ["info", "competencies"]
def parse_document(doc_id, path):
    document = Document(path)
    doc_dict = {} #dictionary of all tables
    doc_data = pd.DataFrame({'doc_id':[doc_id], 'doc_path':[path]})
    table_id = 0 #tracks the table index
    for table in document.tables:
        table_id += 1
        table_type = which_table(table)
        if table_type not in tables_to_parse:
            print(table_type, "not parsed")
            continue
        print(table_type, "parsing")
        table_dict, table_data = parse_table(table, doc_id) #parse the table and its dictionary and dataframe form
        # add the table to the document dictionary and document table
        doc_dict[table_type]=table_dict 
        doc_data = pd.merge(doc_data, table_data, on='doc_id', how='outer')
    return (doc_dict, doc_data)












