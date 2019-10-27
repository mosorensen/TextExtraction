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

comp_cat = ["intellectual ability", "personal effectiveness", 
            "team leadership", "interpersonal effectiveness", 
            "contextual skills/experience"]
# loops through the table and decides what table it is based on if any of
# its cells contains certain strings. To add robustness, we could
# check if multiple cells fulfill certain criteria.
def which_table(table):
    table_name="other"
    num_comp = 0
    for row in table.rows:
        last_text = ""
        for cell in row.cells:
            curr_text = cell.text
            #Check 1: does the table give identifying information
            if "Candidate Name" == curr_text.strip():
                table_name="info"
            #Check 2: does the table give competency information
            if curr_text.lower() in comp_cat:
                if last_text=="":
                    last_text = curr_text
                else:
                    if last_text==curr_text: #a merged cell is interpreted as two cells with the same contents
                        num_comp+=1 #keep track of the number of listed competencies
                    else:
                        last_text=curr_text #if not a merged cell, update last_text
    if num_comp == 5: #the table must identify all five competency categories
        table_name = "competencies"
            
    return table_name









