# TextExtraction
Code that takes word documents as inputs and outputs the useful information

process_all.py is a python script that iterates over all the documents that 
	are explicitly listed in the document_paths list and used parse_document
	to extract and store the information in these documents once as a dictionary
	and once as a dataframe. This file also determines which tables we want to parse
	in the underlying documents.
parse_document.py is a python file containing the parse_document function. The
	function takes as inputs document id, path, and the types of tables it should parse.
	In then imports the table, and parses only the relevant information by first
	identifying a table based on the output of which_table. It then stores this
	information in a document specific dictionary and document specifc dataframe row.
which_table.py reads in a table and outputs the type of information that it contains.
	This is helpful if different types of tables vary in their format and, thus,
	should be parsed accordingly. It is also helpful in assigning a name to a given
	table in the document dictionary.

parse_table.py reads in a table and extracts the relevant information. In the future,
	we can either write seperate functions for parse_table that parse in only specific
	tables, or we can add an input that tells parse_table what kind of table it is
	being asked to parsed. It returns the table information both as a dictionary and
	as a dataframe row.


process.pl is a perl file that Prof. Sorenson created to read in PDF interviews and parse their contents.
Extract.py is an old file that contains all the functions written at the time. I do not suggest
	using it, as collaboration is easier when people are focused on seperate files.