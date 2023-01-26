# docx-parser-converter

This script is capable of gathering information from .docx and .txt files for the purpose of collecting that data and inserting into a new word document template.

## Guide:

The script is setup to run in batches per folder directory and configured in `config.py`

Edit the `base_path` and `area` variables in `config.py` to adjust the program to whatever directory structure you have. This script requires the .docx and .txt files to be in the area folder, and for there to be two additional folders: NEW and TEMPLATE where the new docx files are saved and the EMPTY_TEMPLATE.docx resides respectively.

The program outputs two csv files in the main directory. The purpose of these files is to review the data parsed from the old documents and then inserted into the new template. 

A new parsing template can be created in the `parsers` folder and also an inserter template in the `inserters` folder. The current ones in use are `default.py` for each one.
