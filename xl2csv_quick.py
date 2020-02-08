"""
xl2csv_quick.py

Reads an excel file, and for every sheet named SHEET:
 - writes a csv file SHEET.csv

Quick version, with no switches or checking

The main work is done by Python pandas module, so you will need that installed.

(c) Peter Ballard 2020
Free to reuse and modify under terms of GPL
"""

import sys
import pandas
# might need to install xlrd with "pip install xlrd"
# but no need to import it; pandas uses it internally

def main():
    xlfile = sys.argv[1]

    # read the excel file. 
    # By specifying sheet_name=None, this returns a dictionary,
    #  where each key is the sheet name
    #  and each value is a pandas data frame
    pdict = pandas.read_excel(xlfile, sheet_name=None)

    for sheet_name in pdict.keys():
        ofile = sheet_name + ".csv"
        
        # by default pandas adds an index number as the first field of every row,
        #  which is undesirable because that is not what you get when you "save as CSV" from inside Excel
        #  so "index=False" turns this off and gives output in the same format as Excel "save as CSV"
        #
        # Excel only stores strings or floats,
        # and float_format="%g" strips the decimal points from numbers which have integer values,
        # so e.g. 7.0 is displayed as 7,
        # which is usually what we want.
        #
        s = pdict[sheet_name].to_csv(ofile, index=False, float_format="%g") 
        
main()
