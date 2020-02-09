"""
xl2csv.py

Reads an excel file, and for every sheet named SHEET, writes a csv file SHEET.csv

The main work is done by Python pandas module, so you will need that installed.

See usage() function below for all options.

(c) Peter Ballard 2020
Free to reuse and modify under terms of GPL
"""

import sys
import os
import pandas
# might need to install xlrd with "pip install xlrd"
# but no need to import it; pandas uses it internally

def desc():
    sys.stdout.write(
        """Reads an Excel file, and writes one CSV file per sheet (SHEET_NAME.csv),
The main work is done by Python pandas module, so you will need that installed.
""")

def usage():
    sys.stdout.write(
        """Usage:
    python xl2csv.py <xlfile> [-f=fmt] [-n] [-q]
      -f=fmt = float_format parameter for pandas to_csv method, default is "%g".
      -n = not overwrite existing files (default is overwrite).
      -q = quiet (default is to echo some messages).
      -h or -u to display this message and exit.
""")
    sys.exit(1)

def main():
    # float_format="%g" strips the decimal points from numbers have integer values,
    # so e.g. 7.0 is displayed as 7,
    # which is usually what we want.
    float_format = "%g"
    nov = 0
    quiet = 0
    xlfile = ""

    for arg in sys.argv[1:]:
        if arg[0]=="-":
            if arg[:3]=="-f=":
                float_format = arg[3:]
            elif arg=="-n":
                nov = 1
            elif arg=="-q":
                quiet = 1
            elif arg in ("-h", "-u"):
                desc()
                usage()
            else:
                sys.stdout.write("Error: unsupported arg %s\n\n" % arg)
                usage()
        else:
            if xlfile!="":
                sys.stdout.write("Error: two Excel files, %s and %s\n" % (xlfile, arg))
                usage()
            xlfile = arg

    if xlfile=="":
        sys.stdout.write("Error: Excel file not specified\n")
        usage()

    # read the excel file. 
    # By specifying sheet_name=None, this returns a dictionary,
    #  where each key is the sheet name
    #  and each value is a pandas data frame
    pdict = pandas.read_excel(xlfile, sheet_name=None)

    ccount = 0
    for sheet_name in pdict.keys():
        ofile = sheet_name + ".csv"
        
        if nov and os.path.exists(ofile):
            sys.stdout.write("Not overwriting %s\n" % ofile)
            continue
        
        # by default pandas adds an index number as the first field of every row,
        #  which is undesirable because that is not what you get when you "save as CSV" from inside Excel
        #  so "index=False" turns this off and gives output in the same format as Excel "save as CSV"
        #
        # Excel only stores strings or floats,
        # and float_format="%g" strips the decimal points from numbers which have integer values,
        # so e.g. 7.0 is displayed as 7,
        # which is usually what we want.
        #
        s = pdict[sheet_name].to_csv(ofile, index=False, float_format=float_format) 

        if not quiet:
            sys.stdout.write(sheet_name + ", ")
        ccount += 1

    if not quiet:
        sys.stdout.write("\nWrote %d csv files\n" % (ccount))
        
main()
