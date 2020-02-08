"""
xl2csv.py

Reads an excel file, and for every sheet named SHEET:
 - writes a csv file SHEET.csv
 - writes a TAB-separated file SHEET.tsv

The main work is done by Python pandas module, so you will need that installed.

See usage() function below for all options.

(c) Peter Ballard 2020
Free to reuse and modify under terms of GPL
"""

import sys
import os
import csv
import pandas
# might need to install xlrd with "pip install xlrd"
# but no need to import it; pandas uses it internally

def desc():
    sys.stdout.write(
        """Reads an Excel file, and writes one CSV file per sheet (SHEET_NAME.csv),
and optionally also one TSV (TAB separated values) file per sheet (SHEET_NAME.tsv).
The main work is done by Python pandas module, so you will need that installed.
""")
    
def usage():
    sys.stdout.write(
"""
Usage:
    python xl2csv.py <xlfile> [-f=fmt] [-n] [-q] [-t] [-t7]
      -f=fmt = float_format parameter to pandas to_csv method. Default is "%g".
      -n = not overwrite existing files (default is overwrite).
      -q = quiet (default is to echo some messages).
      -s = strip leading and trailing whitespace from sheet names
      -t = write TSV files (in addition to CSV files).
      -t7 = write TSV files with fields truncated to 7 chars, so the columns line up. This overrides -t.
      -h or -u to display this message and exit.
""")

def csv2tsv(ifile, nov, t7):
    if ifile.count(".csv"):
        ofile = ifile.replace(".csv", ".tsv", 1)
    else:
        ofile = ifile + ".tsv"

    if nov and os.path.exits(ofile):
        sys.stdout.write("Not overwriting %s\n" % ofile)
        return 0 # indicating 0 files written
        
    ip = open(ifile, "r", newline="") # csv.reader should open files using newline='', see https://docs.python.org/3/library/csv.html
    csvobj = csv.reader(ip)
    op = open(ofile, "w")
  
    for row in csvobj:
         for j in range(len(row)):
            if t7 and j < len(row)-1:
                # truncate to 7 chars
                s = row[j][:7]
            else:
                # no need to truncate last column
                s = row[j]
            # replace any TABS with spaces
            s = s.replace("\t", " ")
            if j < len(row)-1:
                op.write(s + "\t")
            else:
                op.write(s + "\n")
    ip.close()
    op.close()
    return 1 # indicating 1 file written

def main():
    # float_format="%g" strips the decimal points from numbers have integer values,
    # so e.g. 7.0 is displayed as 7,
    # which is usually what we want.
    float_format = "%g"
    nov = 0
    quiet = 0
    strip = 0
    tsv = 0
    t7 = 0
    xlfile = ""

    for arg in sys.argv[1:]:
        if arg[0]=="-":
            if arg[:3]=="-f=":
                float_format = arg[3:]
            elif arg=="-n":
                nov = 1
            elif arg=="-n":
                nov = 1
            elif arg=="-q":
                quiet = 1
            elif arg=="-s":
                strip = 1
            elif arg=="-t":
                tsv = 1
            elif arg=="-t7":
                tsv = 1
                t7 = 1
            elif arg in ("-h", "-u"):
                desc()
                usage()
                sys.exit(0)
            else:
                sys.stdout.write("Error: unsupported arg %s\n" % arg)
                usage()
                sys.exit(1)
        else:
            xlfile = arg

    if xlfile=="":
        sys.stdout.write("Error: Excel file not specified\n")
        usage()
        sys.exit(1)

    # read the excel file. 
    # By specifying sheet_name=None, this returns a dictionary,
    #  where each key is the sheet name
    #  and each value is a pandas data frame
    pdict = pandas.read_excel(xlfile, sheet_name=None)

    used = []
    for sheet_name in pdict.keys():
        used.append(sheet_name)

    ccount = 0
    tcount = 0
    for sheet_name in pdict.keys():
        if strip:
            if sheet_name.strip() != sheet_name and sheet_name.strip() in used:
                sys.stdout.write('Keeping sheet name as "%s" (note extra spaces), because sheet name "%s" already exists\n'
                                 % (sheet_name, sheet_name.strip()))
                oname = sheet_name
            else:
                oname = sheet_name.strip()
                used.append(oname)
        else:
            oname = sheet_name
        ofile = oname + ".csv"
        
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

        # optionally also write TAB-separated file
        if tsv:
            tcount += csv2tsv(ofile, nov, t7)

        if not quiet:
            sys.stdout.write(oname + ", ")
        ccount += 1

    if not quiet:
        sys.stdout.write("\nWrote %d csv files and %d tsv files\n" % (ccount, tcount))
        
main()
