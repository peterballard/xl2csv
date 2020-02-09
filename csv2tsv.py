"""
csv2tsv.py

Converts csv to TAB-separated (tsv)

See usage() function below for all options.

(c) Peter Ballard 2020
Free to reuse and modify under terms of GPL
"""

import sys
import os
import csv

def usage():
    sys.stdout.write(
"""
Usage:
    python csv2tsv.py <files> [-n] [-7]
      -n = not overwrite existing files (default is overwrite).
      -t7 = write TSV files with fields truncated to 7 chars, so the columns line up.
      -h or -u to display this message and exit.
""")
    sys.exit(1)

def csv2tsv(ifile, nov, t7):
    if ifile.count(".csv"):
        ofile = ifile.replace(".csv", ".tsv", 1)
    else:
        ofile = ifile + ".tsv"

    if nov and os.path.exists(ofile):
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
    nov = 0 # not overwrite (default is overwrite)
    t7 = 0  # limit to 7 chars (default is no limiting)
    ifiles = []

    for arg in sys.argv[1:]:
        if arg[0]=="-":
            if arg=="-n":
                nov = 1
            elif arg in ("-t7", "-7"):
                t7 = 1
            elif arg in ("-h", "-u"):
                usage()
            else:
                sys.stdout.write("Error: unsupported arg %s\n" % arg)
                usage()
        else:
            ifiles.append(arg)

    if len(ifiles)==0:
        sys.stdout.write("Error: No input files specified\n")
        usage()

    for ifile in ifiles:
        csv2tsv(ifile, nov, t7)
        
main()
