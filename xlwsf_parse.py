#-------------------------------------------------------------------------------
# Name:        xlwsf_parse
# Purpose:     Parse Microsoft webpage source for Excel functions and page
#               addresses
#
# Author:      Brian Skinn
#                bskinn@alum.mit.edu
#
# Created:     24 May 2016
# Copyright:   (c) Brian Skinn 2016
# License:     The MIT License; see "license.txt" for full license terms
#                   and contributor agreement.
#
#       This file is a helper script for rapid generation of an objects.inv
#       inventory file for intersphinx cross-referencing of Excel worksheet
#       functions.
#
#       http://www.github.com/bskinn/intersphinx-xlwsf
#
#-------------------------------------------------------------------------------


import re, sys, time


def main():

    # Try loading the input file
    try:
        with open(sys.argv[1], 'r') as f:
            datastr = f.read()

    except OSError:
        print("\nIndicated file not found.  Exiting...\n")
        sys.exit(1)

    except IndexError:
        print("\nNo input filename passed. Exiting...\n")
        sys.exit(1)

    # Define the regex pattern
    pat = re.compile('a href="(?P<uri>[^"]+)" class="[^"]+" '
            'title="(?P<fname>[^ ]+) [^"]+">', re.I)

    # Pull the function names and URIs
    dic = {m.group("fname"): m.group("uri") for m in pat.finditer(datastr) if
            m.group("fname") == m.group("fname").upper()}

    # Open and write the objects.txt file
    with open('objects.txt', 'w') as f:
        # First header line
        f.write("# Sphinx inventory version 2\n")

        # Second header line, with project name
        f.write("# Project: Excel WSF\n")

        # Third header line, with timestamp
        f.write("# Version: {0}\n".format(time.strftime('%Y-%m-%d %H:%M:%S')))

        # Fourth header line, verbatim
        f.write("# The remainder of this file is compressed using zlib.\n")

        # Dump the data lines
        for func, uri in sorted(dic.items()):
            f.write("{0} py:function 1 article/{1} -\n".format(func, uri))


if __name__ == '__main__':
    main()

