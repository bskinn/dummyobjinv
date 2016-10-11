# ------------------------------------------------------------------------------
# Name:        xlwsf_statuscheck
# Purpose:     Parse objects.txt for Excel worksheet functions to validate
#                its list of page links.
#
# Author:      Brian Skinn
#                bskinn@alum.mit.edu
#
# Created:     11 Oct 2016
# Copyright:   (c) Brian Skinn 2016
# License:     The MIT License; see "license.txt" for full license terms
#                   and contributor agreement.
#
#       This file is a helper script for rapid validation of the list of page links
#       contained in objects.txt for the Excel worksheet function intersphinx
#       directory. It only checks for the *EXISTENCE* of the linked pages, not 
#       accuracy of content (i.e., whether the page is the correct one for the
#       function indicated in objects.txt).
#
#       http://www.github.com/bskinn/intersphinx-xlwsf
#
# ------------------------------------------------------------------------------


def main():

    import os
    from urllib.error import HTTPError
    import wget

    dlfile = 'dl.tmp'

    with open('objects.txt', 'r') as ifile:
        with open('log.txt', 'w') as ofile:
            for line in ifile.readlines()[4:]:
                fxn_name = line.split()[0]
                try:
                    wget.download('https://support.office.com/en-us/' + line[line.find('article/'):-3],
                                  out=dlfile)
                except HTTPError:
                    ofile.write('Failed on {0}\n'.format(fxn_name))
                except Exception as e:
                    ofile.write('Unknown problem on {0}: {1}\n'.format(fxn_name, str(e)))
                else:
                    ofile.write('OK on {0}\n'.format(fxn_name))
                finally:
                    ofile.flush()
                    if os.path.isfile(dlfile):
                        os.remove(dlfile)


if __name__ == '__main__':
    main()


