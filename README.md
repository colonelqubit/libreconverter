libreconverter
==============

Convert spreadsheets to CSV text files

Based on code from
www.linuxjournal.com/content/convert-spreadsheets-csv-files-python-and-pyuno-part-1v2

WHAT WORKS?

Nothing, yet. This code is from 2009, so it might need a little
dust-off to run with modern LibreOffice and python3.


USAGE (per the article):

Provide pairs of SPREADSHEET OUTPUT-FILE like this:

  $ python ssconverter.py file1.xls file1.csv file2.ods file2.csv

To select a particular sheet, you may append a number or a sheetname to the input filepath using a colon or @ sign:

  $ python ssconverter.py file1.xls:1      file1.csv
  $ python ssconverter.py file1.xls:Sheet1 file1.csv
  $ python ssconverter.py file2.ods@1      file2.csv
  $ python ssconverter.py file2.ods@Sheet2 file2.csv

To convert all the things, use %d or %s -- those will spit out files named by sheet number or by sheet name, respectively:

  $ python ssconverter.py file1.xls file1-%d.csv
  $ python ssconverter.py file1.xls file1-%s.csv

When using the %d format, you may include zero pad and width specifiers (e.g. %02d).
