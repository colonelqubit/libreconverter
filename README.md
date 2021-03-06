libreconverter
==============

Convert spreadsheets to CSV text files

The killer feature of this particular script is that it can convert
multi-sheet spreadsheets into CSV files, either sheet-by-sheet, or all
sheets at once.

Based on code from
www.linuxjournal.com/content/convert-spreadsheets-csv-files-python-and-pyuno-part-1v2

WHAT WORKS?
-----------

General conversion.... enjoy!


USAGE
-----

*(per the article)*

Provide pairs of SPREADSHEET OUTPUT-FILE like this:

```
  $ LO_PYTHON libreconverter.py file1.xls file1.csv file2.ods file2.csv
```

Note that you'll need to determine the right path to the python executable bundled with LibreOffice to run this script properly. E.g. if libreoffice is located at

```
/usr/lib/libreoffice/program/soffice
```

...then you'll want to use the python executable located at

```
/usr/lib/libreoffice/program/python
```

To select a particular sheet, you may append a number or a sheetname to the input filepath using a colon or @ sign:

```
  $ LO_PYTHON libreconverter.py file1.xls:1      file1.csv
  $ LO_PYTHON libreconverter.py file1.xls:Sheet1 file1.csv
  $ LO_PYTHON libreconverter.py file2.ods@1      file2.csv
  $ LO_PYTHON libreconverter.py file2.ods@Sheet2 file2.csv
```

To convert all the things, use %d or %s -- those will spit out files named by sheet number or by sheet name, respectively:

```
  $ LO_PYTHON libreconverter.py file1.xls file1-%d.csv
  $ LO_PYTHON libreconverter.py file1.xls file1-%s.csv
```

When using the %d format, you may include zero pad and width specifiers (e.g. %02d).

-----

Running LibreOffice Headless
----------------------------

You may either run LibreOffice headless first (as a daemon), and then call this script, or you may let the script run LibreOffice itself.

To run the latest version of LibreOffice you have available:

```
walrus@lo:~/LibreOfficeDev4.2$ /usr/lib/libreoffice/program/soffice --nologo --headless --nofirststartwizard --accept='socket,host=127.0.0.1,port=8100,tcpNoDelay=1;urp'
```

In a separate terminal window, invoke the libreconverter script using the path to the python (LO_PYTHON) bundled with *the same version* of LibreOffice.

If you haven't already started up LibreOffice, make sure that the variable *_lopaths* in *loutils.py* has been set to something reasonable that can find LibreOffice. Here's our current default:

```
# Find LibreOffice.
_lopaths=(
    ('/usr/lib/libreoffice/program', '/usr/lib/libreoffice/program')
    )
```

Now invoke libreconverter:

```
walrus@lo:~/some/test/dir$ /usr/lib/libreoffice/program/python libreconverter.py multi-sheet-spreadsheet.ods:2 output.csv
```

Good luck!