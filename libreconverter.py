#!/usr/bin/python3
#
# Requires Python3
# *Please* make sure to use the version of Python included with
# your copy of LibreOffice.
#
# Convert spreadsheet to CSV file.
#
# Based on:
#   PyODConverter (Python OpenDocument Converter) v1.0.0 - 2008-05-05
#   Copyright (C) 2008 Mirko Nasato <mirko@artofsolving.com>
#   Licensed under the GNU LGPL v2.1 - or any later version.
#   http://www.gnu.org/licenses/lgpl-2.1.html
#

import os
import re
import loutils

import uno
from com.sun.star.task import ErrorCodeIOException



class LibreConverter:
    """
    Spreadsheet converter class.
    Converts spreadsheets to CSV files.
    """

    def __init__(self, lorunner=None):
        self.desktop  = None
        self.lorunner = None


    def convert(self, inputFile, outputFile, verbose=False):
        """
        Convert the input file (a spreadsheet) to a CSV file.

        The input file name can contain a sheet specification to specify a particular sheet.
        The sheet specification is either a number or a sheet name.
        The sheet specification is appended to the file name separated by a colon
        or an at sign: ":" or "@".

        If the output file name contains a %d or %s format specifier, then all the sheets
        in the input file are converted, otherwise only the first sheet is converted.

        If the output file name contains a %d format specifier then the sheet number
        is used when formatting the output file name.
        The format can contain a width specifier (eg %02d).

        If the output file name contains a %s specifier then the sheet name is used
        when formatting the output file name.
        """

        # Start openoffice if needed.
        if not self.desktop:
            if not self.lorunner:
                self.lorunner = loutils.LORunner()

            self.desktop = self.lorunner.connect()

        # Check for sheet specification in input file name.
        match = re.search(r'^(.*)[@:](.*)$', inputFile)
        if os.path.exists(inputFile) or not match:
            inputUrl   = uno.systemPathToFileUrl(os.path.abspath(inputFile))
            inputSheet = '1'   # Convert fist sheet.
        else:
            inputUrl   = uno.systemPathToFileUrl(os.path.abspath(match.group(1)))
            inputSheet = match.group(2)


        # NOTE:
        #   Sheet activation does not work properly when Hidden is specified.
        #   Although the sheet does become the active sheet, it's not the sheet that
        #   gets saved if the spreadsheet is loaded with Hidden=True.
        #
        #   Removing Hidden=True doesn't seem to change anything: nothing appears
        #   on the screen regardless of the Hidden value.
        #
        # document  = self.desktop.loadComponentFromURL(inputUrl, "_blank", 0, loutils.lo_properties(Hidden=True))
        document  = self.desktop.loadComponentFromURL(inputUrl, "_blank", 0, loutils.lo_properties())

        try:
            props = loutils.lo_properties(FilterName="Text - txt - csv (StarCalc)")
            #
            # Another useful property option:
            #   FilterOptions="59,34,0,1"
            #     59 - Field separator (semicolon), this is the ascii value.
            #     34 - Text delimiter (double quote), this is the ascii value.
            #      0 - Character set (system).
            #      1 - First line number to export.
            #
            # For more information see:
            #   http://wiki.services.openoffice.org/wiki/Documentation/DevGuide/Spreadsheets/Filter_Options

            # To convert a particular sheet, the sheet needs to be active.
            # To activate a sheet we need the spreadsheet-view, to get the spreadsheet-view
            # we need the spreadsheet-controller, to get the spreadsheet-controller
            # we need the spreadsheet-model.
            #
            # The spreadsheet-model interface is available from the document object.
            # The spreadsheet-view interface is available from the controller.
            #
            controller = document.getCurrentController()
            sheets     = document.getSheets()

            # If the output file name contains a %d or %s format specifier, convert all sheets.
            # Use the sheet number if the format is %d, otherwise the sheet name.
            dfmt = re.search(r'%[0-9]*d', outputFile)
            sfmt = re.search(r'%s', outputFile)

            if dfmt  or  sfmt:
                i = 0
                while i < sheets.getCount():
                    # Activate the sheet.
                    sheet = sheets.getByIndex(i)
                    controller.setActiveSheet(sheet)

                    # Create output file name.
                    if dfmt:
                        ofile = outputFile % (i+1)
                    else:
                        ofile = outputFile % sheet.getName().replace(' ', '_')

                    if verbose: print( "    %s" % ofile)

                    # Save the sheet to the output file.
                    outputUrl = uno.systemPathToFileUrl(os.path.abspath(ofile))
                    document.storeToURL(outputUrl, props)
                    i += 1

            else:
                # Activate the sheet to be converted.
                if re.search(r'^\d+$', inputSheet):
                    sheet = sheets.getByIndex(int(inputSheet)-1)
                else:
                    sheet = sheets.getByName(inputSheet)

                controller.setActiveSheet(sheet)
                outputUrl = uno.systemPathToFileUrl(os.path.abspath(outputFile))
                document.storeToURL(outputUrl, props)
        finally:
            if document: document.close(True)


if __name__ == "__main__":
    from sys import argv
    from os.path import isfile

    if len(argv) == 2  and  argv[1] == '--shutdown':
        loutils.lo_shutdown_if_running()
    else:
        if len(argv) < 3  or  len(argv) % 2 != 1:
            print("USAGE:")
            print("  python %s INPUT-FILE[:SHEET] OUTPUT-FILE ..." % argv[0])
            print("OR")
            print("  python %s --shutdown" % argv[0])
            exit(255)

        try:
            i = 1
            converter = LibreConverter()

            while i+1 < len(argv):
                print('%s => %s' % (argv[i], argv[i+1]))
                converter.convert(argv[i], argv[i+1], True)
                i += 2

        except ErrorCodeIOException as exception:
            print("ERROR! ErrorCodeIOException %d" % exception.ErrCode)
            exit(1)
