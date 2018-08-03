#   `               NOTES FROM JOHNNY:
#This program takes in an xlsm file and outputs the vba code.
#It is run from the command line by typing in "py vba_extract.py xlsm_file"
#If successful, the program will print "Extracted _____.bin" and the vba code
# will be saved in the folder containing vba_extract.py and the xlsm file.
#I wrote 0% of this code. The person below wrote everything and I found
# his code on Google
##############################################################################
#
# vba_extract - A simple utility to extract a vbaProject.bin binary from an
# Excel 2007+ xlsm file for insertion into an XlsxWriter file.
#
# Copyright 2013-2018, John McNamara, jmcnamara@cpan.org
#
import sys
import shutil
from zipfile import ZipFile
from zipfile import BadZipfile

def extractVBA(macro_file):
    print("Extracting vba from ", macro_file)

    # The VBA project file we want to extract.
    vba_filename = 'vbaProject.bin' #you can change this name but make sure to keep the ".bin" extension

    # Get the xlsm file name from the commandline.
    xlsm_file = macro_file
    #This portion actually extracts the vba code
    try:
        # Open the Excel xlsm file as a zip file.
        xlsm_zip = ZipFile(xlsm_file, 'r')

        # Read the xl/vbaProject.bin file.
        vba_data = xlsm_zip.read('xl/' + vba_filename)

        # Write the vba data to a local file.
        vba_file = open(vba_filename, "wb")
        vba_file.write(vba_data)
        vba_file.close()
        return(vba_filename)

    except:
        # Catch any other exceptions.
        print("File error: %s" % str(macro_file))
        exit()




def main():
    vba_filename = 'vbaProject.bin' #you can change this name but make sure to keep the ".bin" extension

    # Get the xlsm file name from the commandline.
    if len(sys.argv) > 1:
        xlsm_file = sys.argv[1]
    else:
        print("\nUtility to extract a vbaProject.bin binary from an Excel 2007+ "
              "xlsm macro file for insertion into an XlsxWriter file."
              "\n"
              "See: https://xlsxwriter.readthedocs.io/working_with_macros.html\n"
              "\n"
              "Usage: vba_extract file.xlsm\n")
        exit()

    #This portion actually extracts the vba code
    try:
        # Open the Excel xlsm file as a zip file.
        xlsm_zip = ZipFile(xlsm_file, 'r')

        # Read the xl/vbaProject.bin file.
        vba_data = xlsm_zip.read('xl/' + vba_filename)

        # Write the vba data to a local file.
        vba_file = open(vba_filename, "wb")
        vba_file.write(vba_data)
        vba_file.close()

    #All of the "except" blocks of code are for if the input file is in a bad format
    except IOError:
        # Use exc_info() for Python 2.5+ compatibility.
        e = sys.exc_info()[1]
        print("File error: %s" % str(e))
        exit()

    except KeyError:
        # Usually when there isn't a xl/vbaProject.bin member in the file.
        e = sys.exc_info()[1]
        print("File error: %s" % str(e))
        print("File may not be an Excel xlsm macro file: '%s'" % xlsm_file)
        exit()

    except BadZipfile:
        # Usually if the file is an xls file and not an xlsm file.
        e = sys.exc_info()[1]
        print("File error: %s: '%s'" % (str(e), xlsm_file))
        print("File may not be an Excel xlsm macro file.")
        exit()

    except:
        # Catch any other exceptions.
        e = sys.exc_info()[1]
        print("File error: %s" % str(e))
        exit()

    print("Extracted: %s" % vba_filename)


if __name__ == "__main__":
    main()
