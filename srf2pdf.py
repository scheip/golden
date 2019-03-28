#------------------------------
# script name: srf2pdf.py
# project name: golden
#
#
# Author: Corey Scheip, AECOM Technical Services of North Carolina
#
# Creation date: 2018/12/18
#
# Instructions: Either input a directory as argument to this script or it will default
# to the directory containing the script (e.g. C:\SRF_Files). When run, a PDf folder will
# be generated in your working directory (e.g. C:\SRF_Files\pdf and PDF files exported to
# that location.
#
#------------------------------

import sys, os
try:
    import win32com.client   # if this fails, you must install module pywin32
except:
    print('Failed to import win32com.client, try installing module pywin32')
    sys.exit(0)

# Create application object
surf = win32com.client.Dispatch("Surfer.Application")
surf.Visible = False

# Check for user input directory. If this doesn't exist, default to directory containing script
if sys.argv[1]:
    pwd = sys.argv[1]
else:
    pwd = os.path.dirname(sys.argv[0])

# Get list of files from working directory
fileList = os.listdir(pwd)

# Check for PDF directory
outDir = os.path.join(pwd, 'pdf')
if not os.path.isdir(outDir):
    os.mkdir(outDir)
    print ('PDF Directory created')

# Now go through files, open SRF and print to PDF
for fil in fileList:
    print 'Checking file... ' + fil
    if fil.endswith('.srf'):
        # Name to use for printing
        printName = os.path.basename(fil)[:-4]

        # Open surfer file
        print ' ...opening file...'
        CurrentDoc = surf.Documents.Open(os.path.join(pwd, fil))
        pdfFile = os.path.join(outDir, printName + '.pdf')
        print ' ...printing file...'
        CurrentDoc.Export2(pdfFile, False,  "HDPI=300, VDPI=300", "pdfi")

        # Close surfer file
        CurrentDoc.Close()

# Delete COM object
del surf
