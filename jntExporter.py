"""
This script accepts a directory as a parameter and will recursively search for
and convert .jnt files to .pdf files.
It needs the path of Windows Journal. Further it needs the path of pdfcreator.
Another variable can be set to exclude specific directory names from search.
"""

__author__ = ["Dietmar Malli"]
__copyright__ = "Copyright 2017, Dietmar Malli"
__credits__ = []
__license__ = "GPLv3"
__version__ = "1.0.5"
__maintainer__ = ["Dietmar Malli"]
__email__ = ["git.commits@malli.co.at"]
__status__ = "Production"

import os
import subprocess
import time
import shutil
import sys
from sharedExporterFunctions import get_output_path, notefile_needs_update, get_recursive_filelist

# Folders to exclude from jnt scan:
folders_to_exclude = ['.dropbox.cache']
# Windows journal.exe path:
journal = 'C:\\Program Files\\Windows Journal\journal.exe'
# PDF-Creator path:
pdfcreator = 'C:\\Program Files\\PDFCreator\\PDFCreator.exe'

# Find root folder:
root_folder = './'
if len(sys.argv) > 1:
    root_folder = sys.argv[1]

# Set PDFCreator PDF Printer as default printer:
os.system('rundll32 printui.dll,PrintUIEntry /y /n "PDFCreator"')

# Start AutoIT script which will answer the annoying PopUp introduced
# with the update in background (should exit together with this script
# once execution is finished):
subprocess.Popen('killJournalOpenDialogue.exe')

# Initialize PDF-Creator settings:
subprocess.call([pdfcreator, '/InitializeSettings'])
subprocess.call(['reg.exe', 'ADD', 'HKEY_CURRENT_USER\\Software\\pdfforge\\PDFCreator\\Settings\\ConversionProfiles\\0', '/f', '/v', 'OpenViewer', '/d', 'false'])
subprocess.call(['reg.exe', 'ADD', 'HKEY_CURRENT_USER\\Software\\pdfforge\\PDFCreator\\Settings\\ConversionProfiles\\0', '/f', '/v', 'SkipPrintDialog', '/d', 'true'])
subprocess.call(['reg.exe', 'ADD', 'HKEY_CURRENT_USER\\Software\\pdfforge\\PDFCreator\\Settings\\ConversionProfiles\\0\\AutoSave', '/f', '/v', 'Enabled', '/d', 'true'])
subprocess.call(['reg.exe', 'ADD', 'HKEY_CURRENT_USER\\Software\\pdfforge\\PDFCreator\\Settings\\ConversionProfiles\\0\\AutoSave', '/f', '/v', 'EnsureUniqueFilenames', '/d', 'false'])
subprocess.call(['reg.exe', 'ADD', 'HKEY_CURRENT_USER\\Software\\pdfforge\\PDFCreator\\Settings\\ConversionProfiles\\0\\AutoSave', '/f', '/v', 'TargetDirectory', '/d', '<Environment:TEMP>'])
subprocess.call(['reg.exe', 'ADD', 'HKEY_CURRENT_USER\\Software\\pdfforge\\PDFCreator\\Settings\\ConversionProfiles\\0', '/f', '/v', 'FileNameTemplate', '/d', 'temp']) # .pdf is added by pdfcreator..

# Optional PDF-Creator settings:
subprocess.call(['reg.exe', 'ADD', 'HKEY_CURRENT_USER\\Software\\pdfforge\\PDFCreator\\Settings\\ConversionProfiles\\0\\PdfSettings', '/f', '/v', 'DocumentView', '/d', 'ThumbnailImages'])
subprocess.call(['reg.exe', 'ADD', 'HKEY_CURRENT_USER\\Software\\pdfforge\\PDFCreator\\Settings\\ConversionProfiles\\0\\PdfSettings', '/f', '/v', 'PageView', '/d', 'OneColumn'])
subprocess.call(['reg.exe', 'ADD', 'HKEY_CURRENT_USER\\Software\\pdfforge\\PDFCreator\\Settings\\ConversionProfiles\\0\\JpegSettings', '/f', '/v', 'Quality', '/d', '100'])
subprocess.call(['reg.exe', 'ADD', 'HKEY_CURRENT_USER\\Software\\pdfforge\\PDFCreator\\Settings\\ConversionProfiles\\0\\JpegSettings', '/f', '/v', 'Dpi', '/d', '300'])
subprocess.call(['reg.exe', 'ADD', 'HKEY_CURRENT_USER\\Software\\pdfforge\\PDFCreator\\Settings\\ConversionProfiles\\0\\PngSettings', '/f', '/v', 'Color', '/d', 'Color32BitTransp'])
subprocess.call(['reg.exe', 'ADD', 'HKEY_CURRENT_USER\\Software\\pdfforge\\PDFCreator\\Settings\\ConversionProfiles\\0\\PngSettings', '/f', '/v', 'Dpi', '/d', '300'])

# Scan for files:
jnt_list = get_recursive_filelist(filebase=root_folder, filetype='.jnt', exclude_folders=folders_to_exclude)

# Print the .jnt files:
for jnt_file in jnt_list:
    output_path = get_output_path(jnt_file)

    if notefile_needs_update(jnt_file, output_path):
        print('Exporting {}'.format(jnt_file))
        subprocess.call([journal, '/p', jnt_file])
        temp_pdf = os.path.join(os.environ['TEMP'], 'temp.pdf')
        copy_worked = False
        while not copy_worked:
            try:
                # wait until pdf is created
                time.sleep(0.5)
                shutil.move(temp_pdf, output_path)
                copy_worked = True
            except OSError as error_info:
                print('Waiting for pdf to arrive at ' + temp_pdf)
                print(error_info)

# Kill the AutoIT Tool again:
os.system('taskkill /f /im killJournalOpenDialogue.exe')
