"""
This script accepts a directory as a parameter and will recursively search for
and convert .one (OneNote) files to .pdf files.
A variable can be set to exclude specific directory names from search.
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
import sys
from sharedExporterFunctions import get_output_path, notefile_needs_update, get_recursive_filelist

# Folders to exclude from OneNote file scan:
folders_to_exclude = ['.dropbox.cache']

# Find root folder:
root_folder = './'
if len(sys.argv) > 1:
    root_folder = sys.argv[1]

# Scan current folder recursively for .one and .jnt files:
one_list = get_recursive_filelist(filebase=root_folder, exclude_folders=folders_to_exclude)

# Print the .one files:
for one_file in one_list:
    output_path = get_output_path(one_file)

    if notefile_needs_update(one_file, output_path):
        print('Exporting {}'.format(one_file))
        # Remove old .pdf file first. AutoIT script assumes it to be removed:
        if os.path.exists(output_path):
            os.remove(output_path)
        subprocess.call(['exportPDF.exe', one_file, output_path])

# Close OneNote if it is opened:
subprocess.call(['closeOneNote.exe'])
