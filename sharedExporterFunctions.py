"""
This script provides functions used together by jntExporter and oneExporter.
"""

__author__ = ["Dietmar Malli"]
__copyright__ = "Copyright 2017, Dietmar Malli"
__credits__ = []
__license__ = "GPLv3"
__version__ = "1.0.1"
__maintainer__ = ["Dietmar Malli"]
__email__ = ["git.commits@malli.co.at"]
__status__ = "Production"

import os

def get_output_path(filename):
    basename, _ = os.path.splitext(filename)
    return os.path.abspath(basename + '.pdf')

def newer_as(note_file, pdf_file):
    if not os.path.exists(note_file):
        raise Exception('At least file1 must exist.')

    note_file_time = os.path.getmtime(note_file)
    pdf_file_time = os.path.getmtime(note_file)  # init var
    if os.path.exists(pdf_file):
        pdf_file_time = os.path.getmtime(pdf_file)

    if pdf_file_time < note_file_time:
        return True
    else:
        return False

def get_recursive_filelist(filebase='./', filetype='.one', exclude_folders=[]):
    list = []
    for root, dirs, files in os.walk(filebase, topdown=True):
        for file in files:
            dirs[:] = [d for d in dirs if d not in exclude_folders]
            if file.endswith(filetype):
                list.append(os.path.abspath(os.path.join(root, file)))
    return list
