# NotesExporter project
Some scripts to recursively export .jnt and or .one files to .pdf in the same directory. They will also keep the .pdfs up to date with the notefiles by comparing the modification times. PDF export will
be triggered if the notefile is newer.

## Dependencies:
The Python 3 scripts used to export depend on some AutoIT scripts. Download a release archive from the github.com project page to get them as precompiled executables or compile them yourself but be sure to place them next to the .py scripts with the same name as their source files.

For .jnt export you additionally need:
* PDF Creator (Tested with version 2.4.1)
* Windows Journal (The new english only, annoying startup-popup version which was made available after the update which originally removed the software.)

## Installation:
* Install dependencies first.
* Place the scripts in the root of the directory structure you want to export or place them anywhere else. If you place them somewhere else you need to supply the root of your directory structure as first parameter to the scripts.

## Usage:
Just double-click them if they are at the right place or:
* ```python jntExporter.py C:\Dropbox\```
* ```python oneExporter.py C:\Dropbox\```

## Example use cases:
* Microsoft discontinued Windows Journal and you want to save your notes from the filetype-graveyard.
* You want to view the notefiles using Linux or some other device where OneNote isn't available. Just sync the drives and use these scripts regularly on the notefile creating device.
