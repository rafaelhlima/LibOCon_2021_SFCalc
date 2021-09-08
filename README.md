# LibreOffice Conference 2021

This repository contains the files used in the talk "Python scripts in LibreOffice Calc using the ScriptForge library" given by Rafael Lima during the LibreOffice Conference 2021.

To run the examples in the Python file you need to use LibreOffice >= 7.2.

## Running macros from LibOCon_2021.py

The macros were creating as user scripts. To be able to run them, the file LibOCon_2021.py needs to be placed into the LibreOffice user scripts folder in your machine.

On **Linux** machines, the folder is:

`/home/user/.config/4/user/Scripts/python/` (if it does not exist, create it)

On **Windows** machines, the folder is:

`%APPDATA%\LibreOffice\4\user\Scripts\python`

Read [this help page](https://help.libreoffice.org/latest/en-US/text/sbasic/python/python_locations.html) to learn more about where Python scripts are located.
