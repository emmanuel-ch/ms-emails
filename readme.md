# Outlook Mini-toolkit

Outlook Mini-toolkit is a Python package that can run some elementary admin task on Outlook mailboxes (Windows).

## Installation

Installation not needed. Just the right Python configuration.  
It is also possible to make it an .EXE file thanks to pyinstaller.

## Usage

```bash
python outlook-minitoolkit.py
```

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## Requirements

Developped for Python 3.10.
Module win32com is needed (package name pywin32)

You may use the below command to create conda environment with required packages:
```bash
conda env create -f environment.yml
```