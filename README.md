# Project Name

Simple Python script that extracts latin species names from cyrilic texts in docx files.

## Features

- Processes Word (.docx) files from the `docx_files folder`.
- Extracts Latin species names and saves them to a file in the `output` folder.

## Run script

If the script is run using the command `python main.py`, the resulted file will be named `species_list` and saved to `output` folder. You can change the output filename using the `-f` or `--filename` argument, for example:  `python main.py -f my-species-list`. Note: the filename should not include a file extension.

## Requirements

- Python 3.x
- Required packages:
  - python-docx==1.1.2


## Installation

Clone the repo:
   git clone https://github.com/Egoart/parse-latin-species-names.git