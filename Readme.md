# Trimark Readings File Import

## Overview

This Python project is designed to read Excel files sent from Trimark Service 
and write the data to a SQL Server database. It is particularly useful 
for automating the process of importing data from Excel spreadsheets into a database.

## Requirements

- Python 3.x
- openpyxl
- python_decouple
- pyodbc

To install the required packages, run:

```bash
pip install -r requirements.txt
```

## Configuration

Create a `.env` file based on the `sample.env` and fill 
in the required database credentials.  This application uses 
Windows Trusted Authentication to connect to the database, so 
the logged in user running this application must have access
to the database.

``` .env
DB_SERVER=SQL-SERVER
DB_INSTANCE=InstanceName
DB_NAME=DatabaseName
```

## Usage

### Main Script

The `main.py` script is the entry point of the application. It uses `argparse` to handle command-line arguments.

``` bash
python main.py -i input.xlsx|input-folder -a process|view -d destination_folder
```

The `-i` argument is used to specify a single Excel file or a folder containing files to process. 
The `-a` argument is used to specify the action to perform. 
The `-d` argument is used to specify the destination folder for the processed files.  

Required arguments:
* `-i`
* `-a`
* `-d` if `-a` is `process`

### Excel Reader

The `excel_reader.py` class ExcelReader reads Excel files and stores the data in a list of dictionaries.

### Data Writer

The `data_writer.py` class DataWriter writes the data to a SQL Server database.

