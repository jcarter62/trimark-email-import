# Trimark Email Import

## Overview

This Python project is designed to read Excel files and write the data to a SQL Server database. It is particularly useful for automating the process of importing data from Excel spreadsheets into a database.

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

Create a `.env` file based on the `sample.env` and fill in the required database credentials.

```env
DB_SERVER=SQL-SERVER
DB_INSTANCE=InstanceName
DB_NAME=DatabaseName
```

## Usage

### Main Script

The `main.py` script is the entry point of the application. It uses `argparse` to handle command-line arguments.

``` bash
python main.py -i input.xlsx -f input-folder -a process -d destination_folder
```

The `-i` argument is used to specify a single Excel file to process. The `-f` argument is used to specify a folder containing Excel files to process. The `-a` argument is used to specify the action to perform. The `-d` argument is used to specify the destination folder for the processed files.  

Required arguments:
* `-i` or `-f`
* `-a`
* `-d` if `-a` is `process`

### Excel Reader

The `excel_reader.py` script reads Excel files and stores the data in a list of dictionaries.

### Data Writer

The `data_writer.py` script writes the data to a SQL Server database.

## Contributing

Feel free to fork the project and submit a pull request.
