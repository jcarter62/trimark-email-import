import argparse
import os
import shutil
from datetime import datetime
from excel_reader import ExcelReader
from data_writer import DataWriter

"""
main.py - Entry point for the Trimark Email Import application.

Usage:
    python main.py -i input.xlsx -a process -d destination_folder

Arguments:
    -i, --input: The input Excel file to read.
    -f, --folder: The folder containing the Excel files to read.
    -a, --action: The action to perform (e.g., 'process').
    -d, --destination: The destination folder for processed files.

Description:
    This script serves as the entry point for the application.
    It uses argparse to handle command-line arguments and calls
    the appropriate functions for reading Excel files and writing
    to a SQL Server database.
"""


dataobj = {
    "file": "",
    "source": "",
    "destination": "",
    "data": []
}


def process_file(filename, destination_folder):
    dataobj["file"] = filename
    dataobj["destination"] = destination_folder
    rd = ExcelReader(dataobj["file"])
    dataobj["data"] = rd.records
    #
    dw = DataWriter(dataobj)
    dw.execute()
    move_input_file(dataobj["file"], dataobj["destination"])


def view_file(filename):
    dataobj["file"] = filename
    rd = ExcelReader(dataobj["file"])
    dataobj["data"] = rd.records
    print(dataobj["data"])


def extract_file_name(file_path) -> str:
    return os.path.basename(file_path)


def calculate_destination_file(file_path, destination_folder) -> str:
    file_name = extract_file_name(file_path)
    # prepend YYYYMMDD-HHMM to the new_name
    new_name = f"{datetime.now().strftime('%Y%m%d-%H%M')}-{file_name}"
    # calculate the new file name
    new_file_name = os.path.join(destination_folder, new_name)
    return new_file_name


def move_input_file(file_path, destination_folder) -> bool:
    new_file_name = calculate_destination_file(file_path, destination_folder)
    shutil.move(file_path, new_file_name)
    # determine if move succeeded
    if os.path.exists(new_file_name):
        return True
    else:
        return False


def main(exec_file_name, version):
    parser = argparse.ArgumentParser(description="Trimark File Import Application - " + exec_file_name + " - v" + version)
    parser.add_argument("-i", "--input", required=True, help="Input filename or Folder containing .xlsx file(s)")
    parser.add_argument("-a", "--action", required=True, choices=["process", "view"], help="Action to perform: process or view")
    parser.add_argument("-d", "--destination", required=False, help="Destination folder (required for 'process' action)")
    parser.add_help = True

    args = parser.parse_args()

    # check to see if x.env file exists
    if not os.path.exists(".env"):
        print("Error: .env file not found.")
        return

    # check if args.input is a folder
    isFolder = os.path.isdir(args.input)

    if isFolder:
        # need to inspect folder to see if it contains .xlsx files
        folder = args.input
        if os.path.isdir(folder):
            # get list of files in the folder
            files = os.listdir(folder)
            #
            # sort files by name
            files.sort()
            # check if there are .xlsx files
            excel_files = [f for f in files if f.endswith(".xlsx")]
            if len(excel_files) == 0:
                return
            else:
                # process each file
                for file in excel_files:
                    file_path = os.path.join(folder, file)
                    if args.action == "process":
                        process_file(file_path, args.destination)
                    elif args.action == "view":
                        view_file(file_path)
    else:
        # we have one file to process or view
        if args.action == "process":
            process_file(args.input, args.destination)
        elif args.action == "view":
            view_file(args.input)


if __name__ == "__main__":
    exe_file_name = os.path.basename(__file__)
    version = "2023.10.13.0"
    main(exe_file_name, version)
