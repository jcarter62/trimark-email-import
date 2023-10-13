import argparse
import os
import shutil
from datetime import datetime

from excel_reader import ExcelReader
from data_writer import DataWriter


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


def main():
    parser = argparse.ArgumentParser(description="File Operations")
    parser.add_argument("-i", "--input", required=False, help="Input filename")
    parser.add_argument("-f", "--folder", required=False, help="Input Folder containing .xlsx file(s)")
    parser.add_argument("-a", "--action", required=True, choices=["process", "view"], help="Action to perform: process or view")
    parser.add_argument("-d", "--destination", required=False, help="Destination folder (required for 'process' action)")

    args = parser.parse_args()

    print(args)

    if args.input is None and args.folder is None:
        print("Error: Either input file or input folder is required.")
        return

    if args.folder is not None:
        # need to inspect folder to see if it contains .xlsx files
        folder = args.folder
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
        # we have one file to process
        if args.action == "process":
            process_file(args.input, args.destination)
        elif args.action == "view":
            view_file(args.input)


if __name__ == "__main__":
    main()
