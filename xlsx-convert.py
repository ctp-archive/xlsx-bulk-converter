import argparse
import pathlib
from collections import deque
import openpyxl
import pandas

parser = argparse.ArgumentParser()
parser.add_argument("directory", help="the directory to be crawled through")

args = parser.parse_args()

root = pathlib.Path(args.directory)
folders = deque()
files = deque()
folders.append(root)

def dir_crawl(folders):
    if len(folders) > 0:
        current_dir = folders.popleft()
        for child in current_dir.iterdir():
            if child.is_dir():
                folders.append(child)
            else:
                if ".xlsx" in str(child):
                    files.append(child)
        dir_crawl(folders)

def sheet_names(file):
    print(file)
    workbook = openpyxl.load_workbook(filename=str(file), read_only = True)
    return workbook.sheetnames

def workbook_mkdir(file):
    parent = file.parent
    filename = file.stem
    new_dir = parent.joinpath(filename)
    new_dir.mkdir()

def convert(files):
    if len(files) > 0:
        file = files.popleft()
        workbook_mkdir(file)
        sheets = sheet_names(file)
        for sheet in sheets:
            xlsx = pandas.read_excel(file, sheet_name = sheet)
            new_file = str(file.parent.joinpath(file.stem).joinpath(sheet)) + ".csv"
            xlsx.to_csv(new_file, index = False)
        convert(files)

dir_crawl(folders)
convert(files)