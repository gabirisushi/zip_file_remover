# -*- coding: utf-8 -*-
"""
Created on Wed Jan 29 17:41:50 2020

@author: MOG1CA
"""
import sys
import shutil
import zipfile
import pandas as pd


def parse_the_file(in_file, shift = 5):

    print("Parsing the file to find the files to remove : \t", end="", flush=True)
    # read the excel file
    df = pd.read_excel(in_file)

    # remove the headers
    df = df.shift(-shift)

    # extract the features columns
    df_new = pd.DataFrame()
    df_new['file'] = df.iloc[:, 0]
    df_new['action'] = df.iloc[:, 6]

    # drop the unwanted lines
    df_new.dropna(inplace=True)
    print("[ DONE ]")

    files_to_remove = list(df_new['file'])
    return files_to_remove


def removes_files_from_zip(files_to_remove, src_zip, dest_zip):
    zin = zipfile.ZipFile(src_zip, 'r')
    zout = zipfile.ZipFile(dest_zip, 'w')

    print("Remaining files : ")
    count = 0

    for item in zin.infolist():
        buffer = zin.read(item.filename)
        if not item.filename in files_to_remove:
            zout.writestr(item, buffer)
        else:
            count = count + 1
            print("{} out of {}".format(count, len(files_to_remove)))
    zout.close()
    zin.close()


if __name__ == "__main__":

    # application starts here
    input_file = sys.argv[1]
    zip_input_file = sys.argv[2]
    zip_output_file = "new.zip"

    shift_up = 5

    files_to_remove = parse_the_file(input_file, shift_up)

    removes_files_from_zip(files_to_remove, zip_input_file, zip_output_file)
