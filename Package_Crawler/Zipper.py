#!/usr/bin/python3
"""
@author Ben Karabinus

Background: The package crawler project is meant to search Dynamics 365 Finance and Operations data packages for
legal entity dependence and print the results of the search to an Excel workbook. The project is far rom from finished.
As of yet there is no exception handling or unit tests built in to the project. For now package crawler is a simple
Python project the gets a monotonous task done quickly. If I've shared the GitHub repo with you feel free to contribute.
Any and all help/ideas are welcome.

"""

from zipfile import ZipFile

import pandas as pd

import xlsxwriter

import os

# Class Zipper.py is designed with main method to run as stand alone module for now
def main():
    filePath = input("Please enter working directory of the zip file you would like to process: ")
    workbookName = input("Please input the workbook name: ")
    directoryContents = get_path(filePath)
    dataFrames = create_dataframes(directoryContents)
    package_dict = test_dependence(dataFrames)
    write_data_to_workbook(package_dict, workbookName)


def get_path(filePath):

    #convert to appropriate path type for client os
    filePath = os.path.realpath(filePath)

    # check if path exists if yes change to specified directory and return contents as list
    if os.path.isdir(filePath):
        os.chdir(filePath)
        directoryContents = os.listdir(filePath)

    return directoryContents


def create_workbook(workbookName, sheetNames):

    # create a new excel workbook
    wb = xlsxwriter.Workbook(workbookName + '.xlsx')
    sheet = wb.add_worksheet('Entity List')
    sheet.write(0, 0, 'Package Name')
    sheet.write(0, 1, 'Legal Entity Dependent')

    for sheetName in sheetNames:
        # Excel sheet names cannot be longer than 31 characters
        if len(sheetName) > 31:
            sheetName = sheetName[:31]
            sheet = wb.add_worksheet(sheetName)
            sheet.write(0, 0, sheetName)
            sheet.write(0, 1, 'Data entities')
            sheet.write(0, 2, 'Legal entity dependent')
        else:
            sheet = wb.add_worksheet(sheetName)
            sheet.write(0, 0, sheetName)
            sheet.write(0, 1, 'Data entities')
            sheet.write(0, 2, 'Legal entity dependent')
    return wb


def write_data_to_workbook(package_dict, workbookName):

    sheetNames = package_dict.keys()
    wb = create_workbook(workbookName, sheetNames)
    # init count for sheet index control variable
    count = 1
    for key, value in package_dict.items():
        sheet = wb.get_worksheet_by_name(str(key))
        sheet.write_column(1,1,value.keys())
        sheet.write_column(1,2, value.values())
        if 'Yes' in value.values():
            sheet = wb.get_worksheet_by_name('Entity List')
            sheet.write(count, 0, str(key))
            sheet.write(count, 1, 'Yes')
        else:
            sheet = wb.get_worksheet_by_name('Entity List')
            sheet.write(count, 0, str(key))
            sheet.write(count, 1, 'No')
        count += 1

    wb.close()


def create_dataframes(directoryContents):
    """
    init dicitonary to store package contents as dictionaries of pandas dataframe objects. This is resource intensive
    but allows for versatility when manipulating data.
    """
    package_dict = {}

    for file in directoryContents:
        with ZipFile(file, 'r')as zip:
            contents = zip.namelist()
            print('The package contains the following entities:')
            df_dict = {}

            for entity in contents:
                if entity.endswith('.xlsx'):
                    print(entity)
                    # open the excel file for the entity and read to dataframe subtracting .xlsx file extension
                    entityName = entity[0:len(entity)-5]
                    entity = zip.open(entity)
                    entity = pd.read_excel(entity)
                    df_dict[entityName] = entity
            package_dict[file] = df_dict

    return package_dict

def test_dependence(package_dict):

    # sheetnames must be 31 characters or less so store truncated package names for later
    sheetNames = []

    for packageName, entity in package_dict.items():
        # remove the .zip file extension for sheetname then truncate if necessary
        sheetName = packageName[0:len(packageName)-4]
        if len(sheetName) > 31:
            sheetName = sheetName[0:31]
        sheetNames.append(sheetName)
        dependencies_dict = {}
        df_dict = entity
        testVariable = 'TMPL'

        for key, value in df_dict.items():
            dependencies_dict[key] = 'N0'
            if testVariable.casefold() in value.values or testVariable in value.values:
                dependencies_dict[key] = 'Yes'
        package_dict[packageName] = dependencies_dict

    # zip truncated sheet names
    package_dict = dict(zip(sheetNames, list(package_dict.values())))

    return package_dict

if __name__ == '__main__':
    main()