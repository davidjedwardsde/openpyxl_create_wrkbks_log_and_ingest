import pandas as pd
import openpyxl
import os


def check_files(path):

    headers = {'Account List': ['Management Account ID', 'Management Name', 'Management Company ALN ID',
                                'Management Owner Name', 'Sales Region', 'Manager Name', 'Validated?'],
               'Property List': ['Management Account ID', 'Management Name', 'Management Company ALN ID',
                                 'Management Owner Name', 'Sales Region', 'Manager Name', 'Property Account ID',
                                 'Property Name', 'Property ALN ID', 'Current Property Units', 'Current Property Type',
                                 'New Property Units', 'New Property Type'],
               'COM Properties': ['Property Account ID', 'Property Name', 'New Management Company ID',
                                  'New Management Company Name'],
               'Property Add': ['Property Name', 'Property Type', 'Property Units', 'Shipping Street Address',
                                'Shipping City', 'Shipping State', 'Shipping Zip', 'Billing Street Address', 'Billing City',
                                'Billing State', 'Billing Zip', 'Management Account ID',	'Management Account Name',
                                'Property ALN ID'],
               'Duplicate Properties': ['Property Account ID In Book', 'Property Name In Book',
                                        'Duplicate Property Account ID', 'Duplicate Property Name']
               }
    # initialize an error log as a blank list
    error_log = []

    # iterate through the first row of each file in path to check headers in target sheet
    for file in os.listdir(path):
        wrkbk = openpyxl.load_workbook(f'{path}/{file}')
        print(f'Processing {file}')

        for sheet in wrkbk.worksheets:
            if sheet.title not in headers:
                # if the workbook has an extra sheet, record the name of the file and what the sheet
                # is called in the error log
                error_log.append(f'{file} has extra sheet: {sheet}')
                continue
            cols = len(headers[sheet.title])
            for rows in sheet.iter_rows(min_row=1, max_row=1, min_col=1, max_col=cols, values_only=True):
                wrksht_headers = list(rows)
            # when a sheet has an extra column, log it as an issue in the error log
            if wrksht_headers != headers[sheet.title]:
                error_log.append(f'Issue on {file}, {sheet}')
        print(f'File {file} complete')

    # create a dataframe for the error log and export it as a csv
    log = pd.DataFrame(error_log)
    log.to_csv('Error_log.csv')
