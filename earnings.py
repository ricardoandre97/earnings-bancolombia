import xmltodict
import json
import os
from re import match


def consolidated_earnings(folder, files):

    """
    Reads the total earnings from Bancolombia xls files for a given year.
    Go to https://sucursalpersonas.transaccionesbancolombia.com and
    generate the extracts that you need.
    Then, just download the files, put them in the same directory
    and put some easy names.
    Finally, just provide the full path to the directory where the files
    are stored, and the name of the files.
    """

    # Get the files with full path
    full_files = [os.path.join(folder, i) for i in files]

    # Define the regex to match in the file
    result, regex = 0, '^\d{1,2}\/\d{1,2}$'

    for file in full_files:

        # Loads the file in memory
        with open(file) as fp:

            # From xml to json
            f = json.loads(json.dumps(xmltodict.parse(fp.read())))

            # Loop over the records and find the matches
            for row in f['ss:Workbook']['ss:Worksheet']['ss:Table']['ss:Row']:
                cell = row['ss:Cell']
                if isinstance(cell, list):
                    data = cell[0]['ss:Data']

                    # If regex & conditions match, append to the result
                    if '#text' in data:
                        if match(regex, data['#text']):
                            n = cell[-2]['ss:Data']['#text'].replace(',', '')
                            try:
                                n = int(n)
                            except ValueError:
                                n = float(n)
                            if n > 0:
                                result += n
    return result


if __name__ == '__main__':

    folder = '/home/ricardo/Downloads'
    files = ['marzo.xls', 'junio.xls', 'septiembre.xls', 'diciembre.xls']
    print(f'Total earnings: $ {consolidated_earnings(folder, files):,.2f} COP')
