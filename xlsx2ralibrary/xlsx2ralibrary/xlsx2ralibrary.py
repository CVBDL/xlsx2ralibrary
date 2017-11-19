import sys
import argparse
import json
import requests
from openpyxl import load_workbook

import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

import logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger('RaLibraryImportTool')

def parse_cli_args():
    # process command line parameters
    parser = argparse.ArgumentParser(description='Import books to RaLibrary.')
    parser.add_argument('--user-name', help='RA-INT account name without ra-int perfix.')
    parser.add_argument('--password', help='RA-INT account password.')
    parser.add_argument('--path', help='Input file path.')
    args = parser.parse_args()
    # validate authentication info
    if not args.user_name or not args.password:
        logger.info('[Abort] Missing accout name or password')
        raise Exception
    # validate existence of input excel file
    if not args.path:
        logger.info('[Abort] Missing input file path.')
        raise Exception
    return args

def login(username, password):
    logger.info('Identifying...')
    endpoint = r'https://apcndaec3ycs12.ra-int.com/raauthentication/api/user'
    payload = { 'UserName': username, 'Password': password }
    req = requests.post(endpoint, data=payload, verify=False)

    if req.status_code == 200:
        logger.info('Identify successfully.')
        return req.json()['IdToken']
    elif req.status_code == 401:
        raise Exception('Unauthorized.')
    else:
        raise Exception('Identify failed.')

def query_book(isbn):
    """Query books details via books open API."""
    if not isbn or not isinstance(isbn, str):
        raise Exception

    endpoint = r'https://apcndaec3ycs12.ra-int.com/ralibrary/api/book/isbn/'
    query_endpoint = '{0}{1}'.format(endpoint, isbn)
    req = requests.get(query_endpoint, verify=False)

    if req.status_code == 200:
        return req.json()
    else:
        raise Exception

def read_excel(file_path):
    """Process Excel file."""
    wb = load_workbook(filename=file_path, read_only=True)
    ws = wb.active
    rows = iter(ws.rows)

    # skip header row
    # next(rows)

    for row in rows:
        isbn = str(row[0].value)
        code = row[1].value
        title = row[2].value

        if not isbn or not code or not title:
            pass
        else:
            try:
                book = query_book(isbn)
                book['Code'] = code
                save_book(book)
            except:
                if len(isbn) == 10:
                    isbn_10 = isbn
                    isbn_13 = ''
                else:
                    isbn_10 = ''
                    isbn_13 = isbn
                save_book({
                    'ISBN10': isbn_10,
                    'ISBN13': isbn_13,
                    'Code': code,
                    'Title': title})

def save_book(book):
    """Save book to RaLibrary."""
    endpoint = r'https://apcndaec3ycs12.ra-int.com/ralibrary/api/books'
    headers = { 'Authorization': 'Bearer ' + id_token }
    req = requests.post(endpoint, headers=headers, data=book, verify=False)

    if req.status_code == 200:
        logger.info('Added {0}'.format(req.json()['Title']))
    else:
        logger.info('Failed to add {0}'.format(book['Title']))

def main():
    # Parse command line arguments.
    try:
        args = parse_cli_args()
    except:
        return 1

    # Authentication.
    try:
        id_token = login(args.user_name, args.password)
    except Exception as e:
        logger.info('[Abort] {0}'.format(e))
        return 1

    read_excel(args.path)

    return 0
