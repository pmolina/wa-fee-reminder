# -*- coding: utf-8 -*-

import gmail
import gspread
import ConfigParser
import logging


def main():
    logging.basicConfig(
        format='%(levelname)s - %(asctime)s: %(message)s',
        datefmt='%m/%d/%Y %I:%M:%S %p',
        level=logging.DEBUG
    )
    logging.debug('Reading configuration file ...')
    config = ConfigParser.RawConfigParser()
    config.read('config.cfg')
    gmail_username = config.get('Gmail', 'username')
    gmail_password = config.get('Gmail', 'password')
    # if gmail_username and gmail_password:
    #     logging.debug(
    #         'Using username "%s" and password "%s" ...'
    #         % (gmail_username, '*'*len(gmail_password))
    #     )
    #     g = gmail.GMail(username=gmail_username, password=gmail_password)
    #     g.connect()
    spreadsheet_name = config.get('General', 'spreadsheet_name')
    worksheet_name = config.get('General', 'worksheet_name')
    logging.debug(
        'Trying to use worksheet "%s" from spreadsheet "%s" ...'
        % (worksheet_name, spreadsheet_name)
    )
    # Using the same username and password from GMail
    gc = gspread.login(gmail_username, gmail_password)
    spreadsheet = gc.open('Nueva planilla de socios de Wikimedia Argentina')
    worksheet = spreadsheet.worksheet(worksheet_name)
    for col_value in worksheet.col_values(1):
        print col_value.encode('utf-8')

if __name__ == '__main__':
    main()
