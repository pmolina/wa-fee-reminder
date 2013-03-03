# -*- coding: utf-8 -*-

import gmail
import gspread
import ConfigParser
import logging
import codecs


class WAException(Exception):
    pass


class FormatMessage(object):
    message = None

    def __init__(self, *args, **kwargs):
        if not self.__class__.message:
            f = codecs.open('template.html', encoding='utf-8')
            self.__class__.message = f.read()
            f.close()

    def get_message(self, d):
        return self.__class__.message % d


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
    spreadsheet_name = config.get('General', 'spreadsheet_name')
    worksheet_name = config.get('General', 'worksheet_name')
    min_debt = int(config.get('General', 'min_debt'))
    max_debt = int(config.get('General', 'max_debt'))
    logging.debug(
        'Trying to use worksheet "%s" from spreadsheet "%s" ...'
        % (worksheet_name, spreadsheet_name)
    )
    # Using the same username and password from GMail
    gs = gspread.login(gmail_username, gmail_password)
    spreadsheet = gs.open(spreadsheet_name)
    worksheet = spreadsheet.worksheet(worksheet_name)
    # Here starts a customized behaviour. [1:] avoids header.
    names = worksheet.col_values(1)[1:]
    mails = worksheet.col_values(3)[1:]
    amounts = worksheet.col_values(8)[1:]
    debts = worksheet.col_values(9)[1:]
    to_send = []
    for name, mail, amount, debt in zip(names, mails, amounts, debts):
        if amount and amount.isdigit() and int(amount) >= min_debt and int(amount) <= max_debt:
            to_send.append({
                'nombre': name,
                'mail': mail,
                'cuotas': amount,
                'deuda': debt,
            })
    if to_send:
        fm = FormatMessage()
        logging.debug(
            'Logging into GMail using username "%s" and password "%s" ...'
            % (gmail_username, '*'*len(gmail_password))
        )
        gm = gmail.GMail(username=gmail_username, password=gmail_password)
        gm.connect()
        for d in to_send:
            to = d['mail']
            logging.debug('Sending message to %s ...' % to)
            msg = gmail.Message(
                'Pago de cuotas de Wikimedia Argentina',
                to=to,
                html=fm.get_message(d)
            )
            gm.send(msg)


if __name__ == '__main__':
    try:
        main()
    except Exception, e:
        raise WAException(e)
