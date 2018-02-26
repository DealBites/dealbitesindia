from __future__ import print_function

import sys, os, time
import ast
import email
import imaplib
import os
import sys
import re
from datetime import date,datetime,timedelta
from dateutil.parser import parse
import csv

prop_path = './credentials.property'
fh = open(prop_path)
dictionary = ast.literal_eval(fh.readline())
subject_to_parse='Cuelinks Deals and Campaign Updates'
final = 'In case you have any queries, you can reach out to us at sales@cuelinks.com'

def parse_gmail(b):
    if b.is_multipart():
        for payload in b.get_payload():
            return payload.get_payload()
    else:
        return b.get_payload()


def isTitle(line):
    length = len(line)
    return line[0] == '*' and line[length-1] == '*'

# soup_text: String // Content of mail
# dest: String // destination file
def parse_soup_text(soup_text, dest):
    # open file
    try:
        destFile = open(dest, 'w+')
    except:
        print('An error opening file occurred')
    # rows
    fields = ['Store', 'Offer', 'Coupon Code', 'Landing Page', 'Affiliate URL', 'Valid till', 'T&C']
    # variable to return
    ret = ""
    lines = soup_text.splitlines()
    assert(len(lines) > 0)
    for i in range(len(lines)):
        line = lines[i]
        length = len(line)
        if (length == 0):
            continue
        # if it is a title
        if (isTitle(line)):
            title = line
            # ignore empty lines
            while (i+1 < len(lines) and len(lines[i+1]) == 0):
                i += 1
            # if it is followed by an offer, aka, it is a company name
            if (i+1 < len(lines) and (lines[i+1].split(' :')[0] == 'Offer')):
                # add company to ret
                writer = csv.DictWriter(destFile, fieldnames=fields)
                row = {x: '' for x in fields}
                # store
                currentField = 'Offer'
                value = ''
                for j in range(i+1, len(lines)):
                    if (len(lines[j]) == 0):
                        continue

                    # if all offers have been read
                    if (lines[j] == final):
                        return

                    field = lines[j].split(' :')[0]

                    if (field in fields or isTitle(lines[j])):
                        # store on row
                        row[currentField] = value

                        if (isTitle(lines[j])):
                            break

                        # new offer found
                        if (field == 'Offer'):
                            if row != {x: '' for x in fields}:
                                writer.writerow(row)
                            row = {x: '' for x in fields}
                            row[fields[0]] = title[1:-1]

                        # next field
                        currentField = field
                        value = ''

                    line = ''
                    if (len(lines[j].split(' :')) == 2):
                        line = lines[j].split(' :')[1]
                    else:
                        line = lines[j]

                    value += line

                writer.writerow(row)

    destFile.close()

def get_imap_session():
    try:
        imap_session = imaplib.IMAP4_SSL('imap.gmail.com')
        typ, account_details = imap_session.login(dictionary['username'], dictionary['password'])
        # print(typ)
        # print(account_details)
        # print(imap_session)
        if typ != 'OK':
            print('Not able to sign in!')
            # TODO send email if the login attempt fails
            logger.exception("Fatal error in sign-in")
        # else:
        #    print('OK')

        #print('TYPE',typ)
        imap_session.select('INBOX')
        print('Subject to parse: ', subject_to_parse)
        if len(subject_to_parse) > 0:
            header_subject = '(HEADER Subject "' + str(subject_to_parse) + '")'
            typ, data = imap_session.search(None, '(UNSEEN)', header_subject)
            #print(typ)
        else:
            typ, data = imap_session.search(None, '(UNSEEN)')

        if typ != 'OK':
            print('Error searching Inbox.')
            # TODO send email if the error or use logger
            raise

        #print(data[0].split())
        for msgId in data[0].split():
            typ, message_parts = imap_session.fetch(msgId, '(RFC822)')
            if typ != 'OK':
                print('Error fetching mail.')
                raise

            msg = email.message_from_string(message_parts[0][1])
            body = parse_gmail(msg)
            
            # pip install bs4 --user
            # pip install lxml --user
            from bs4 import BeautifulSoup
            soup = BeautifulSoup(body, 'lxml')
            soup_text=soup.get_text()
            #print(soup_text)

            parse_soup_text(soup_text, 'mailtext.csv')


            # Remove all newline characters
            #soupgood=re.sub(r'In case you have any queries*$','',re.sub(r'[\n\r]+', '', soup.get_text()))

    
    except imaplib.IMAP4.error as e:
        # TODO send email if download attachment fails
        print(e.message)
        raise imaplib.IMAP4.error
    except Exception as e:
        print(e.message)
        raise Exception
    

if __name__ == '__main__':
    get_imap_session()
