#!/usr/bin/env python
#
# Very simple Python script to dump all emails in ZJU server to local files.  
# This code is released into the public domain.
#
# RKI Nov 2013
# By Nan Xu 2021

import os
import re
import sys
import eml_parser
import imaplib
import getpass

IMAP_SERVER = 'imap.zju.edu.cn'
EMAIL_ACCOUNT = input("Input your email adddress:").strip()

PASSWORD = getpass.getpass()

def process_mailbox(M,dirs):
    """
    Dump all emails in the folder to files in output directory.
    """

    rv, data = M.search(None, "ALL")
    if rv != 'OK':
        print("No messages found!")
        return

    for num in data[0].decode("utf-8").split():
        rv, data = M.fetch(num, '(RFC822)')
        if rv != 'OK':
            print("ERROR getting message", num)
            return

        sys.stdout.write('\r' + "Writing message %s" %num)
        sys.stdout.flush()
        contents = eml_parser.eml_parser.decode_email_b(data[0][1])
        subject = contents['header']['subject']
        tobe_substituted=subject if subject else "(无主题)"
        normalized_name = re.sub('[*"/:?|<>\n]', '', tobe_substituted, 0).strip("\0")
        dates_ = str(contents['header']['date']).split("-")
        dirpath = dirs+"/"+dates_[0]+"年"+dates_[1]+"月"
        file_path = dirpath+"/"+normalized_name
        
        if not os.path.exists(dirpath):
            os.makedirs(dirpath)
        
        if not os.path.exists(file_path):
            pass
        elif "无主题" in file_path:
            tmp = file_path
            count = 1
            while os.path.exists(tmp):
                tmp = file_path+"%d" %count
                count +=1
            file_path=tmp
        else:
            continue       
        
        f = open('%s.eml' %(file_path), 'wb')
        f.write(data[0][1])
        f.close()

def back_up_inbox():
    M = imaplib.IMAP4_SSL(IMAP_SERVER)
    M.login(EMAIL_ACCOUNT, PASSWORD)
    EMAIL_FOLDER="INBOX"# "Sent Items"
    rv, data = M.select('"%s"' %EMAIL_FOLDER)
    if rv == 'OK':
        print("Backing up INBOX: ", EMAIL_FOLDER)
        process_mailbox(M,"INBOX")
        M.close()
        print()
    else:
        print("ERROR: Unable to open mailbox ", rv)
    M.logout()

def back_up_sent():
    M = imaplib.IMAP4_SSL(IMAP_SERVER)
    M.login(EMAIL_ACCOUNT, PASSWORD)
    EMAIL_FOLDER="Sent Items"
    rv, data = M.select('"%s"' %EMAIL_FOLDER)
    if rv == 'OK':
        print("Backing up Sent ITEMS: ", EMAIL_FOLDER)
        process_mailbox(M,"SENT")
        M.close()
        print()
        print("Finished.")
    else:
        print("ERROR: Unable to open mailbox ", rv)
    M.logout()

if __name__ == "__main__":
    back_up_inbox()
    back_up_sent()
    