'''Program to download an SSRS report as a CSV from an intranet report server, then save to an FTP folder
https://www.reddit.com/r/Python/comments/67xak4/enterprise_intranet_authentication_and_ssrs/ kudos to u/indepndnt for help with retrieving credentials and ntlm auth
note:config.ini file stored separately
'''

import configparser
import os
import glob
from datetime import datetime, timedelta
from ftplib import FTP

import win32cred
from requests import Session
from requests_ntlm import HttpNtlmAuth

# global variables read from config file
config = configparser.ConfigParser()
config.read('config.ini')
base_url = config['REPORT']['url']
report_path = config['REPORT']['report_path']
host = config['FTP']['host']
user = config['FTP']['user']
ftp_password = config['FTP']['ftp_password']
ftp_path = config['FTP']['ftp_path']
# retrieves the current date in 'yyyy-mm-dd' format
current_date = (datetime.now()).strftime('%Y-%m-%d')


def csv_download():

    # retrieves the user domain, username and password from windows credential manager
    username = "{}\\{}".format(os.environ.get(
        'USERDOMAIN'), os.environ.get('USERNAME'))
    credential = win32cred.CredRead(
        'TERMSRV/PDCPTMSDB', win32cred.CRED_TYPE_GENERIC, 0)
    password = credential.get('CredentialBlob').decode('utf-16')
    # using repr(object) we can see there is a null character ('\x00') at the end of password. use str.replace to remove
    password = password.replace('\x00', '')

    # option below to add parameters to report URL if required, make sure to update the 'response' variable if doing so
    # url = base_url + '?/Report&Date=' + current_date + '&rs:Format=csv'

    # open a session using Ntlm authorisation and save to specified path
    session = Session()
    session.auth = HttpNtlmAuth(username=username, password=password)
    response = session.get(base_url, verify=False)
    with open('{}\\KeyImport-{}.csv'.format(report_path, current_date), 'wb') as fout:
        fout.write(response.content)
    session.close()


def ftp_upload():

    ftp = FTP(host=host)

    # print login status
    login_status = ftp.login(user=user, passwd=ftp_password)
    print(login_status)

    # navigate to target folder on the ftp server
    ftp.cwd(ftp_path)

    # read the downloaded file
    #fp = open('C:\\Users\\wmayne\\Downloads\\KeyImporttest.csv', 'rb')
    fp = open('{}\\KeyImport-{}.csv'.format(report_path, current_date), 'rb')

    # upload file
    ftp.storbinary('STOR %s' % os.path.basename(
        '{}\\KeyImport-{}.csv'.format(report_path, current_date)), fp, 1024)
    fp.close()
    # OPTIONAL: remove all files in report path starting with 'KeyImport'
    for f in glob.glob('{}\\KeyImport*'.format(report_path)):
        os.remove(f)

    print(ftp.dir())



