# Author: Junying Chen
# Date: 14/07/2017
# Description:  This python script downloads youtube notifications sent to a folder on your gmail account. It is meant to be runned by
#               Windows task scheduler during off peak periods so that you do not use up your on peak download quota defined by your
#               internet subscription.
from __future__ import print_function
from __future__ import unicode_literals
import httplib2
import sys
import os
import re
import base64
import email
from html.parser import HTMLParser
import datetime
import subprocess

from apiclient import discovery
from apiclient import errors
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

import youtube_dl

# Get commandline arguments
try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at <HOME FOLDER>\.credentials\gmail-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/gmail.modify'
CLIENT_SECRET_FILE = 'client_secret.json' # Client secret file is meant to be in the same folder is this script
APPLICATION_NAME = 'Gmail API Python Quickstart'

SCRIPT_PATH = os.path.realpath(__file__)
HOME_DIR = os.path.expanduser('~') + "\\OffPeakYt" #os.path.dirname(SCRIPT_PATH)
WORK_DIR = os.path.join(HOME_DIR, 'work') # Define path where all downloaded videos go
LOG_PATH = os.path.join(HOME_DIR, 'offpeakyt.log') # Define path of log file
LOG_PTR = open(LOG_PATH, 'a') # Open file in append mode

# Get exception details
def get_excdetails(e):
    exc_type, exc_obj, exc_tb = sys.exc_info()
    fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
    
    excMsg = str(e) + "\n" + fname + ":" + str(exc_tb.tb_lineno)
    return excMsg

# print errors out on to screen for debugging and into log file
def print_logscreen(msg):
    try:
        print(msg)
        print(msg, file = LOG_PTR)
    except Exception as e:
        errorMsg = "Failed to log message:\n" + get_excdetails(e)
        print(errorMsg)
        print(errorMsg, file = LOG_PTR)
    LOG_PTR.flush()

# Get google account OAuth2 credential file path from <HOME FOLDER>\.credentials
def get_credentials():
    credential_dir = os.path.join(HOME_DIR, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'gmail-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

# Get gmail email pointers, download email contents, extract youtube video ID and passes it to youtube_dl to download video
def process_emails(emailPointers, service):
    print("Getting emails and sorting them by date")
    emailList = []
    for emailPointer in emailPointers:
        emailObj = service.users().messages().get(userId='me',
                                                  id=emailPointer['id'],
                                                  format='raw').execute()
        emailList.append(emailObj)
    emailList.sort(key=lambda email: email['internalDate'])

    print_logscreen("Videos to download:")
    [print_logscreen("- " + email['snippet']) for email in emailList]

    print("Downloading videos")
    for email in emailList:
        try:
            emailBuffer = base64.urlsafe_b64decode(email['raw'].encode('ASCII'))
            rawEmailStr = str(emailBuffer)
            emailStr = rawEmailStr.replace(r'=\r\n', '') # Remove email body newline. Need to do this as youtube links get split by them
            reResults = []

            # Record email subject and date for log file
            subjectRes = re.search('Subject: (.+?)From', emailStr)
            subject = subjectRes.group(1).strip() if hasattr(subjectRes, 'group') else emailStr
            dateRecvRes = re.search('Date: (.+?)List', emailStr)
            dateRecv = dateRecvRes.group(1).strip() if hasattr(dateRecvRes, 'group') else 'UNKNOWN'
            
            # There's variation in how google embeds youtube links in thier notification emails so we have to use multiple regex strings
            # to extract the youtube video ID
            reResult1 = re.search('\/watch\?v=(.+?)&feature', emailStr)
            reResults.append(reResult1)
            reResult2 = re.search('\/watch\?v%3D(.+?)%26feature', emailStr)
            reResults.append(reResult2)
            reResult3 = re.search('\/watch%3Fv%3D(.+?)%26feature', emailStr)
            reResults.append(reResult3)
            reResult4 = re.search('\/watch\?v%(.+?)%26feature', emailStr)
            reResults.append(reResult4)
            reResult5 = re.search('\/watch\?v=(.+?)&', emailStr)
            reResults.append(reResult5)

            # Get youtube video ID from regex object, join with the youtube watch URL and it to youtube_dl to download. If regex finds nothing
            # record it in log file so we know that we have to update our regular expressions
            videoUrl = None
            while reResults:
                if len(reResults) > 0:
                    try:
                        reRes = reResults.pop(-1)

                        videoId = reRes.group(1).strip()
                        
                        if not os.path.exists(WORK_DIR):
                            os.makedirs(WORK_DIR)
                        os.chdir(WORK_DIR)
                        ydl_opts = {}
                        videoUrl = 'https://www.youtube.com/watch?v=' + videoId
                        
                        with youtube_dl.YoutubeDL(ydl_opts) as ydl:
                            ydl.download([videoUrl])
                        break
                    except Exception as e:
                        print("Download failed:\n" + get_excdetails(e))
                        if videoUrl is not None:
                            print("Video source URL: " + videoUrl)

                    os.chdir(HOME_DIR)
                else:
                    raise ValueError("Video link in email id = " +
                                     email['id'] + " not found or not " +
                                     "recognized. Have to download manually " +
                                     "from " + str(subject) + " (" +
                                     str(dateRecv) + ")")
            
            # Once the video is downloaded, we can delete it from our folder in gmail
            print_logscreen("Successfully downloaded video id = " + videoId)
            print("Deleting email id = " + email['id'])
            service.users().messages().trash(userId='me', id=email['id']
                                             ).execute()
        except Exception as e:
            print_logscreen("Processing email failed:\n" + get_excdetails(e))
            print_logscreen("Skipping email id = " + email['id'])

# This will be used in a future project to extend the functionality of this script to be able to seperate videos by description
def get_video_length(filename):
    result = subprocess.Popen(["ffprobe", filename], stdout = subprocess.PIPE,
                              stderr = subprocess.STDOUT)
    return [x for x in result.stdout.readlines() if "Duration" in x]

if __name__ == '__main__':
    try:
        print_logscreen("===========================================")
        print_logscreen("Start time: " + str(datetime.datetime.now()))

        # OAuth authentication with google API
        print("Authenitcating with Google API")
        credentials = get_credentials()
        http = credentials.authorize(httplib2.Http())
        service = discovery.build('gmail', 'v1', http=http)

        # Get all messages in folder in gmail ID Label_24 (folder containing all youtube notifications)
        results = service.users().messages().list(userId='me', labelIds="Label_24",
                                                maxResults=1000).execute()
        emailPointers = results.get('messages', [])
        # Process emails and download videos
        process_emails(emailPointers, service)
        print_logscreen("End time: " + str(datetime.datetime.now()))
    except Exception as e:
        print_logscreen("Unhandled exception:\n" + get_excdetails(e))
LOG_PTR.close()
