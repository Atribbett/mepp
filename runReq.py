import os
import math
import time
import requests

import gspread

from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas

from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

from oauth2client.service_account import ServiceAccountCredentials

# CONSTANTS
# --------------
# Constants only to be changed if new database is used.
GDOCS_OAUTH_JSON = "MePurchasing-117856214d4f.json"
GDOCS_SPREADSHEET_NAME = "Rowan Mechanical Purchase Request (Responses)"
WORKSHEET_NAME = "Testing"
#WORKSHEET_COLUMN_COUNT = 3


# FUNCTIONS
# --------------
def login_open_sheet(oauth_key_file, spreadsheet, sheet):
    """ Grant access to database worksheet using OAuth key in json file. """

    try:
        scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/forms']
        creds = ServiceAccountCredentials.from_json_keyfile_name(oauth_key_file, scope)
        gc = gspread.authorize(creds)
        requestedSheet = gc.open(spreadsheet).worksheet(sheet)
        return requestedSheet
    except Exception as ex:
        print('Unable to login and get spreadsheet. Check OAuth credentials, spreadsheet name, and make sure spreadsheet is shared to the client_email address in the OAuth .json file!')
        print('Google sheet login failed with error:', ex)
        sys.exit(1)

def connectToDrive():
    """Connect to Google Drive based on client_secret.json file
    
    Checks credentials and creates them if none are found so user
    does not need to allow access each time program runs
    
    Returns: Drive object
    """
    
    gauth = GoogleAuth()
    # Try to load saved client credentials
    gauth.LoadCredentialsFile("mycreds.txt")
    if gauth.credentials is None:
        # Authenticate if they're not there
        gauth.LocalWebserverAuth()
    elif gauth.access_token_expired:
        # Refresh them if expired
        gauth.Refresh()
    else:
        # Initialize the saved creds
        gauth.Authorize()
    # Save the current credentials to a file
    gauth.SaveCredentialsFile("mycreds.txt")
    drive = GoogleDrive(gauth)
    return drive

def downloadFile(fileId,localFilename = 'temp.pdf'):
    """Download file from drive using file Id
    
    Return the filename as it was listed in the drive
    """
    
    tempPdf = drive.CreateFile({'id':fileId})
    driveFilename = tempPdf['title']
    tempPdf.GetContentFile(localFilename)
    #print(driveFilename + ' has been downloaded and saved as ' + localFilename)
    return driveFilename

def watermarkCorners(watermarkText,localFilename = 'temp.pdf'):
    """Watermark four corners of every page with text"""

    output = PdfFileWriter()
    originalPdf = PdfFileReader(open(localFilename, "rb"),strict=False)
    
    # Measure dims of original document
    x, y = originalPdf.getPage(0).mediaBox.getUpperRight()
    x = int(x)
    y = int(y)

    # Create watermark
    c = canvas.Canvas('tempMark.pdf', pagesize=(x,y))
    c.setFont("Helvetica",14)
    c.setFillColorRGB(.5,.25,0)
    c.drawString(10,10,watermarkText)
    c.drawString(10,y-24,watermarkText)
    c.drawRightString(x-10,y-24,watermarkText)
    c.drawRightString(x-10,10,watermarkText)
    c.save()
    
    # Watermark all pages
    watermark = PdfFileReader(open('tempMark.pdf', "rb"))
    for p in range(0,originalPdf.getNumPages()):
        page = originalPdf.getPage(p)
        page.mergePage(watermark.getPage(0))
        output.addPage(page)
        
    # Output merged document
    outputStream = open('mergedPdf.pdf', "wb")
    output.write(outputStream)

def uploadFile(driveFilename, folderId, localFilename = 'mergedPdf.pdf'):
    fileToUpload = drive.CreateFile({"title":["M" + driveFilename], "parents":[{"kind":"drive#fileLink", "id":folderId}]})
    fileToUpload.SetContentFile(localFilename)
    fileToUpload.Upload()
    return fileToUpload['id']
    
    
#Main
quoteList = None
saveFolderId = "0BzZSztlM2pZYS0JEWjg0Slc0QVU"
quoteCol = 26       # Col "Z"
processedCol = 32   # Col "AF"
orderTypeCol = 43   # Col "AQ"
markedQuoteCol = 45 # Col "AS"
reqNumberCol = 33   # Col "AG"
submitUrlCol = 58   # Col "BF"
pollTime = 360      # sleep time in seconds  


os.environ['TZ'] = 'EST5EDT'
#time.tzset()

while True:
    quoteList = login_open_sheet(GDOCS_OAUTH_JSON,GDOCS_SPREADSHEET_NAME,WORKSHEET_NAME)
  
    colValues = quoteList.col_values(markedQuoteCol)

    for r in range(1,len(colValues)):
        if colValues[r] == "":   
            # Look up job information from spreadsheet
            rowData = quoteList.row_values(r+1)
            if rowData[orderTypeCol-1] == "Approved Vendor":
                if rowData[reqNumberCol-1] != "":
                    reqNum = rowData[reqNumberCol-1]
                    fileId = rowData[quoteCol-1].split("=",1)[1] 

                    # Perform watermarking service
                    drive = connectToDrive()     
                    driveFilename = downloadFile(fileId)
                    watermarkCorners(reqNum)
                    print("New Merge Completed - %s" % (driveFilename))

                    # Save file to drive and place link into spreadsheet
                    uploadId = uploadFile(driveFilename, saveFolderId)
                    quoteList.update_acell('AS'+str(r+1), str("https://drive.google.com/open?id="+uploadId))
            elif rowData[orderTypeCol-1] == "Non-Approved Vendor" or rowData[orderTypeCol-1] == "One-Click Vendor":
                quoteList.update_acell('AS'+str(r+1), str("N/A"))
            else:
                continue
            
            # # Look up job information from spreadsheet
            # rowData = quoteList.row_values(r+1)
            # if rowData[processedCol-1] == "":
                # #print('Row' + str(r+1) + ': Order Not Reviewed Yet')
                # continue                
            # elif rowData[processedCol-1] != "" and rowData[quoteCol-1] == "":
                # quoteList.update_acell('AS'+str(r+1), str("N/A"))
            # else:
                # reqNum = rowData[reqNumberCol-1]
                # fileId = rowData[quoteCol-1].split("=",1)[1] 

                # # Perform watermarking service
                # drive = connectToDrive()     
                # driveFilename = downloadFile(fileId)
                # watermarkCorners(reqNum)
                # print("New Merge Completed - %s" % (driveFilename))

                # # Save file to drive and place link into spreadsheet
                # uploadId = uploadFile(driveFilename, saveFolderId)
                # quoteList.update_acell('AS'+str(r+1), str("https://drive.google.com/open?id="+uploadId))
                
                # Resubmit form to trigger email
                #submitUrl = rowData[submitUrlCol-1]
                #url = submitUrl.split('?')[0]
                #user_agent = {'Referer':url}
                #url.replace('viewform','formResponse')
                #formData = submitUrl.split('?')[1]
                
                #submitUrl = submitUrl + '&submit=Submit'
                
                #r = requests.post(url,data=formData,headers=user_agent)
                #print(r.status_code)
                
    print('Merges up to date ' + str(time.strftime('%x %X %Z')))
    time.sleep(pollTime)