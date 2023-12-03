#Written by Katri Keränen on Nov 28th 2023 =) with Python ver 3.11.6
#Requires Python version 3.7 or higher so that dictionaries are Ordered.
import requests
import json
import xlsxwriter
from PIL import Image
import os
from datetime import datetime

#Fetches emojilist from Slack API and saves it to dictionary
def get_list():
    response = requests.post('https://slack.com/api/emoji.list', {
        'token':'TOKENHERE'})
    data = response.json()
    return data

#Creates subdirectory for emojis and calls to save_list to save them to said subdirectory
def save_emojis(emojis):
    print('Getting date and creating subdirectory..')
    #Gets the date and time for subdirectory name
    timeNow = datetime.now()
    dtt = f'{timeNow.day}{timeNow.month}{timeNow.year}{"_"}{timeNow.hour}{timeNow.minute}{timeNow.second}'
    path = 'emojis_' + dtt
    #Creates subdirectory
    os.mkdir(path)
    print('Subdirectory created.')
    save_list(emojis, path)
    print('Completed saving to folder ', path)
    
#Goes through emojis.items, saves images as is
def save_list(emojis, path):
    print('Starting to go through emojis..')
    for key, value in emojis.items():
        #Skips aliases, don't want to save those
        if value.startswith('alias'):
            continue
        #For URLS calls to save_image to the received path
        else:
            save_image(key, value, path)    

#Fetches the images from URLS and saves them with emoji name + file extension
def save_image(key, url, path):
    res = requests.get(url)
    ext = url[-4:]
    t = path+'\\'+key+ext
    with open(t, 'wb') as f:
        f.write(res.content)

#Creates an Excel file for emojis and fills it with emoji names and images.        
def excel_emojis(emojis):
    name = input('Give a name to the Excel file: ') + '.xlsx'
    #Creates an Excel workbook by given name
    workbook = xlsxwriter.Workbook(name)
    #Creates a worksheet to workbook
    worksheet = workbook.add_worksheet()
    #Sets row height and column width higher to accommodate emojis
    worksheet.set_default_row(110)
    worksheet.set_column("B:B",22)
    worksheet.set_column("A:A",18)
    
    #Sets a simple header
    worksheet.write(0,0,'Name')
    worksheet.write(0,1,'Image')
    
    print('Excel file created.')
    
    #Sets the place to begin filling the file.
    row = 1
    col = 0
    
    #Creates a temporary folder for the images.
    path = 'tempemojis'
    os.mkdir(path)
    print('Temporary image directory created.')
    
    print('Starting to go through emojis..')
    #Goes through each emoji object's key and value and adds them to Excel file.
    for key, value in emojis.items():
        worksheet.write(row, col, key)
        #Aliases are added as text.
        if value.startswith('alias'):
            worksheet.write(row, col + 1, value)
        #For URLS calls to get_image to temporarily save image and adds it to Excel.
        else:
            fname = get_image(path, value)
            worksheet.insert_image(row, col + 1, fname)
        row += 1
        
    workbook.close()
    print('Excel file closed.')
    dest_temp(path)
    print('Temporary image directory deleted.')

#Gets image by URL, saves it to temporary image directory, converts it to PNG (kind of) and returns filename
def get_image(path, url):
    res = requests.get(url)
    fname = extract_name(url)
    t = path+'\\'+fname
    with open(t, 'wb') as f:
        f.write(res.content)
    #I had to add this convert to PNG-ish part to deal with emojis that were saved as PNGs and whatever but were actually webp format, 
    # as they caused errors. It doesn't actually transform all to PNGs but it prevents errors so I kept it. =)
    im = Image.open(t).convert("RGBA")
    im.save(t,"png")
    return t

#Splits URL to get filename for temporary images
def extract_name(url):
    file_name = url.split("/")[-1]
    return file_name

#Destroys temporary image directory and the images in it.
def dest_temp(path):
    print('Starting to delete temporary files..')
    files = os.listdir(path)
    #Goes through files in directory and deletes them one by one
    for fname in files:
        t = path + '\\' + fname
        if os.path.isfile(t):
            os.remove(t)
        else:
            print('File', fname, 'couldn''t be removed')
    #Deletes empty folder
    os.rmdir(path)

#Calls to get_list to get emoji JSON from Slack API, saves emoji part to it's own dictionary
print('Fetching emoji list..')
data = get_list()
emojis = data['emoji']
print('Emoji list fetched.', len(emojis) ,'emojis and aliases found.')

#Loop for repeating commands.
pyorii = True
while pyorii:
    print('Choose action:\n1 - Save emoji images to a new subdirectory.\n2 - Put emoji names and images to a new Excel file.\n3 - Exit')
    act = input()
    if act == '1':
        save_emojis(emojis)
    elif act ==  '2':
        excel_emojis(emojis)
    elif act == '3':
        print('Kiitos näkemiin thank you and bye')
        pyorii = False
    else:
        print("Not acceptable action, please retry.")