import sys
import time
import pyautogui
import os
import shutil
import pytesseract
import json
from datetime import datetime, date, timedelta
from PIL import ImageGrab
import PIL.ImageOps  
from textblob import TextBlob
from tqdm import tqdm
import re
import glob

pytesseract.pytesseract.tesseract_cmd = 'D:/programming/PyAutoGUI_session/tesseract/tesseract.exe'

save_path = "E:/lowriskdata_raw"

today = date.today()
date = f"{today.day:02.0f}_{today.month:02.0f}_{today.year:04.0f}"
tomorrow = today + timedelta(days=1)
tomorrowdate = f"{tomorrow.day:02.0f}_{tomorrow.month:02.0f}_{tomorrow.year:04.0f}"

with open('D:/programming/ExtractDataFromPACS/lowriskextractiondict.json') as json_file:
    alldata = json.load(json_file) 

def click_by_img(img_path, shift=None):
    btn_search_loc = pyautogui.locateOnScreen(img_path)
    if btn_search_loc:
        btn_search_loc = pyautogui.center(btn_search_loc)
        x,y = btn_search_loc

        if shift is not None:
            x += shift[0]
            y += shift[1]

        pyautogui.click(x, y)
        #pyautogui.moveTo(x,y)

        return [x,y]
    else:
        return None

stringrule = re.compile(r'PACS[0-9]*')

# Define number of longitudinal exams to be included
inclexams = [1, 2, 3, 4, 5]

print(f"Extract data for patients with {inclexams} exams")

# Print number of patients to be extracted
totnumex = 0
data = {}

for n in inclexams:
    totnumex += len(alldata[str(n)])
    for pat in alldata[str(n)]:
        data[pat] = alldata[str(n)][pat]
print(f"Extract {totnumex} patients in total...")    

# Limit number of patients to be extracted for space reasons
pats = list(data.keys())

# Define a batchsize for extraction
batchsize = 25

batchstart = 0
batchend = batchstart + batchsize

num_batches = len(pats)//batchsize
if len(pats)%batchsize > 0:
    num_batches += 1

# Give some time switch to Enterprise Imaging
time.sleep(5)

for num_batch in range(num_batches):
    print(f"Extracting batch {num_batch+1}/{num_batches}")
    failedextractionsdict = {}

    if batchend > len(pats):
        batchend = len(pats)    

    batch = pats[batchstart:batchend]

    print("Extracting data from Enterprise Imaging...")
    for pat_id in tqdm(batch):
        noentry = False
        # Put in patient ID
        _ = click_by_img("D:\programming\ExtractDataFromPACS\imagestomatch\patID.png", [400,0])
        if _ is not None:
            pyautogui.hotkey('ctrl','a')
            pyautogui.write(pat_id, interval=0.01)
            pyautogui.press('enter') 

        time.sleep(1)

        accpos = pyautogui.locateOnScreen("D:/programming/ExtractDataFromPACS/imagestomatch/accessionnr.png")
        accpos = pyautogui.center(accpos)

        box_height = 19
        box_width = 100

        pyautogui.moveTo(accpos[0]-22, accpos[1])
        #pacsnrlist = list(data[pat_id].keys())
        pacsnrlist = list(data[pat_id])

        with pyautogui.hold('ctrl'): 
            while pacsnrlist:
                pyautogui.move(0, box_height)    

                curr_pos = pyautogui.position()

                screen = ImageGrab.grab(bbox=(curr_pos[0]-box_width//2,curr_pos[1]-box_height//2,curr_pos[0]+box_width//2,curr_pos[1]+box_height//2))
                screen = screen.convert('L')
                screen = PIL.ImageOps.invert(screen)

                # Scale up image to increase probability of correct detection
                width, height = screen.size
                scale_factor = 3
                screen = screen.resize((width*scale_factor, height*scale_factor))
                #screen.save('D:/programming/ExtractDataFromPACS/grabbed.png')
            
                text = pytesseract.image_to_string(screen)
                correctedText = TextBlob(text).correct()
            
                puretext = correctedText.split("\n")[0] 

                if puretext == "":
                    if not noentry:
                        noentry = True
                        continue
                    else:    
                        failedextractionsdict[pat_id] = pacsnrlist
                        print(f"Couldn't find all files for {pat_id}: {pacsnrlist} was not found !")
                        break 
                else:
                    noentry = False        
        
                # Deal with accidental space in the PACS number
                if " " in puretext:
                    puretext = "".join(puretext.split(" "))  

                # Find PACS string in found text
                try:
                    puretext = stringrule.search(puretext).group(0)
                except:
                    continue    

                if puretext in pacsnrlist:
                    pyautogui.click(pyautogui.position())
                    pacsnrlist.remove(puretext)
                    noentry = False

        if len(pacsnrlist) > 0:
            continue
        # Export data
        _ = click_by_img("D:/programming/ExtractDataFromPACS/imagestomatch/exportSymbol.png")

        # Wait until export screen has loaded fully
        while click_by_img("D:/programming/ExtractDataFromPACS/imagestomatch/lucmfr.png") is None:
            time.sleep(1)
        _ = click_by_img("D:/programming/ExtractDataFromPACS/imagestomatch/sendData.png")

        # Wait until export is successful
        while True:
            time.sleep(1)
            _ = click_by_img("D:\programming\ExtractDataFromPACS\imagestomatch\coffeeCup.png")
            if _ is None:
                break
                
        time.sleep(1)
        click_by_img("D:\programming\ExtractDataFromPACS\imagestomatch\closeButton.png", [25, -25])
        time.sleep(5)

    # Print out failed extractions
    if len(list(failedextractionsdict.keys())) > 0:
        print("Extractions failed for:")
    else:
        print("Images extracted successfully!")  

    for pat_id in failedextractionsdict:
        print(f"    {pat_id} : {failedextractionsdict[pat_id]}")

    # Move data to local
    print("Moving data to local...")
    for pat_id in tqdm(batch):
        # Skip failed extractions
        if pat_id in failedextractionsdict:
            continue      
        if pat_id in os.listdir(save_path):
            shutil.rmtree(f"{save_path}/{pat_id}")
        for pacsnr in data[pat_id]:
            img_path = []
            for datestring in [date, tomorrowdate]:
                try:
                    img_path += glob.glob(f"//UZ/Data/lucmfr/vol2/gladys/backup/{datestring}/*/{pacsnr}/MG")
                except:
                    continue    
            dropoff_path = f"{save_path}/{pat_id}/{pacsnr}"
            try:
                shutil.copytree(img_path[0], dropoff_path)
            except:
                print(f"{pat_id} {pacsnr} failed to copy!")
                continue    

    print("Data moved successfully!")            

    # Delete all files from LUCMFR folder to prevent cluttering
    print("Delete data from LUCMFR server...")
    for pat_id in tqdm(batch):
        # Skip failed extractions
        if pat_id in failedextractionsdict:
            continue
        failed = False    
        for pacsnr in data[pat_id]: 
            removed = False
            for datestring in [date, tomorrowdate]:
                try:
                    pathtormv = glob.glob(f"//UZ/Data/lucmfr/vol2/gladys/backup/{datestring}/*/{pacsnr}/MG")[0]
                    shutil.rmtree(pathtormv, ignore_errors=True)
                    removed = True
                except:
                    continue
            if not removed:
                failed = True
                print(f"{pat_id} {pacsnr} failed to remove!") 
                         
            for datestring in [date, tomorrowdate]:
                try:  
                    pathtormv2 = glob.glob(f"//UZ/Data/lucmfr/vol2/gladys/backup/{datestring}/{pacsnr}")[0]
                    shutil.rmtree(pathtormv2, ignore_errors=True)
                except:    
                    continue
        if failed:
            try:
                shutil.rmtree(f"{save_path}/{pat_id}")   
            except:
                print(f"Failed to remove {save_path}/{pat_id}")        


    print("Data deleted successfully from LUCMFR server!")   
    batchstart += batchsize
    batchend += batchsize   