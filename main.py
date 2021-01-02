import time
import gspread
import os
import cv2
import pytesseract
import pyautogui
from oauth2client.service_account import ServiceAccountCredentials

pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(creds)
sheet_instance = client.open('Miner')
sheet = sheet_instance.get_worksheet(0)

# get total bitcoin wallet balance
# navigate to tab
os.startfile("C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Microsoft Edge")
time.sleep(1)
pyautogui.write("https://www.nicehash.com/my/wallets/BTC")
pyautogui.press('enter')
time.sleep(3)

wallet_region = pyautogui.locateOnScreen('C:\\Users\\Andrew Stelmach\\Desktop\\screenshot\\total_bitcoin.PNG')
region = (wallet_region[0] - 10, wallet_region[1] + 25, 300, 50)
pyautogui.screenshot('C:\\Users\\Andrew Stelmach\\Desktop\\screenshot\\wallet.PNG',
                     region=region)

# extract number into variable
img = cv2.imread('C:\\Users\\Andrew Stelmach\\Desktop\\screenshot\\wallet.PNG')
text = pytesseract.image_to_string(img)
split_string = text.split(" ", 1)
wallet = float(split_string[0])

# get average payrate per day
# https://www.nicehash.com/my/mining/stats

pyautogui.hotkey("ctrl", "t")
pyautogui.write("https://www.nicehash.com/my/mining/stats")
pyautogui.press('enter')
time.sleep(3)

logo = pyautogui.locateOnScreen('C:\\Users\\Andrew Stelmach\\Desktop\\screenshot\\nicehash_logo.PNG')
pyautogui.moveTo(logo)
pyautogui.scroll(-500)
time.sleep(1)

payrate_region = pyautogui.locateOnScreen('C:\\Users\\Andrew Stelmach\\Desktop\\screenshot\\payrate_pre.PNG')
region = (payrate_region[0], payrate_region[1] + 20, payrate_region[2] - 100, payrate_region[3] + 30)
pyautogui.screenshot('C:\\Users\\Andrew Stelmach\\Desktop\\screenshot\\payrate.PNG',
                     region=region)

img = cv2.imread('C:\\Users\\Andrew Stelmach\\Desktop\\screenshot\\payrate.PNG')
text = pytesseract.image_to_string(img)
split_string = text.split(" ", 1)
payrate = float(split_string[0])

for _ in range(2):
    pyautogui.hotkey("ctrl", "w")

time.sleep(2)

sheet.append_row([sheet.row_count, wallet, "", payrate, ""])

formula = '=$F$2*'
sheet.update_acell('C' + str(sheet.row_count + 1), formula + 'B' + str(sheet.row_count + 1))
sheet.update_acell('E' + str(sheet.row_count + 1), formula + 'D' + str(sheet.row_count + 1))
