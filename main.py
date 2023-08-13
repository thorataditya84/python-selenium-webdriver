import openpyxl #for reading excel
import pandas as pd #for manipulating tabular data 

import time
import autoit
from selenium.webdriver.common.keys import Keys #new tab thing
from selenium import webdriver

from selenium.webdriver.support.ui import WebDriverWait #alert accept
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains

options =Options()
options.add_argument(r'--user-data-dir=C:/Users/91799/AppData/Local/Google/Chrome/User Data/Default')
options.add_argument('--profile-directory=Default')
browser = webdriver.Chrome(options=options)

loc = "C:/Users/91799/Documents/contacts/sample1.xlsx"
df = pd.read_excel(loc, engine = 'openpyxl')
for i in range(len(df)):
	name = df.iloc[i, 0]
	number = df.iloc[i, 1]
	browser.execute_script("window.open('about:blank', 'tab2');")
	browser.switch_to.window("tab2")
	url = 'https://web.whatsapp.com/send?phone=91' + str(number)
	browser.get(url)
	time.sleep(10)

	msz_box = browser.find_element_by_xpath('/html/body/div/div[1]/div[1]/div[4]/div[1]/footer/div[1]/div[2]/div/div[2]')
	msz = "Hi" + name + ",\n Hello!"
	msz_box.send_keys(msz)
	send_button_msz = browser.find_element_by_xpath('/html/body/div/div[1]/div[1]/div[4]/div[1]/footer/div[1]/div[3]')
	send_button_msz.click()
	time.sleep(2)

print("Msz sent!")
time.sleep(5)
browser.close()
