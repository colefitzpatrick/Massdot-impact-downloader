from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import re
import pandas as pd
import os
import time
import openpyxl
import requests
from datetime import date

today = date.today()

url = "https://apps.impact.dot.state.ma.us/cdv/"

# create a new Firefox session
driver = webdriver.Firefox()
driver.implicitly_wait(8)
driver.get(url)

time.sleep(2)

#xpaths for the various steps in the process
guestbutton = '/html/body/app-root/app-landing/div/main/div/div[1]/div[2]/div[1]/button[1]'
limitedbutton = '/html/body/app-root/app-layout/main/div/mat-tab-group/div/mat-tab-body[1]/div/app-define/div/div[2]/div/div/div/div/div/div[2]/div[3]/button/span/img'
typebutton = '/html/body/app-root/app-layout/main/div/mat-tab-group/div/mat-tab-body[1]/div/app-define/div/div[1]/div/div/div[2]/div/button'
basicbutton = '/html/body/app-root/app-layout/main/div/mat-tab-group/div/mat-tab-body[2]/div/app-home/div/div/div[2]/div[1]/app-category-card/div/div[2]/div[3]/button'
visualizebutton = '/html/body/app-root/app-layout/main/div/mat-tab-group/div/mat-tab-body[3]/div/app-data/div/div[1]/div/div[2]/div/button[2]'
downloadbutton = '/html/body/app-root/app-layout/main/div/mat-tab-group/div/mat-tab-body[4]/div/app-results/div/div[2]/div/span[6]/span/button[1]'

#function to click the inputted xpath
def buttonclick( path, sleeptime ):
    clickobject = driver.find_element_by_xpath(path)
    driver.execute_script("arguments[0].click();", clickobject)
    time.sleep(sleeptime)

#navigates through the various button clicks
buttonclick(guestbutton, 1)
buttonclick(limitedbutton, 1)
buttonclick(typebutton, 1)
buttonclick(basicbutton, 1)

#enters the start date
date = driver.find_element_by_id("mat-input-0")
time.sleep(1)
date.send_keys(Keys.CONTROL + "a");
time.sleep(1)
date.send_keys(Keys.DELETE);
time.sleep(1)
date.send_keys("03/01/2020")

#does the final button clicks
buttonclick(visualizebutton, 20)
buttonclick(downloadbutton, 20)

#finds the URL of the download link
href = driver.find_element_by_xpath('/html/body/div[3]/div/div/snack-bar-container/app-snackbar-excel-link/div/div[2]/a')
url2 = href.get_attribute("href")
print(url2)

#downloads the file and puts it in Box, dynamically named with today's date
myfile = requests.get(url2)
filename = "C:\\Users\\Cole Fitzpatrick\\Box\\Covid Daily Crash Data\\impact_2020-03-01 to " + str(today) + ".xlsx"
open(filename, 'wb').write(myfile.content)
