import re
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import time
import datetime
import xlwings as xw
import pytz
from selenium.webdriver.common.by import By

driver = webdriver.Chrome(ChromeDriverManager().install())
driver.get('https://www.ethercrash.io/play')

# delay 5 sec
time.sleep(5)

# click "History" button
button = driver.find_element(By.XPATH, "//*[@class='tab col-2 noselect']")
button.click()

maxId = 0
sheetIndex = 1

# create new excel file
wb = xw.Book()
wb.save("./" + str(datetime.datetime.now(pytz.timezone('Asia/Shanghai')).hour) + "-" + str(datetime.datetime.now(pytz.timezone('Asia/Shanghai')).minute) + "-" + str(datetime.datetime.now(pytz.timezone('Asia/Shanghai')).second) + ".xlsx")

# scrap the bust value
def scrap_bustvalue():
    global maxId
    global sheetIndex
    global wb
    global button
    gameId_arr = []

    # get the max value from href
    lnks = driver.find_elements(By.XPATH, "//a[starts-with(@class, 'games-log-')]")
    connectionState = driver.find_element(By.CLASS_NAME, "connection-state").text
    if connectionState != "Connection Lost ...": 
        # get the all busted value of the table
        for lnk in lnks:
            url = lnk.get_attribute('href')
            gameId = re.split(r"([0-9]+)",url)[1]
            gameId_arr.append(int(gameId))

        # sort the array
        gameId_arr.sort()

        # get the new busted value and write the data to excel file
        for gameId in gameId_arr:
            #get the bust value by gameId 
            if gameId > int(maxId):                
                bustValue = driver.find_element(By.XPATH, '//a[@href="/game/'+str(gameId)+'"]').text
                # print(bustValue)
                # write the data to excel file

                # change the color of text 
                # red color if value is under 2
                # green color if value is over 2
                intValue = re.split(r"(x)",bustValue)[0]
                print(intValue)

                sht1 = wb.sheets['Sheet1']
                colA = 'A' + str(sheetIndex)
                colB = 'B' + str(sheetIndex)
                sht1.range(colA).value = intValue
                sht1.range(colB).value = datetime.datetime.now(pytz.timezone('Asia/Shanghai'))


                emp_str = re.sub("[^\d\.]", "", intValue)
                # print(emp_str)
                
                if float(emp_str) >= float(2):                
                    sht1.range(colA).color = (0, 170, 100)
                    sht1.range(colB).color = (0, 170, 100)
                else:
                    sht1.range(colA).color = (170, 90, 50)
                    sht1.range(colB).color = (170, 90, 50)
                sheetIndex += 1
                maxId = gameId
                wb.save()
            else:
                continue
    else:
        button.click()
        button.click()
        button.click()

# infinite loop            
while True:
    scrap_bustvalue()
    time.sleep(5)
