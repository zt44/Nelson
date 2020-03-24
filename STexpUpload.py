from selenium import webdriver
from selenium.webdriver.common import keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains

import time
import csv
import sys

from datetime import datetime, timedelta

key = keys.Keys

# start chrome
options = webdriver.ChromeOptions()
options.add_argument("user-data-dir=C:\\Users\\ltran\\AppData\\Local\\Google\\Chrome\\User Data")
options.add_argument("profile-directory=Default")
options.add_argument('--ignore-certificate-errors')
options.add_argument('--ignore-ssl-errors') 
driver = webdriver.Chrome(executable_path='chromedriver.exe', options=options)
driver.get("https://smarttrack.gcpnode.com/home")
builder = webdriver.ActionChains(driver)
wait = WebDriverWait(driver, 20)

# Open expense file
exp_file = open('Expense Output.csv', encoding='utf-8-sig')
exp_reader = csv.reader(exp_file, delimiter=',')

today = datetime.today()

today = today - timedelta(days = 6)

d4 = today.strftime("%b-%d-%Y")

d4_list = d4.split('-')

day = d4_list[1] + '/' + d4_list[0] + '/' + d4_list[2]

dt = datetime.strptime(day, '%d/%b/%Y')
start_week = dt - timedelta(days=dt.weekday())
end_week = start_week + timedelta(days=6)

start_week = str(start_week)
end_week = str(end_week)

start_week = start_week.replace("-", "/")

end_week = end_week.replace("-", "/")

start_week = start_week.split(" ")[0]
end_week = end_week.split(" ")[0]

start_week = start_week.split("/")
end_week = end_week.split("/")

start = start_week[1] + "/" + start_week[2] + "/" + start_week[0]

end = end_week[1] + "/" + end_week[2] + "/" + end_week[0]

# Get week begin and end dates from user, this happens once per total week of expenses
print("Please enter the beginning and end dates for the expense period this script will invoice.")
print("Dates must be formatted MM/DD/YYYY to conform to SmartTrack, otherwise this script will crash.")
weekBegin = str(start)
weekEnd = str(end)

time.sleep(4)

side_bar_exp_button = driver.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/div/div[1]/div[1]/div/nav/div/ul/li[6]/a')

side_bar_exp_button.click()

time.sleep(1)

side_bar_exp = driver.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/div/div[1]/div[1]/div/nav/div/ul/li[6]/ul/li[2]/a')

side_bar_exp.click()

currentTVC = ""
roundOne = True
firstItem = True

for exp_row in exp_reader:
    # Check if want to create new expense report
    if exp_row[0] != currentTVC:
        # if first round dont ask to submit but afterwards it means we have reached end of a TVC
        if roundOne:
            roundOne = False
        else:
            input("Please check and submit the current open report and hit enter to continue: ")

        newRI = input("Do you want to create a new report for " + exp_row[0] + " (y/n): ")

        while newRI != 'y' and newRI != 'n':
            newRI = input("Do you want to create a new report for " + exp_row[0] + " (y/n): ")

        currentTVC = exp_row[0]
        firstItem = True

    else:
        newRI = "continuing"

    # Create new report for yes
    if newRI is 'y':
        # prompt to go to new report page
        input("Please navigate to the new report page and hit enter: (enter)")

        xWNumber = driver.find_element_by_xpath('//*[@id="txtSaeTest"]')
        xWNumber.send_keys(exp_row[7])
        time.sleep(2)
        click_xWNumber = driver.find_element_by_xpath('//*[@id="rh_contentContainer"]/div[4]/div/div[3]/div/div/div/div')
        click_xWNumber.click()

        # continue creation code here
        tvcNameBox = driver.find_element_by_xpath('//*[@id="txtCwName"]')
        tvcNameBox.send_keys(exp_row[0].split(' ')[0])

        # select TVC name
        time.sleep(1.75)
        tvcNameDrop = driver.find_element_by_xpath('//*[@id="txtCwName"]/../div')
        builder = webdriver.ActionChains(driver)
        builder.send_keys(key.ARROW_DOWN).perform()
        time.sleep(1)
        loopCount = 0
        while True:
            try:
                tvcName = driver.find_element_by_xpath('//div[@tabindex="0" and contains(.,"'
                                                            + exp_row[0].split(' ')[1] + '")]')
                builder.send_keys(key.ENTER).perform()
                break
            except:
                if loopCount < 6:
                    builder.send_keys(key.ARROW_DOWN).perform()
                    time.sleep(1.5)
                    loopCount += 1
                else:
                    print('failed xpath=//div[contains(@class, "opt") and contains(.,"'
                                                                + exp_row[0].split(' ')[1] + '")]')
                    input("Cant find TVC, please find name manually and hit enter to continue")
                    break

        time.sleep(2)

        chrg_num = driver.find_element_by_xpath('//*[@id="txtChargeNumber_chosen"]/a/span')
        chrg_num.click()
        builder.send_keys(key.ARROW_DOWN).perform()
        builder.send_keys(key.ENTER).perform()

        time.sleep(2)

        #bDateInput = driver.find_element_by_xpath('//*[@id="txtDateFrom"]')
        #eDateInput = driver.find_element_by_xpath('//*[@id="txtDateTo"]')

        bDateInput = driver.find_element_by_id('txtDateFrom')
        eDateInput = driver.find_element_by_id('txtDateTo')

        calender_button1 = driver.find_element_by_xpath('//*[@id="rh_contentContainer"]/div[4]/div/div[10]/div/div/span')
        calender_button1.click()
        calender_button1.click()
        calender_button1.click()
        bDateInput.send_keys(weekBegin)

        calender_button2 = driver.find_element_by_xpath('//*[@id="rh_contentContainer"]/div[4]/div/div[11]/div/div/span')
        calender_button2.click()
        calender_button2.click()
        calender_button2.click()
        eDateInput.send_keys(weekEnd)

        ''' 
        bDateInput = driver.find_element_by_xpath('//*[@id="txtDateFrom"]')
        eDateInput = driver.find_element_by_xpath('//*[@id="txtDateTo"]')
        bDateInput.send_keys(key.LEFT_SHIFT, key.TAB)
        # input("wait")
        builder.send_keys(key.ARROW_DOWN).perform()
        # input("wait")
        builder.send_keys(key.ARROW_DOWN).perform()
        # input("wait")
        builder.send_keys(key.ENTER).perform()
        '''

        # Enter expense summary and purpose
        expSummary = driver.find_element_by_xpath('//*[@id="txtTitle"]')
        expPurpose = driver.find_element_by_xpath('//*[@id="txtPurpose"]')
        expSummary.send_keys('Expenses for Week Ending ' + weekEnd)
        expPurpose.send_keys('Weekly TVC expenses')

        # REMEMBER TO ADD CODE TO CLICK SAVE AND CONTINUE TO CREATE REPORT AFTER TESTING
        #
        #
        # !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        builder = webdriver.ActionChains(driver)
        # click save
        # !!!! MAKE SURE SAVE BUTTON IS IN VIEW, HARD TO SCROLL IN SELENIUM. ZOOM OUT ON PAGE TO FIX
        addExpSubmit = driver.find_element_by_xpath('//*[@id="AddExpSubmit"]')
        addExpSubmit.send_keys(key.ENTER)
        input("Please confirm the new report and hit enter: (enter)")
        time.sleep(1)
        # click yes
        # builder.send_keys(key.TAB)
        # builder.send_keys(key.ENTER).perform()
        # time.sleep(1)
        # # click ok
        # builder.send_keys(key.TAB)
        # builder.send_keys(key.ENTER).perform()
        # time.sleep(3)

    # if no, ask to open report to continue from or quit
    elif newRI is 'n':
        print("Please navigate to the expense report for " + exp_row[0] + " that you would like to continue from.")
        print("Once there enter 'c' to continue, if you would like to quit this script enter 'q'")
        contI = input("Continue or Quit? (c/q): ")
        while contI != 'c' and contI != 'q':
            contI = input("Continue or Quit? (c/q): ")

        # if c just let it exit the else to start entering line items
        if contI is 'q':
            exp_file.close()
            sys.exit()

    # start line item entry code here, expect to be inside report page
    # click add item to add new item
    input('press enter for next item after confirming current item')
    print("Entering " + exp_row[3])
    addItem = driver.find_element_by_xpath('//a[@class="expenseAdd"]')
    moveTo = driver.find_element_by_xpath('//div[@id="contentContainer"]')
    while True:
        try:
            builder = webdriver.ActionChains(driver)
            builder.move_to_element(moveTo).perform()
            # print("builder move worked")
            # moveTo.click()
            # print("moveto click worked")
            addItem.send_keys(key.ENTER)
            # print("add item click worked")
            break
        except:
            input("Script error: Please navigate to details and back to this report, then hit enter: (enter)")
            addItem = driver.find_element_by_xpath('//a[@class="expenseAdd"]')
            moveTo = driver.find_element_by_xpath('//div[@id="contentContainer"]')

    time.sleep(1.25)
    
    time.sleep(30)
    # ----- fill in info
    # date
      
    expDate = driver.find_element_by_xpath('//*[@id="txtExpenseDate"]')
    weekDateB = datetime(int(weekBegin.split('/')[2]), int(weekBegin.split('/')[0]), int(weekBegin.split('/')[1]))
    expDatetime = datetime(int(exp_row[2].split('/')[2]), int(exp_row[2].split('/')[0]), int(exp_row[2].split('/')[1]))
    # this needs to be fixed always hits the if and never the else, sending same for now
    # deal with error on confirm 
    
    # select category//*[@id="txtExpenseDate"]
    cat = exp_row[6]
    categoryDrop = driver.find_element_by_xpath('//*[@id="ExpeneseCategory"]')
    # categoryDrop.click()
    if cat[0] == 'b':
        # Select "Background Check"
        categoryDrop.send_keys("Back")
    elif cat[0] == 'e':
        # Select 'Entertainment'
        # driver.find_element_by_xpath('//*[@id="ExpeneseCategory"]/option[5]').click()
        categoryDrop.send_keys("Entert")
    elif cat[0] == 't':
        # Select 'Travel'
        # driver.find_element_by_xpath('//*[@id="ExpeneseCategory"]/option[14]').click()
        categoryDrop.send_keys("Travel")
    elif cat[0] == 'o':
        # Select 'Other'
        categoryDrop.send_keys("Other")
    else:
        input('Invalid category parsed, press enter to exit.')
        sys.exit()

    # select expense type
    time.sleep(0.25)
    expTypeDrop = driver.find_element_by_xpath('//*[@id="expensetype"]')
    # expTypeDrop.click()
    if cat[0] == 'e':
        # Select 'Client Entertainment'
        # driver.find_element_by_xpath('//*[@id="expensetype"]/option[2]').click()
        expTypeDrop.send_keys("Client")
    elif cat[0] == 'b':
        # Select 'Background check fee'
        expTypeDrop.send_keys("Back")
    elif cat[0] == 't':
        if cat[1] == 'a':
            # Select "Auto Exp"
            # driver.find_element_by_xpath('//*[@id="expensetype"]/option[4]').click()
            expTypeDrop.send_keys("Auto E")
        elif cat[1] == 'f':
            # Select "Airfare (Domestic)"
            # driver.find_element_by_xpath('//*[@id="expensetype"]/option[2]').click()
            expTypeDrop.send_keys("Airfare (D")
        elif cat[1] == 'h':
            # Select "Hotel (Domestic)"
            # driver.find_element_by_xpath('//*[@id="expensetype"]/option[5]').click()
            expTypeDrop.send_keys("Hotel (D")
        elif cat[1] == 'c':
            # Select "Car Rental"
            # driver.find_element_by_xpath('//*[@id="expensetype"]/option[7]').click()
            expTypeDrop.send_keys("Car Ren")
        elif cat[1] == 'm':
            # Select "Meals"
            # driver.find_element_by_xpath('//*[@id="expensetype"]/option[9]').click()
            expTypeDrop.send_keys("Meal")
        elif cat[1] == 'd':
            # Select "Mileage"
            # driver.find_element_by_xpath('//*[@id="expensetype"]/option[11]').click()
            expTypeDrop.send_keys("Mile")
        elif cat[1] == 't':
            # Select "Toll Fees"
            # driver.find_element_by_xpath('//*[@id="expensetype"]/option[12]').click()
            expTypeDrop.send_keys("Toll")
    elif cat[0] == 'o':
        if cat[1] == 't':
            # Select "Telephone"
            expTypeDrop.send_keys("Tele")
        if cat[1] == 'g':
            # Select "General"
            expTypeDrop.send_keys("Gen")
    else:
        input('Invalid category parsed, press enter to exit.')
        sys.exit()

    # Enter quantity and amount
    quantBox = driver.find_element_by_xpath('//*[@id="txtQuantity"]')
    quantBox.send_keys('1')
    amountBox = driver.find_element_by_xpath('//*[@id="txtAmount"]')
    amountBox.send_keys(exp_row[3][7:])

    # Enter comment
    descBox = driver.find_element_by_xpath('//*[@id="txtDescription"]')
    descBox.send_keys(exp_row[5])

    # add if statement to only add incurred part of comment if out of date range
    descBox.send_keys(" incurred on " + exp_row[2])

    # click 'save'
    descBox.send_keys(key.TAB)
    # builder.send_keys(key.TAB)
    builder.send_keys(key.ENTER).perform()
    time.sleep(1.5)
    # click "yes" for are you sure
    # if firstItem:
    #     builder.send_keys(key.TAB)
    # builder.send_keys(key.ENTER).perform()
    # time.sleep(1.5)
    # click 'ok' to confirm
    # oks = driver.find_elements_by_xpath('//button[text()="Ok"]')
    # if firstItem:
    #     builder.send_keys(key.TAB)
    #     firstItem = False
    # builder.send_keys(key.ENTER).perform()
    # time.sleep(2.5)

# end row for loop

input("Press enter to end script (enter)")
driver.close()
exp_file.close()