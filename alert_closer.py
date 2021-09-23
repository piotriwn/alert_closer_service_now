from selenium import webdriver
from selenium.webdriver.common import keys
from selenium.webdriver.common.keys import Keys 
from selenium.webdriver.common.by import By 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.select import Select
import re
import sys
import getpass
from datetime import datetime
from datetime import timedelta
import os

# CONSTANTS
SN_URL = "service now URL"
DRIVER_PATH = r".\chromedriver.exe"
FILE_PATH = r".\data.txt"
BORDER_LINE = "The following Monitoring-Event occured:" 
FILTER_URL = r"service now filter URL"
WORKNOTE_MESSAGE = r"Alert cleared"
SECONDS_ADDED = 10

# log variable
log = ""

# log in to Service Now
def logIn(driver, email):
    global log
    driver.get(SN_URL)
    try:
        usernameBox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "username")))
        usernameBox.send_keys(email)
    except Exception as e:
        driver.quit()
        log += e
        print(e)
        return

    nextLoginPage = driver.find_element_by_name("next")
    nextLoginPage.click()

    try:
        externalLogin = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.LINK_TEXT, "Use external login")))
        externalLogin.click()
        driver.implicitly_wait(1)

        submitButton = driver.find_element_by_name("login")
        submitButton.click()
    except TimeoutException as ex:
        pass

    return None

# extract data from .txt file containing bodies of Manane Now alert messages
def extractData():
    dataDict = {}
    i = -1
    with open (FILE_PATH, "r") as file:
        for line in file:
            if (BORDER_LINE in line):
                i += 1
                dataDict[i] = {}
            elif ("Processed in MN" in line):
                dataDict[i]["date"] = re.search("(\d{4})-(\d{2})-(\d{2})", line).group() # 1111-11-11
                dataDict[i]["time"] = re.search("(\d{2}):(\d{2}):(\d{2})", line).group() # 11:11:11
            elif ("Affected Host" in line):
                dataDict[i]["host"] = re.search("(^Affected Host:)(\s+)(\w+)", line).group(3)
            elif ("Label" in (lineReplaced := line.replace(" ", "") ) ) : # for some reason Label is sometimes randomly displayed in alerts as "Lab el" or "Labe l"
                dataDict[i]["label"] = re.search("(^Label:)(\s*)(\w.+)", lineReplaced).group(3)
            elif ("Message" in line):
                dataDict[i]["message"] = re.search("(^Message:)(\s+)(\w.+)", line).group(3)
    return dataDict

# print dict 
def printDict(dct):
    global log
    for key, value in dct.items():
        line = f"{key}\t{[ dct[key][x]  for x in value]}"
        log += line + '\n'
        print(line)
    line = "\n--------------------\n\n"
    log += line
    print(line)

# go to the page listing all incidents and sort by Opened By attribute
def goToIncPage(driver):
    driver.get(FILTER_URL)

    sortedByOpened = driver.find_elements_by_xpath('//th[@glide_field="incident.opened_at"]//i[@class="sort-icon-padding list-column-icon icon-vcr-down"]')
    while (sortedByOpened == [] ):
        openedButton = driver.find_element_by_xpath('//th[@glide_field="incident.opened_at"]//a[@class="column_head list_hdrcell table-col-header"]')
        openedButton.click()
        sortedByOpened = driver.find_elements_by_xpath('//th[@glide_field="incident.opened_at"]//i[@class="sort-icon-padding list-column-icon icon-vcr-down"]')
        driver.implicitly_wait(3)
    return None

# search for the incidents using keyword query
def searchForIncident(driver, dct, i):
    filterToggle = driver.find_element_by_xpath('//a[@id="incident_filter_toggle_image"]')
    filterToggle.click()
    checkIfKeywordPresent = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '//div[@class="filterContainer"]//span[@id="select2-chosen-2"]')))
    if ("Keywords" not in [ item.text for item in checkIfKeywordPresent ]):
        addNew = driver.find_element_by_xpath('//button[@aria-label="Add a new AND filter condition"]')
        addNew.click()
        driver.implicitly_wait(3)
        chooseOption = driver.find_element_by_xpath('//div[@class="select2-container filerTableSelect select2 form-control filter_type"]//a[@class="select2-choice"]')
        chooseOption.click()
        driver.implicitly_wait(3)
        addNewSelection = Select(driver.find_element_by_xpath("//div['select2-container filerTableSelect select2 form-control filter_type select2-container-active select2-dropdown-open']//select[@aria-label='Choose Field']"))
        addNewSelection.select_by_visible_text("Keywords")
        driver.implicitly_wait(3)
    keywordInput = driver.find_element_by_xpath('//tr[@class="filter_row_condition"]//input[@class="filerTableInput form-control"]')
    keywordInput.send_keys(dct[i]["host"])
    runButton = driver.find_element_by_xpath('//button[@id="test_filter_action_toolbar_run"]')
    runButton.click()
    return None

# establish range for seconds and minutes
def findDateTimePossibilities(dct, i):
    _date = dct[i]["date"]
    _time = dct[i]["time"]
    _datetime = datetime( int(_date[0:4]), int(_date[5:7]), int(_date[8:10]), int(_time[0:2]), int(_time[3:5]), int(_time[6:8]) )
    timePossibilites = []
    for i in range(0, SECONDS_ADDED+1):
        timePossibilites.append( _datetime + timedelta(seconds = i) )
    dateTimePossibilites = [ f"{item.day:02}/{item.month:02}/{item.year} {item.hour:02}:{item.minute:02}:{item.second:02}" for item in timePossibilites ] # f'{value:{width}.{precision}}'

    return dateTimePossibilites

# find incidents matching the criteria
def findIncident(driver, i, dct, dateTimePossibilites):
    global log
    values = ["date", "time", "host", "label", "message"]
    line = f"{i}\t{[ dct[i][x]  for x in values]}\t\t"

    incidentSearchTime = WebDriverWait(driver, 20).until(EC.presence_of_all_elements_located((By.XPATH, '//table//tr//div[@class="datex date-calendar"]')))
    
    incURLs = []
    for incTime in incidentSearchTime:
        # print("incTime.text = ", incTime.text, end=" ")
        if (incTime.text in dateTimePossibilites):
            # print("Match!", end="")
            rowElement = incTime.find_element_by_xpath('../..')
            incLink = rowElement.find_element_by_xpath('.//a[@class="linked formlink"]').get_attribute('href')
            incLink = re.search(  "(http.*?)(&sysparm)" , incLink ).group(1) # ? at the end of first group makes it non-greedy
            incURLs.append(incLink)
        # print("")

    if (incURLs == []):
        line += "No tickets found for that data.\n"
        log += line
        print(line)

    for incURL in incURLs:
        driver.get(incURL)
        driver.implicitly_wait(3) 

        ticketNumber = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'sys_readonly.incident.number'))).get_attribute("value")
        line += ticketNumber + "\t"

        ifAlreadyResolved = Select(WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'incident.state'))))
        state = ifAlreadyResolved.first_selected_option.text
        if (state == "Resolved"):
            line += "Ticket is already resolved.\n"
            
        else:
            ifActResult = checkIfAct(driver, i, dct)
            if (ifActResult):
                closed = closeTicket(driver)
                if (closed):
                    line += "Closed successfully.\n"
                else:
                    line += "Failed to close successfully.\n"
            else:
                line += "The ticket is not-actionable.\n"
        log += line
        print(line)
    driver.get(FILTER_URL)

# check if the ticket is actionable
def checkIfAct(driver, i, dct):
    global log
    shortDescription = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//input[@id="incident.short_description"]'))).get_attribute("value")
    # actually this is too strict of a requirement
    # --
    # if (dct[i]["label"] not in shortDescription ):
    #     line = "The short description of the ticket does not match the content of alert\n"
    #     line = f"Ticket\'s short description is: {shortDescription}\nAlert\'s subject is {dct[i]['label']}\n"
    #     log += line
    #     print(line, end='')
    #     return False
    description = driver.find_element_by_xpath('//textarea[@id="incident.description"]').text
    if (dct[i]["message"] not in description):
        line = "The description of the ticket does not match the content of alert\n"
        line = f"Ticket\'s description is: {description}\nAlert\'s content is {dct[i]['message']}\n"
        log += line
        print(line)        
        return False
    stateSelect =  Select(WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'incident.state'))))
    currentChoice = stateSelect.first_selected_option.text
    if (currentChoice != "Open" and currentChoice != "Acknowledged"):
        line = "The status of the ticket is not Open or Acknowledged\t"
        log += line
        print(line) 
        return False
    openedByValue = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'incident.opened_by_label'))).get_attribute("value")
    if (openedByValue != 'Monitoring'):
        line = "The caller of the ticket is not Monitoring\t"
        log += line
        print(line)
        return False
    assignedToValue = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'sys_display.incident.assigned_to'))).get_attribute("value")
    if (assignedToValue != ""):
        line = "There is already somebody assigned to the ticket\t"
        log += line
        print(line)
        return False
    return True

# close the ticket and perform necessary updates in the ticket
def closeTicket(driver):
    try:
        stateSelect =  Select(WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'incident.state'))))
        stateSelect.select_by_visible_text("Acknowledged")
        saveButton = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'sysverb_update_and_stay')))
        saveButton.click()
        notesTab = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//div[@id="tabs2_section"]//span[contains (text(), "{0}") ]'.format("Notes"))))
        driver.implicitly_wait(3) 
        notesTab.click()
        workNotes = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'activity-stream-work_notes-textarea')))
        driver.implicitly_wait(3) 
        workNotes.send_keys(WORKNOTE_MESSAGE)
        stateSelect =  Select(WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'incident.state'))))
        driver.implicitly_wait(3) 
        stateSelect.select_by_visible_text("Work in Progress")
        saveButton = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'sysverb_update_and_stay')))
        saveButton.click()
        stateSelect =  Select(WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'incident.state'))))
        driver.implicitly_wait(3) 
        stateSelect.select_by_visible_text("Resolved")
        resolveNotesTab = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//div[@id="tabs2_section"]//span[contains (text(), "{0}") ]'.format("Resolution / Closure Information"))))
        driver.implicitly_wait(3) 
        resolveNotesTab.click()
        resolveCodeSelect = Select(WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'incident.close_code'))))
        resolveCodeSelect.select_by_visible_text("No Fault Found")
        resolveComment = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'incident.close_notes')))
        driver.implicitly_wait(3) 
        resolveComment.send_keys("Alert cleared")
        saveButton = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'sysverb_update_and_stay')))
        saveButton.click()
        return True
    except:
        return False

# save logs to a file
def logToFile():
    filename =  datetime.now().strftime(r"%d_%m_%Y_%H_%M_%S") + "_" + getpass.getuser() + ".txt"
    pathname = os.path.abspath(os.path.dirname(__file__))    
    file_path = os.path.join(pathname , 'Logs' , filename)
    with open (file_path , "w") as f:
        f.write(log)

if __name__ == "__main__":
    now = datetime.now()
    dt_string = now.strftime(r"%d/%m/%Y %H:%M:%S")
    log += getpass.getuser() + '\n'
    print(getpass.getuser())
    log += dt_string + '\n\n--------------------\n\n'
    print(dt_string)

    driver = webdriver.Chrome(DRIVER_PATH)
    email = sys.argv[1]
    logIn(driver, email)
    dct = extractData()
    printDict(dct)
    goToIncPage(driver)
    for key in dct.keys():
        searchForIncident(driver, dct, key)
        dateTimePossibilites = findDateTimePossibilities(dct, key)
        findIncident(driver, key, dct, dateTimePossibilites)
        log += "\n\n"
    driver.quit()
    logToFile()

