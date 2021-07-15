import openpyxl
import time
# import logging
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.webdriver.chrome.options import Options
import re

# ***********************Start timer, load workbook, ask input********************************************************
# ***********************Start timer, load workbook, ask input********************************************************
# ***********************Start timer, load workbook, ask input********************************************************

start_time = time.time()
wb = openpyxl.load_workbook('QIL Document_V2_20210518_2.xlsm')
# wb = openpyxl.load_workbook(input("Please enter the name of your QIL Document: "))
surveyhovers = wb['5- Hovers (Optional)']
surveyquest = wb['4- Survey Questions']
surveyinv = wb['2- Survey Invitation']

# ***********************CREATE MULTILINGUAL QIL ARRAY********************************************************
# ***********************CREATE MULTILINGUAL QIL ARRAY********************************************************
# ***********************CREATE MULTILINGUAL QIL ARRAY********************************************************

qarr = []
surveyquest.delete_cols(8)
for r in range(24, 177):
    questions = []
    for c in range(3, 101, 2):
        questions.append(surveyquest.cell(row=r, column=c).value)
    qarr.append(questions)
questionarr = [x for x in qarr if x != []]

for i, category in enumerate(questionarr):
    if category[0] is None:
        replacestring = str(questionarr[i - 1][0])
        questionarr[i][0] = replacestring

# ***********************CREATE HOVER ARRAY*****************************************************************************************
# ***********************CREATE HOVER ARRAY*****************************************************************************************
# ***********************CREATE HOVER ARRAY*****************************************************************************************

hover_words_array = []
hover_texts_array = []
for r in range(4, 100):
    hover_words = []
    hover_texts = []
    for hwc in range(4, 100, 3):
        hword = surveyhovers.cell(row=r, column=hwc).value
        hover_words.append(hword)
    for htc in range(5, 101, 3):
        htext = surveyhovers.cell(row=r, column=htc).value
        hover_texts.append(htext)
# if cell is not None and not cell.startswith('='): If a given row has hovers in some but not all languages,
# the cell is not none argument can cause the rows to not be aligned
    if hover_words != [] and hover_texts != []:
        hover_words_array.append(hover_words)
        hover_texts_array.append(hover_texts)
# hoverarray = [x for x in fhovers if x != []]

# ***********************REPLACE WORD WITH HOVER*******************************************************************************
# ***********************REPLACE WORD WITH HOVER*******************************************************************************
# ***********************REPLACE WORD WITH HOVER*******************************************************************************

# hoverlanguage count actually includes all of the None type objects read in using
# openpyxl (approx 32), even though there may be fewer languages
# TO DO: find the last column of data and stop there (both in the question array and hover array)

hoverlanguagecount = len((hover_words_array[0]))
numtotalhovers = len(hover_words_array) - 1
numtotalquestions = len(questionarr)

langnum = 1
hoverlangnum = 0

while langnum < hoverlanguagecount:
    hovernum = 1
    while hovernum < numtotalhovers:
        word = hover_words_array[hovernum][hoverlangnum]
        text = hover_texts_array[hovernum][hoverlangnum]
        if word is not None and text is not None:
            questionnum = 1
            while questionnum < numtotalquestions:
                # + 1 is to account for the question id column
                quest = questionarr[questionnum][langnum + 1]
                if quest is not None:
                    if word.lower() in quest.lower():
                        if quest.lower().startswith(word.lower()) is True:
                            propercase = word.capitalize()
                            replacehover = "{{" + "\"" + \
                                str(propercase) + " (" + str(text) + ")\" |hover}} "
                            pattern = re.compile('\\b' + word + '\\s', re.IGNORECASE)
                            questionarr[questionnum][langnum + 1] = pattern.sub(replacehover, quest)
                        else:
                            normalcase = word.lower()
                            replacehover = "{{" + "\"" + \
                                str(normalcase) + " (" + str(text) + ")\" |hover}} "
                            pattern = re.compile('\\b' + word + '\\s', re.IGNORECASE)
                            questionarr[questionnum][langnum + 1] = pattern.sub(replacehover, quest)
                questionnum += 1
        hovernum += 1
    langnum += 1
    hoverlangnum += 1


# *************************READ IN SURVEY INVITATION EMAIL**************************************************************
# *************************READ IN SURVEY INVITATION EMAIL**************************************************************
# This section also used to determine total amount of languages in a given QIL.

arr = []
for r in range(16, 37):
    words = []
    for c in range(3, 100, 2):
        if surveyinv.cell(row=r, column=c).value is not None:
            words.append(surveyinv.cell(row=r, column=c).value)
    arr.append(words)
emailinvitationarray = [x for x in arr if x != []]
totallanguagecount = len(emailinvitationarray[0])

# ****************************************************CREATE CHROME DRIVER************************************************************
# ****************************************************CREATE CHROME DRIVER************************************************************
# ****************************************************CREATE CHROME DRIVER*************************************************************

chrome_options = Options()
chrome_options.add_argument(
    "user-data-dir=C:\\Users\\haydr\\AppData\\Local\\Google\\Chrome\\User Data")
driver = webdriver.Chrome(chrome_options=chrome_options)
driver.get('https://www.infotech.com/kip')
driver.get('https://surveys.mcleanco.com/')

linkElem = driver.find_element_by_link_text('McLean & Company')
linkElem.click()
linkElemEng = driver.find_element_by_link_text('Engagement')
linkElemEng.click()

# surveyname = input("Enter the name of your survey as it appears in Sergeant: ")
surveyname = "Survey Automation Testing"

# ***********************FIND SURVEY NAME / START EDITING QUESTIONS********************************************************
# ***********************FIND SURVEY NAME / START EDITING QUESTIONS********************************************************
# ***********************FIND SURVEY NAME / START EDITING QUESTIONS********************************************************


searchforsurvey = driver.find_element_by_xpath(
    "//*[@id='q_translations_name_or_reseller_name_or_user_company_name_cont']")
searchforsurvey.send_keys(surveyname)
# click the search button after sending surveyname
driver.find_element_by_xpath('//*[@id="survey_search"]/input[4]').click()
findsurveys = driver.find_element_by_xpath("//tbody//a[text()='" + surveyname + "']").click()
driver.find_element_by_link_text("Edit Survey").click()
editquestions = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.LINK_TEXT, 'Questions'))).click()


# *************************************CUSTOM FUNCTIONS***************************************************************************
# *************************************CUSTOM FUNCTIONS************************************************************************************************************
# *************************************CUSTOM FUNCTIONS***************************************************************************

# //*[@id='survey-edit-questions']/child::div[3]/div/ul/child::li[@class='next']/a (Next button at top of page)
#  //*[@id='survey-edit-questions']/child::div[4][@class='row']/div/div/ul/li[@class='next']/a (Next button at bottom of page)
def clicknext():
    try:
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='survey-edit-questions']/child::div[3]/div/ul/child::li[@class='next']/a"))).click()
    except ElementClickInterceptedException or StaleElementReferenceException or NoSuchElementException as c:
        WebDriverWait(driver, 20).until(EC.invisibility_of_element(
            (By.XPATH, "//sergeant-uploads1.s3.amazonaws.com/sergeant/brands/production/2/hr-logo.svg?1499654030")))
        driver.execute_script("arguments[0].click();", WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='survey-edit-questions']/child::div[3]/div/ul/child::li[@class='next']/a"))))
        print(str(c))


def questions_returntopageone():
    try:
        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='survey-edit-questions']/div[3]/div/ul/li[1]/a"))).click()
    except ElementClickInterceptedException or StaleElementReferenceException or NoSuchElementException as q:
        WebDriverWait(driver, 20).until(EC.invisibility_of_element(
            (By.XPATH, "//sergeant-uploads1.s3.amazonaws.com/sergeant/brands/production/2/hr-logo.svg?1499654030")))
        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='survey-edit-questions']/div[3]/div/ul/li[1]/a"))).click()
        print(str(q))

# McLean Logo typically intercepts the click when we go to save, return to page one and occassionally when we click next
# depending on where we are on the page.
# selenium.common.exceptions.ElementClickInterceptedException: Message: element click intercepted:
# Element <a data-remote="true" href="/engagement/surveys/39515/edit?tab=questions">...</a> is not clickable at point
#  (368, 15). Other element would receive the click:
#  <img src="//sergeant-uploads1.s3.amazonaws.com/sergeant/brands/production/2/hr-logo.svg?1499654030" alt="Hr logo">


def savepage():
    try:
        WebDriverWait(driver, 60).until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='survey-edit-questions']/child::div[5]/input[@type='submit']"))).click()
    except ElementClickInterceptedException or StaleElementReferenceException or NoSuchElementException as s:
        print(str(s))
        driver.execute_script("arguments[0].click();", WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='survey-edit-questions']/child::div[5]/input[@type='submit']"))))


def changelanguage():
    try:
        for language in emailinvitationarray[0]:
            if language == emailinvitationarray[0][languagedropdownposition]:
                # WebDriverWait(driver, 20).until(
                #     EC.visibility_of_element_located((By.XPATH, '//*[@id="language"]')))
                languagedropdown = Select(driver.find_element_by_xpath('//*[@id="language"]'))
                languagedropdown.select_by_visible_text(language)
    except Exception as changefail:
        print("unable to select the language dropdown option: ", changefail)


def clickswitch():
    WebDriverWait(driver, 60).until(EC.visibility_of_element_located((
        By.XPATH, "//*[@class='tab-pane active']/div[@class='row']/div/form/child::input[@type='submit']"))).click()


def logginginfodelete():
    print("----------------------------------------------------------------")
    print("Question ID: " + str(x) + " at QIL array position: " +
          str(QILarrayenumerated) + " is None and was toggled to 'Delete'")
    print("Question Counter = " + str(questionid_count), " Sergeant Toggle Pos = " +
          str(sergeantidpos) + " Total Delete Buttons: " + str(len(toggledeletelist)))
    print("Page Chg Count: ", pagechangecount)
    print("----------------------------------------------------------------")


def logginginfoedit():
    print("----------------------------------------------------------------")
    print("Question ID: " + str(x) + " at QIL array position: " +
          str(QILarrayenumerated) + " was edited")
    print("Question Counter = " + str(questionid_count), " Sergeant ID Pos = " +
          str(sergeantidpos))
    print("Page Chg Count: ", pagechangecount, 'Language: ',
          emailinvitationarray[0][languagedropdownposition])
    print("----------------------------------------------------------------")

# ***********************REMOVE SENIOR MANAGEMENT QUESTIONS GROUPING*****************************************************************************************
# ***********************REMOVE SENIOR MANAGEMENT QUESTIONS GROUPING*****************************************************************************************
# ***********************REMOVE SENIOR MANAGEMENT QUESTIONS GROUPING*****************************************************************************************
# Check for the presence of senior management relationships question groups and delete them from page 2.
# Required in order to equalize the length of the delete toggles list with that of id and text.


start_time_measure_seniormgt_deletion = time.time()
clicknext()
sergeantlistdeletetoggleposition = 0
# secondarytextlist = []
scanningforquestiongroups = True
while scanningforquestiongroups is True:
    try:
        questiongroupsobjects = WebDriverWait(driver, 10).until(EC.visibility_of_element_located(
            (By.XPATH, "//h6[contains(text(),'Question Group Title')]")))
        findquestiongroups = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, "//h6[contains(text(),'Question Group Title')]")))
        seniormgmtdeletetoggles = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located(
            (By.XPATH, "//h6[contains(text(),'Question Group Title')]/following::div[@class='row'][3]/div/div")))
        questiongroupspresent = questiongroupsobjects.is_displayed()
        while questiongroupspresent is True:
            secondarytextlist = [x.text for x in findquestiongroups]
            for i in secondarytextlist:
                if i == "Question Group Title":
                    seniormgmtdeletetoggles[sergeantlistdeletetoggleposition].click()
                    sergeantlistdeletetoggleposition += 1
                    print(i)
            questiongroupspresent = False
            scanningforquestiongroups = False
        savepage()
        questions_returntopageone()
        break
    except Exception:
        scanningforquestiongroups = False
        questiongroupspresent = False
        print("No question group objects found")
        questions_returntopageone()
        break
print("--- %s seconds ---" % (time.time() - start_time_measure_seniormgt_deletion))

# ***********************COMPARE QUESTION IDS IN SERGEANT TO MULTI-QIL*****************************************************************************************
# ***********************COMPARE QUESTION IDS IN SERGEANT TO MULTI-QIL*****************************************************************************************
# ***********************COMPARE QUESTION IDS IN SERGEANT TO MULTI-QIL*****************************************************************************************
# This code block is responsible for:
# Sending text from the multilingual QIL to Sergeant Question text area fields.
# Deleting Questions (Via toggledelete())
# Adding new questions to Sergeant
# Multiple time.sleep()'s have been introduced to allow for the DOM/page to load properly. Although inefficient,
# it is difficult to improve the XPATH's for these objects as unique attribute identifiers are not used in the XML DOM.
# Also responsible for deleting questions where multilingual QIL cells are of NoneType.


# TO DO: Likely going to have to delete questions first based on how questions are being matched to the QIL from Sergeant
# TO DO: After the code runs it should print out operations to a log file for review by the Project Coordinator for QA
time.sleep(1)
condition = True
questionid_count = 0
pagechangecount = 0
sergeantidpos = 0
languagenumber = 2
languagedropdownposition = 0

while condition is True:
    # This checks to see if we have finished processing a language
    # totallanguagecount - 1 accounts for the fact we start on the English language page
    if pagechangecount == 2 and languagedropdownposition < totallanguagecount - 1:
        questionid_count = 0
        sergeantidpos = 0
        pagechangecount = 0
        languagenumber += 1
        languagedropdownposition += 1
        changelanguage()
        clickswitch()
        time.sleep(2)
        print("Language changed to: ",
              emailinvitationarray[0][languagedropdownposition])
    # This checks that all languages in the dropdown have been processed
    elif pagechangecount == 2 and languagedropdownposition >= totallanguagecount - 1:
        condition = False
        questions_returntopageone()
        print('Done!')
        break
    while questionid_count <= len(questionarr):
        secondaryidlist = []
        toggledeletelist = WebDriverWait(driver, 20).until(EC.presence_of_all_elements_located(
            (By.XPATH, "//*[starts-with(@class,'move-content sortable-disabled')]/div[@class='row']/div/div")))
        sergeantquestionidlist = WebDriverWait(driver, 20).until(EC.presence_of_all_elements_located(
            (By.XPATH, "//*[starts-with(@class,'switch-container')]/div[3]/span/strong")))
        sergeantquestiontextlist = WebDriverWait(driver, 20).until(EC.presence_of_all_elements_located(
            (By.XPATH, "//*[starts-with(@class,'switch-container')]/div[3]/span/strong/ancestor::div/textarea[starts-with(@class,'question-text-area sortable-disabled question')]")))
        # Creating a secondary list of id's to improve speed when we go to match against the ID in the QIL
        for questionid in sergeantquestionidlist:
            secondaryidlist.append(questionid.text)
        for x in secondaryidlist:
            QILarrayenumerated = [(index, row.index(int(x)))
                                  for index, row in enumerate(questionarr) if int(x) in row]
            QILquestionposition = QILarrayenumerated[0][0]
            # + 1 is to account for Question ID column, since languages are counted from survey invitation page
            if questionarr[QILquestionposition][languagenumber] is not None:
                replacetext = questionarr[QILquestionposition][languagenumber]
                sergeantquestiontextlist[sergeantidpos].clear()
                sergeantquestiontextlist[sergeantidpos].send_keys(replacetext)
                questionid_count += 1
                sergeantidpos += 1
                logginginfoedit()
                # This checks to see if we are on first page, then saves and clicks next.
                if int(sergeantquestionidlist[0].text) == int(16595):
                    savepage()
                    clicknext()
                    pagechangecount += 1
                    questionid_count = 0
                    sergeantidpos = 0
                    time.sleep(8)
                # This checks if we have edited the last question of page 2, then saves and clicks next
                elif questionid_count == len(sergeantquestionidlist) and pagechangecount < 2:
                    savepage()
                    clicknext()  # Introduce loop here to start adding questions?
                    pagechangecount += 1
                    questionid_count = 0
                    sergeantidpos = 0
                    time.sleep(3)
                # This checks to see if we finished editing questions on page (3), saves and returns to first
                elif questionid_count == len(sergeantquestionidlist) and pagechangecount == 2:
                    questionid_count = 1000000000000
                    savepage()
                    questions_returntopageone()
                    time.sleep(2)
                    break
            # This checks to ensure we are only deleting questions once and that the cell value object type is None.
            elif languagenumber == 1 and questionarr[QILquestionposition][languagenumber] is None:
                toggledeletelist[sergeantidpos].click()
                questionid_count += 1
                sergeantidpos += 1
                # This checks to see if we are deleting the last question on page 3 then returns us to first page
                # and breaks out of the loop altogether
                if questionid_count == len(sergeantquestionidlist) and pagechangecount == 2:
                    savepage()
                    questions_returntopageone()  # Introduce loop here to start adding questions?
                    logginginfodelete()
                    time.sleep(2)
                    questionid_count = 1000000000000
                    break
                # Checks to see if we are deleting the last question of page 2 and then saves and clicks next
                elif questionid_count == len(sergeantquestionidlist) and pagechangecount == 1:
                    savepage()
                    clicknext()  # Introduce loop here to start adding questions?
                    time.sleep(2)
                    pagechangecount += 1
                    questionid_count = 0
                    sergeantidpos = 0
                logginginfodelete()
print("--- %s seconds ---" % (time.time() - start_time))

# https://stackoverflow.com/questions/27175400/how-to-find-the-index-of-a-value-in-2d-array-in-python
# https://stackoverflow.com/questions/11223011/attributeerror-list-object-has-no-attribute-click-selenium-webdriver
# https://stackoverflow.com/questions/17385419/find-indices-of-a-value-in-2d-matrix/17388914

# ****************************************************ADD QUESTIONS************************************************************
# ****************************************************ADD QUESTIONS************************************************************
# ****************************************************ADD QUESTIONS************************************************************
# All questions have an element inside the question text area field with a Boolean value True or False
# Will need to rebuild the list on each iteration to include the questions from before in order to edit the question in the next language
# the script should automatically rebuild the web elements object list on future languages, but the QIL will not be able to reference
# the corresponding ID number since this is created afterwards.
# Custom question translation will need separate logic to handle how to check for matches without question ID.
# May have to write the new question id from created question back to multilingual QIL to then properly edit over them on future loops.

# for i, category in enumerate(questionarr):
#     if category[0] is None:
#         replacestring = str(questionarr[i - 1][0])
#         questionarr[i][0] = replacestring

# find driver name
#
# Driver name field
# //fieldset/descendent::h4
# Add Question Button
# //div[@class='span12']/ul/div/li/button
# Add Question text area (after fade in (XPATH))
# //form[@id='add-custom-question-form']/div[@class='modal-body']/child::div[@class='fields']/div/div/select[@class='grouped_select optional selectized']
# Add Question Save Button (after fade in)
# //form[@id='add-custom-question-form']/div[@class='modal-footer']/input
