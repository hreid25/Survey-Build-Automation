import openpyxl
# import time
# from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
# from selenium.common.exceptions import NoSuchElementException
# from selenium.common.exceptions import StaleElementReferenceException
# from selenium.common.exceptions import ElementClickInterceptedException
# from selenium.webdriver.chrome.options import Options

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

arr = []
for r in range(16, 37):
    words = []
    for c in range(3, 100, 2):
        if surveyinv.cell(row=r, column=c).value is not None:
            words.append(surveyinv.cell(row=r, column=c).value)
    arr.append(words)
emailinvitationarray = [x for x in arr if x != []]
totallanguagecount = len(emailinvitationarray[0])
# print(questionarr[5][2])


def addcustomquestions():
    questiondrivernames = WebDriverWait(driver, 20).until(
        EC.presence_of_all_elements_located((By.XPATH, "//*[@id='survey_pages_attributes_0_page_questions_attributes_0_title']/following::h4[not(@*)]")))
    # You may have to 'click' the question text area first and then call send_keys for the question text
    for i, excelrowlistobject in enumerate(questionarr):
        if excelrowlistobject[1] is None and excelrowlistobject[2] is not None:
            # OR condition to check for Y/N toggle for custom question where ID is present
            if excelrowlistobject[2] == "Empty Slot":
                continue
            else:
                # print("***************************************************")
                # print("Driver Name: ", excelrowlistobject[0])
                # print("Row Number: ", i, " Question Text: ", excelrowlistobject[2])
                # print("***************************************************")
                for drivername in questiondrivernames.text:
                    if excelrowlistobject[0] == drivername:
                        # Click the Add Question Button following the driver name match
                        WebDriverWait(driver, 20).until(EC.visibility_of_element_located(
                            (By.XPATH, f"//h4[not(@*) and contains(text(),{excelrowlistobject[0]})]/following::button[position()=1]"))).click()
                        # AddQuestionTextArea
                        WebDriverWait(driver, 20).until(
                            EC.element_to_be_clickable((By.XPATH, "//form[@id='add-custom-question-form']/div[@class='modal-body']/child::div[@class='fields']/div/div/select[@class='grouped_select optional selectized']/option"))).send_keys(excelrowlistobject[2])
                        # Click Save
                        WebDriverWait(driver, 20).until(
                            EC.element_to_be_clickable((By.XPATH, "//form[@id='add-custom-question-form']/div[@class='modal-footer']/input"))).click()
                        # May need to time.sleep() here and wait for the ID to get pulled
                        # Grab the Question ID of the new question and add it to questionarr[questionnum][1]
                        newcustomquestionid = WebDriverWait(driver, 20).until(
                            EC.element_to_be_clickable((By.XPATH, f"//*[starts-with(@class='question-text-area') and contains(text(),{excelrowlistobject[2]})]/following::strong[position()=2]")))
                        # Match the text from excelrowlistobject to questionarr[counter][2] (the QIL question text) and then insert the Question ID to prev column
                        for counter, qilrowlistobject in enumerate(questionarr):
                            if excelrowlistobject[2] == questionarr[counter][2]:
                                questionarr[counter][1] = int(newcustomquestionid)
        print(questionarr[i][1])
# Example below of matching the newly saved question and grabbing that new Question's ID:
# "//*[starts-with(@class,'question-text-area') and contains(text(),'My friends outside work would describe me as having a very positive attitude.')]/following::strong[position()=2]"


# pass the name of the driver back to the web element object to be found and clicked
# f//h4[not(@*)contains(text()='{excelrowlistobject[0]}')]
# f"//h4[not(@*) and contains(text(),{excelrowlistobject[0]})]/following::button[position()=1]")))

# need to return the match position in the QIL
# print(addquestion[2])
# print(questionstobeadded)
# print(drivername)

# print(questionarr[i][1])

# list = [i for x in questionarr[i][rownum] if x is not None]
# print(list)
# def addquestions():
#     if questionarr[questionnum][1] is None and questionarr[questionnum][languagenumber] is not None:
#         drivername = [x for x in questionarr[questionnum][0]]
#         addquestionlist = [i for question in questionarr if question[0] is not None]
