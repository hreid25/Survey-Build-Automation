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


def letsaddquestions():
    questiondrivernames = WebDriverWait(driver, 20).until(
        EC.presence_of_all_elements_located((By.XPATH, "//*[@id='survey_pages_attributes_0_page_questions_attributes_0_title']/following::h4[not(@*)]")))
    # You may have to 'click' the question text area first and then call send_keys for the question text
    for i, excelrowlistobject in enumerate(questionarr):
        if excelrowlistobject[1] is None and excelrowlistobject[2] is not None:
            if excelrowlistobject[2] == "Empty Slot":
                continue
            else:
                # print("***************************************************")
                # print("Driver Name: ", excelrowlistobject[0])
                # print("Row Number: ", i, " Question Text: ", excelrowlistobject[2])
                # print("***************************************************")
                for drivername in questiondrivernames.text:
                    if excelrowlistobject[0] == drivername:
                        # Click the following Add Question Button
                        WebDriverWait(driver, 20).until(EC.visibility_of_element_located(
                            (By.XPATH, f"//h4[not(@*) and contains(text(),{excelrowlistobject[0]})]/following::button[position()=1]"))).click()
                        # AddQuestionTextArea
                        WebDriverWait(driver, 20).until(
                            EC.element_to_be_clickable((By.XPATH, "//form[@id='add-custom-question-form']/div[@class='modal-body']/child::div[@class='fields']/div/div/select[@class='grouped_select optional selectized']/option"))).send_keys(excelrowlistobject[2])
                        # Click Save
                        WebDriverWait(driver, 20).until(
                            EC.element_to_be_clickable((By.XPATH, "//form[@id='add-custom-question-form']/div[@class='modal-footer']/input"))).click()

                        # match excelrowlistobject[2] text to newly added question and obtain the question id, take that value
                        # and insert it into the questionarr position relative to that text.
                        # //*[starts-with(@class,'switch-container')]/div[3]/span/strong

# f"//*[starts-with(@class='question-text-area') and contains(text(),{excelrowlistobject[2]})]/following::strong[position()=1]"
# "//*[starts-with(@class='question-text-area') and contains(text(),'My friends outside work would describe me as having a very positive attitude.')]"
# need to update the questions ID and add it to the array, that way future iterations are updating those newly added questions


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
