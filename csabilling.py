from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
import xlrd
import xlwt

driver = webdriver.Firefox()
driver.get("your site")
#assert "Python" in driver.title
xpaths = { 'user' : "//input[@id='TextBox1']",
           'pass' : "//input[@id='TextBox2']",
           'submitButton' :   "//input[@name='loginBtn']",
           'bill' :   "//input[@name='Button1']",
           'billcode' : "//a[@id='Wizard1_SideBarContainer_SideBarList_ctl00_SideBarButton']",
           'billname' : "//input[@id='Wizard1_codeTxt']",
            'billnext' : "//input[@id='Wizard1_StartNavigationTemplateContainerID_StartNextButton']",
            'id' : "//input[@id='Wizard1_idnoTxt']",
           'amount' : "//input[@id='Wizard1_amountTxt']",
           'add' : "//input[@id='Wizard1_addBtn']",
           'addnext' : "//input[@id='Wizard1_StepNavigationTemplateContainerID_StepNextButton']",
           'finish' : "//input[@id='Wizard1_FinishNavigationTemplateContainerID_FinishButton']",
         }
username = ""
password = ""
billnam = "TSHIRT_2015"
amount = "220"
book = xlrd.open_workbook('your excel sheet')
first_sheet = book.sheet_by_index(0)
driver.find_element_by_xpath(xpaths['user']).clear()

#Write Username in Username TextBox
driver.find_element_by_xpath(xpaths['user']).send_keys(username)

#Clear Password TextBox if already allowed "Remember Me" 
driver.find_element_by_xpath(xpaths['pass']).clear()

#Write Password in password TextBox
driver.find_element_by_xpath(xpaths['pass']).send_keys(password)

#Click Login button
driver.find_element_by_xpath(xpaths['submitButton']).click()
driver.find_element_by_xpath(xpaths['bill']).click()
driver.find_element_by_xpath(xpaths['billcode']).click()
driver.find_element_by_xpath(xpaths['billname']).send_keys(billnam)
driver.find_element_by_xpath(xpaths['billnext']).click()
for i in range(174,221):
    cell = first_sheet.cell(i,0)
    test = str(cell.value)
    driver.find_element_by_xpath(xpaths['id']).clear()
    driver.find_element_by_xpath(xpaths['id']).send_keys(test)
    driver.find_element_by_xpath(xpaths['amount']).clear()
    driver.find_element_by_xpath(xpaths['amount']).send_keys(amount)
    driver.find_element_by_xpath(xpaths['add']).click()

