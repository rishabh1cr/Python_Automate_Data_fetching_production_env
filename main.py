import openpyxl
import xlwings as xw
import openpyxl

from selenium import webdriver
from time import sleep
from selenium.webdriver.support.ui import Select
username = "************"
password = "********"
mots = "*****"
url = "****"

driver =webdriver.Chrome("C:\chromedriver.exe")

driver.get(url)

driver.find_element_by_id("GloATTUID").send_keys(username)
driver.find_element_by_id("GloPassword").send_keys(password)
driver.find_element_by_id("GloPasswordSubmit").click()
driver.find_element_by_id("successButtonId").click()

print("Logged in Successfully")

sleep(5)

driver.get("https://cana.web.com/?type=cr-list-by-mots-id")

print("Mots id page loaded")

sleep(10)

driver.find_element_by_id("successButtonId").click()

sleep(5)

driver.find_element_by_class_name("rbt-input-main").send_keys(mots)

sleep(10)
driver.find_element_by_class_name("dropdown-item").click()

sleep(50)


driver.find_element_by_xpath('//button[text()="Export CSV"]').click()

print("Clicked successfully")
#driver.find_element_by_xpath('//*[@id="resize-1"]/div[1]/div/div[6]/div/div[2]/div/div/div/div[1]/div[2]').click()
#drp=Select(element)
#drp.select_by_value('In range')
sleep(50)
driver.find_element_by_xpath('//button[text()="Export CSV"]').click()

print("File downloaded successfully")

sleep(5)

filename =r"C:\Users\rishabh\Downloads\export.csv"

CRexcel = xw.Book(filename)

#CRexcel.sheets[0].api.Range("A1:J5685").AutoFilter(Field:= 6,Criteria1:="Approved")


CRexcel.sheets[0].api.Range("A1:J5685").AutoFilter(6,["Approved","Closed"],Operator:=7)

CRexcel.save('newCana.xlsx')

CRnew = openpyxl.load_workbook('newCana.xlsx')
CRsheet = CRnew['export']

CRsheet.delete_cols(idx=2, amount = 1)
CRsheet.delete_cols(idx=5, amount = 1)
CRsheet.delete_cols(idx=5, amount = 1)
CRsheet.delete_cols(idx=5, amount = 1)
CRsheet.delete_cols(idx=5, amount = 1)

CRnew.save('finalCRsheet.xlsx')




print("CR List Generated Successfully")
openpyxl.load_workbook('finalCRsheet.xlsx')
