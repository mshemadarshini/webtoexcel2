
import time,datetime, zipfile, io, os, win32com.client as win32
from selenium import webdriver
from selenium.webdriver.common.by import By
# from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from docx import Document
from docx.shared import Inches
from openpyxl import Workbook

# Launch Chrome Driver
browser = webdriver.Chrome("C:\Selenium\chrome\chromedriver_win32\chromedriver.exe")
browser.implicitly_wait(50)

# Launch the Link
# time.sleep(1)

time.sleep(1)
url = 'https://apj-i.svcs.hp.com/sm/index.do?lang=en&mode=index.do'
browser.get(url)
browser.switch_to.alert.accept()
browser.maximize_window()




# Click the Incident Management Tab
# browser.fin
incident_management = browser.find_element_by_id('ext-gen-top157')
incident_management.click()

# Click the 'Search Incident' tabs
browser.set_page_load_timeout(120)
wait = WebDriverWait(browser,50)
searchIncidentClick = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ROOT/Incident Management/Search Incidents"]')))
searchIncidentClick.click()

print('==========================================')

frame = browser.find_elements_by_tag_name('iframe')
for each in frame:
    print(each)
    eachid = each.get_attribute('id')
    eachsrc = each.get_attribute('src')
    print(eachid)
    print(eachsrc)
    url = 'https://apj-i.svcs.hp.com/sm/cwc/nav.menu?name=navStart&id=ROOT%2FIncident%20Management%2FSearch%20Incidents'
    if eachsrc==url:
        print(eachid)
        newid = eachid

print("The New ID : " + newid)
browser.switch_to.frame(newid)

# === Automatically enter the key values into the text field
# ==== would need to loop for another 2 values to be entered automatically
xpath_info = browser.find_element_by_id('X78').send_keys('A-INCFLS-AFNB-SOC-MY')
time.sleep(2)

# === Select the Radio Button
select_radiobutton = browser.find_element_by_xpath('//*[@id="X55"]')
select_radiobutton.send_keys(Keys.SPACE)


# === Select the Date
wait = WebDriverWait(browser,50)
open_before = wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="X108"]')))
# open_before = browser.find_element_by_xpath('/html/body/div[1]/div[1]/form/div[4]/div[2]/div/div/div/div[2]/div/div[30]/div/div[2]/a/img')
open_before.send_keys(Keys.RETURN)

open_after = browser.find_element_by_xpath('//*[@id="X108"]')
open_after.click()

# === Switch to Parent Frame which has the 'Saeach Icon' at Top
browser.switch_to.parent_frame()
time.sleep(1)

# === Find 'Search' button at top
button = browser.find_elements_by_tag_name('button')
print("The current window is 3 : " + str(browser.current_window_handle))
for each in button:
    area_label = each.get_attribute('aria-label')
    print(area_label)
    label = 'Start this Search (Ctrl+Shift+F6)'
    if area_label == label:
        label_id =each.get_attribute('id')
        newlabelid = label_id
    else:
        print(' None Exist ')

print(newlabelid)

Seach_End = browser.find_element_by_id(newlabelid)
Seach_End.click()
print("Clicj search end done ")

printapge_iframe = browser.find_elements_by_tag_name('iframe')
print(printapge_iframe)
for each in printapge_iframe:
    printapge_iframe_id = each.get_attribute('id')
    print("The iframe id is : " + printapge_iframe_id)
    printapge_iframe_src = each.get_attribute('src')
    print("The iframe src is : " + printapge_iframe_src)

browser.switch_to.frame(1)
button2 = browser.find_elements_by_tag_name('button')
for each3 in button2:
    each4 = each3.get_attribute('aria-label')
    aria2 = 'Print Page'
    if  aria2 == each4:
        print("The print page area label is : " + each4)
        each4_id = each3.get_attribute('id')
        print("The print page area ID is : " + each4_id)
        print_page_correct_id = 'ext-gen-listdetail-1-43'
        if each4_id == print_page_correct_id:
            each5 = each4_id
print(each5)

win2 = browser.window_handles
print(win2)

a_href = browser.find_elements_by_tag_name('a')
print(a_href)

cur_win = browser.current_window_handle
print("Current Window Before Page Clicked " + cur_win)

printpage5 = browser.find_element_by_id(each5)
printpage5.click()
print("Print Page has been clicked !")

win1 = browser.window_handles

for each in win1:
    newwin = browser.switch_to.window(each)
    cur_win = browser.current_window_handle
    print("Current Window For This Switch " + cur_win)
    win21 = browser.window_handles
    print("Win 2 New : " + str(win21))


print(win1)
cur_win = browser.current_window_handle
print("Current Window After Page Clicked " + cur_win)

# http://openpyxl.readthedocs.io/en/default/tutorial.html
print("Going to Create WorkBook ")
wb = Workbook()
wb_title = wb['New Book']
w_active = wb.active
w_active.title = ' This New HEma'
ws1 = wb.create_sheet('New Sheet')
print("Work Book Done Created ")

tableheadingtext = browser.find_elements_by_class_name('TableHeadingText')
for each in tableheadingtext:
    print(each.text)


td_data_headers = browser.find_elements_by_class_name('TableCellResults ')
for each_2 in td_data_headers:
    print(each_2.text)

print("Work Done !")


# === Save webpage print into excel













# browser.switch_to.frame(2)
# button2 = browser.find_element_by_tag_name('button')
# button2_id = button2.get_attribute('id')
# print(button2_id)


