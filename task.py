"""Template robot with Python."""
import time
from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF
import os


browser = Selenium()
http = HTTP()
file = Files()
pdf = PDF()
table = []

"""__________Tasks___________"""
"""___For Log In ___"""
"""___Download___"""
"""___Open & Read WorkBook___"""
"""___For Submit___"""
"""___Collect Results___"""
"""___PDF File___"""

"""___For Log In___"""
browser.open_available_browser("https://robotsparebinindustries.com/", maximized=False)
browser.input_text("//*[@id='username']", "maria")
browser.input_text('//*[@id="password"]', "thoushallnotpass")
browser.wait_and_click_button('//*[@id="root"]/div/div/div/div[1]/form/button')

"""___Download___"""
http.download(url="https://robotsparebinindustries.com/SalesData.xlsx",target_file=f"{os.getcwd()}/test.xlsx", overwrite=True)

"""___Open & Read WorkBook___"""
file.open_workbook("test.xlsx")
data = file.read_worksheet_as_table(header=True)
file.close_workbook()
for x in data:
    table.append(x)

"""___For Submit___"""
for data in table:
    # browser.open_available_browser("https://robotsparebinindustries.com/", maximized=False)
    browser.input_text('//*[@id="firstname"]', f'{data["First Name"]}')
    browser.input_text('//*[@id="lastname"]', f'{data["Last Name"]}')
    browser.select_from_list_by_value('//*[@id="salestarget"]', f'{data["Sales Target"]}')
    browser.input_text('//*[@id="salesresult"]', f'{data["Sales"]}')
    browser.wait_and_click_button('//*[@id="sales-form"]/button')

"""___Collect Results___"""
browser.screenshot(locator='//*[@id="root"]/div/div/div/div[2]/div[1]')

"""___PDF File___"""
browser.wait_until_element_is_visible('//*[@id="sales-results"]')
filework = browser.get_element_attribute(locator='//*[@id="sales-results"]', attribute='outerHTML')
pdf.html_to_pdf(filework, output_path=f'{os.getcwd()}/output/test.pdf')
"""___Log Out___"""
browser.wait_and_click_button('//*[@id="logout"]')
browser.close_browser()
print("Done")
