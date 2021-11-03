# +
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.FileSystem import FileSystem
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from pathlib import Path
import threading
import time
import shutil

url = "https://itdashboard.gov"

browser = Selenium()
excel_handler = Files()
file_sys = FileSystem()
list_of_agency = []
list_of_link = {}
tableData = []
default_download = str(Path.home()) + "/Downloads/"


# -

def initial_setup():
    global test_agency
    if file_sys.does_directory_exist("./output") is False:
        file_sys.create_directory("./output", parents=False, exist_ok=True)
    config = file_sys.read_file("config").split("\n")
    test_agency = config[0].split("=")[1]
    browser.open_available_browser(url, maximized=True)


def scrape_agency_list():
    browser.click_link("#home-dive-in")
    browser.wait_until_element_is_visible('xpath://*[@id="agency-tiles-widget"]')
    agencies = browser.find_elements(
         'xpath://*[@id="agency-tiles-widget"]/div/div/div/div/div/div/div/a/span[1]'
        )
    spendings = browser.find_elements(
         'xpath://*[@id="agency-tiles-widget"]/div/div/div/div/div/div/div/a/span[2]'
        )
    for index in range(0, len(agencies), 1):
        agency = agencies[index].text
        spending = spendings[index].text
        tempTuple = (agency, spending)
        list_of_agency.append(tempTuple)


def write_agencies_to_excel():
    if excel_handler.get_active_worksheet() != "Agencies":
        excel_handler.set_active_worksheet("Agencies")
    excel_handler.set_cell_value(1, 1, "Agency")
    excel_handler.set_cell_value(1, 2, "Spending")
    row_count = 2
    for data_tuple in list_of_agency:
        agency = data_tuple[0]
        spending = data_tuple[1]
        excel_handler.set_cell_value(row_count, 1, agency)
        excel_handler.set_cell_value(row_count, 2, spending)
        row_count += 1


def write_investment_to_excel():
    if excel_handler.worksheet_exists(test_agency) is False:
        excel_handler.create_worksheet(test_agency)
    if excel_handler.get_active_worksheet() != test_agency:
        excel_handler.set_active_worksheet(test_agency)
    for header in tableHeaders:
        excel_handler.set_cell_value(1, tableHeaders.index(header) + 1,
                                     header.text)
    row_count = 2
    for data_set in tableData:
        for data in data_set:
            excel_handler.set_cell_value(row_count, data_set.index(data) + 1,
                                         data)
        row_count += 1


def scrape_table_data():
    global list_of_link
    tableDataRaw = []
    browser.click_link(test_agency)
    browser.wait_for_condition("return document.readyState=='complete'")
    browser.wait_until_element_is_visible(
        'xpath://*[@id="investments-table-object"]'
        , 15)
    next_page = browser.find_element(
        'xpath://*[@id="investments-table-object_next"]'
    ).get_attribute("class").find("disabled")
    browser.find_element('xpath://*[@id="investments-table-object_last"]').click()
    table_body = browser.find_element('xpath://*[@id="investments-table-object"]/tbody')
    time.sleep(10)
    browser.find_element('xpath://*[@id="investments-table-object_first"]').click()
    while next_page == -1:
        tableElem = browser.find_elements(
            'xpath://*[@id="investments-table-object"]/tbody/tr/td'
        )
        time.sleep(0.5)
        for elem in tableElem:
            try:
                if elem.find_element(By.TAG_NAME, 'a') is not None:
                    list_of_link[elem.text] =\
                        elem.find_element(By.TAG_NAME, 'a').get_attribute('href')
            except Exception:
                pass
            retry = 0
            while retry < 5:
                try:
                    tableDataRaw.append(elem.text)
                    break
                except Exception:
                    if retry == 5:
                        print("Failed fetching element after 5 retries")
                    else:
                        print("Problem fetching element. Retrying ...")
                        time.sleep(0.25)
                        retry += 1
        next_page = browser.find_element(
            'xpath://*[@id="investments-table-object_next"]'
        ).get_attribute("class").find("disabled")
        if next_page == -1:
            browser.find_element('xpath://*[@id="investments-table-object_next"]').click()
    organize_elements(tableDataRaw)


def organize_elements(tableDataRaw):
    global tableHeaders, tableData
    tableHeaders = browser.find_elements(
        'xpath://*[@id="investments-table-object_wrapper"]/div/div/div/table/thead/tr/th'
    )
    columnCount = len(tableHeaders)
    totalData = len(tableDataRaw) / columnCount
    counter = 0
    tempList = []
    for index in range(0, int(totalData), 1):
        while len(tempList) < columnCount:
            tempList.append(tableDataRaw[counter])
            counter += 1
        if len(tempList) == columnCount:
            tableData.append(tempList)
            tempList = []
    if len(tableData) == totalData:
        print("Done organizing data")


def write_data_to_excel():
    if file_sys.does_file_exist("./output/agency_data.xlsx"):
        file_sys.remove_file("./output/agency_data.xlsx", missing_ok=True)
    excel_handler.create_workbook(path='./output/agency_data.xlsx', fmt='xlsx')
    excel_handler.rename_worksheet("Sheet", "Agencies")
    write_agencies_to_excel()
    write_investment_to_excel()
    excel_handler.save_workbook()
    excel_handler.close_workbook()


def get_pdfs_from_links():
    browser.set_download_directory(directory="./output", download_pdf=True)
    thread1 = threading.Thread(target=get_pdf, name="Download Thread")
    thread2 = threading.Thread(target=move_pdf_to_output, name="Move file thread")
    thread1.start()
    thread2.start()
    thread2.join()


def get_pdf():
    for file, link in list_of_link.items():
        browser.go_to(link)
        try:
            file_downloaded = False
            browser.wait_until_element_is_visible('link:Download Business Case PDF', 10)
            browser.click_link("Download Business Case PDF")
            while file_downloaded is not True:
                if file_sys.does_file_exist(default_download + file + ".pdf") is True:
                    file_downloaded = True
        except Exception:
            print('Cannot locate the download button for {file}'.format(file=file))
            browser.reload_page()


def move_pdf_to_output():
    for file in list_of_link.keys():
        source = default_download + file + ".pdf"
        destination = "./output/" + file + ".pdf"
        move_success = False
        while move_success is False:
            try:
                shutil.move(source, destination)
                move_success = True
            except FileNotFoundError as fe:
                time.sleep(3)


def wrap_and_clean_up():
    browser.close_all_browsers()


if __name__ == "__main__":
    try:
        initial_setup()
        scrape_agency_list()
        scrape_table_data()
        write_data_to_excel()
        get_pdfs_from_links()
        wrap_and_clean_up()
    finally:
        print("Finished Task")
