from asyncio import DefaultEventLoopPolicy
from cgitb import text
from distutils.spawn import find_executable
from logging import raiseExceptions
from re import search
from libraries.common import log_message, capture_page_screenshot, act_on_element, files, file_system, check_file_download_complete, pdf
from config import OUTPUT_FOLDER
import random, os
from selenium.webdriver.common.keys import Keys

class Gobpe():

    def __init__(self, rpa_selenium_instance, credentials:dict):
        self.browser = rpa_selenium_instance
        self.gobpe_url = credentials["url"]

    def access_gobpe(self):
        """
        Access Gob.pe from the browser
        """
        log_message("Start - Access Gobpe")
        self.browser.go_to(self.gobpe_url)
        log_message("End - Access Gobpe")

    def open_onpe_page(self):
        """
        Reach the ONPE url by moving through the site
        """
        log_message("Start - Open ONPE page")
        act_on_element('//ul[@class="list-footer"]//span[text()="El Estado Peruano"]', "click_element")
        act_on_element('//a[text()="Organismos Autónomos"]', "click_element")    
        act_on_element('//a[@href="/onpe"]', "click_element")  
        act_on_element('//div[@class="institutions-show"]', "find_element")
          
        log_message("End - Open ONPE page")

    def read_input_file(self):
        """
        Read the data from the input file received
        """
        log_message("Start - Read Input File")

        with open("Category.txt", "r", encoding="utf-8") as file:
            self.category = file.read()
            
        log_message("End - Read Input File")

    def go_to_category(self):
        """
        Click on the Category specified on the txt
        """
        log_message("Start - Go To Category")
        buttons= act_on_element('//button', "find_elements")  
        found = False
        count = 1
        while found == False and count < len(buttons):
            try:        
                act_on_element(buttons[count], "click_element")  
                count +=1
                act_on_element('//h2[text()="{}"]'.format(self.category.strip()), "find_element")
                act_on_element('//a[@data-origin="menu-onpe-todos-los-informes-y-publicaciones"]', "click_element",1)
                found = True
            
            except Exception as e:
                pass
        act_on_element('//footer[@aria-label="Pie de página"]', "find_element")

        log_message("End - Go To Category")

    def read_download_file(self):
        """
        Read the data from the input file received
        """
        log_message("Start - Read Download File")

        files.open_workbook("Files_To_Download.xlsx")
        excel_data_dict_list = files.read_worksheet(name = "Sheet1", header=True)
        files.close_workbook()
        self.excel_data_dict_list = excel_data_dict_list

        log_message("End - Read Download File")

    def download_files(self):
        """
        Downloads files from Search Bar
        """
        log_message("Start - Download Files")

        documents_to_download = []

        for excel_data_dict in self.excel_data_dict_list:
            if excel_data_dict["Download Required"].strip().upper() == "YES":
                documents_to_download.append(excel_data_dict["Name"].strip().upper())

        for document in documents_to_download:
            search_bar = act_on_element('//input[@id="filter_terms"]', "find_element")
            self.browser.input_text_when_element_is_visible('//input[@id="filter_terms"]', document)
            search_bar.send_keys(Keys.ENTER)            
            search_results = act_on_element('//li//div[child::a[@class="mb-2 block text-primary text-xl font-bold card__mock mr-10"]]', "find_elements")
            results_buttons = act_on_element('//li//a[@class="btn btn--secondary download md:w-auto"]', "find_elements")
            gobpe_file_name = document.strip()
            for result, button in zip(search_results, results_buttons):
                if gobpe_file_name == result.text.upper():
                    act_on_element(button, "click_element")
                    check_file_download_complete("pdf", 20)

        log_message("End - Download Files")     

    def read_pdf(self):
        """
        Read Pdf File
        """   
        log_message("Start - Read PDFs")

        files_downloaded = file_system.find_files("{}/*.{}".format(OUTPUT_FOLDER, "pdf"))
        files.create_workbook(path = "{}/Results.xlsx".format(OUTPUT_FOLDER))
        files.create_worksheet(name = "Documents", content= None, exist_ok = True, header = False)
        content_data = ""
        excel_data = {}
        for file_downloaded in files_downloaded:
            text_dict = pdf.get_text_from_pdf(file_downloaded)
            pages_amount = len(text_dict)
            excel_data[str(file_downloaded[1])] = pages_amount
            files.append_rows_to_worksheet(excel_data, name = "Documents", header = False, start= None)
            #files.append_rows_to_worksheet(lenght_data, name = "Documents", header = False, start= None)
            if pages_amount > 50:
                content_data = content_data + "File Name: " + str(file_downloaded[1]) + "\n-------\n"
                file_system.create_file("{}/Results.txt".format(OUTPUT_FOLDER), content=content_data, encoding = "utf-8", overwrite = False)
                file_content = file_system.read_file("{}/Results.txt".format(OUTPUT_FOLDER), encoding = "utf-8")
            log_message("Finished {}".format(str(file_downloaded[1])))
        files.remove_worksheet(name = "Sheet")
        files.save_workbook(path = None)
        files.close_workbook()
        log_message("End - Read PDFs")        