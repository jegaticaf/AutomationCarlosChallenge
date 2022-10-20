from libraries.common import log_message, files, file_system, check_file_download_complete, pdf
from config import OUTPUT_FOLDER
import os
from selenium.webdriver.common.keys import Keys
import SeleniumLibrary.errors

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
        self.browser.click_element('//ul[@class="list-footer"]//span[text()="El Estado Peruano"]')
        self.browser.click_element('//a[text()="Organismos Autónomos"]')
        self.browser.click_element('//a[@href="/onpe"]')  
        self.browser.find_element('//div[@class="institutions-show"]')
          
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
        buttons= self.browser.find_elements('//button')  
        found = False
        count = 1
        while found == False and count < len(buttons):
            try:        
                self.browser.click_element(buttons[count])  
                count +=1
                self.browser.find_element('//h2[text()="{}"]'.format(self.category.strip()))
                self.browser.click_element('//a[@data-origin="menu-onpe-todos-los-informes-y-publicaciones"]')
                found = True
            
            except SeleniumLibrary.errors.ElementNotFound as ex:
                log_message(str(ex)) 
                raise Exception(str(ex)) 
        self.browser.find_element('//footer[@aria-label="Pie de página"]')

        log_message("End - Go To Category")

    def read_download_file(self):
        """
        Read the data from the input file received
        """
        log_message("Start - Read Download File")

        place = os.environ.get("Environment", "Locally")

        if place == "Robocloud":
            files.open_workbook("Files_To_Download.xlsx")
            excel_data_dict_list = files.read_worksheet(name = "Robocloud", header=True)
            files.close_workbook()
            self.excel_data_dict_list = excel_data_dict_list
        else:
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
            search_bar = self.browser.find_element('//input[@id="filter_terms"]')
            self.browser.input_text_when_element_is_visible('//input[@id="filter_terms"]', document)
            search_bar.send_keys(Keys.ENTER)            
            search_results = self.browser.find_elements('//li//div[child::a[@class="mb-2 block text-primary text-xl font-bold card__mock mr-10"]]')
            results_buttons = self.browser.find_elements('//li//a[@class="btn btn--secondary download md:w-auto"]')
            gobpe_file_name = document.strip()
            for result, button in zip(search_results, results_buttons):
                if gobpe_file_name == result.text.upper():
                    self.browser.click_element(button)
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
        excel_data = []
        for file_downloaded in files_downloaded:
            text_dict = pdf.get_text_from_pdf(file_downloaded)
            pages_amount = len(text_dict)
            excel_data.append({"File Name": str(file_downloaded[1]), "Amount of Pages": pages_amount})
            #files.append_rows_to_worksheet(lenght_data, name = "Documents", header = False, start= None)
            if pages_amount > 50:
                content_data = content_data + "File Name: " + str(file_downloaded[1]) + "\n-------\n"
                file_system.create_file("{}/Results.txt".format(OUTPUT_FOLDER), content=content_data, encoding = "utf-8", overwrite = True)
            log_message("Finished {}".format(str(file_downloaded[1])))
        files.append_rows_to_worksheet(excel_data, name = "Documents", header = True, start= None)
        files.remove_worksheet(name = "Sheet")
        files.save_workbook(path = None)
        files.close_workbook()
        log_message("End - Read PDFs")        