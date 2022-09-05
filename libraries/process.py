from libraries.common import log_message, capture_page_screenshot, browser
from config import OUTPUT_FOLDER, tabs_dict
from libraries.gobpe.gobpe import Gobpe

class Process():
    
    def __init__(self, credentials: dict):
        log_message("Initialization")

        prefs = {
            "profile.default_content_setting_values.notifications": 2,
            "profile.default_content_setting_popups": 0,
            "directory_upgrade": True,
            "download.default_directory": OUTPUT_FOLDER,
            "plugins.always_open_pdf_externally": True,
            "download.prompt_for_download": False
        }

        browser.open_available_browser(preferences = prefs)
        browser.set_window_size(1920, 1080)
        browser.maximize_browser_window()

        gobpe = Gobpe(browser, {"url": "https://www.gob.pe/"})
        tabs_dict["Gobpe"] = len(tabs_dict)
        gobpe.access_gobpe()
        self.gobpe = gobpe

    def start(self):
        """
        main
        """
        self.gobpe.open_onpe_page()
        self.gobpe.read_input_file()
        self.gobpe.go_to_category()
        self.gobpe.read_download_file()
        self.gobpe.download_files()
        self.gobpe.read_pdf()

        pass
    
    def finish(self):
        log_message("DW Process Finished")
        browser.close_browser()