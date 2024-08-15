import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from RPA.Browser.Selenium import Selenium
from SeleniumLibrary.base import keyword


class ExtendedSelenium(Selenium):
    """
    ExtendedSelenium is a subclass of Selenium that includes custom functionality for specific needs,
    such as setting up the Chrome browser with particular options and custom keywords.
    """

    def __init__(self, *args, **kwargs):
        """
        Initializes the ExtendedSelenium class with a ChromeDriver installation.
        """
        super().__init__(*args, **kwargs)
        self.driver_path = ChromeDriverManager().install()

    @keyword
    def open_site(self, url, **kwargs):
        """
        Opens a site using Chrome with custom download and privacy settings.

        Args:
            url (str): The URL of the site to open.
            **kwargs: Additional arguments to pass to the `open_browser` method.
        """
        download_dir = os.path.abspath("output/images")

        chrome_options = Options()
        chrome_prefs = {
            "profile.default_content_settings.popups": 0,
            "download.default_directory": download_dir,
            "directory_upgrade": True,
            "safebrowsing.enabled": True,
            "download.prompt_for_download": False,
            "plugins.always_open_pdf_externally": True,
            "profile.default_content_setting_values.automatic_downloads": 1,
            "profile.content_settings.exceptions.automatic_downloads": {
                "[*.]example.com,*": {"setting": 1}
            }
        }
        chrome_options.add_experimental_option("prefs", chrome_prefs)
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--headless")
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-web-security')
        chrome_options.add_argument("--remote-debugging-port=9222")

        # Create the WebDriver instance with the service object
        service = Service(self.driver_path)
        driver = webdriver.Chrome(service=service, options=chrome_options)

        # Register the WebDriver with the WebDriverCache
        self._drivers.register(driver, alias="chrome")
        self._current_browser = driver

        # Navigate to the specified URL
        self.go_to(url)
