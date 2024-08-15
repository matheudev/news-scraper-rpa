from RPA.Browser.Selenium import Selenium
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from SeleniumLibrary.base import keyword
from selenium.webdriver.common.by import By
import json
import re
import os
import openpyxl
import time

class ExtendedSelenium(Selenium):

    def __init__(self, *args, **kwargs):
        Selenium.__init__(self, *args, **kwargs)
        self.driver_path = ChromeDriverManager().install()
        
    @keyword
    def looking_at_element(self, locator):
        element = self.get_webelement(locator)
        self.logger.warn(dir(element))

    @keyword
    def open_site(self, url, **kwargs):
        download_dir = os.path.abspath("output/images")  # Set your download directory here
        
        chrome_options = Options()
        chrome_prefs = {
            "profile.default_content_settings.popups": 0,
            "download.default_directory": download_dir,
            "directory_upgrade": True,
            "safebrowsing.enabled": True,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True,  # Disable Chrome PDF Viewer
            "profile.default_content_setting_values.automatic_downloads": 1,  # Allow automatic downloads
            "profile.content_settings.exceptions.automatic_downloads": {
                "[*.]example.com,*": {
                    "setting": 1
                }
            }
        }
        chrome_options.add_experimental_option("prefs", chrome_prefs)
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--ignore-certificate-errors")
        self.open_browser(
            url=url,
            browser="chrome",
            executable_path=self.driver_path,
            options=chrome_options,
            **kwargs
        )

    @keyword
    def print_webdriver_log(self, logtype):
        print(f"\n{logtype.capitalize()} Log")
        return self.driver.get_log(logtype)

class NewsScraperBot:
    def __init__(self, config):
        self.browser = ExtendedSelenium()
        self.config = config
        self.search_phrase = config['search_phrase']
        self.news_category = config['news_category']
        self.months = config['months']
        self.output_file = os.path.join('output', 'news_data.xlsx')
        self.images_dir = os.path.join('output', 'images')

    def start_browser(self, url):
        print(f"Opening browser and navigating to '{url}'...")
        self.browser.open_site(url)

    def search_news(self):
        print(f"Searching for news related to '{self.search_phrase}'...")

        # Click the button to make the search input visible
        search_button_selector = "//button[@data-element='search-button']"
        self.browser.wait_until_element_is_visible(search_button_selector)
        self.browser.click_element(search_button_selector)

               
        # Locate the search input field on the website
        search_input_selector = "//input[@type='text']"
        self.browser.wait_until_element_is_visible(search_input_selector)
        print(f"Typing '{self.search_phrase}' into the search input field...")
        self.browser.input_text(search_input_selector, self.search_phrase)
        
        # Trigger the search (either by pressing Enter or clicking the search button)
        self.browser.press_keys(search_input_selector, "ENTER")
        
        # Wait for the results page to load
        results_page_selector = "//ul[contains(@class, 'search-results-module-results-menu')]"
        self.browser.wait_until_element_is_visible(results_page_selector)

       # Filter by category if applicable
        if self.news_category:
            category_filter_selector = (
                f"//div[contains(@class, 'search-filter-input')]"
                f"//label/span[text()='{self.news_category}']"
                "/preceding::input[@type='checkbox'][1]"
            )
            self.browser.wait_until_element_is_visible(category_filter_selector, timeout=15)
            self.browser.click_element(category_filter_selector)
            self.browser.wait_until_element_is_visible(results_page_selector, timeout=15)

        
        # Sort by newest using the select dropdown
        sort_by_newest_selector = "//select[@class='select-input']"
        self.browser.wait_until_element_is_visible(sort_by_newest_selector)
        self.browser.select_from_list_by_value(sort_by_newest_selector, "1")

        self.browser.wait_until_element_is_visible(results_page_selector)

        print("Search results are ready!")

    def extract_news_data(self):
        articles_selector = "//ul[contains(@class, 'search-results-module-results-menu')]/li"
        self.browser.wait_until_element_is_visible(articles_selector)
        
        articles = self.browser.find_elements(articles_selector)
        news_data = []

        for article in articles:
            title_selector = ".//h3[@class='promo-title']/a"
            date_selector = ".//p[@class='promo-timestamp']"
            description_selector = ".//p[@class='promo-description']"
            image_selector = ".//picture/img[contains(@class, 'image')]"

            title = article.find_element(By.XPATH, title_selector).text
            date = article.find_element(By.XPATH, date_selector).text
            description = article.find_element(By.XPATH, description_selector).text
            image_url = article.find_element(By.XPATH, image_selector).get_attribute("src")
            
            # Download the image
            if image_url:
                image_filename = self.download_image(image_url, title)
            else:
                image_filename = ""

            count_search_phrases = self.count_occurrences(self.search_phrase, title, description)
            contains_money = self.contains_money(title, description)
            news_data.append([title, date, description, image_filename, count_search_phrases, contains_money])
        
        return news_data

    def download_image(self, image_url, title):
        os.makedirs(self.images_dir, exist_ok=True)

        # Sanitize the title to create a valid filename
        sanitized_title = re.sub(r'[^\w\-_\. ]', '_', title)
        extension = '.png'
        image_filename = f"{sanitized_title}{extension}"
        image_path = os.path.join(self.images_dir, image_filename)
        
        # Open a new tab with the image URL
        self.browser.execute_javascript(f"window.open('{image_url}', '_blank');")
        self.browser.switch_window("NEW")

        # Wait for the image to load
        image_element_selector = "//img"
        self.browser.wait_until_element_is_visible(image_element_selector)
        
        # Save the image using Selenium's screenshot functionality
        image_element = self.browser.find_element(image_element_selector)
        image_data = image_element.screenshot_as_png

        with open(image_path, "wb") as file:
            file.write(image_data)

        print(f"Image saved as '{image_filename}'")

        # Close the image tab and return to the original window
        self.browser.close_window()
        self.browser.switch_window('MAIN')

        print(f"Tab closed, returning to the main window...")
        
        return image_filename  # Returning the filename for reference


    def count_occurrences(self, phrase, *texts):
        pattern = re.compile(re.escape(phrase), re.IGNORECASE)
        return sum(len(pattern.findall(text)) for text in texts)

    def contains_money(self, *texts):
        money_patterns = [
            r"\$\d+(\.\d{1,2})?",  # e.g., $10 or $10.99
            r"\d+(,\d{3})*(\.\d{1,2})?\s*dollars",  # e.g., 1,000 dollars
            r"\d+(,\d{3})*(\.\d{1,2})?\s*USD",  # e.g., 1,000 USD
        ]
        combined_pattern = re.compile("|".join(money_patterns), re.IGNORECASE)
        return any(combined_pattern.search(text) for text in texts)

    def save_to_excel(self, data):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['Title', 'Date', 'Description', 'Image Filename', 'Count of Search Phrases', 'Contains Money'])
        for row in data:
            ws.append(row)
        wb.save(self.output_file)

    def close_browser(self):
        self.browser.close_all_browsers()

    def run(self):
        try:
            self.start_browser(self.config['url'])
            self.search_news()
            data = self.extract_news_data()
            self.save_to_excel(data)
        finally:
            self.close_browser()

def load_config(config_file='config.json'):
    with open(config_file) as f:
        return json.load(f)

if __name__ == "__main__":
    config = load_config()
    bot = NewsScraperBot(config)
    bot.run()
