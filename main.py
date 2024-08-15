from RPA.Browser.Selenium import Selenium
from webdriver_manager.chrome import ChromeDriverManager
import json
import re
import os
import openpyxl
from datetime import datetime

class NewsScraperBot:
    def __init__(self, config):
        self.browser = Selenium()
        self.config = config
        self.search_phrase = config['search_phrase']
        self.news_category = config['news_category']
        self.months = config['months']
        self.output_file = os.path.join('output', 'news_data.xlsx')
        self.images_dir = os.path.join('output', 'images')

    def start_browser(self, url):
        chrome_driver_path = ChromeDriverManager().install()

        # Pass the driver path and options to RPA.Browser.Selenium
        self.browser.open_user_browser(url)

    def search_news(self):
        # Locate the search input field on the website
        search_input_selector = "input[name='q']"  # Adjust the selector based on the site's structure
        self.browser.wait_until_element_is_visible(search_input_selector)
        self.browser.input_text(search_input_selector, self.search_phrase)
        
        # Trigger the search (either by pressing Enter or clicking the search button)
        self.browser.press_keys(search_input_selector, "ENTER")
        
        # Wait for the results page to load
        results_page_selector = "ul.search-results-module-results-menu" 
        self.browser.wait_until_element_is_visible(results_page_selector)

        # Filter by category if applicable
        if self.news_category:
            # Use the value attribute to select the corresponding checkbox
            category_filter_selector = f"input[type='checkbox'][value='{self.news_category}']"
            self.browser.wait_until_element_is_visible(category_filter_selector)
            self.browser.click_element(category_filter_selector)

            # Wait for the results to reload after applying the filter
            self.browser.wait_until_element_is_visible(results_page_selector)
        
        # Sort by newest using the select dropdown
        sort_by_newest_selector = "select.select-input"  # Selector for the dropdown
        self.browser.wait_until_element_is_visible(sort_by_newest_selector)
        self.browser.select_from_list_by_value(sort_by_newest_selector, "1")  # '1' corresponds to the 'Newest' option

        # Wait for the results to reload after sorting
        self.browser.wait_until_element_is_visible(results_page_selector)

    def extract_news_data(self):
        articles_selector = "ul.search-results-module-results-menu li"  # Selector for the list of articles
        self.browser.wait_until_element_is_visible(articles_selector)
        
        articles = self.browser.find_elements(articles_selector)
        news_data = []

        for article in articles:
            title_selector = ".promo-title a"
            date_selector = "p.promo-timestamp"
            description_selector = "p.promo-description"
            image_selector = "picture img.image"

            title = self.browser.get_text(f"{article} {title_selector}")
            date = self.browser.get_text(f"{article} {date_selector}")
            description = self.browser.get_text(f"{article} {description_selector}")
            image_url = self.browser.get_element_attribute(f"{article} {image_selector}", "src")
            
            # Download the image
            if image_url:
                image_filename = self.download_image(image_url)
            else:
                image_filename = ""

            # Count occurrences of the search phrase in title and description
            count_search_phrases = self.count_occurrences(self.search_phrase, title, description)
            
            # Check if the title or description contains any amount of money
            contains_money = self.contains_money(title, description)
            
            # Append the data to the list
            news_data.append([title, date, description, image_filename, count_search_phrases, contains_money])
        
        return news_data

    def download_image(self, url):
        image_filename = os.path.join(self.images_dir, os.path.basename(url))
        self.browser.download(url, image_filename)
        return os.path.basename(image_filename)

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
