from RPA.Browser.Selenium import Selenium
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
        self.browser.open_available_browser(url)

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

    def filter_by_category(self):
        pass

    def extract_news_data(self):
        pass

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
            if self.news_category:
                self.filter_by_category()
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
