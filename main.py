import json
import os
import re
import time
import logging
from datetime import datetime, timedelta

import openpyxl
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.by import By
from RPA.Robocorp.WorkItems import WorkItems

from extended_selenium import ExtendedSelenium

class NewsScraperBot:
    """
    NewsScraperBot automates the process of extracting news data from a website.

    Attributes:
        config (dict): Configuration dictionary containing search parameters.
        browser (ExtendedSelenium): Instance of the ExtendedSelenium class to interact with the browser.
        search_phrase (str): The phrase to search for in the news.
        news_category (str): The category or section of the news to filter by.
        months (int): Number of months of news to retrieve.
        output_file (str): Path to the Excel file where the results will be saved.
        images_dir (str): Directory where downloaded images will be stored.
    """

    def __init__(self):
        """
        Initializes the NewsScraperBot with the given configuration.

        Args:
            config (dict): Configuration dictionary containing search parameters.
        """
        self.logger = logging.getLogger(__name__)
        logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

        self.work_items = WorkItems()
        self.work_items.get_input_work_item()
        self.config = self.work_items.get_work_item_variables()

        # If no configuration is provided via work items, load from config.json
        if not self.config:
            self.logger.info("No configuration found in work items. Loading from config file.")
            self.config = load_config()

        self.browser = ExtendedSelenium()
        self.search_phrase = self.config['search_phrase']
        self.news_category = self.config.get('news_category', '').capitalize()
        self.months = self.config.get('months', 1)
        self.output_file = os.path.join('output', 'news_data.xlsx')
        self.images_dir = os.path.join('output')

    def start_browser(self, url):
        """
        Starts the browser and navigates to the specified URL.

        Args:
            url (str): The URL of the news site to scrape.
        """
        self.logger.info("Opening browser and navigating to '%s'", url)
        self.browser.open_site(url)
        self.logger.info("URL loaded, starting interaction with the page.")

    def search_news(self):
        """
        Performs a search on the news site using the specified search phrase.
        Filters results by category and sorts by newest if applicable.
        """
        self.logger.info("Searching for news related to '%s'", self.search_phrase)

        # Make the search input field visible and type the search phrase
        search_button_selector = "//button[@data-element='search-button']"
        self.browser.wait_until_element_is_visible(search_button_selector)
        self.browser.click_element(search_button_selector)

        search_input_selector = "//input[@type='text']"
        self.browser.wait_until_element_is_visible(search_input_selector)
        self.logger.info("Typing '%s' into the search input field", self.search_phrase)
        self.browser.input_text(search_input_selector, self.search_phrase)

        # Trigger the search by pressing Enter
        self.browser.press_keys(search_input_selector, "ENTER")

        results_page_selector = "//ul[contains(@class, 'search-results-module-results-menu')]"
        self.browser.wait_until_element_is_visible(results_page_selector)

        self.logger.info("Search results are almost ready!")

        # Click "See All" to expand results before filtering by category
        see_all_button_selector = (
            "//button[contains(@class, 'see-all-button') and @data-toggle-trigger='see-all']"
        )

        try:
            # xpath of this search-results-module-filters-overlay div
            self.browser.wait_until_element_is_visible("//div[@class='search-results-module-filters-overlay']", timeout=10)

            if self.browser.is_element_visible(see_all_button_selector):
                self.browser.click_element(see_all_button_selector)
                self.browser.wait_until_element_is_visible("//span[@class='see-less-text']", timeout=10)
        except Exception as e:
            self.logger.error("An error occurred while expanding topics filter: %s", e)

        # Filter by the specified news category
        if self.news_category:
            category_filter_selector = (
                f"//div[contains(@class, 'search-filter-input')]"
                f"//label/span[text()='{self.news_category}']"
                "/preceding::input[@type='checkbox'][1]"
            )
            self.browser.wait_until_element_is_visible(category_filter_selector, timeout=15)
            try:
                self.browser.click_element(category_filter_selector)
            except Exception as e:
                self.logger.error("An error occurred while filtering by category: %s", e)
            self.browser.wait_until_element_is_visible(results_page_selector, timeout=15)

        # Sort results by newest
        sort_by_newest_selector = "//select[@class='select-input']"
        self.browser.wait_until_element_is_visible(sort_by_newest_selector)
        self.browser.select_from_list_by_value(sort_by_newest_selector, "1")

        self.browser.wait_until_element_is_visible(results_page_selector)

        self.logger.info("Search results are ready!")

    def extract_news_data(self):
        """
        Extracts news data from the search results and stores them in a list.
        
        Returns:
            list: A list of lists, where each inner list contains data about a single news article.
        """
        articles_selector = "//ul[contains(@class, 'search-results-module-results-menu')]/li"
        news_data = []

        # Wait for the page to fully load before extraction
        time.sleep(2)

        while True:
            self.browser.wait_until_element_is_visible(articles_selector)
            articles = self.browser.find_elements(articles_selector)

            self.logger.info("Extracting data from %d articles", len(articles))

            for index, article in enumerate(articles):
                title_selector = ".//h3[@class='promo-title']/a"
                date_selector = ".//p[@class='promo-timestamp']"
                description_selector = ".//p[@class='promo-description']"
                image_selector = ".//picture/img[contains(@class, 'image')]"

                retry_count = 0
                max_retries = 3

                while retry_count < max_retries:
                    try:
                        title = article.find_element(By.XPATH, title_selector).text
                        try:
                            date_text = article.find_element(By.XPATH, date_selector).text
                        except Exception:
                            date_text = None
                        if not date_text:
                            self.logger.warning("No date found for the article, skipping it.")
                            break
                        try:
                            description = article.find_element(By.XPATH, description_selector).text
                        except Exception:
                            description = "No description available"
                        try:
                            image_url = article.find_element(By.XPATH, image_selector).get_attribute("src")
                        except Exception:
                            image_url = None

                        # Check if the article's date is within the specified range
                        if not self.is_within_date_range(date_text):
                            self.logger.info("Date out of range, stopping extraction")
                            return news_data

                        # Download the image and save the filename
                        image_filename = self.download_image(image_url, title) if image_url else "No image available"

                        # Count occurrences of the search phrase and check for money references
                        count_search_phrases = self.count_occurrences(self.search_phrase, title, description)
                        contains_money = self.contains_money(title, description)

                        # Append the extracted data to the list
                        news_data.append([
                            title, date_text, description, image_filename,
                            count_search_phrases, contains_money
                        ])

                        break

                    except StaleElementReferenceException:
                        self.logger.warning("Encountered a stale element reference, refreshing the element...")
                        try:
                            article = self.browser.find_element(By.XPATH, articles_selector + f"[{index + 1}]")
                        except Exception:
                            self.logger.error("Failed to re-fetch the article, skipping to the next one.")
                            break
                        retry_count += 1

                    except IndexError:
                        self.logger.error("Article list index out of range, stopping extraction.")
                        break

                if retry_count == max_retries:
                    self.logger.warning("Max retries reached, moving to the next article")
                    continue

            self.logger.info("Data extracted from %d articles", len(articles))

            # Check for a next page and navigate if available
            next_page_selector = "//div[contains(@class, 'search-results-module-next-page')]/a"
            if self.browser.is_element_visible(next_page_selector):
                self.browser.click_element(next_page_selector)
                self.browser.wait_until_page_contains_element(articles_selector)
            else:
                self.logger.info("No more pages available")
                break

        return news_data

    def is_within_date_range(self, date_text):
        """
        Checks if the article date falls within the specified date range.

        Args:
            date_text (str): The date string extracted from the article.

        Returns:
            bool: True if the article date is within the range, False otherwise.
        """
        # Handle the case where "Sept." is used
        date_text = date_text.replace("Sept.", "Sep.")
        date_formats = ["%b. %d, %Y", "%B %d, %Y"]

        article_date = None
        for date_format in date_formats:
            try:
                article_date = datetime.strptime(date_text, date_format)
                break
            except ValueError:
                continue

        if not article_date:
            # Handle relative date formats like "2 days ago"
            relative_time_match = re.match(
                r'(\d+)\s+(minutes?|hours?|days?)\s+ago', date_text
            )
            if relative_time_match:
                quantity = int(relative_time_match.group(1))
                unit = relative_time_match.group(2)
                article_date = datetime.now() - {
                    'minute': timedelta(minutes=quantity),
                    'hour': timedelta(hours=quantity),
                    'day': timedelta(days=quantity)
                }.get(unit, timedelta(days=0))

            if not article_date:
                raise ValueError(f"Date format not recognized: {date_text}")

        current_date = datetime.now()
        earliest_year = current_date.year
        earliest_month = current_date.month - self.months + 1

        if earliest_month <= 0:
            earliest_year -= (abs(earliest_month) // 12) + 1
            earliest_month = 12 - (abs(earliest_month) % 12)

        return (
            (article_date.year > earliest_year or
             (article_date.year == earliest_year and article_date.month >= earliest_month)) and
            (article_date.year < current_date.year or
             (article_date.year == current_date.year and article_date.month <= current_date.month))
        )

    def download_image(self, image_url, title):
        """
        Downloads an image from the specified URL and saves it to the images directory.

        Args:
            image_url (str): The URL of the image to download.
            title (str): The title of the news article (used for naming the image file).

        Returns:
            str: The filename of the downloaded image.
        """
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

        self.logger.info("Image saved as '%s'", image_filename)

        # Close the image tab and return to the original window
        self.browser.close_window()
        self.browser.switch_window('MAIN')

        self.logger.info("Tab closed, returning to the main window")

        return image_filename

    def count_occurrences(self, phrase, *texts):
        """
        Counts the occurrences of a search phrase in the given texts.

        Args:
            phrase (str): The search phrase to count.
            *texts (str): The texts in which to search for the phrase.

        Returns:
            int: The total number of occurrences of the phrase in the given texts.
        """
        pattern = re.compile(re.escape(phrase), re.IGNORECASE)
        return sum(len(pattern.findall(text)) for text in texts)

    def contains_money(self, *texts):
        """
        Checks if any of the given texts contain references to money.

        Args:
            *texts (str): The texts in which to search for money references.

        Returns:
            bool: True if any of the texts contain money references, False otherwise.
        """
        money_patterns = [
            r"\$\d+(\.\d{1,2})?",  # e.g., $10 or $10.99
            r"\d+(,\d{3})*(\.\d{1,2})?\s*dollars",  # e.g., 1,000 dollars
            r"\d+(,\d{3})*(\.\d{1,2})?\s*USD",  # e.g., 1,000 USD
        ]
        combined_pattern = re.compile("|".join(money_patterns), re.IGNORECASE)
        return any(combined_pattern.search(text) for text in texts)

    def save_to_excel(self, data):
        """
        Saves the extracted news data to an Excel file.

        Args:
            data (list): The list of news data to save.
        """

        os.makedirs(os.path.dirname(self.output_file), exist_ok=True)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['Title', 'Date', 'Description', 'Image Filename',
                   'Count of Search Phrases', 'Contains Money'])
        for row in data:
            ws.append(row)
        wb.save(self.output_file)
        self.logger.info("Data saved to Excel file '%s'", self.output_file)

    def close_browser(self):
        """
        Closes all browser windows.
        """
        self.logger.info("Closing all browser windows")
        self.browser.close_all_browsers()

    def run(self):
        """
        Runs the complete news scraping process.
        """
        try:
            self.start_browser(self.config['url'])
            self.search_news()
            data = self.extract_news_data()
            self.save_to_excel(data)
        finally:
            self.close_browser()


def load_config(config_file='config.json'):
    """
    Loads the configuration from a JSON file.

    Args:
        config_file (str): The path to the configuration file.

    Returns:
        dict: The configuration dictionary.
    """
    with open(config_file) as f:
        return json.load(f)


if __name__ == "__main__":
    bot = NewsScraperBot()
    bot.run()
