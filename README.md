# News Scraper Bot

## Overview

The **News Scraper Bot** is an automated script that extracts news data from a specified website, filters the results by category, and saves the data to an Excel file. This project was created as part of a challenge with specific requirements, such as avoiding the use of web requests. To fulfill this requirement, the bot uses Selenium's screenshot functionality to download images from the website.

## Features

- **Automated News Search**: The bot performs searches on the news site using a specified phrase.
- **Category Filtering**: It can filter news results by a specified category (e.g., Business, Politics).
- **Date Range Filtering**: The bot checks the publication date of articles to ensure they fall within a specified date range.
- **Image Download**: Using Selenium's screenshot functionality, the bot downloads images associated with the articles.
- **Excel Output**: The extracted data is saved to an Excel file, including the article title, date, description, image filename, and additional information such as phrase occurrence count and money references.
