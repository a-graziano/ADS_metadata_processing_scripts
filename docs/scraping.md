# Projects Scraping

This Python script scrapes project name from a specific organization URLs, processes it, and exports it to an Excel file.

## About the Project

This script uses `requests`, `BeautifulSoup`, and `pandas` to scrape project data from organization URLs. It then compiles the data into a pandas DataFrame and exports it to an Excel file.

## Getting Started

### Prerequisites

Make sure you have the following Python libraries installed:

- `requests`
- `BeautifulSoup`
- `pandas`

You can install them using the following command:

```bash
pip install requests beautifulsoup4 pandas
```
### Usage

Search for the organization where you want to find projects on the [ADS website](https://archaeologydataservice.ac.uk/library/browse/organisations.xhtml).  Be careful, as one organization could have multiple names, which means you'll need to add multiple URLs.

Run the script

```bash
python projects_scraping.py
```