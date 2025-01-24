Test Status Checker & Excel Updater
This Python script automates the process of scraping test results from a web page, checking the status of tests, and updating an Excel sheet with the test results. The script uses Selenium for web scraping, FuzzyWuzzy for string matching, and OpenPyXL for handling Excel files.

**Key Features:**
Web Scraping with Selenium:

The script automatically navigates through a provided URL and retrieves the status of tests listed on the webpage.
It fetches test details, such as test names and their statuses (e.g., "Pass", "Fail"), by interacting with the webpage elements using Selenium.
Excel File Update:

After gathering the test results, the script compares them with the data present in an existing Excel sheet.
It uses FuzzyWuzzy to compare test names and status from the webpage to the corresponding names in the Excel sheet, ensuring that it can handle minor mismatches (e.g., typos or variations in formatting).
The script updates the Excel sheet, filling in the results with the corresponding "Pass" or "Fail" status, and colors the cells accordingly (green for "Pass", red for "Fail").
**Fuzzy String Matching:**

To handle variations in test names or statuses between the webpage and the Excel file, the script uses FuzzyWuzzy, a library for fuzzy string matching. It ensures that even if there are minor differences in the test names, a match will still be found.
Customizable for Different Use Cases:

The script prompts the user to input the URL of the test results page, the file path, and the sheet name of the Excel file to update, making it customizable to different testing environments and datasets.
Excel Color Coding:

When updating the Excel sheet, the "Pass" and "Fail" statuses are highlighted with color:
Green (00b050) for "Pass"
Red (FF0000) for "Fail"
**How It Works:**
_Scraping Data:_ The user provides a URL where test results are available. The script scrapes test details and their statuses (e.g., "Pass" or "Fail").

_String Matching_: The script compares the test names and statuses from the webpage to the ones in an existing Excel file, using fuzzy matching to account for possible discrepancies.

_Updating the Excel File_: If a match is found, the script updates the Excel sheet with the corresponding status ("Pass" or "Fail"), and colors the cells based on the status.

_Logging Results_: The script logs which rows in the Excel sheet were successfully matched and updated.

**The script will prompt you for the following inputs:**

URL: Enter the URL where the test results are located.
File path: Enter the path to your existing Excel file.
Sheet name: Enter the name of the sheet where test data is stored.
The script will scrape the test statuses from the URL, match them with the Excel sheet, and update the sheet with "Pass" or "Fail" statuses. The results will be color-coded accordingly.
