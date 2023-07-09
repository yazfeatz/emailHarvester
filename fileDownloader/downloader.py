import os
import re
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Set the path to the wordlist file
wordlist_file = "wordlist.txt"

# Set the path to the folder to save the downloaded files
output_folder = "../sources"

# Create the output folder if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

# Configure Selenium Firefox driver options
options = webdriver.FirefoxOptions()
# options.headless = True  # Run the browser in headless mode
options.binary_location = "/usr/bin/firefox"  # Set the path to the Firefox binary

# Initialize the Firefox driver
driver = webdriver.Firefox(options=options)

# Read the wordlist from the file
wordlist = []
with open(wordlist_file, "r") as file:
    for line in file:
        word = line.strip()
        wordlist.append(word)

# Loop through each word in the wordlist
for word in wordlist:
    try:
        # Construct the search query for xlsx and xls files
        query = f'site:lk "{word}" (filetype:xlsx OR filetype:xls)'

        # Perform the Google search
        search_url = f"https://www.google.com/search?q={query}"
        driver.get(search_url)

        # Loop through the search result pages
        while True:
            # Wait for the search results to load
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.yuRUbf > a')))

            # Get the search result links
            search_results = driver.find_elements(By.CSS_SELECTOR, 'div.yuRUbf > a')

            # Check if there are search results
            if not search_results:
                print(f"No search results for '{word}'. Skipping...")
                break

            # Loop through the search results
            for result in search_results:
                link = result.get_attribute("href")

                # Check if the link points to a .xlsx or .xls file
                if link.endswith(".xlsx") or link.endswith(".xls"):
                    # Open the search result link in a new tab
                    driver.execute_script(f"window.open('{link}', '_blank');")

            # Check if there is a next page
            next_page_link = driver.find_element(By.LINK_TEXT, "Next")
            if not next_page_link:
                break

            # Go to the next page of search results
            next_page_link.click()

        # Switch to each opened tab and download files
        for tab in driver.window_handles[1:]:
            driver.switch_to.window(tab)

            # Check if the file extension is .xlsx or .xls
            file_extension = os.path.splitext(driver.current_url)[1].lower()
            if file_extension in [".xlsx", ".xls"]:
                # Generate the filename for the downloaded file
                filename = os.path.join(output_folder, f"{word}_{os.path.basename(driver.current_url)}")

                # Download the file
                driver.save_screenshot(filename)  # Save the downloaded file
                print(f"Downloaded: {filename}")

            driver.close()  # Close the current tab

        driver.switch_to.window(driver.window_handles[0])  # Switch back to the main tab

    except Exception as e:
        print(f"An error occurred for '{word}': {str(e)}")

# Close the browser
driver.quit()
