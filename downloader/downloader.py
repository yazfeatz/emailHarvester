import os
import re
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


wordlist_file = "wordlist.txt"
output_folder = "../sources"
visited_links_file = "visited_links.txt"
searched_words_file = "searched_words.txt"

os.makedirs(output_folder, exist_ok=True)

options = webdriver.ChromeOptions()
options.binary_location = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"  # Set the path to the Chrome binary

# Set the download folder
prefs = {
    "download.default_directory": os.path.abspath(output_folder),
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": False
}
options.add_experimental_option("prefs", prefs)

visited_links = set()

if os.path.exists(visited_links_file):
    with open(visited_links_file, "r") as file:
        for line in file:
            visited_links.add(line.strip())

searched_words = set()

if os.path.exists(searched_words_file):
    with open(searched_words_file, "r") as file:
        for line in file:
            searched_words.add(line.strip())

with open(wordlist_file, "r") as file:
    wordlist = [line.strip() for line in file if line.strip() not in searched_words]

for word in wordlist:
    try:
        driver = webdriver.Chrome(options=options)  # Start a new browser instance for each word

        query = f'site:lk "{word}" (filetype:xlsx OR filetype:xls)'

        # Perform the Google search
        search_url = f"https://www.google.com/search?q={query}"
        driver.get(search_url)

        while True:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.yuRUbf > a')))

            search_results = driver.find_elements(By.CSS_SELECTOR, 'div.yuRUbf > a')

            if not search_results:
                print(f"No search results for '{word}'. Skipping...")
                break

            for result in search_results:
                link = result.get_attribute("href")

                if link.endswith(".xlsx") or link.endswith(".xls") and link not in visited_links:
                    # Open the search result link in a new tab
                    driver.execute_script(f"window.open('{link}', '_blank');")

                    visited_links.add(link)

                    with open(visited_links_file, "a") as file:
                        file.write(link + "\n")

            next_page_link = driver.find_element(By.LINK_TEXT, "Next")
            if not next_page_link:
                break

            next_page_link.click()

        for tab in driver.window_handles[1:]:
            driver.switch_to.window(tab)

            file_extension = os.path.splitext(driver.current_url)[1].lower()
            if file_extension in [".xlsx", ".xls"]:
                filename = os.path.join(output_folder, f"{word}_{os.path.basename(driver.current_url)}")

                driver.save_screenshot(filename)  # Save the downloaded file
                print(f"Downloaded: {filename}")

            driver.close()  # Close the current tab

        driver.switch_to.window(driver.window_handles[0])  # Switch back to the main tab
        driver.quit()  # Quit the browser instance after processing a word

        searched_words.add(word)
        with open(searched_words_file, "a") as file:
            file.write(word + "\n")

    except Exception as e:
        print(f"An error occurred for '{word}': {str(e)}")
