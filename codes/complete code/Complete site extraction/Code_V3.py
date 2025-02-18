import random
import time
import re
import pandas as pd
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from fake_useragent import UserAgent
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# Generate random user-agent
def random_user_agent():
    ua = UserAgent()
    return ua.random

# Setup undetected Chrome driver
def setup_driver():
    options = uc.ChromeOptions()
    options.add_argument(f"user-agent={random_user_agent()}")  # Random User-Agent
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-gpu")
    options.add_argument("--incognito")
    options.add_argument("--disable-extensions")

    driver = uc.Chrome(options=options, headless=False)

    return driver

# Human-like delays
def human_delay(min_delay=2, max_delay=5):
    time.sleep(random.uniform(min_delay, max_delay))

# Random scrolling
def human_scroll(driver):
    scroll_script = "window.scrollBy(0, {})"
    for _ in range(random.randint(2, 5)):
        driver.execute_script(scroll_script.format(random.randint(300, 700)))
        human_delay(1, 3)


def click_next_page(driver):
    try:
        next_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "pnnext")))
        ActionChains(driver).move_to_element(next_button).perform()
        human_delay(2, 4)
        next_button.click()
        human_delay(3, 6)
        return True
    except (NoSuchElementException, TimeoutException):
        print("No 'Next' button found. Ending search.")
        return False

# Function to get Google search URLs by clicking "Next" button
def get_google_search_urls(query, num_results):
    driver = setup_driver()
    search_urls = []

    try:
        driver.get("https://www.google.com")
        search_box = driver.find_element(By.NAME, "q")
        search_box.send_keys(query)
        human_delay(2, 5)
        search_box.send_keys(Keys.RETURN)

        page_count = 0  # Track pages visited

        while len(search_urls) < num_results and page_count < 5:  # Limit to 5 pages to avoid detection
            try:
                # Wait for search results
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.yuRUbf a")))

                # Extract URLs from search results
                results = driver.find_elements(By.CSS_SELECTOR, "div.yuRUbf a")
                search_urls.extend([result.get_attribute("href") for result in results])

                # Stop if enough results are collected
                if len(search_urls) >= num_results:
                    break

                # Scroll before navigating to the next page
                human_scroll(driver)

                # Click "Next" button
                if not click_next_page(driver):
                    break

                page_count += 1  # Increment page counter

            except Exception as e:
                print(f"Error navigating: {e}")
                break

    finally:
        driver.quit()

    return search_urls

# Function to extract text from a list of URLs
def extract_text_from_urls(urls):
    driver = setup_driver()
    data = []
    print(len(urls))
    for url in urls:
        random_delay = random.uniform(1, 5)
        time.sleep(random_delay)
        try:
            print(f"Scraping URL: {url}")
            driver.get(url)
            time.sleep(3)  # Allow page to load
            
            # Get page source
            page_source = driver.page_source
            soup = BeautifulSoup(page_source, "html.parser")
            
            # Remove script and style elements
            for script in soup(["script", "style"]):
                script.decompose()
                
            # Extract visible text
            text = soup.get_text(separator=" ")
            
            # Clean the text
            clean_text = re.sub(r'\s+', ' ', text).strip()
            
            # Store the extracted data
            data.append({"URL": url, "Extracted_Text": clean_text})

        except Exception as e:
            print(f"Error extracting text from {url}: {e}")
            data.append({"URL": url, "Extracted_Text": None})  # Add None if extraction fails

    driver.quit()
    
    # Convert to DataFrame
    df = pd.DataFrame(data)
    return df

# Main execution
if __name__ == "__main__":
    location = r"C:\Users\MRK\Downloads\20 Parts run wo prior web & line.xlsx"
    part_list = pd.read_excel(location)
    
    combined_list = part_list.apply(lambda row: f"{row['Part Num to search']} {row['Description']}", axis=1).tolist()
    
    for i in range(len(combined_list)):  
        query = combined_list[i]
        print(f"Searching for: {query}")
        urls = get_google_search_urls(query, 20)
        df = extract_text_from_urls(urls)
        
        file_name = f"{query.replace(' ', '_')}_project_part_details.xlsx"
        df.to_excel(file_name, index=False)

        print(f"{file_name} - Scraping Done!\n")
