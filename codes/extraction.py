from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from bs4 import BeautifulSoup
import re
import time


def setup_driver():
    """Setup Selenium Chrome driver with default options."""
    options = Options()
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('--incognito')
    return webdriver.Chrome(options=options)


def get_google_search_urls(query, num_results):
    """
    Retrieves URLs from Google search results for the given query.
    Returns a list of URLs.
    """
    driver = setup_driver()
    search_urls = []
    try:
        driver.get("https://www.google.com")
        search_box = driver.find_element(By.NAME, "q")
        search_box.send_keys(query)
        search_box.send_keys(Keys.RETURN)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.yuRUbf a")))

        # Extract URLs from the first page
        results = driver.find_elements(By.CSS_SELECTOR, "div.yuRUbf a")
        search_urls.extend([result.get_attribute("href") for result in results])

        # Fetch next pages if needed
        while len(search_urls) < num_results:
            try:
                next_button = driver.find_element(By.ID, "pnnext")
                next_button.click()
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.yuRUbf a")))
                results = driver.find_elements(By.CSS_SELECTOR, "div.yuRUbf a")
                search_urls.extend([result.get_attribute("href") for result in results])
            except Exception as e:
                print(f"Error navigating to the next page: {e}")
                break
    finally:
        driver.quit()

    return search_urls[:num_results]


def price_comparison(query, num_results):
    """
    Performs price comparison for the given query by extracting product details from multiple websites.
    Returns a pandas DataFrame containing the product details.
    """
    urls = get_google_search_urls(query, num_results)
    all_product_details = []

    for url in urls:
        driver = setup_driver()
        try:
            driver.get(url)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
            soup = BeautifulSoup(driver.page_source, 'html.parser')

            # Extract all class names
            all_elements = soup.find_all()
            class_names = []
            for element in all_elements:
                classes = element.get('class', [])
                ids = element.get('id', '')
                for cls in classes:
                    class_names.append(cls)
                class_names.append(ids)

            #print(class_names)
            # Gather product details
            details = {
                'url': url,
            'Product Name': extract_product_details(soup,class_names, ['heading','productTitle','productName', 'product-title', 'mainTitle', 'product-detail-title','product-basic-details__text--title', 'productTitle', 'titleSection','Product header','product title','item tilte','item name','product name']),
            'Price': extract_product_details(soup,class_names, ['corePrice', 'product-price', 'price-primary','item price','price','mrp']),
            'Part No': extract_product_details(soup,class_names, ['part_number', 'partNumSection', 'part-number', 'product_part-info', 'item-part-number','sku','sku number','item part number','part number','item no','product-basic-details__text--part']),
            'Taxonomy': extract_product_details(soup,class_names, ['brdcrmb','breadcrumb', 'bread-crumb']),
            'Cross Reference': extract_product_details(soup,class_names, ['interchange','Crossreference','replaces', 'product superseded','superseded','interchages']),
            'Details': extract_product_details(soup,class_names,['item-description','description','prodDetails', 'ProductDetail', 'item-desc', 'product_description', 'techSpec', 'item-description', 'product-details', 'product_info', 'desc','detail']),
            'Specification': extract_product_details(soup,class_names,['spec', 'specs', 'details', 'product-spec','specification','tech']),
            'Warranty': extract_product_details(soup, class_names,['warranty', 'guarantee']),
            'availability': extract_product_details(soup,class_names,['availability', 'stock', 'in-stock'])
            }

            all_product_details.append(details)
            print(f"Website done: {url}")

        except Exception as e:
            print(f"Error processing URL {url}: {e}")
        finally:
            driver.quit()

    return pd.DataFrame(all_product_details)


def extract_product_details(soup, class_names, substrings):
    extracted_data = []
    # For each data field, find all class names that contain any of the specified substrings
    for substr in substrings:
        matched_classes= []
        for cls in class_names:
            cls1 = re.sub(r'[\s+\-\_+]', '', cls).lower()
            substr1 = re.sub(r'[\s+\-\_+]', '', substr).lower()
            if substr1 in cls1:
                matched_classes.append(cls)
                break
        #print(matched_classes)
        elements = soup.find_all()  # Find all elements in the soup
        for element in elements:
             if element.get('id', '') in matched_classes:
                extracted_text = re.sub(r'\s+', ' ', element.get_text()).strip()
                extracted_data.append(extracted_text)
                    # Check if any class substring is present in the 'class' attribute
             if (' '.join(element.get('class', [])) in matched_classes):
                extracted_text = re.sub(r'\s+', ' ', element.get_text()).strip()
                extracted_data.append(extracted_text)
    return extracted_data


if __name__ == "__main__":
    location = r"/home/bama/Documents/_2024/apa_engg/codes/sites.xlsx"
    list = pd.read_excel(location)
    print(list.columns)
    combined_list = list.apply(lambda row: f"{row['Part num']} {row['Description']}" + " Price", axis=1).tolist()
    for i in range(30,41):
        query = combined_list[i]
        df = price_comparison(query, num_results=25)
        file_name = f"{query.replace(' ', '_')}_substr_part_details.xlsx"
        df.to_excel(file_name, index=False)
        wb = load_workbook(file_name)
        ws = wb.active
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, horizontal='left', vertical='top')
        wb.save(file_name)
        print(df)
        print(file_name,"Done")