from selenium_stealth import stealth
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException, WebDriverException
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pandas as pd
from openpyxl import load_workbook
from bs4 import BeautifulSoup
import re
import time
import nltk
from nltk import pos_tag, word_tokenize
from flask import Flask, request, render_template, jsonify
import os

# Create Flask app with explicit template folder
app = Flask(__name__, 
            template_folder='templates',  # Specify template folder
            static_folder='static')       # Specify static folder

# Create necessary folders
os.makedirs('templates', exist_ok=True)
os.makedirs('static', exist_ok=True)
os.makedirs('output', exist_ok=True)

class CrossReferenceExtractor:
    def __init__(self):
        # Download NLTK data
        try:
            nltk.download('punkt')
            nltk.download('averaged_perceptron_tagger')
        except Exception as e:
            print(f"NLTK download error: {e}")
    
    def safe_tokenize(self, text):
        """Safely tokenize text with fallback option"""
        try:
            return word_tokenize(text)
        except LookupError:
            return text.replace('(', ' ( ').replace(')', ' ) ').replace(',', ' , ').split()

    def extract_cross_references_nlp(self, text):
        """Extract cross references using NLP approach"""
        tokens = self.safe_tokenize(text)
        tagged = pos_tag(tokens)
        cross_ref = []
        i = 0
        
        while i < len(tagged):
            if tagged[i][1] in ('NN', 'NNS', 'NNP', 'NNPS'):
                noun_phrase = []
                while i < len(tagged) and tagged[i][1] in ('NN', 'NNS', 'NNP', 'NNPS'):
                    noun_phrase.append(tagged[i][0])
                    i += 1

                num_phrase = []
                while i < len(tagged):
                    if tagged[i][0] == ',':
                        i += 1
                        continue
                    
                    if tagged[i][1] == 'CD' or re.match(r'^[A-Z0-9]+$', tagged[i][0]):
                        num_phrase.append(tagged[i][0])
                    else:
                        break
                    i += 1

                if noun_phrase and num_phrase:
                    cross_ref.append((' '.join(noun_phrase), ', '.join(num_phrase)))
            else:
                i += 1
        
        return cross_ref

    def extract_cross_references_regex(self, text):
        """Extract cross references using regex as fallback"""
        pattern = r'([A-Za-z]+)\s+([A-Z0-9]+(?:\s*,\s*[A-Z0-9]+)*)'
        matches = re.finditer(pattern, text)
        return [(match.group(1), match.group(2)) for match in matches]

    def get_cross_references(self, text):
        """Main method to extract cross references with fallback"""
        try:
            refs = self.extract_cross_references_nlp(text)
            return refs if refs else self.extract_cross_references_regex(text)
        except Exception as e:
            print(f"Error in NLP processing, using regex fallback: {e}")
            return self.extract_cross_references_regex(text)

def setup_driver():
    options = Options()
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('--incognito')
    options.add_argument('--lang=en-US')
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_argument('--headless')  # Added headless mode for server deployment

    driver = webdriver.Chrome(options=options)
    stealth(driver,
            languages=["en-US", "en"],
            vendor="Google Inc.",
            platform="Win32",
            webgl_vendor="Intel Inc.",
            renderer="Intel Iris OpenGL Engine",
            fix_hairline=True)
    return driver

def get_google_search_urls(query, num_results):
    driver = setup_driver()
    search_urls = []
    sites = []
    required_results = (len(sites) * 2) + num_results
    
    try:
        # Site-specific searches
        for site in sites:
            site_query = f"{query} {site}"
            driver.get("https://www.google.com")
            search_box = driver.find_element(By.NAME, "q")
            search_box.send_keys(site_query)
            search_box.send_keys(Keys.RETURN)

            try:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.yuRUbf a")))
                results = driver.find_elements(By.CSS_SELECTOR, "div.yuRUbf a")[:2]
                site_urls = [result.get_attribute("href") for result in results]
                search_urls.extend(site_urls)
            except Exception as e:
                print(f"Error in site search: {e}")

        # General search
        if len(search_urls) < required_results:
            driver.get("https://www.google.com")
            search_box = driver.find_element(By.NAME, "q")
            search_box.send_keys(query)
            search_box.send_keys(Keys.RETURN)
            
            while len(search_urls) < required_results:
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.yuRUbf a")))
                    results = driver.find_elements(By.CSS_SELECTOR, "div.yuRUbf a")
                    search_urls.extend([result.get_attribute("href") for result in results])
                    
                    next_button = driver.find_element(By.ID, "pnnext")
                    next_button.click()
                except Exception:
                    break

    finally:
        driver.quit()

    return list(set(search_urls[:num_results]))  # Remove duplicates

def extract_product_details(url, driver, substrings):
    """Extract product details and cross references from webpage"""
    try:
        driver.get(url)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
        
        # Extract main content
        main_soup = BeautifulSoup(driver.page_source, 'html.parser')
        substrings_set = {re.sub(r'[\s+\-\_+]', '', s).lower() for s in substrings}
        extracted_text = extract_data_from_elements(main_soup.find_all(), substrings_set, url, driver)
        
        # Extract from iframes
        for iframe in driver.find_elements(By.TAG_NAME, 'iframe'):
            try:
                driver.switch_to.frame(iframe)
                iframe_soup = BeautifulSoup(driver.page_source, 'html.parser')
                iframe_text = extract_data_from_elements(iframe_soup.find_all(), substrings_set, url, driver)
                if iframe_text:
                    extracted_text += ' ' + iframe_text
            finally:
                driver.switch_to.default_content()
        
        # Extract cross references
        extractor = CrossReferenceExtractor()
        cross_refs = extractor.get_cross_references(extracted_text)
        
        return {
            'extracted_text': extracted_text,
            'cross_references': cross_refs
        }
    
    except Exception as e:
        print(f"Error extracting from {url}: {e}")
        return {'extracted_text': '', 'cross_references': []}

def extract_data_from_elements(elements, substrings_set, url, driver):
    """Extract text from matching elements"""
    extracted_data = []
    for element in elements:
        element_classes = [cls.lower() for cls in element.get('class', [])]
        element_id = element.get('id', '').lower()
        
        if any(substr in re.sub(r'[\s+\-\_+]', '', cls) 
              for cls in element_classes + [element_id] 
              for substr in substrings_set):
            text = ' '.join(element.stripped_strings)
            if text:
                extracted_data.append(text)
    
    return ' /'.join(extracted_data)

def price_comparison(query, num_results, part_num, description, mfr):
    """Main function to compare prices and extract data"""
    driver = setup_driver()
    urls = get_google_search_urls(query, num_results)
    all_product_details = []

    try:
        for url in urls:
            details = extract_product_details(url, driver, [
                'part detail title number',
                'product-line-sku',
                'product_info__description_list',
                'part_number',
                'product-details-information',
                'OECROSSREFERENCE',
                'product-specifications'
            ])
            
            product_info = {
                'MFR Line': mfr,
                'Part Num to search': part_num,
                'Description': description,
                'url': url,
                'Extracted_Text': details['extracted_text'],
                'Cross_References': '; '.join([f"{brand}: {nums}" for brand, nums in details['cross_references']])
            }
            
            all_product_details.append(product_info)
            print(f"Processed: {url}")

    finally:
        driver.quit()

    return pd.DataFrame(all_product_details)

@app.route('/')
def index():
    try:
        return render_template('index.html')
    except Exception as e:
        app.logger.error(f"Error rendering template: {str(e)}")
        return jsonify({"error": "Template not found. Please ensure templates directory exists."}), 500

@app.route('/extraction', methods=['POST'])
def compare_prices_endpoint():
    try:
        if not request.is_json:
            return jsonify({"error": "Request must be JSON"}), 400

        data = request.json
        input_str = data.get('input_str')
        
        if not input_str:
            return jsonify({"error": "input_str is required"}), 400

        queries = input_str.split('|')
        
        if len(queries) != 3:
            return jsonify({"error": "Invalid input format. Expected: part_number|description|brand"}), 400
        
        part_num = queries[0].strip()
        description = queries[1].strip()
        brand = queries[2].strip()
        query = f"{part_num} {description} {brand}"

        df = price_comparison(query, num_results=5, part_num=part_num, description=description, mfr=brand)
        
        # Save to file and return results
        file_name = f"output/{query.replace(' ', '_')}_part_details.json"
        result = df.to_json(orient='records', indent=4)
        
        with open(file_name, 'w') as f:
            f.write(result)
        
        return jsonify({"success": True, "data": df.to_dict(orient='records')})

    except Exception as e:
        app.logger.error(f"Error in extraction endpoint: {str(e)}")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)