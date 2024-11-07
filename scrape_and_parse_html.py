import os
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from flask import Flask, render_template, request, send_file
from bs4 import BeautifulSoup
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd

# Initialize the Flask application
app = Flask(__name__)

# Define the field mapping to standardize the labels
field_mapping = {
    'Project Name': ['Project Name', 'Project Title', 'Name of Project'],
    'Project Owner': ['Project Owner', 'Owner', 'Client'],
    'Role In Project': ['Role In Project', 'Role', 'Project Role'],
    'Individual': ['Individual', 'Staff Member', 'Team Member'],
    'Percent': ['Percent', 'Percentage', '% Time'],
    'Completed': ['Completed', 'Completion Date', 'Date Completed'],
    'Work Class': ['Work Class', 'Work Classification', 'Project Class'],
    'Non Bridge Complexity': ['Non Bridge Complexity', 'Non-Bridge Complexity'],
    'Complexity': ['Complexity', 'Bridge Complexity'],
    'Non Bridge Cost ($1,000\'s)': ['Non Bridge Cost ($1,000\'s)', 'Non-Bridge Cost'],
    'Bridge Cost ($1,000\'s)': ['Bridge Cost ($1,000\'s)', 'Bridge Cost']
}

# Reverse mapping for quick lookups
reverse_mapping = {alt: std for std, alts in field_mapping.items() for alt in alts}

def parse_html_content(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    project_data = {}
    
    # Find all spans with class f5 or f6 in reverse order
    spans = soup.find_all('span', class_=['f5', 'f6'])[::-1]
    
    def clean_text(text):
        return text.strip().replace(':', '').replace('\n', ' ').strip()
    
    def get_standardized_field(field):
        cleaned_field = clean_text(field)
        return reverse_mapping.get(cleaned_field, cleaned_field)
    
    i = 0
    while i < len(spans):
        if 'f6' in spans[i].get('class', []):
            label_text = clean_text(spans[i].get_text())
            standardized_label = get_standardized_field(label_text)
            
            if standardized_label:
                # Find the nearest preceding f5 element
                j = i + 1
                while j < len(spans) and 'f5' not in spans[j].get('class', []):
                    j += 1
                
                if j < len(spans) and 'f5' in spans[j].get('class', []):
                    value_text = clean_text(spans[j].get_text())
                    project_data[standardized_label] = value_text
                    print(f"Collected {standardized_label}: {value_text}")
        
        i += 1

    return project_data  # Return the complete dictionary of parsed fields

def create_project_dataframe(all_parsed_data):
    # Create a DataFrame from all parsed data without restricting columns
    df = pd.DataFrame(all_parsed_data)
    
    # Transpose the DataFrame for the desired format
    df_transposed = df.transpose()
    df_transposed.columns = [f"Project {i+1}" for i in range(df_transposed.shape[1])]
    
    # Format specific columns if needed
    numeric_columns = [
        "Non Bridge Cost ($1,000's)",
        "Bridge Cost ($1,000's)",
        "Percent"
    ]
    
    for col in df_transposed.columns:
        for idx in numeric_columns:
            if idx in df_transposed.index:
                value = df_transposed.loc[idx, col]
                if isinstance(value, str) and value:
                    numeric_value = ''.join(c for c in value if c.isdigit() or c == '.')
                    try:
                        df_transposed.loc[idx, col] = float(numeric_value)
                    except ValueError:
                        pass
    
    return df_transposed

@app.route('/')
def index():
    return render_template('search_form.html')

@app.route('/search', methods=['POST'])
def search_firm():
    firm_name = request.form['firm_name']
    firm_number = request.form['firm_number']
    
    # Set up the Chrome WebDriver
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    url = 'https://www.dot.ny.gov/main/business-center/consultants/architectural-engineering/view-inventory-of-consultant-firm-information-experience'
    driver.get(url)

    try:
        view_inventory_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.LINK_TEXT, "View The Inventory of Consultant Firm Information and Experience"))
        )
        view_inventory_link.click()

        driver.switch_to.window(driver.window_handles[1])

        if firm_name:
            search_box = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, 'firmName'))
            )
            search_box.click()
            search_box.clear()
            search_box.send_keys(firm_name)

        elif firm_number:
            search_box = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, 'fedIDNumber'))
            )
            search_box.click()
            search_box.clear()
            search_box.send_keys(firm_number)

        search_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.NAME, 'forward'))
        )
        search_button.click()

        driver.find_element(By.LINK_TEXT, "Select Firm").click()
        driver.find_element(By.LINK_TEXT, "Active Reports").click()

        inventory_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@value='Construction Inspection Project Inventory']"))
        )
        inventory_button.click()

        html_links = driver.find_elements(By.LINK_TEXT, "HTML")
        all_parsed_data = []

        original_tab = driver.current_window_handle

        for index, link in enumerate(html_links, start=1):
            action = ActionChains(driver)
            action.key_down(Keys.CONTROL if os.name == 'nt' else Keys.COMMAND).click(link).key_up(Keys.CONTROL if os.name == 'nt' else Keys.COMMAND).perform()
    
            WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(len(driver.window_handles)))
            driver.switch_to.window(driver.window_handles[-1])

            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "html")))
    
            html_page_source = driver.page_source
            parsed_data = parse_html_content(html_page_source)
            all_parsed_data.append(parsed_data)

            driver.close()
            driver.switch_to.window(original_tab)
            time.sleep(1)

        driver.quit()

        df_transposed = create_project_dataframe(all_parsed_data)

        if not os.path.exists('downloads'):
            os.makedirs('downloads')

        excel_file = os.path.join('downloads', 'formatted_project_data.xlsx')
        df_transposed.to_excel(excel_file, index=True)

        return send_file(excel_file, as_attachment=True, download_name='formatted_project_data.xlsx')

    except Exception as e:
        return f"An error occurred: {e}"
    finally:
        driver.quit()

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=49152, debug=True)
