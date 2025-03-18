from flask import Flask, render_template, request, send_file
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import zipfile

app = Flask(__name__)


@app.route('/')
def home():
    return render_template('index.html')

@app.route('/scrape', methods=['POST'])
def scrape():
    files = request.files.getlist('files')
    input_folder = 'input_files'
    output_folder = 'output_files'

    # Create input and output folders if they don't exist
    os.makedirs(input_folder, exist_ok=True)
    os.makedirs(output_folder, exist_ok=True)

    # Save uploaded files to the input folder
    for file in files:
        file_path = os.path.join(file.filename)
        print(f"Saving file to: {file_path}")  # Debugging information
        file.save(file_path)

    service = Service("C:\\Users\\WS7888\\OneDrive - SKF\\Documents\\ms\\updated\\msedgedriver.exe")
    edge_options = Options()
    edge_options.add_argument("--start-maximized")
    edge_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")

    driver = webdriver.Edge(service=service, options=edge_options)

    try:
        driver.get("https://apps.motor.com/login")
        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.NAME, "username"))).send_keys("wstuck")
        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.NAME, "password"))).send_keys("India@123")
        WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']"))).click()

        time.sleep(5)
        WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.CLASS_NAME, 'card-img-top'))).click()
        time.sleep(5)

        for filename in os.listdir(input_folder):
            if filename.endswith('.xlsx'):
                file_path = os.path.join(input_folder, filename)
                print(f"Processing file: {file_path}")  # Debugging information
                df = pd.read_excel(file_path, engine='openpyxl')

                df['Superseded By'] = ''
                df['Current Part Number'] = ''

                for index, row in df.iterrows():
                    part_number = row['Part Number']
                    search_input = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.NAME, "enterPartNumber")))
                    search_input.clear()
                    search_input.send_keys(part_number)
                    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//button[@id='submitQuickSearchButton']"))).click()

                    time.sleep(5)

                    page_source = driver.page_source
                    soup = BeautifulSoup(page_source, 'html.parser')
                    rows = soup.find_all('tr')

                    superseded_by_list = []
                    current_part_number_list = []

                    for row in rows:
                        columns = row.find_all('td')
                        if len(columns) > 4:
                            superseded_by = columns[4].text.strip()
                            current_part_number = columns[5].text.strip()
                            superseded_by_list.append(superseded_by)
                            current_part_number_list.append(current_part_number)

                    df.at[index, 'Superseded By'] = ', '.join(sorted(set(superseded_by_list)))
                    df.at[index, 'Current Part Number'] = ', '.join(sorted(set(current_part_number_list)))

                    specific_url = "https://apps.motor.com/supersessions/researchparts"
                    driver.get(specific_url)

                output_file_path = os.path.join(output_folder, filename)
                print(f"Saving processed file to: {output_file_path}")  # Debugging information
                df.to_excel(output_file_path, index=False)

    except Exception as e:
        print("An error occurred:", e)
        driver.save_screenshot("error_debug.png")

    finally:
        driver.quit()

    # Create a zip file of the output folder
    zip_filename = 'output_files.zip'
    with zipfile.ZipFile(zip_filename, 'w') as zipf:
        for root, dirs, files in os.walk(output_folder):
            for file in files:
                zipf.write(os.path.join(root, file), file)

    return send_file(zip_filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)