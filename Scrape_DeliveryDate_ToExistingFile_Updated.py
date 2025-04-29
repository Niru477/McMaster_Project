import os
import pandas as pd
from seleniumbase import SB
from selenium.webdriver.common.by import By
import time
from datetime import datetime
import requests
import re

# Get address from IP info
# def get_address():
#     """Gets the user's approximate address via IP."""
#     try:
#         response = requests.get("https://ipinfo.io/json", timeout=5)
#         if response.status_code == 200:
#             data = response.json()
#             city = data.get("city", "City not available")
#             region = data.get("region", "Region not available")
#             country = data.get("country", "Country not available")
#             return f"{city}, {region}, {country}"
#     except:
#         pass
#     return "Address Not Found"
#
def get_address():
    """Get the approximate address (city, region, country) of the current machine using its IP."""
    try:
        # Request geolocation information from ip-api.com
        response = requests.get("http://ip-api.com/json", timeout=5)
        
        if response.status_code == 200:
            data = response.json()  # Convert the response to a JSON dictionary
            
            # Extract the required details from the JSON response
            city = data.get("city", "City not available")
            region = data.get("regionName", "Region not available")
            country = data.get("country", "Country not available")
            
            # Combine city, region, and country into a single address string
            address = f"{city}, {region}, {country}"
            return address
        else:
            return "Could not fetch address data."
    except requests.RequestException as e:
        return f"Error occurred: {str(e)}"

# Check if the file has already been processed by checking if it exists in the output folder
def is_file_processed(file_name, output_folder):
    """Checks if the file has been processed by looking for the file in the output folder."""
    output_path = os.path.join(output_folder, file_name)
    return os.path.exists(output_path)

# Clean the first row of the DataFrame by removing "Unnamed: X" columns
def clean_dataframe(df):
    # Remove 'Unnamed' columns or rows with 'Unnamed' in the column name
    df.columns = df.columns.str.replace(r'^Unnamed.*$', '', regex=True)
    return df

# Process each Excel file in the folder
input_folder = r"C:\Users\Lenevo\Desktop\McM - Second\Spring Plungers"
output_folder = os.path.join(input_folder, "Output")  # Output folder path

# Create the Output folder if it doesn't exist
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Process each Excel file in the folder
for filename in os.listdir(input_folder):
    if filename.endswith(".xlsx"):
        file_path = os.path.join(input_folder, filename)
        print(f"\nProcessing file: {filename}")

        # Check if the file has already been processed (exists in the output folder)
        if is_file_processed(filename, output_folder):
            print(f"File {filename} has already been processed. Skipping.")
            continue

        # Load all sheets from the Excel file into a dictionary of DataFrames
        all_sheets = pd.read_excel(file_path, sheet_name=None)

        # Dictionaries to hold data
        delivery_dates = {}
        extracted_dates = {}
        addresses = {}

        # Start SeleniumBase browser
        with SB(uc=True, incognito=True, maximize=True, locale_code="en", skip_js_waits=True, headless=False) as sb:
            # Process each sheet
            for sheet_name, df in all_sheets.items():
                print(f"Processing sheet: {sheet_name}")
                
                # Extract part numbers from all cells
                part_numbers = set()
                for row in df.values:
                    for cell in row:
                        if isinstance(cell, str):
                            match = re.match(r"\d{4,5}[A-Z]\d{1,3}", cell.strip())
                            if match:
                                part_numbers.add(match.group())
                
                part_numbers = list(part_numbers)

                # Process each part number
                for part in part_numbers:
                    url = f"https://www.mcmaster.com/{part}"
                    print(f"Visiting {url}")
                    sb.driver.open(url)
                    time.sleep(5)

                    try:
                        try:
                            choose_type = sb.driver.find_element(By.XPATH,"//div[starts-with(@class,'SpecChoiceParameterContainer_parameter')]//span[starts-with(@class,'TextDisplay_plainText')]")
                            choose_type.click() 
                            time.sleep(3)
                            # print("Selected the type")
                        except :   
                            # print("No type displayed")
                            pass
                        try:
                            quantity_input = sb.driver.find_element(By.XPATH, "//input[contains(@class,'input-simple--qty')]")
                            quantity_input.send_keys("1")
                            time.sleep(2)

                            add_to_order_button = sb.driver.find_element(By.XPATH, "//button[contains(@class,'add-to-order-pd button-add-to-order-pd')]")
                            add_to_order_button.click()
                            time.sleep(3)

                            delivery_date_element = sb.driver.find_element(By.XPATH, "//div[contains(@class,'alert--ord-conf productDetailATOCopy InLnOrdWebPartLayout_ItmAddedMsg')]")
                            delivery_text = delivery_date_element.text
                            delivery_date_lines = delivery_text.split('\n')
                            delivery_date = delivery_date_lines[1].strip() if len(delivery_date_lines) > 1 else "Not found"
                                                                  
                        except:
                            # choose_type = sb.driver.find_element(By.XPATH,"//div[contains(@class,'SpecChoiceParameterContainer_parameter__3JRqL')]")
                            # choose_type.click() 
                               
                            # add_to_order_button = sb.driver.find_element(By.XPATH, "//button[contains(@class,'AddToOrderButton_addToOrderButton__335hl ProductDetailOrderBoxFrame_productDetailAddToOrderButton__3UwT7 ProductDetailOrderBoxFrame_productDetailAddToOrderButtonBottomMargin__x9yfE')]")
                            # add_to_order_button.click()
                            # time.sleep(3)

                            # Find all elements matching the XPath
                            delivery_date_elements = sb.driver.find_elements(By.XPATH, "//span//span[starts-with(@class,'ProductDetailOrderBoxFrame_productDetailDeliveryMessage')]")

                            # Check if the second element is visible
                            if len(delivery_date_elements) > 1 and delivery_date_elements[1].is_displayed():
                                delivery_date_element = delivery_date_elements[1]  # Second visible element
                                delivery_date = delivery_date_element.text
                                # print(f"Delivery Date: {delivery_date}")
                            else:
                                # Handle the case when the second element is not visible
                                print("Second delivery date element is either not found or not visible.")


                        time.sleep(2)
                        extracted_date = int(datetime.now().timestamp())
                        address = get_address()

                    except Exception as e:
                        delivery_date = "Error"
                        extracted_date = int(datetime.now().timestamp())
                        address = "Address Error"
                        print(f"Error getting delivery date for {part}: {e}")

                    delivery_dates[part] = delivery_date
                    extracted_dates[part] = extracted_date
                    addresses[part] = address

                # Add new columns to the DataFrame
                df["Delivery Date"] = None
                df["Extracted Date"] = None
                df["Address"] = None

                # Write scraped data into the DataFrame
                for idx, row in df.iterrows():
                    for cell in row:
                        if isinstance(cell, str):
                            match = re.match(r"\d{4,5}[A-Z]\d{1,3}", cell.strip())
                            if match:
                                part = match.group()
                                if part in delivery_dates:
                                    df.at[idx, "Delivery Date"] = delivery_dates[part]
                                    df.at[idx, "Extracted Date"] = extracted_dates[part]
                                    df.at[idx, "Address"] = addresses[part]
                                    break  # Stop at the first matching part in the row

                # Clean the DataFrame header (remove 'Unnamed' columns)
                df = clean_dataframe(df)

        # Save the updated DataFrame for each sheet into a new Excel file with the same name as the input file
        output_path = os.path.join(output_folder, filename)

        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            # Save each DataFrame (sheet) to the new Excel file
            for sheet_name, df in all_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"✅ Updated Excel saved to: {output_path}")
