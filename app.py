import time
import os
import shutil
import zipfile
import openpyxl
import streamlit as st
from io import BytesIO
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

# --- STREAMLIT UI SETUP ---
st.set_page_config(page_title="GST E-Invoice Automation", layout="wide")
st.title("GST E-Invoice Downloader")

# Initialize session states to handle Streamlit's top-down rerun behavior
if 'driver' not in st.session_state:
    st.session_state.driver = None
if 'captcha_needed' not in st.session_state:
    st.session_state.captcha_needed = False
if 'current_username' not in st.session_state:
    st.session_state.current_username = None

def setup_headless_chrome(download_dir):
    """Sets up an invisible Chrome browser for the cloud server."""
    chrome_options = Options()
    chrome_options.add_argument("--headless=new") # Crucial for Cloud Servers
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": False,
        "safebrowsing.disable_download_protection": True,
        "profile.default_content_setting_values.automatic_downloads": 1 
    }
    chrome_options.add_experimental_option("prefs", prefs)

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

# --- MAIN APP UI ---
uploaded_file = st.file_uploader("Upload GST Excel Data File", type=["xlsx", "xlsm"])

if uploaded_file:
    # Save the uploaded file to the server temporarily
    with open("temp_data.xlsx", "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    if st.button("Start Processing"):
        st.session_state.processing = True
        st.session_state.status_area = st.empty()

# --- PROCESSING LOGIC ---
if getattr(st.session_state, 'processing', False):
    
    # Setup Server Directories
    base_dir = os.path.dirname(os.path.abspath(__file__))
    temp_dl_dir = os.path.join(base_dir, "_temp_gst_downloads")
    final_output_dir = os.path.join(base_dir, "_final_outputs")
    os.makedirs(temp_dl_dir, exist_ok=True)
    os.makedirs(final_output_dir, exist_ok=True)

    wb = openpyxl.load_workbook("temp_data.xlsx")
    data_sheet = wb['Data']
    
    for row_idx in range(2, data_sheet.max_row + 1):
        username = data_sheet.cell(row=row_idx, column=1).value
        if not username: continue 
            
        current_status = str(data_sheet.cell(row=row_idx, column=7).value).strip()
        if current_status not in ["None", "", "File in Progress", "No e-invoice generated"]:
            st.write(f"Skipping Row {row_idx} ({username}): Status is '{current_status}'.")
            continue

        password = data_sheet.cell(row=row_idx, column=2).value
        fin_year = str(data_sheet.cell(row=row_idx, column=3).value).strip()
        quarter = str(data_sheet.cell(row=row_idx, column=4).value).strip()
        period = str(data_sheet.cell(row=row_idx, column=5).value).strip()
        
        # WE IGNORE EXCEL COLUMN F ON THE CLOUD. We route everything to our server folder.
        client_output_dir = os.path.join(final_output_dir, username)
        os.makedirs(client_output_dir, exist_ok=True)

        st.write(f"### Processing: {username} | {period}")

        # --- CAPTCHA LOGIN HANDLING ---
        if username != st.session_state.current_username:
            if st.session_state.driver is not None:
                st.session_state.driver.quit()

            st.write(f"Launching invisible browser for {username}...")
            driver = setup_headless_chrome(temp_dl_dir)
            st.session_state.driver = driver
            
            driver.get("https://services.gst.gov.in/services/login")
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "username")))
            driver.find_element(By.ID, "username").send_keys(username)
            driver.find_element(By.ID, "user_pass").send_keys(password)
            
            # 1. Take a screenshot of the CAPTCHA image
            time.sleep(1) # Let image load
            captcha_element = driver.find_element(By.XPATH, "//img[contains(@id, 'captcha')]")
            captcha_png = captcha_element.screenshot_as_png
            
            # 2. Display it on the Web Page
            st.image(captcha_png, caption="Please enter the CAPTCHA below:")
            
            # 3. Wait for User Input
            captcha_input = st.text_input(f"CAPTCHA for {username}:", key=f"cap_{username}")
            
            if captcha_input:
                driver.find_element(By.ID, "captcha").send_keys(captcha_input)
                # Click Login...
                st.write("Logging in...")
                # Add your login button click logic here
                
                st.session_state.current_username = username
            else:
                st.warning("Waiting for CAPTCHA input. Script paused.")
                st.stop() # Stops execution until user types the captcha

        # ---------------------------------------------------------
        # ALL OF YOUR EXISTING NAVIGATION AND DOWNLOAD LOGIC GOES HERE
        # (Services -> Returns -> Dropdowns -> Trigger / Phase 2)
        # Note: Change `safe_target_dir` to `client_output_dir`
        # ---------------------------------------------------------
        
        st.write(f"Completed {username} for {period}.")

    # --- FINAL ZIP CREATION ---
    if st.session_state.driver:
        st.session_state.driver.quit()
        
    st.success("All processes complete! Zipping files for you...")
    
    shutil.make_archive("Final_GST_Downloads", 'zip', final_output_dir)
    
    with open("Final_GST_Downloads.zip", "rb") as f:
        st.download_button(
            label="Download All Completed Files (ZIP)",
            data=f,
            file_name="Final_GST_Downloads.zip",
            mime="application/zip"
        )
    
    # Provide the updated Excel file back to the user
    wb.save("Updated_Data.xlsx")
    with open("Updated_Data.xlsx", "rb") as f:
        st.download_button("Download Updated Excel Status Sheet", f, "Updated_Data.xlsx")
