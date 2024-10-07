# OK Select the Pascal view 
# OK Hit the button to review senders and receivers
# OK Analyze the parties 
# OK Decide if there will be a reply based on the forbiden @s list (clients and standart vendors)
# Fix the code flow to skip forbidden emails. I guess is that we have to separete the functions within the expand one and put them listed on the main() so that we ca break if the email domain is forbidden. 

import time
import logging
import openai
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium.webdriver.common.action_chains import ActionChains
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    WebDriverException
)
import random  # If you still want typing simulation delays

# ==============================
# Configuration and Setup
# ==============================

# OpenAI API Key
OPENAI_API_KEY = "The Key"

# Path to your Chrome user data directory
chrome_user_data_dir = r"C:\\Users\\55199\\AppData\\Local\\Google\\Chrome\\User Data"

# Specify the profile directory (e.g., "Profile 1")
chrome_profile_dir = "Default"  # Update with your profile directory

# Path to your ChromeDriver
chromedriver_path = r"C:\Users\55199\OneDrive\Área de Trabalho\chromedriver-win64\chromedriver-win64\chromedriver.exe"


# Initialize logging
logging.basicConfig(level=logging.INFO, format='%(message)s')

# Initialize the Chrome options
options = Options()
options.add_experimental_option("detach", True)
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument(f"--user-data-dir={chrome_user_data_dir}")
options.add_argument(f"--profile-directory={chrome_profile_dir}")

# Initialize the WebDriver
driver = None
try:
    driver_service = Service(chromedriver_path)
    driver = webdriver.Chrome(service=driver_service, options=options)
except Exception as e:
    logging.error(f"Error initializing ChromeDriver: {e}")

# Define your HubSpot Inbox URL
HUBSPOT_INBOX_URL = "https://app.hubspot.com/live-messages/21130119/inbox/8129652150"  
# Initialize OpenAI API
openai.api_key = OPENAI_API_KEY

# ==============================
# Updated Functions
# ==============================

def click_last_message_box():
    """
    Locates the last message container and clicks to select it.
    """
    try:
        # Step 1: Locate all message containers
        MESSAGE_CONTAINER_SELECTOR = "//div[contains(@class, 'CommonMessage__StyledContentColumn')]"
        logging.info("Locating all message containers...")

        message_containers = driver.find_elements(By.XPATH, MESSAGE_CONTAINER_SELECTOR)

        if message_containers:
            # Get the number of message containers
            num_containers = len(message_containers)
            logging.info(f"Found {num_containers} message containers.")

            # Step 2: Select the last message container (newest message)
            logging.info("Selecting the last message container (newest message)...")
            last_message_container = message_containers[-1]

            # Click the last message container to show the ellipsis button
            last_message_container.click()
            time.sleep(1)

            return last_message_container
        else:
            logging.warning("No message containers found.")
            return None
    except Exception as e:
        logging.error(f"Error while clicking the last message container: {e}")
        return None

def click_ellipsis_expand(last_message_container):
    """
    Clicks the ellipsis button inside the last message container to reveal full message history.
    """
    try:
        logging.info("Looking for the ellipsis expand button inside the last message container...")

        # Find all ellipsis buttons inside the last message container
        ellipsis_buttons = last_message_container.find_elements(
            By.XPATH,
            ".//button[.//span[@data-icon-name='ellipses']]"
        )

        if ellipsis_buttons:
            # Click the first ellipsis button found in the last message container
            ellipsis_button = ellipsis_buttons[0]
            logging.info("Ellipsis expand button found. Clicking it to reveal full message history.")
            driver.execute_script("arguments[0].click();", ellipsis_button)
            time.sleep(2)  # Wait for the full message history to load
        else:
            logging.warning("Ellipsis expand button not found inside the last message container.")
    except Exception as e:
        logging.error(f"Error while clicking the ellipsis expand button: {e}")

def is_email_forbidden(from_email):
    """
    Checks if the extracted 'From' email contains any forbidden domains.
    """
    forbidden_domains = [
        "@onfrontiers.com",   # Example of a forbidden domain
        "@client.com",        # Add other forbidden domains here
        "@vendor.com"         # Add more domains as needed
    ]

    for domain in forbidden_domains:
        if from_email.endswith(domain):
            return True
    return False

def wait_for_element(selector, by=By.CSS_SELECTOR, timeout=15):
    """
    Waits for an element to be present in the DOM and visible on the page.
    """
    try:
        element = WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located((by, selector))
        )
        return element
    except Exception:
        logging.error(f"Error: Element with selector '{selector}' not found within {timeout} seconds.")
        return None

def click_review_parties_button(last_message_container):
    """
    Locates and clicks the button to review the email involved parties (downCarat icon)
    within the provided last message container.
    """
    try:
        logging.info("Looking for the downCarat button to review email parties inside the last message container...")

        # Locate the button within the last message container by its data-icon-name
        review_button = last_message_container.find_element(By.XPATH, ".//span[@data-icon-name='downCarat']")

        if review_button:
            logging.info("Found the downCarat button inside the last message container. Clicking to review email parties.")
            review_button.click()
            time.sleep(2)  # Small delay to allow the section to expand
        else:
            logging.warning("Review button (downCarat) not found inside the last message container.")
    except NoSuchElementException:
        logging.error("Failed to find the downCarat button inside the last message container.")
    except Exception as e:
        logging.error(f"An error occurred while clicking the review parties button: {e}")

def get_from_email():
    """
    Locates and extracts the 'From' email after expanding the involved parties section.
    Also prints whether the email is from the forbidden list or not.
    """
    try:
        logging.info("Looking for the 'From' email address...")

        # Locate the element with data-message-metadata-type="from"
        from_element = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, "//dt[@data-message-metadata-type='from']//b"))
        )

        if from_element:
            # Extract the email address from the <b> tag
            from_email = from_element.text
            logging.info(f"Extracted 'From' email: {from_email}")

            # Check if the email is from a forbidden domain
            if is_email_forbidden(from_email):
                logging.info("Email from forbidden list, skipping reply.")
                return None  # Skip further processing if forbidden
            else:
                logging.info(f"Email {from_email} is not from forbidden list.")
                return from_email

        else:
            logging.warning("'From' email element not found.")
            return None

    except TimeoutException:
        logging.error("Failed to find the 'From' email element within the given time.")
        return None
    except Exception as e:
        logging.error(f"An error occurred while extracting the 'From' email: {e}")
        return None

def get_message_history():
    """
    Locate and return the full message history as a string, using the data-test-id attribute.
    """
    try:
        # Wait for the page to load completely
        time.sleep(2)

        # Use XPath to find the element with the data-test-id attribute
        MESSAGE_CONTAINER_XPATH = "//div[@data-test-id='primary-message-body-content']"

        # Locate the message container using the data-test-id attribute
        logging.info("Locating the message container to extract the full message...")
        message_container = wait_for_element(MESSAGE_CONTAINER_XPATH, by=By.XPATH)

        if message_container:
            # Extract the entire message from the container
            full_message = message_container.text.strip()

            logging.info(f"Full message extracted: {full_message}")
            return full_message
        else:
            logging.warning("Message container not found.")
            return None

    except Exception as e:
        logging.error(f"Error while getting message history: {e}")
        return None

def categorize_message(message_history):
    """
    Categorizes the message into one of the two categories:
    1. Technical issues and blockers
    2. Interested/Not interested/Questions
    """
    try:
        # Adjust the categorization logic based on the subjects provided in the table
        system_instructions = '''
        You are an assistant that categorizes messages into the following categories:
        1. Technical issues or blockers
        2. Interested, not interested, or questions

        Here are some examples to guide your categorization:

        Category 1:
        - Login issues (The platform is not accepting email or phone number)
        - Technical problems (bugs) (Technical issues blocking expert from applying)
        - Rescheduling (Expert needs to reschedule a call ASAP or not)
        - Not getting paid (Expert didn't get paid yet)

        Category 2:
        - Interested (Asking how to apply, Saying will apply soon, Will take a look and revert soon)
        - Not interested (Don't want to apply, Is not interested)
        - Asking for information about the client (Wants to know who the client is, Needs to know their name or area of work)
        - Conflict of interest worries (Is worried about getting into a conflict of interest, Needs to check with their company first)
        - How to get paid (Payment process, Amount, Method)
        - Call confirmation (Expert is not sure how to confirm a call)

        Provide only the category number (1 or 2) based on the message content.
        '''
        
        user_prompt = f"Categorize the following message:\n\n{message_history}\n\nCategory Number:"

        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": system_instructions},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0,
            max_tokens=100
        )

        category_number = response.choices[0].message["content"].strip()

        # Validate the category number
        if category_number not in ["1", "2"]:
            logging.warning(f"Unexpected category number received: '{category_number}'. Defaulting to 2.")
            category_number = "2"  # Default to Category 2 for general messages or error

        logging.info(f"Message categorized as: {category_number}")

        return category_number

    except Exception as e:
        logging.error(f"Error categorizing message: {e}")
        return "2"  # Default to Category 2 on error

def generate_response(message_history, category_number):
    """
    Generates a response based on the message history and the category determined by the categorize_message function.
    """
    try:
        # Define category-specific reply instructions based on the table
        if category_number == "1":
            assistant_instructions = '''
                You are a helpful assistant replying to an expert who is experiencing technical issues with our platform.

                Instructions:
                - For login issues or technical problems, request more details and a screenshot, and assure them our team will assist soon.
                - For rescheduling requests, ask for more details about their availability.
                - For payment delays, inform them that our finance team will be in touch and to check for emails from yzhang@onfrontiers.com.
            '''
        elif category_number == "2":
            assistant_instructions = '''
                You are a helpful assistant replying to an expert with general inquiries or interest in a project.

                Instructions:
                - For interested experts (e.g., asking how to apply or saying will apply soon), provide the project link and encourage their application.
                - For not interested experts, kindly ask for referrals and suggest they sign up as an expert for future opportunities.
                - For questions about the client or conflict of interest, respond based on the details available in the project description or assure them they can check with their company.
                - For payment inquiries, explain the process, amount, and method, and provide relevant links or contact emails.
                - For call confirmation questions, guide them on how to confirm or reschedule a call using the provided OnFrontiers platform link.
            '''
        
        # Prepare the messages for the API call
        messages = [
            {"role": "system", "content": assistant_instructions},
            {"role": "user", "content": message_history}
        ]

        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=messages,
            temperature=0,
            max_tokens=500,
            frequency_penalty=0,
            presence_penalty=0
        )

        response_text = response.choices[0].message["content"].strip()
        return response_text

    except Exception as e:
        logging.error(f"Error generating response: {e}")
        return None

def log_message_to_sheet(worksheet, message_history, response_text, category_number, message_link):
    """
    Logs the data to Google Sheets.
    """
    try:
        # Prepare the data to be inserted
        data_row = [category_number, message_link, message_history, response_text]
        # Insert the data at the top of the sheet (row 1)
        worksheet.insert_row(data_row, 2)
        logging.info("Data logged to Google Sheet.")
    except Exception as e:
        logging.error(f"An error occurred while logging the data to the sheet: {e}")

def setup_google_sheets(sheet_name, worksheet_name):
    """
    Setup Google Sheets API.
    """
    try:
        # Define the scope
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

        # Provide the path to your credentials JSON file
        creds = ServiceAccountCredentials.from_json_keyfile_name(r"C:\Users\55199\OneDrive\Área de Trabalho\Automations\Pascal\pascal-436120-736e1b83f257.json", scope)

        # Authorize the client
        client = gspread.authorize(creds)

        # Open the Google Sheet by name
        sheet = client.open(sheet_name)

        # Select the worksheet by name
        worksheet = sheet.worksheet(worksheet_name)

        return worksheet
    except Exception as e:
        logging.error(f"Error setting up Google Sheets: {e}")
        return None

def type_message_with_javascript(driver, message):
    try:
        # Use the CSS selector we have already identified
        EDITABLE_MESSAGE_BOX_SELECTOR = "div.ProseMirror[contenteditable='true']"

        logging.info("Waiting for the editable message box to be visible...")
        # Locate the message box
        editable_message_box = WebDriverWait(driver, 15).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, EDITABLE_MESSAGE_BOX_SELECTOR))
        )
        logging.info("Editable message box found.")

        # Use JavaScript to click and focus the message box
        driver.execute_script("arguments[0].focus(); arguments[0].click();", editable_message_box)
        logging.info("Clicked and focused on the editable message box using JavaScript.")

        # Optional: Add a delay to ensure the element is ready for typing
        time.sleep(1)

        # Clear any existing text using JavaScript (optional, if needed)
        driver.execute_script("arguments[0].innerHTML = '';", editable_message_box)

        # Type the message using JavaScript
        driver.execute_script("arguments[0].innerText = arguments[1];", editable_message_box, message)
        logging.info("Message typed using JavaScript.")

    except TimeoutException:
        logging.error("Timed out waiting for the editable message box.")
        driver.save_screenshot("timeout_editable_message_box.png")
    except Exception as e:
        logging.error(f"An unexpected error occurred while typing the message: {e}")
        driver.save_screenshot("unexpected_error_typing.png")

def send_message():
    """
    Locate the send button and click it to send the message.
    """
    try:
        SEND_BUTTON_SELECTOR = "i18n-string[data-key='composer-ui.send-button.capabilities-editor-submit-button.send']" 

        logging.info("Looking for the send button...")
        send_button = wait_for_element(SEND_BUTTON_SELECTOR)

        if send_button:
            logging.info("Send button found. Clicking it to send the message.")
            send_button.click()
            time.sleep(2)  # Small delay to allow the message to send
        else:
            logging.warning("Send button not found.")
    except Exception as e:
        logging.error(f"Error while trying to send the message: {e}")

def click_new_dropdown_item():
    try:
        # Wait for the dropdown element to be visible
        logging.info("Waiting for the dropdown element with label 'New'...")
        
        # Locate the element by its class and visible text
        new_item = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//span[@class='private-dropdown__item__label' and text()='New']"))
        )
        
        # Click on the element
        logging.info("Clicking the 'New' dropdown item.")
        new_item.click()
        
    except TimeoutException:
        logging.error("Failed to find the 'New' dropdown item within the given time.")
    except Exception as e:
        logging.error(f"An error occurred while clicking the 'New' dropdown item: {e}")

def select_closed_option():
    try:
        # Wait for the 'Closed' option to appear in the dropdown
        logging.info("Waiting for the 'Closed' option to appear in the dropdown...")
    
        # Locate the "Closed" option in the dropdown by its visible text
        closed_option = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//span[@class='private-dropdown__item__label' and text()='Closed']"))
        )
        
        # Click the "Closed" option
        logging.info("Clicking the 'Closed' option.")
        closed_option.click()
        
    except Exception as e:
        logging.error(f"An error occurred while selecting the 'Closed' option: {e}")

def select_waiting_on_us_option():
    """
    Locates and clicks the 'Waiting on us' option in the dropdown.
    """
    try:
        # Wait for the 'Waiting on us' option to appear in the dropdown
        logging.info("Waiting for the 'Waiting on us' option to appear in the dropdown...")

        # Locate the "Waiting on us" option in the dropdown by its visible text
        waiting_on_us_option = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//span[@class='private-dropdown__item__label' and text()='Waiting on us']"))
        )

        # Click the "Waiting on us" option
        logging.info("Clicking the 'Waiting on us' option.")
        waiting_on_us_option.click()

    except Exception as e:
        logging.error(f"An error occurred while selecting the 'Waiting on us' option: {e}")

# ==============================
# Updated Main Function
# ==============================

def main():
    # Initialize the Google Sheet
    worksheet = setup_google_sheets('Pascal', 'Messages')

    try:
        logging.info("Opening the HubSpot Inbox once at the start...")
        driver.get(HUBSPOT_INBOX_URL)

        # Adding delay to ensure the page loads fully
        time.sleep(5)

        while True:
            try:
                # Re-locate the Email List Container
                EMAIL_LIST_SELECTOR = ".ReactVirtualized__Grid__innerScrollContainer"
                logging.info("Locating the email list container...")

                email_list_container = wait_for_element(EMAIL_LIST_SELECTOR)

                if not email_list_container:
                    logging.error("Failed to locate the email list container. Please check the selector.")
                    break

                # Fetch the first email item only (the one at the top of the list)
                EMAIL_ITEM_SELECTOR = ".PreviewCardStyles__PreviewCardWrapper-m1ryjz-11"
                logging.info("Fetching the first email item in the queue...")
                email_items = email_list_container.find_elements(By.CSS_SELECTOR, EMAIL_ITEM_SELECTOR)

                if not email_items:
                    logging.warning("No emails found in the inbox.")
                    time.sleep(5)  # Wait a bit before checking again for new emails
                    continue

                # Get the first email item (the newest one at the top)
                first_email_item = email_items[0]

                # Click the first email to open it
                first_email_item.click()

                # Adding delay to wait for the email content to load
                time.sleep(3)

                # Step 1: Click the last message container
                last_message_container = click_last_message_box()

                if not last_message_container:
                    logging.warning("No last message container found. Skipping email processing.")
                    continue

                # Step 2: Click to review email involved parties
                click_review_parties_button(last_message_container)

                # Step 3: Extract the "From" email after expanding the involved parties
                from_email = get_from_email()

                # Step 4: If the email is from the forbidden list, skip further processing
                if from_email is None or is_email_forbidden(from_email):
                    logging.info("Skipping this email as it is from the forbidden list. Selecting 'Waiting on us'.")
                    
                    # Hit the 'New' dropdown
                    click_new_dropdown_item()
                    time.sleep(2)

                    # Select 'Waiting on us'
                    select_waiting_on_us_option()
                    
                    # Wait for 2 seconds before processing the next email
                    time.sleep(2)
                    continue

                # Step 5: Click the ellipsis to expand the full message history
                click_ellipsis_expand(last_message_container)

                # Get the message history
                message_history = get_message_history()

                if message_history:
                    # Proceed with categorization and response generation
                    category_number = categorize_message(message_history)

                    # Generate a response
                    response_text = generate_response(message_history, category_number)

                    if response_text:
                        logging.info("Generated Response:")
                        logging.info(response_text)

                        # Type the message into the message box
                        #type_message_with_javascript(driver, response_text)
                        #time.sleep(10)

                        # Send the message
                        #send_message()

                        # Log the data to Google Sheets
                        #if worksheet:
                        #    log_message_to_sheet(worksheet, message_history, response_text, category_number, driver.current_url)
                        #else:
                        #    logging.warning("Worksheet not available; cannot log data.")
                        
                        # Close the ticket
                        #click_new_dropdown_item()
                        #time.sleep(2)

                        #select_closed_option()
                        #time.sleep(2)

                    else:
                        logging.warning("Failed to generate a response.")
                else:
                    logging.error("Failed to get message content.")

                # Wait a moment before checking the top email again (in case new emails have arrived)
                logging.info("Waiting 2 seconds before checking the next email at the top of the queue...")
                time.sleep(2)

            except Exception as e:
                logging.error(f"An error occurred while processing the email: {e}")
                continue

    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")

    finally:
        # Close the browser after completion
        logging.info("Closing the browser.")
        driver.quit()

if __name__ == "__main__":
    main()