from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC#
from selenium.webdriver.firefox.options import Options
from urllib.parse import urlparse, parse_qs
import pyperclip

# Initialize Firefox browser with default settings
options = Options()
options.add_argument("-headless")  # Run in headless mode

# Initialize the Firefox browser driver
driver = webdriver.Firefox(options=options)
wait = WebDriverWait(driver, 10)

# Replace these values with your login credentials
email = "modvwbo297@cardev.net"
password = "Test777#"

# URL to start the login process
start_url = "https://identity.vwgroup.io/oidc/v1/authorize?client_id=0fa5ae01-ebc0-4901-a2aa-4dd60572ea0e@apps_vw-dilab_com&scope=openid%20cars%20vin&response_type=code&redirect_uri=http%3A%2F%2Flocalhost%3A8080%2Foauth2%2Fcallback&state=state"

try:
    # Open the login page
    driver.get(start_url)

    # Wait for the email input field to appear and enter the email
    email_input = wait.until(EC.presence_of_element_located((By.NAME, "identifier")))
    email_input.send_keys(email)
    email_input.send_keys(Keys.RETURN)

    # Wait for the password input field to appear and enter the password
    password_input = wait.until(EC.presence_of_element_located((By.NAME, "password")))
    password_input.send_keys(password)
    password_input.send_keys(Keys.RETURN)

    # Wait for the redirect to the localhost URL and check the URL contains the expected callback
    wait.until(EC.url_contains("localhost:8080/oauth2/callback"))

    # Extract the URL with the token
    final_url = driver.current_url

    # Extract the token from the URL
    parsed_url = urlparse(final_url)
    token = parse_qs(parsed_url.query).get("code", [None])[0]
    print(f"swaggerUI_Token: {token}")
# Copy the token to the clipboard
    if token:
        pyperclip.copy(token)
        print("Token copied to clipboard.")

except Exception as e:
    print(f"An error occurred: {e}")

finally:
    # Close the browser
    driver.quit()
