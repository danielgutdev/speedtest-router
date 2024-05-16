import argparse
import configparser
import time
from datetime import datetime
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def create_unique_filename(base_path):
    """Generate a unique filename based on the current datetime."""
    current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{base_path}_{current_time}.xlsx"


def append_data_to_excel(file_path, data_dict):
    """Append a row of data to an Excel file."""
    try:
        df = pd.read_excel(file_path)
    except FileNotFoundError:
        df = pd.DataFrame(columns=data_dict.keys())

    new_row = pd.DataFrame([data_dict])
    df = pd.concat([df, new_row], ignore_index=True)
    df.to_excel(file_path, index=False)
    print(f"Data appended to {file_path}")


def save_dataframe_to_excel(df, filename, retries=3, delay=2):
    """Save a DataFrame to an Excel file with retry on failure."""
    for attempt in range(retries):
        try:
            df.to_excel(filename, index=False)
            print(f"Data successfully saved to {filename}.")
            break
        except PermissionError as e:
            print(f"Attempt {attempt + 1}: Unable to save the DataFrame - {e}")
            time.sleep(delay)
    else:
        print("Failed to save the DataFrame after several attempts.")


def load_configuration(file_path):
    """Load interval, router website URL, and loop count from configuration file."""
    config = configparser.ConfigParser()
    config.read(file_path)
    interval = config['Settings'].getint('Interval')
    router_website = config['Settings'].get('RouterWebsite')
    loop_count = config['Settings'].getint('LoopCount')
    return interval, router_website, loop_count


def parse_arguments():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(description='Run Speedtest.net tests at regular intervals.')
    parser.add_argument('-i', '--interval', type=int, help='Interval in seconds between tests')
    parser.add_argument('-l', '--loops', type=int, help='Number of loops to run')
    return parser.parse_args()


def setup_browser():
    """Setup and return a Selenium WebDriver instance."""
    browser = webdriver.Chrome()
    browser.maximize_window()
    browser.execute_script("window.open('');")
    browser.switch_to.window(browser.window_handles[0])
    return browser


def perform_speedtest(browser):
    """Perform a speed test on Speedtest.net and save a screenshot."""
    browser.get('https://www.speedtest.net')
    try:
        accept_button = WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.ID, "onetrust-accept-btn-handler")))
        accept_button.click()
        dismiss_button = WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.notification-dismiss.close-btn[title='Dismiss'][role='button']")))
        dismiss_button.click()
    except:
        pass

    go_button = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.CLASS_NAME, 'start-text')))
    go_button.click()
    WebDriverWait(browser, 60).until(EC.visibility_of_element_located((By.CLASS_NAME, 'result-label')))
    time.sleep(5)

    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    speedtest_screenshot_path = f'speedtest_result_{timestamp}.png'
    browser.save_screenshot(speedtest_screenshot_path)


def perform_router_status_check(browser, pathfile, timestamp, router_website):
    """Perform a status check on the router and save a screenshot."""
    browser.switch_to.window(browser.window_handles[1])
    browser.get(router_website)

    try:
        login_content = WebDriverWait(browser, 5).until(EC.visibility_of_element_located((By.CLASS_NAME, "pc-login-content")))
        password_input = WebDriverWait(browser, 2).until(EC.visibility_of_element_located((By.ID, "pc-login-password")))
        password_input.send_keys("admin1")
        login_button = WebDriverWait(browser, 2).until(EC.element_to_be_clickable((By.ID, 'pc-login-btn')))
        login_button.click()
        confirm_button = WebDriverWait(browser, 2).until(EC.element_to_be_clickable((By.ID, "confirm-yes")))
        confirm_button.click()
    except:
        pass

    advanced_button = WebDriverWait(browser, 5).until(EC.visibility_of_element_located((By.XPATH, "//span[@class='T_adv text' and contains(text(), 'Advanced')]")))
    advanced_button.click()
    time.sleep(2)
    scroll_element = WebDriverWait(browser, 5).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "label.label-title.m.T_c_band")))
    browser.execute_script("arguments[0].scrollIntoView(true);", scroll_element)
    time.sleep(1)

    router_screenshot_path = f'status_{timestamp}.png'
    browser.save_screenshot(router_screenshot_path)

    data = {
        'timestamp': timestamp,
        'nr_band': browser.find_element(By.ID, "nr_band").get_attribute('value'),
        'nr_ssrsrp': browser.find_element(By.ID, "nr_ssrsrp").get_attribute('value'),
        'lte_band': browser.find_element(By.ID, "lte_band").get_attribute('value'),
        'lte_rsrp': browser.find_element(By.ID, "lte_rsrp").get_attribute('value')
    }
    append_data_to_excel(pathfile, data)

    browser.switch_to.window(browser.window_handles[0])


def main():
    args = parse_arguments()
    interval, router_website, config_loop_count = load_configuration('config.ini')
    
    interval = args.interval or interval
    loop_count = args.loops or config_loop_count

    pathfile = create_unique_filename('report')

    browser = setup_browser()
    try:
        for _ in range(loop_count):
            timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
            perform_speedtest(browser)
            perform_router_status_check(browser, pathfile, timestamp, router_website)
            time.sleep(interval)
    finally:
        browser.quit()


if __name__ == '__main__':
    main()