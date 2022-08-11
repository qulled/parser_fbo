import datetime as dt
import logging
import os
import pickle
import time
from logging.handlers import RotatingFileHandler

from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium_stealth import stealth
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


options = Options()

prefs = {'download.default_directory': '/home/lotelove/wb/pars_fbo/excel_docs/'}

options.add_experimental_option('prefs', prefs)
options.add_argument("--disable-blink-features")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
options.add_argument('--headless')


driver = webdriver.Chrome(options=options,service=Service(ChromeDriverManager().install()))
driver.maximize_window()

driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
driver.execute_cdp_cmd('Network.setUserAgentOverride', {
    "userAgent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.5060.53 Safari/537.36'})

stealth(driver,
        languages=["ru-Ru", "en"],
        vendor="Google Inc.",
        platform="Win32",
        webgl_vendor="Intel Inc.",
        renderer="Intel Iris OpenGL Engine",
        fix_hairline=True,
        )


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
log_dir = os.path.join(BASE_DIR, 'logs/')
log_file = os.path.join(BASE_DIR, 'logs/get_report_FBO.log')
console_handler = logging.StreamHandler()
file_handler = RotatingFileHandler(
    log_file,
    maxBytes=100000,
    backupCount=3,
    encoding='utf-8'
)
logging.basicConfig(
    level=logging.CRITICAL,
    format='%(asctime)s, %(levelname)s, %(message)s',
    handlers=(
        file_handler,
        console_handler
    )
)


def auth(url, name):
    driver.get(url)
    cookies = pickle.load(open(f'/home/lotelove/wb/SERVICE/cookie/cookies-{name}.py', 'rb'))
    for cookie in cookies:
        driver.add_cookie(cookie)
    driver.execute_script('arguments[0].click();', WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/div/div/div/div[3]/button'))))
    return time.sleep(1)
    

def get_report(name):
    calendar_button = driver.find_element(By.CLASS_NAME,'Date-input__input__1Z3Ndswoo9')
    calendar_button.click()
    time.sleep(2)
    find_date_start = driver.find_element(By.XPATH, '//*[@id="startDate"]')
    retrieval = f'{day}.{month}.{year}'
    find_date_start.send_keys(f'{retrieval}')
    time.sleep(0.5)
    find_date_end = driver.find_element(By.XPATH, '//*[@id="endDate"]')
    find_date_end.send_keys(f'{retrieval}')
    time.sleep(0.5)
    find_save_calendar = driver.find_element(By.CLASS_NAME,'DatePickerMenu__save-button__2A3XXwbVBF')
    find_save_calendar.click()
    time.sleep(15)
    save_report = driver.find_element(By.CLASS_NAME,'Export-button__lEpcIK0gfJ')
    save_report.click()
    time.sleep(5)
    name_files = [f for f in os.listdir(dirparth)]
    for i in name_files:
        if f'{name}-{year}-{month}-{day}.xlsx' in name_files:
            path = os.path.join(dirparth, f'{name}-{year}-{month}-{day}.xlsx')
            if len(i) > 25:
                file_oldname = os.path.join(dirparth, i)
                file_newname = os.path.join(dirparth, f'{name}-{year}-{month}-{day}.xlsx')
                os.remove(path)
                os.rename(file_oldname, file_newname)
                break
        if len(i) > 25:
            file_oldname = os.path.join(dirparth, i)
            file_newname = os.path.join(dirparth, f'{name}-{year}-{month}-{day}.xlsx')
            os.rename(file_oldname, file_newname)
            break
    return


if __name__ == '__main__':
    day = (dt.datetime.now()-dt.timedelta(days=1)).strftime('%d')
    month = (dt.datetime.now()-dt.timedelta(days=1)).strftime("%m")
    year = (dt.datetime.now()-dt.timedelta(days=1)).strftime("%Y")
    dirparth = 'excel_docs/'
    try:
        name = 'Белотелов'
        auth('https://seller.wildberries.ru/analytics/sales',name)
        get_report(name)
    except Exception as e:
        print(e)
    finally:
        driver.quit()
        exit()
