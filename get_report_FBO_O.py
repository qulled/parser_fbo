import logging
from logging.handlers import RotatingFileHandler
from selenium import webdriver
from selenium.webdriver import DesiredCapabilities
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options
import pickle
import json
import datetime as dt
import random
import time
import os


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

options = Options()
# options.set_preference('permissions.default.stylesheet', 2)
# options.set_preference('permissions.default.image', 2)
# options.set_preference('dom.ipc.plugins.enabled.libflashplayer.so','false')
# options.set_preference("http.response.timeout", 10)
# options.set_preference("dom.max_script_run_time", 10)
options.set_preference("browser.download.folderList", 2)
options.set_preference("browser.download.manager.showWhenStarting", False)
options.set_preference("browser.download.dir", '~//wb/parser_fbo//excel_docs')
options.set_preference("browser.download.downloadDir", '~//wb//parser_fbo//excel_docs')
options.set_preference("browser.download.defaultFolder", '~//wb//parser_fbo//excel_docs')
options.headless = True





def auth(url,name):
    driver.get(url)
    cookies = pickle.load(open(f'cookies-{name}.py', 'rb'))
    for cookie in cookies:
        driver.add_cookie(cookie)
    time.sleep(30)
    attempt = driver.find_element(By.XPATH,'/html/body/div[10]/div/div/div/div[3]/button')
    attempt.click()
    return time.sleep(5)


def get_report(name):
    analitic_button = driver.find_element(By.XPATH,'/html/body/div[5]/div/div/div[2]/div/ul[2]/li[5]/div/a')
    analitic_button.click()
    time.sleep(3)
    sales_button = driver.find_element(By.XPATH,'/html/body/div[1]/div/div[2]/div/div[1]/div/div[2]/div[3]/button')
    sales_button.click()
    time.sleep(10)
    calendar_button = driver.find_element(By.XPATH,
                                          '/html/body/div[1]/div/div[2]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/button')
    calendar_button.click()
    time.sleep(random.randrange(2, 3))
    '''внесенеие данных сразу в форму'''
    find_date_start = driver.find_element(By.XPATH, '//*[@id="startDate"]')
    retrieval = f'{day}.{month}.{year}'
    find_date_start.send_keys(f'{retrieval}')
    time.sleep(random.randrange(3, 4))
    find_date_end = driver.find_element(By.XPATH, '//*[@id="endDate"]')
    find_date_end.send_keys(f'{retrieval}')
    time.sleep(random.randrange(3, 4))
    find_save_calendar = driver.find_element(By.XPATH,
                                             '/html/body/div[1]/div/div[2]/div/div[1]/div/div/div[2]/div[1]/div/div[2]/form/div[3]/div[2]/button/span')
    time.sleep(random.randrange(2, 3))
    find_save_calendar.click()
    time.sleep(10)
    save_report = driver.find_element(By.XPATH,
                                      '/html/body/div[1]/div/div[2]/div/div[1]/div/div/div[2]/div[3]/div[2]/div[1]/div/div[4]/div/button/span')
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
        elif len(i) > 25:
            file_oldname = os.path.join(dirparth, i)
            file_newname = os.path.join(dirparth, f'{name}-{year}-{month}-{day}.xlsx')
            os.rename(file_oldname, file_newname)
            break
    return


if __name__ == '__main__':
    day = dt.datetime.now().strftime('%d')
    month = dt.datetime.now().strftime("%m")
    year = dt.datetime.now().strftime("%Y")
    driver = webdriver.Firefox(options=options)
    dirparth = 'excel_docs/'
    try:
        name = 'Орлова'
        auth('https://seller.wildberries.ru/',name)
        get_report(name)
    except Exception as e:
        print(e)
    finally:
        driver.close()
        exit()
