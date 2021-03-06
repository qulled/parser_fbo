import json
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pickle
import os


def get_cookie_DynamicCode(url,name):
    driver.get(url)
    time.sleep(10)
    phone = driver.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[1]/div[1]/div[2]/section/div[2]/form/div[1]/div[2]/div/div[2]/input')
    phone.click()
    phone.send_keys('926 996-66-76')
    time.sleep(2)
    send_code = driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[1]/div[1]/div[2]/section/div[2]/form/div[2]/button/span')
    send_code.click()
    time.sleep(60)
    switch_button = driver.find_element(By.XPATH,
                                        '/html/body/div[5]/div/div/div[1]/div/div[2]/span/div/div/div/button/span[1]/span')
    driver.execute_script("arguments[0].click();", switch_button)
    time.sleep(1)
    if name == 'Белотелов':
        ip_button = driver.find_element(By.XPATH,
                                          '/html/body/div[5]/div/div/div[1]/div/div[2]/span/div/div/div/div/div/div[2]/div[1]/div/ul/li[1]/div')
    elif name == 'Орлова':
        ip_button = driver.find_element(By.XPATH,
                                          '/html/body/div[5]/div/div/div[1]/div/div[2]/span/div/div/div/div/div/div[2]/div[1]/div/ul/li[3]/div')
    elif name == 'Кулик':
        ip_button = driver.find_element(By.XPATH,
                                      '/html/body/div[5]/div/div/div[1]/div/div[2]/span/div/div/div/div/div/div[2]/div[1]/div/ul/li[2]/div')
    time.sleep(1)
    ip_button.click()
    time.sleep(10)
    switch_button.click()
    time.sleep(3)
    pickle.dump(driver.get_cookies(), open(f'cookies-{name}.py', 'wb'))
    return driver.close()

if __name__ == '__main__':
    cred_file = os.path.join('credentials.json')
    try:
        with open(cred_file, 'r', encoding="utf-8") as f:
            cred = json.load(f)
            for i in cred:
                print(i)
                if i != 'Савельева':
                    driver = webdriver.Firefox()
                    get_cookie_DynamicCode('https://seller.wildberries.ru/',i)
    except Exception as e:
        print(e)
    finally:
        driver.close()
        exit()
