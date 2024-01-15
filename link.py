from selenium import webdriver
from selenium.common.exceptions import *
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.webdriver.common.actions.action_builder import ActionBuilder
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support import ui
from selenium.webdriver.support.ui import Select
from time import sleep
from threading import Thread
import pyautogui
import json

def Find_Element(driver : webdriver.Chrome, by, value : str) -> WebElement:
    while True:
        try:
            element = driver.find_element(by, value)
            break
        except:
            pass
        sleep(0.1)
    return element

def Find_Elements(driver : webdriver.Chrome, by, value : str) -> list[WebElement]:
    while True:
        try:
            elements = driver.find_elements(by, value)
            if len(elements) > 0:
                break
        except:
            pass
        sleep(0.1)
    return elements

def Send_Keys(element : WebElement, content : str):
    element.clear()
    for i in content:
        element.send_keys(i)
        sleep(0.1)

def pressTab(count : int):
    for i in range(count):
        pyautogui.hotkey('tab')
        sleep(0.1)                    

def pressSpace():
    pyautogui.press('space')
    sleep(0.1)

def pressEnter():
    pyautogui.press('enter')
    sleep(1)

def pressDown(count : int):
    for i in range(count):
        pyautogui.press('down')
        sleep(0.1)

service = Service(executable_path = "C:\chromedriver-win64\chromedriver.exe")
options = Options()
options.add_experimental_option("debuggerAddress", "127.0.0.1:9030")
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.6099.217 Safari/537.3")
driver = webdriver.Chrome(options = options, service = service)
# driver.get('https://casadosdados.com.br/solucao/cnpj/pesquisa-avancada')

# state = Find_Element(driver, By.XPATH, '//*[@id="__nuxt"]/div/div[2]/section/div[3]/div[2]/div/div/div/div/div[1]/input')
# Send_Keys(state, 'São Paulo')
# pressDown(1)
# pressEnter()

# county = Find_Element(driver, By.XPATH, '//*[@id="__nuxt"]/div/div[2]/section/div[3]/div[3]/div/div/div/div/div[1]/input')
# Send_Keys(county, 'MOGI MIRIM')
# pressDown(1)
# pressEnter()

# cnaes = Find_Elements(driver, By.XPATH, '//*[@id="__nuxt"]/div/div[2]/section/div[2]/div[2]/section/div/div/div/div/div[2]/div')

# print(len(cnaes))
# print(cnaes[1].text)
output = []

links = Find_Elements(driver, By.XPATH, '//*[@id="__nuxt"]/div/div[2]/section/div[9]/div[1]/div/div/div/div/div')
for link in links:
    url = link.find_element(By.TAG_NAME, 'a').get_attribute('href')
    # print(url)
    output.append({"link" : url})
        
with open('output.json', 'w') as file:
    json.dump(output, file)

next_btn = Find_Element(driver, By.XPATH, '//*[@id="__nuxt"]/div/div[2]/section/div[8]/div/nav/a[2]')
driver.execute_script("arguments[0].click();", next_btn)

# //*[@id="__nuxt"]/div/div[2]/section/div[2]/div[2]/section/div/div/div/div/div[2]/div[1352]