import art, time, os, platform
from colorama import init, Fore
from termcolor import colored
import pandas as pd
import datetime
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import configparser 


FILE_LOC = "data.xlsx"
SRNO = 'SRNO'
NAME = 'Name'
PHONENUMBER = 'Phone Number'
MESSAGE = 'Message'
REPLACENAME = '{{name}}'
URL = 'https://web.whatsapp.com/'
WAITER_ELEMENT = "landing-title _3-XoE"

PHONE_NUMER_INPUT =  '' #'//*[@id="side"]/div[1]/div/div/div[2]/div/div[2]' #"//*[@id='side']/div[1]/div/div/div[2]/div/div[2]" #"//*[@id='side']/div[1]/div/label/div/div[2]"
PERSON_DIV = "//*[@id='pane-side']/div[1]/div/div/div[1]/div/div/div[1]/div/div/div/img" #"//*[@id='pane-side']/div[1]/div/div/div[1]/div/div/div[1]/div/div/img"
MESSAGE_INPUT = '' #"//*[@id='main']/footer/div[1]/div/span[2]/div/div[2]/div/div/div[1]" #"//*[@id='main']/footer/div[1]/div[1]/span[2]/div/div[2]/div/div/div/p"
SEND_BUTTON = "//*[@id='main']/footer/div[1]/div[2]/div/div[2]/button"
NAME = ''

#PHONE_NUMER_INPUT = "//*[@id='side']/div[1]/div/div/div[2]/div/div[2]"
#PERSON_DIV = "//*[@id='pane-side']/div[1]/div/div/div[1]/div/div/div[1]/div/div/img"
#MESSAGE_INPUT = "//*[@id='main']/footer/div[1]/div/span[2]/div/div[2]/div/div/div[1]"
#SEND_BUTTON = "//*[@id='main']/footer/div[1]/div[2]/div/div[2]/button"


config = configparser.ConfigParser()

try:
    with open('config_whatsapp.ini') as f:
        config.readfp(f)
        PHONE_NUMER_INPUT = config.get('whatsapp_parameters','PHONE_NUMER_INPUT')
        MESSAGE_INPUT = config.get('whatsapp_parameters','MESSAGE_INPUT')

except IOError:
    raise MyError()

driver = webdriver.Chrome('chromedriver.exe')
driver.implicitly_wait(10)
waiter = WebDriverWait(driver, 10)
data = []

def printData(message, type):
    if type == 'INFO':
        print('[' + colored(datetime.datetime.now().strftime('%H:%M:%S'), 'cyan') + '][' + colored('INFO', 'green') +  '] ' + message)
    elif type == 'WARNING':
        print('[' + colored(datetime.datetime.now().strftime('%H:%M:%S'), 'cyan') + '][' + colored('WARNING', 'yellow') +  '] ' + message)
    elif type == 'ERROR':
        print('[' + colored(datetime.datetime.now().strftime('%H:%M:%S'), 'cyan') + '][' + colored('ERROR', 'red') +  '] ' + message)

def read_data_from_excel():
    try:
        df = pd.read_excel(FILE_LOC)
        print(df)
        printData("Retrieving data from excel", 'INFO')
    except:
        printData("Excel 'data.xlsx' not found", 'ERROR')
    printData("Found {0} messages to be send".format(len(df.index)), 'INFO')
    for i in df.index:
        if '+' not in str(df[PHONENUMBER][i]):
            number = '+521' + str(df[PHONENUMBER][i])
        else:
            number = str(df[PHONENUMBER][i])
        output = {
            'SrNo': df[SRNO][i],
            'Name': df[NAME][i],
            'PhoneNumber': number,
            'Message': df[MESSAGE][i]
        }
        data.append(output)


def send_whatsapp_message():
    global driver
    driver.get(URL)
    printData("Loading site...", 'INFO')
    waiter.until(EC.title_is("WhatsApp"))
    printData("Site loaded successfully...", 'INFO')
    printData("Waiting for user to log in using WhatsApp Web", 'INFO')
    waitCounter = 0
    while 1:
        try:
            waiter.until(EC.presence_of_element_located((By.XPATH, "//canvas[@aria-label='Scan me!']")))
            waitCounter+=1
            if waitCounter%1000 == 0:
                printData("Waiting for user to log in...", 'WARNING')
        except:
            printData("Logged in to WhatsApp", 'INFO')
            break

    for entry in data:
        driver.find_element(By.XPATH,PHONE_NUMER_INPUT).send_keys(str(entry['PhoneNumber']))
        time.sleep(5)
        driver.find_element(By.XPATH,PHONE_NUMER_INPUT).send_keys(Keys.ENTER)
        time.sleep(5)
        driver.find_element(By.XPATH, MESSAGE_INPUT).send_keys(str(entry['Message']))
        time.sleep(5)
        driver.find_element(By.XPATH, MESSAGE_INPUT).send_keys(Keys.ENTER)
        time.sleep(5)
        printData("Mensaje Exitosamente enviado {0}, nombre: {1}".format(str(entry['PhoneNumber']), str(entry['Name'])), 'INFO')

if __name__ == '__main__':
    # Initialize colorama
    init()

    # Clear the screen
    if platform.system() == 'Windows':
        os.system('cls')
    else:
        os.system('clear')

    # Display ASCII art
    print(art.text2art("WhatsApp Sender"))
    print(Fore.CYAN + "\nImplemented By:" + Fore.RESET + "jgv")

    # Read data from 'data.xlsx' file
    read_data_from_excel()

    # Send whatsapp message 
    send_whatsapp_message()

    # Close chromedriver
    driver.close()