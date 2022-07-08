print("STARTING CODE")
import argparse
import time
import os
import sqlite3

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
    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.ID, "i0116"))).send_keys(email_text)
    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.ID, "i0118"))).send_keys(password_text)
    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
    return browser


def process_security_email(browser, email_protect_text):

    WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.ID, "idDiv_SAOTCS_Proofs"))).click()
    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.ID, "idTxtBx_SAOTCS_ProofConfirmation"))).send_keys(email_protect_text)
    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.ID, "idSubmit_SAOTCS_SendCode"))).click()
    for i in range(10):
        time.sleep(2)
        code = get_verify_code(mail=email_protect_text, seen=False)
        print("process_security_email:", code, email_protect_text)
        if code is not None and code != -1:
            break

    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.ID, "idTxtBx_SAOTCC_OTC"))).send_keys(code)
    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.ID, "idSubmit_SAOTCC_Continue"))).click()
    try:
        WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.ID, "iCancel"))).click()
    except Exception as e:
        print(f"Exception: {e}")
    return browser


def add_mail_protect(browser, email_protect_text):
    WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.ID, "EmailAddress"))).send_keys(email_protect_text)
    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.ID, "iNext"))).click()
    for i in range(10):
        time.sleep(2)
        code = get_verify_code(mail=email_protect_text, seen=False)
        print("add_mail_protect", code)
        if code is not None and code != -1:
            break

    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.ID, "iOttText"))).send_keys(code)
    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.ID, "iNext"))).click()
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
    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Enable forwarding']"))).click()

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass
    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Keep a copy of forwarded messages']"))).click()

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass
    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys(email_protect_text.split("@")[0])

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass
    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("@")

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass
    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("m")

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass
    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("i")

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass
    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("n")

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass
    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("h")

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass
    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys(".")

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass
    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("l")

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass
    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("i")

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass
    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("v")

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass
    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("e")

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass
    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Save']"))).click()

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
                    WebDriverWait(browser, 3).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='I didn’t get a code']"))).click()
                except:
                    pass
                WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.XPATH, "//input"))).clear()
                time.sleep(2)
                phone_num1, phone_num_id1 = get_phone()
                print("rebuy_phone_th1", phone_num1, phone_num_id1)
                WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.XPATH, "//input"))).send_keys(phone_num1)
                WebDriverWait(browser, 50).until(
                    EC.element_to_be_clickable((By.XPATH, "//*[text()='Send code']"))).click()
                return process_code_verify(browser, phone_num_id, 0, nums1 + 1)
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
                WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.ID, "iBtn_close"))).click()
                WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.ID, "idAddPhoneAliasLink"))).click()
                WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.XPATH, "//select/option[@value='VN']"))).click()
                phone_num1, phone_num_id1 = get_phone()
                print("rebuy_SDT truong hop 2-3: ", phone_num1, phone_num_id1)
                WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.ID, "DisplayPhoneNumber"))).send_keys(phone_num1)
                WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.ID, "iBtn_action"))).click()
                return process_code_verify23(browser, phone_num_id, 0, nums1 + 1)
            return process_code_verify23(browser, phone_num_id, nums0 + 1, nums1)
    except Exception as e:
        print("Exception when process_code_verify: ", e)
        return None


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


def insert_db(mail):
    with sqlite3.connect("database_mail.db") as con:
        cur = con.cursor()
        cur.execute(f"INSERT INTO mail_tabel_all VALUES ('{mail}') ")
        con.commit()


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
    if mail_used == "Mail used":
        status = "forwarded"
    else:
        browser = login_with_account(email_text, password_text)
        time.sleep(1)
        try:
            # gap truong hop 1 chua lam
            WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.ID, "StartAction"))).click()
            print("gap truong hop 1 chua lam: ", email_text)
            try:
                WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.XPATH, "//select/option[@value='VN']"))).click()
                phone_num, phone_num_id = get_phone()
                print("SDT truong hop 1: ", phone_num, phone_num)
                WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.XPATH, "//input"))).send_keys(phone_num)
                WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Send code']"))).click()

                phone_code_verify = process_code_verify(browser, phone_num_id, 0, 0)
                print("phone_code_verify th1: ", phone_code_verify)
                if phone_code_verify is None:
                    browser.close()
                    results.append([email_text, password_text, email_protect_text, "Chua thue dc sdt"])
                    print(f"Chua thue dc sdt_1: {results}")
                    return

                WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.XPATH, "//input[@aria-label='Enter the access code']"))).send_keys(phone_code_verify)
                try:
                    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.ID, "ProofAction"))).click()
                except:
                    browser.close()
                    results.append([email_text, password_text, email_protect_text, "Chua thue dc sdt"])
                    return
                time.sleep(5)

                try:
                    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.ID, "FinishAction"))).click()
                except Exception as e:
                    pass
                try:
                    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.ID, "idBtn_Back"))).click()
                except Exception as e:
                    pass

                time.sleep(5)
                browser.get("https://outlook.live.com/mail/0/options/mail/layout")
                try:
                    browser = add_mail_protect(browser, email_protect_text)
                except Exception as e:
                    pass

                time.sleep(3)
                browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")

                try:
                    try:
                        browser = process_security_email(browser, email_protect_text)
                    except Exception as e:
                        pass
                    try:
                        WebDriverWait(browser, 30).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
                    except Exception as e:
                        pass
                    browser = enter_email_forwarding(browser, email_protect_text)

                    status = "success"
                    insert_db(email_text)
                    print("Lam xong truong hop 1: ", email_protect_text)
                except Exception as e:
                    print("Exception1: ", e)
                    status = "Error: web can't loaded"
                    print("Truong hop 1 bi loi: ", email_protect_text)
            except Exception as e:
                print("Exception1: ", e)
                status = "Error"
                print("Truong hop 1 bi loi: ", email_protect_text)

        except:
            try:
                # gap truong hop 2 hoac 3 chua lam
                browser.get("https://outlook.live.com/mail/0/options/mail/layout")
                print("gap truong hop 2 hoac 3 chua lam: ", email_text)
                WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Sign in']"))).click()

                try:
                    add_mail_protect(browser, email_protect_text)
                except:
                    print("Chua hoan thanh add_mail_protect hoac add_mail_protect ko xuat hien: ", email_protect_text)
                    pass

                try:
                    process_security_email(browser, email_protect_text)
                except Exception as e:
                    print("Chua hoan thanh process_security_email lan 1 hoac process_security_email ko xuat hien: ", email_protect_text)
                    pass

                try:
                    browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")
                except:
                    print("Khong click duoc Forwarding", email_text)
                time.sleep(5)
                try:
                    process_security_email(browser, email_protect_text)
                except Exception as e:
                    print("Chua hoan thanh process_security_email lan 2: ", email_protect_text)

                browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")
                time.sleep(5)

                if "Unable to load these settings. Please try again later" in browser.page_source:
                    print("Chua thue dc sdt: ", email_text)

                    browser.get("https://account.live.com/names/manage?mkt=en-US&refd=account.microsoft.com&refp=profile")
                    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.ID, "idAddPhoneAliasLink"))).click()
                    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.XPATH, "//select/option[@value='VN']"))).click()
                    phone_num23, phone_num_id23 = get_phone()
                    print("SDT truong hop 2-3: ", phone_num23, phone_num_id23)
                    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.ID, "DisplayPhoneNumber"))).send_keys(phone_num23)  # nhap sdt
                    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.ID, "iBtn_action"))).click()
                    phone_code_verify = process_code_verify23(browser, phone_num_id23, 0, 0)
                    print("phone_code_verify th 2-3: ", phone_code_verify)
                    if phone_code_verify is None:
                        browser.close()
                        results.append([email_text, password_text, email_protect_text, "Chua thue dc sdt"])
                        print(f"Chua thue dc sdt_23: {results}")
                        return

                    WebDriverWait(browser, 120).until(EC.element_to_be_clickable((By.ID, "iOttText"))).send_keys(phone_code_verify)  # nhap code
                    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.ID, "iBtn_action"))).click()

                    time.sleep(2)
                    browser.close()
                    browser = login_with_account(email_text, password_text)
                    time.sleep(1)

                    browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")
                    try:
                        WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Sign in']"))).click()
                    except:
                        pass
                    browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")

                    try:
                        browser = process_security_email(browser, email_protect_text)
                    except:
                        print("Chua hoan thanh process_security_email lan 3: ", email_protect_text)
                        pass

                browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")
                try:
                    browser = process_security_email(browser, email_protect_text)
                except:
                    print("Chua hoan thanh process_security_email lan 4: ", email_protect_text)
                    pass

                try:
                    WebDriverWait(browser, 30).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
                except Exception as e:
                    pass
                browser = enter_email_forwarding(browser, email_protect_text)

                insert_db(email_text)
                time.sleep(2)
                try:
                    browser.get("https://account.live.com/names/manage?mkt=en-US&refd=account.microsoft.com&refp=profile")
                    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.ID, "idRemoveAssocPhone"))).click()
                    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.ID, "iBtn_action"))).click()
                except:
                    print("Chua hoan thanh xoa sdt 23: ", email_protect_text)

                status = "success"
                # print("Lam xong truong hop 2-3: ", email_protect_text)
            except Exception as e:
                print("Exception23: ", e)
                status = "Error"
                print("Truong hop 2-3 bi loi", email_text)

        browser.close()
    results.append([email_text, password_text, email_protect_text, status])
    print("RESULTS: ", results)


if __name__ == '__main__':

    future = [pool.submit(process, mail) for mail in emails]
    wait(future)
    print("All task done.")
    results_df = pd.DataFrame(results, columns=["email", "password", "email_protect", "status"])
    results_df.to_excel(f"results_{time.time()}.xlsx")
