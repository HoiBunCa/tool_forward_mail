from threading import Thread

from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from concurrent.futures import ThreadPoolExecutor, ProcessPoolExecutor
from concurrent.futures import wait
from tkinter import *
from tkinter.filedialog import askopenfile
import time
import pandas as pd
import sqlite3
import requests


def init_mailbox(options, main_mail, main_pass):
    driver = webdriver.Firefox(options=options)
    driver.get("https://passport.yandex.ru/auth?retpath=https%3A%2F%2Fmail.yandex.ru%2F&backpath=https%3A%2F%2Fmail.yandex.ru%2F%3Fnoretpath%3D1&from=mail&origin=hostroot_homer_auth_ru")
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "passp-field-login"))).send_keys(main_mail)
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "passp:sign-in"))).click()
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "passp-field-passwd"))).send_keys(main_pass)
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "passp:sign-in"))).click()
    time.sleep(5)
    driver.get('https://mail.yandex.ru/?uid=1130000057343225#inbox')
    WebDriverWait(driver, 100).until(
        EC.element_to_be_clickable((By.XPATH, "//*[text()='Microsoft account security code']"))).click()
    print("Init Mail box browser")
    return driver


def page_has_loaded(browser):
    try:
        page_state = browser.execute_script('return document.readyState;')
        return page_state == 'complete'
    except:
        return False


def wait_until_page_success(browser):
    check_page_loaded = False
    for i in range(200):
        time.sleep(1)
        check_page_loaded = page_has_loaded(browser)
        if check_page_loaded:
            print("PAGE LOAD COMPLETE1")

            break
        else:
            pass
    if not check_page_loaded:
        print("PAGE LOAD ERROR1")

    return browser


def mask_as_read(driver, code_of_mail):
    for i in range(50):
        try:
            emails = driver.find_elements(By.XPATH,
                                          f'//div[contains(@class, "ns-view-messages-item-inner ")]/div/div/div/div[contains(@aria-label, "{code_of_mail}")]/div/a/div/span/span[contains(@title, "Mark as read")]')
            if len(emails):
                WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.XPATH,
                                                                           '//div[contains(@class, "ns-view-messages-item-inner ")]/div/div/div/div[contains(@aria-label, "{}")]/div/a/div/span/span[contains(@title, "Mark as read")]'.format(
                                                                               code_of_mail)))).click()
            break
        except Exception as e:
            print("Exception when mask_as_read: ", code_of_mail, e)


def get_code_by_selenium(driver, mail_input):
    code = None
    emails = None
    mail_s = mail_input.split("@")[0]
    mail_n = mail_s[0] + mail_s[1] + "**" + mail_s[-1]
    sent1 = "unread"
    sent2 = mail_n

    sent_all = f'//div[contains(@class, "ns-view-messages-item-inner ")]/div/div/div/div[contains(@aria-label, "{sent1}") and contains(@aria-label, "{sent2}")]'
    emails = driver.find_elements(By.XPATH, sent_all)
    if type(emails) == list:
        if len(emails) == 0:
            return None
        else:
            security_code_str = "Security code: "
            if_you_str = "If you"
            ind1 = emails[0].text.find(security_code_str)
            ind2 = emails[0].text.find(if_you_str)
            code = emails[0].text[ind1 + len(security_code_str): ind2].strip()
            mask_as_read(driver, code)
            return code
    else:
        return None


def init_test(options):
    driver = webdriver.Firefox(options=options)
    driver.get("https://www.google.com.vn/")
    return driver


def login_with_account(options, email_text, password_text):
    browser = webdriver.Firefox(options=options)
    browser.get('https://login.live.com')
    browser = wait_until_page_success(browser)
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "i0116"))).send_keys(email_text)
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "i0118"))).send_keys(password_text)
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
    return browser


def check_mail_used(mail):
    with sqlite3.connect('database_mail.db') as con:
        cur = con.cursor()
        record = cur.execute('''select * from mail_tabel_all where mail = "{}" '''.format(mail))

        if len(record.fetchall()):
            return "Mail used"
        else:
            return "New mail"


def reactive_mail(browser, browser_, email_protect_text):
    code = ""
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iProofEmail"))).send_keys(
        email_protect_text.split("@")[0])
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iSelectProofAction"))).click()
    for i in range(50):
        time.sleep(2)
        # code = get_verify_code(mail=email_protect_text, seen=False)
        code = get_code_by_selenium(driver=browser_, mail_input=email_protect_text)
        print("re_active mail: ", code)
        if code is not None:
            break
    if code == "":
        return browser
    browser = wait_until_page_success(browser)
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iOttText"))).send_keys(code)
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iVerifyCodeAction"))).click()


def reactive_mail2(browser, browser_, email_protect_text):
    code = ""
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iProofEmail"))).send_keys(
        email_protect_text.split("@")[0])
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iSelectProofAction"))).click()
    for i in range(50):
        time.sleep(2)
        # code = get_verify_code(mail=email_protect_text, seen=False)
        code = get_code_by_selenium(driver=browser_, mail_input=email_protect_text)
        print("re_active mail: ", code)
        if code is not None:
            break
    if code == "":
        return browser

    browser = wait_until_page_success(browser)
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iOttText"))).send_keys(code)
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iVerifyCodeAction"))).click()


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
    url = "https://chothuesimcode.com/api?act=code&apik=3fbef40e&id={}".format(id_phone)
    print(url)
    payload = {}
    headers = {}
    response = requests.request("GET", url, headers=headers, data=payload)
    print("response", response.json())
    if response.json()["Msg"] == "Đã nhận được code":
        return response.json()["Result"]["Code"], response.json()["ResponseCode"]
    elif response.json()["Msg"] == "không nhận được code, quá thời gian chờ":
        return -2, -2
    else:
        return -1, -1


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
                        EC.element_to_be_clickable((By.XPATH, "//*[text()='I didn’t get a code']"))).click()
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


def add_mail_protect(browser, browser_, email_protect_text):
    code = ""
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "EmailAddress"))).send_keys(email_protect_text)
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iNext"))).click()
    print("MAIL::: ", email_protect_text)
    for i in range(50):
        time.sleep(2)
        try:
            # code = get_verify_code(mail=email_protect_text, seen=False)
            code = get_code_by_selenium(driver=browser_, mail_input=email_protect_text)
            print("add_mail_protect::: ", code)
            if code is not None and code != -1:
                break
        except:
            pass

    if code == "":
        return browser

    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iOttText"))).send_keys(code)
    print("send_keys add_mail_protect code: ", code, email_protect_text)
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iNext"))).click()
    print("add_mail_protect done ::: ", email_protect_text)
    return browser


def process_security_email(browser, browser_, email_protect_text):
    code = ""
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "idDiv_SAOTCS_Proofs"))).click()
    print("click to receive process_security_email code: ", code, email_protect_text)
    WebDriverWait(browser, 100).until(
        EC.element_to_be_clickable((By.ID, "idTxtBx_SAOTCS_ProofConfirmation"))).send_keys(email_protect_text)
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "idSubmit_SAOTCS_SendCode"))).click()
    WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "footer")))
    browser = wait_until_page_success(browser)
    if "We couldn't send the code. Please try again." in browser.page_source:
        return browser

    for i in range(50):
        time.sleep(2)
        try:
            # code = get_verify_code(mail=email_protect_text, seen=False)

            code = get_code_by_selenium(driver=browser_, mail_input=email_protect_text)

            print("process_security_email:::", code, email_protect_text)
            if code is not None:
                break
        except:
            pass
    if code == "":
        print("No code: ", email_protect_text)
        return browser

    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "idTxtBx_SAOTCC_OTC"))).send_keys(code)
    print("send_keys process_security_email code: ", code, email_protect_text)
    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "idSubmit_SAOTCC_Continue"))).click()

    for i in range(20):
        time.sleep(2)
        check_loaded_page = page_has_loaded(browser)
        if check_loaded_page:
            break
        else:
            pass
    if "The wrong code was entered" in browser.page_source:
        browser = process_security_email(browser, browser_, email_protect_text)
    try:
        WebDriverWait(browser, 3).until(EC.element_to_be_clickable((By.ID, "iCancel"))).click()
    except Exception as e:
        pass
    print("process_security_email done::: ", email_protect_text)
    return browser


def enter_email_forwarding(browser, email_protect_text):
    WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "app")))
    browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")

    browser = wait_until_page_success(browser)

    try:
        WebDriverWait(browser, 2).until(EC.element_to_be_clickable((By.ID, "iCancel"))).click()
    except Exception as e:
        pass

    try:
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass

    try:
        print("00000000", email_protect_text)
        WebDriverWait(browser, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).is_enabled()

        try:
            WebDriverWait(browser, 3).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
        except:
            pass

        WebDriverWait(browser, 100).until(
            EC.element_to_be_clickable((By.XPATH, "//*[text()='Keep a copy of forwarded messages']"))).click()
        try:
            WebDriverWait(browser, 3).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
        except:
            pass

        WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Save']"))).click()

    except:
        print("1111111", email_protect_text)
        try:
            WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
        except Exception as e:
            pass

        try:
            WebDriverWait(browser, 200).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Enable forwarding']")))
            browser.execute_script("arguments[0].click();",
                                   browser.find_element(By.XPATH, "//*[text()='Enable forwarding']"))
        except Exception as e:
            print("#EEEEEEEEEEEE", email_protect_text, e)
        print("aaaaaaaaaaaaaaaaaaaaaa", email_protect_text)
        try:
            WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
        except Exception as e:
            pass

        WebDriverWait(browser, 200).until(
            EC.element_to_be_clickable((By.XPATH, "//*[text()='Keep a copy of forwarded messages']")))
        browser.execute_script("arguments[0].click();",
                               browser.find_element(By.XPATH, "//*[text()='Keep a copy of forwarded messages']"))
        try:
            WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
        except Exception as e:
            pass
        print("bbbbbbbbbbbbbbbbbbbb", email_protect_text, email_protect_text.split("@")[0])

        browser = wait_until_page_success(browser)
        for i in range(20):
            try:
                WebDriverWait(browser, 200).until(
                    EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys(
                    email_protect_text.split("@")[0])
                break
            except Exception as e:
                WebDriverWait(browser, 200).until(
                    EC.element_to_be_clickable((By.XPATH, "//*[text()='Enable forwarding']")))
                browser.execute_script("arguments[0].click();",
                                       browser.find_element(By.XPATH, "//*[text()='Enable forwarding']"))
                WebDriverWait(browser, 200).until(
                    EC.element_to_be_clickable((By.XPATH, "//*[text()='Keep a copy of forwarded messages']")))
                browser.execute_script("arguments[0].click();", browser.find_element(By.XPATH,
                                                                                     "//*[text()='Keep a copy of forwarded messages']"))
                print("Exception Enter an email address: ", str(e), email_protect_text)

        try:
            WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
        except Exception as e:
            pass
        print("cccccccccccccccccccc", email_protect_text)
        WebDriverWait(browser, 200).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("@")

        try:
            WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
        except Exception as e:
            pass
        print("ddddddddddddddddddddddd", email_protect_text)
        WebDriverWait(browser, 200).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("m")

        try:
            WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
        except Exception as e:
            pass
        print("eeeeeeeeeeeeeeeeeeee", email_protect_text)
        WebDriverWait(browser, 200).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("i")

        try:
            WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
        except Exception as e:
            pass
        print("ffffffffffffffffffff", email_protect_text)
        WebDriverWait(browser, 200).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("n")

        try:
            WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
        except Exception as e:
            pass
        print("ggggggggggggggggggggg", email_protect_text)
        WebDriverWait(browser, 200).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("h")

        try:
            WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
        except Exception as e:
            pass
        print("hhhhhhhhhhhhhhhhhhh", email_protect_text)
        WebDriverWait(browser, 200).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys(".")

        try:
            WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
        except Exception as e:
            pass
        print("iiiiiiiiiiiiiiiiiiii", email_protect_text)
        WebDriverWait(browser, 200).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("l")

        try:
            WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
        except Exception as e:
            pass
        print("kkkkkkkkkkkkkkkkkkkkkk", email_protect_text)
        WebDriverWait(browser, 200).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("i")

        try:
            WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
        except Exception as e:
            pass
        print("lllllllllllllllllllll", email_protect_text)
        WebDriverWait(browser, 200).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("v")

        try:
            WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
        except Exception as e:
            pass
        print("mmmmmmmmmmmmmmmmmmmmmmmm", email_protect_text)
        WebDriverWait(browser, 200).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys("e")

        try:
            WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
        except Exception as e:
            pass
        print("oooooooooooooooooooooooooo", email_protect_text)

        WebDriverWait(browser, 200).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Save']")))
        # todo: uncomment
        browser.execute_script("arguments[0].click();", browser.find_element(By.XPATH, "//*[text()='Save']"))

    print("SUCCESS add forwarding mail: ", email_protect_text)
    return browser


def insert_db(mail):
    try:
        with sqlite3.connect("database_mail.db") as con:
            cur = con.cursor()
            cur.execute("INSERT INTO mail_tabel_all VALUES ('{}') ".format(mail))
            con.commit()
        print("success --- insert into database: ", mail)
    except:
        print("error --- insert into database: ", mail)


class MyWindow:
    def __init__(self, win):
        self.pool = None
        self.win = win
        self.emails = []
        self.results = []
        self.diver1 = None

        self.main_email = "nhanmailao@minh.live"
        self.main_email_password = "Team1234@"
        self.api_key = "3fbef40e"
        self.num_worker = 1
        self.HEADLESS = False

        self.lbl1 = Label(win, text='File data')
        self.lbl1.place(x=50, y=50)
        self.lbl11 = Label(win, text='')
        self.lbl11.place(x=300, y=50)

        self.lbl2 = Label(win, text='Main mail')
        self.lbl2.place(x=50, y=100)
        self.MAIN_EMAIL = Entry(bd=3, width=30)
        self.MAIN_EMAIL.place(x=200, y=100)

        self.lbl3 = Label(win, text='Password main mail')
        self.lbl3.place(x=50, y=150)
        self.MAIN_EMAIL_PASSWORD = Entry(bd=3, width=30)
        self.MAIN_EMAIL_PASSWORD.place(x=200, y=150)

        self.lbl4 = Label(win, text='Key chothuesim')
        self.lbl4.place(x=50, y=200)
        self.API_KEY = Entry(bd=3, width=30)
        self.API_KEY.place(x=200, y=200)

        self.lbl5 = Label(win, text='Num thread')
        self.lbl5.place(x=50, y=250)
        self.NUM_WORKER = Entry(bd=3, width=30)
        self.NUM_WORKER.place(x=200, y=250)

        self.var = IntVar()
        self.lbl6 = Label(win, text='Headless')
        self.lbl6.place(x=50, y=300)
        self.HEADLESS_TRUE = Radiobutton(win, text="True", variable=self.var, value=True, command=self.sel)
        self.HEADLESS_TRUE.place(x=200, y=300)
        self.HEADLESS_FALSE = Radiobutton(win, text="False", variable=self.var, value=False, command=self.sel)
        self.HEADLESS_FALSE.place(x=350, y=300)

        self.lbl7 = Label(win, text='Firefox path')
        self.lbl7.place(x=50, y=350)
        self.FIREFOX_PATH = Entry(bd=3, width=30)
        self.FIREFOX_PATH.place(x=200, y=350)

        self.processBtn = Button(win, text='Process', command=self.main, state=DISABLED)
        self.processBtn.place(x=200, y=400)
        self.resultsLabel = Label(win, text='')
        self.resultsLabel.place(x=300, y=400)

        self.b2 = Button(win, text='Open file', command=lambda: self.open_file())
        self.b2.place(x=200, y=50)

        self.thread_check_fill_data = Thread(target=self.check_input_fill)
        self.thread_check_fill_data.daemon = True
        self.thread_check_fill_data.start()

    def sel(self):
        if self.var.get():
            self.HEADLESS = True
        else:
            self.HEADLESS = False

    def open_file(self):
        file = askopenfile(mode='r', filetypes=[('Python Files', '*.xlsx')])
        if file is not None:

            excel_data = pd.read_excel(file.name, "Data")
            data = pd.DataFrame(excel_data, columns=['email', 'password', 'email_protect'])
            for ind, row in data.iterrows():
                self.emails.append([row['email'], row['password'], row['email_protect']])
            print(excel_data)
            self.lbl11 = Label(self.win, text='Load OK!')
            self.lbl11.place(x=300, y=50)

    def check_input_fill(self):
        while True:
            if len(self.MAIN_EMAIL.get()) and len(self.MAIN_EMAIL_PASSWORD.get()) and len(self.API_KEY.get()) and len(
                    self.NUM_WORKER.get()) and len(self.FIREFOX_PATH.get()):
                self.processBtn['state'] = NORMAL
            else:
                self.processBtn['state'] = DISABLED

    def main(self):
        print(self.MAIN_EMAIL.get())
        print(self.MAIN_EMAIL_PASSWORD.get())
        print(self.API_KEY.get())
        print(self.NUM_WORKER.get())
        print(self.HEADLESS)
        print(str(self.FIREFOX_PATH.get().replace("\\", "/")))

        options = Options()
        options.set_preference('intl.accept_languages', 'en-GB')
        options.set_preference("permissions.default.image", 2)
        # options.binary_location = str(self.FIREFOX_PATH.get().replace("\\", "/"))
        # C:\Program Files\Mozilla Firefox\firefox.exe
        options.headless = False

        diver1 = None
        for i in range(10):
            try:
                diver1 = init_mailbox(options, self.MAIN_EMAIL.get(), self.MAIN_EMAIL_PASSWORD.get())
                # diver1 = init_test(options)
                break
            except Exception as e:
                print(e)
                pass

        def process(account):
            print("-" * 100)
            print(account)

            browser = None

            email_text = account[0]
            password_text = account[1]
            email_protect_text = account[2]

            # check email used
            mail_used = check_mail_used(email_text)
            if mail_used == "Mail used":
                status = "forwarded"
                print("forwarded: ", email_text)
            else:

                for i in range(20):
                    try:
                        browser = login_with_account(options, email_text, password_text)
                        break
                    except:
                        print("LOGIN FAIL")
                        pass
                print("LOGIN SUCCESS")
                # time.sleep(8)
                WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "footer")))
                if "We've detected something unusual about this sign-in. For example, you might be signing in from a new location, device or app." in browser.page_source:
                    try:
                        WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "proofDiv0"))).click()
                        reactive_mail(browser, diver1, email_protect_text)
                        status = "re_active mail1"
                    except:
                        status = "error"
                        browser.close()

                    time.sleep(5)
                elif "For your security and to ensure that only you have access to your account, we will ask you to verify your identity and change your password." in browser.page_source:
                    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iLandingViewAction"))).click()
                    try:
                        WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "proofDiv0"))).click()
                        reactive_mail2(browser, diver1, email_protect_text)
                        status = "change password"
                        password_text = password_text + "!"
                        WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iPassword"))).send_keys(
                            password_text)
                        WebDriverWait(browser, 100).until(
                            EC.element_to_be_clickable((By.ID, "iPasswordViewAction"))).click()
                        try:
                            WebDriverWait(browser, 100).until(
                                EC.element_to_be_clickable((By.ID, "iReviewProofsViewAction"))).click()
                        except:
                            pass
                        status = "re_active mail by phone number"
                    except:
                        status = "error"
                    time.sleep(2)
                    browser.close()
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
                            WebDriverWait(browser, 100).until(
                                EC.element_to_be_clickable((By.XPATH, "//input"))).send_keys(
                                phone_num)
                            WebDriverWait(browser, 100).until(
                                EC.element_to_be_clickable((By.XPATH, "//*[text()='Send code']"))).click()

                            phone_code_verify = process_code_verify(browser, phone_num_id, 0, 0)
                            print("phone_code_verify th1: ", phone_code_verify)
                            if phone_code_verify is None:
                                browser.close()
                                results = [email_text, password_text, email_protect_text, "Chua thue dc sdt"]
                                print("Chua thue dc sdt_1: {}".format(results))
                                return

                            WebDriverWait(browser, 100).until(EC.element_to_be_clickable(
                                (By.XPATH, "//input[@aria-label='Enter the access code']"))).send_keys(
                                phone_code_verify)
                            try:
                                WebDriverWait(browser, 100).until(
                                    EC.element_to_be_clickable((By.ID, "ProofAction"))).click()
                            except:
                                browser.close()
                                results = [email_text, password_text, email_protect_text, "Chua thue dc sdt"]
                                return
                            time.sleep(2)

                            try:
                                WebDriverWait(browser, 100).until(
                                    EC.element_to_be_clickable((By.ID, "FinishAction"))).click()
                            except Exception as e:
                                pass
                            try:
                                WebDriverWait(browser, 100).until(
                                    EC.element_to_be_clickable((By.ID, "idBtn_Back"))).click()
                            except Exception as e:
                                pass

                            time.sleep(2)
                            browser.get("https://outlook.live.com/mail/0/options/mail/layout")
                            try:
                                browser = add_mail_protect(browser, diver1, email_protect_text)
                            except Exception as e:
                                pass

                            WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "app")))
                            browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")

                            for i in range(50):
                                try:

                                    browser = wait_until_page_success(browser)

                                    WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "footer")))
                                    browser = process_security_email(browser, diver1, email_protect_text)
                                    break
                                except Exception as e:
                                    browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")
                                    pass
                            try:
                                WebDriverWait(browser, 3).until(
                                    EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
                            except Exception as e:
                                pass
                            WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "app")))
                            browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")

                            browser = wait_until_page_success(browser)

                            try:
                                browser = enter_email_forwarding(browser, email_protect_text)

                                status = "success"
                                # todo: uncomment
                                insert_db(email_text)
                                print("Lam xong truong hop 1: ", email_protect_text)
                            except Exception as e:
                                browser.save_full_page_screenshot(
                                    "html/{}.png".format(email_protect_text.split('@')[0]))
                                status = "Error: web can't loaded"
                                print("Truong hop 1 bi loi: ", email_protect_text)
                        except Exception as e:
                            browser.save_full_page_screenshot("html/{}.png".format(email_protect_text.split('@')[0]))
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
                            except:
                                print("Can't click sign in")

                            browser = wait_until_page_success(browser)

                            if "Help us protect your account" in browser.page_source:
                                try:
                                    browser = add_mail_protect(browser, diver1, email_protect_text)
                                    time.sleep(3)

                                    browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")

                                    browser = wait_until_page_success(browser)

                                    browser = process_security_email(browser, diver1, email_protect_text)
                                    time.sleep(3)
                                    print("Hoan thanh add email protect ", email_protect_text)
                                except:
                                    print("Chua hoan thanh add_mail_protect: ", email_protect_text)

                                    pass

                            else:
                                print("Nhap mail bao ve: ", email_protect_text)
                                browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")

                                browser = wait_until_page_success(browser)

                                for i in range(50):
                                    try:
                                        WebDriverWait(browser, 100).until(
                                            EC.presence_of_element_located((By.ID, "footer")))
                                        break
                                    except:
                                        browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")

                                browser = wait_until_page_success(browser)

                                try:
                                    browser = process_security_email(browser, diver1, email_protect_text)

                                    browser = wait_until_page_success(browser)

                                    if "We couldn't send the code. Please try again." in browser.page_source:
                                        browser.close()
                                        return [email_text, password_text, email_protect_text,
                                                "Couldn't send the code."]
                                    print("Hoan thanh process_security_email: ", email_protect_text)
                                except Exception as e:
                                    print(
                                        "Chua hoan thanh process_security_email: ",
                                        email_protect_text)
                                    pass

                            # them process_security_email vao day
                            browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")

                            browser = wait_until_page_success(browser)
                            # time.sleep(2)
                            print("11111111", email_protect_text)
                            try:
                                WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.XPATH, "//input[@title='Search settings']")))
                            except Exception as e:
                                print("AAAAAAa", e)
                            print("@222222222", email_protect_text)
                            if "Unable to load these settings. Please try again later" in browser.page_source:
                                print("Chua thue dc sdt: ", email_text)

                                browser.get(
                                    "https://account.live.com/names/manage?mkt=en-US&refd=account.microsoft.com&refp=profile")
                                for i in range(10):
                                    try:
                                        WebDriverWait(browser, 30).until(
                                            EC.presence_of_element_located((By.ID, "idAddPhoneAliasLink"))).click()
                                        break
                                    except:
                                        browser.get(
                                            "https://account.live.com/names/manage?mkt=en-US&refd=account.microsoft.com&refp=profile")
                                WebDriverWait(browser, 100).until(
                                    EC.presence_of_element_located((By.XPATH, "//select/option[@value='VN']"))).click()
                                phone_num_th23, phone_num_id_th23 = get_phone()
                                print("SDT truong hop 2-3: ", phone_num_th23, phone_num_id_th23)
                                WebDriverWait(browser, 100).until(
                                    EC.presence_of_element_located((By.ID, "DisplayPhoneNumber"))).send_keys(phone_num_th23)  # nhap sdt
                                WebDriverWait(browser, 100).until(
                                    EC.presence_of_element_located((By.ID, "iBtn_action"))).click()
                                phone_code_verify23 = process_code_verify23(browser, phone_num_id_th23, 0, 0)
                                print("phone_code_verify th 2-3: ", phone_code_verify23)
                                if phone_code_verify23 is None:
                                    browser.close()
                                    results = [email_text, password_text, email_protect_text, "Chua thue dc sdt"]
                                    print("Chua thue dc sdt_23: {}".format(results))
                                    return results

                                WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "iOttText"))).send_keys(phone_code_verify23)  # nhap code
                                WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "iBtn_action"))).click()

                                time.sleep(2)

                            elif "You can forward your email to another account." in browser.page_source:
                                print("Da thue duoc so dien thoai", email_text)
                                browser = enter_email_forwarding(browser, email_protect_text)
                                insert_db(email_text)
                                browser.close()
                                return [email_text, password_text, email_protect_text, "success"]
                            else:
                                browser.save_full_page_screenshot(
                                    "html/{}.png".format(email_protect_text.split('@')[0]))
                                browser.close()
                                return [email_text, password_text, email_protect_text, "error unknown"]

                            browser.close()

                            for i in range(50):
                                try:
                                    browser = login_with_account(options, email_text, password_text)
                                    break
                                except:
                                    pass
                            time.sleep(2)
                            browser.get("https://outlook.live.com/mail/0/options/mail/layout")
                            WebDriverWait(browser, 60).until(
                                EC.element_to_be_clickable((By.XPATH, "//*[text()='Sign in']"))).click()

                            WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "app")))
                            browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")

                            browser = wait_until_page_success(browser)

                            for i in range(50):
                                try:
                                    WebDriverWait(browser, 30).until(EC.presence_of_element_located((By.ID, "footer")))
                                    break
                                except:
                                    browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")

                            browser = wait_until_page_success(browser)

                            browser = process_security_email(browser, diver1, email_protect_text)

                            time.sleep(5)

                            browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")

                            browser = wait_until_page_success(browser)

                            for i in range(50):

                                if "Verify your identity" in browser.page_source:
                                    browser = process_security_email(browser, diver1, email_protect_text)
                                else:
                                    break

                            browser = wait_until_page_success(browser)

                            try:
                                browser = enter_email_forwarding(browser, email_protect_text)
                                status = "success"
                                # todo: uncomment
                                insert_db(email_text)
                                time.sleep(2)
                            except:
                                browser.save_full_page_screenshot(
                                    "html/{}.png".format(email_protect_text.split('@')[0]))
                                print("Error when add email forwarding", email_text)
                                status = "Error when add email forwarding"

                            try:
                                browser.get(
                                    "https://account.live.com/names/manage?mkt=en-US&refd=account.microsoft.com&refp=profile")
                                # xoa sdt

                                browser = wait_until_page_success(browser)

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
                            browser.save_full_page_screenshot("html/{}.png".format(email_protect_text.split('@')[0]))
                            status = "Error"
                            print("Truong hop 2-3 bi loi", email_text)


            print("Close browser!!!!!!!")
            results = [email_text, password_text, email_protect_text, status]
            return results
            # print("RESULTS: ", results)

        def test_process():
            print("test_process", diver1)
            print("done test_process!")
            diver1.close()

        results_list = []
        try:
            with ThreadPoolExecutor(max_workers=int(self.NUM_WORKER.get())) as pool:
                futures = [pool.submit(process, mail) for mail in self.emails]
                # wait(futures)
                for i in futures:
                    results_list.append(i.result())
        except:
            pass
        # with ThreadPoolExecutor(max_workers=int(self.NUM_WORKER.get())) as pool:
        #     future = [pool.submit(test_process)]

        print("All task done.")
        print("results_list", results_list)
        results_df = pd.DataFrame(results_list, columns=["email", "password", "email_protect", "status"])
        results_df.to_excel("results_{}.xlsx".format(time.time()))

        self.resultsLabel = Label(self.win, text='DONE!')
        self.resultsLabel.place(x=300, y=450)
        diver1.close()


if __name__ == '__main__':
    window = Tk()
    mywin = MyWindow(window)
    window.title('Tool Forward Mail')
    window.geometry("500x500+10+10")
    window.mainloop()