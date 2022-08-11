import sqlite3
from concurrent.futures import ThreadPoolExecutor
from tkinter import Label, NORMAL, DISABLED, Entry, Radiobutton, IntVar, Button
from tkinter.filedialog import askopenfile

import pandas as pd
import requests
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

PATH_FIREFOX = None


class LoginFailException(Exception):
    """Raise if can not log in to mail"""


def create_driver(path_firefox: str = None) -> webdriver.Firefox:
    options = Options()
    options.set_preference('intl.accept_languages', 'en-GB')
    options.set_preference("permissions.default.image", 2)
    if path_firefox is None:
        path_firefox = PATH_FIREFOX
    options.binary_location = path_firefox  # "C:/Program Files/Mozilla Firefox/firefox.exe"
    options.headless = False
    driver = webdriver.Firefox(options=options)
    return driver


def login_mail(driver, mail, password):
    driver.get('https://login.live.com')
    wait_until_page_success(driver)
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "i0116"))).send_keys(mail)
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "i0118"))).send_keys(password)
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()


def buy_phone():
    """Call Api buy phone"""
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


def get_code_by_phone(phone_number, bill_id):
    """Get code from API buy phone"""
    url = "https://chothuesimcode.com/api?act=code&apik=3fbef40e&id={}".format(bill_id)
    payload = {}
    headers = {}
    response = requests.request("GET", url, headers=headers, data=payload)
    print("response", phone_number, response.json())
    if response.json()["Msg"] == "Đã nhận được code":
        return response.json()["Result"]["Code"], response.json()["ResponseCode"]
    elif response.json()["Msg"] == "không nhận được code, quá thời gian chờ":
        return -2, -2
    else:
        return -1, -1


def mask_mail_as_read(driver, code_of_mail):
    for i in range(50):
        try:
            emails = driver.find_elements(By.XPATH, f'//div[contains(@class, "ns-view-messages-item-inner ")]/div/div/div/div[contains(@aria-label, "{code_of_mail}")]/div/a/div/span/span[contains(@title, "Mark as read")]')
            if len(emails):
                WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.XPATH, f'//div[contains(@class, "ns-view-messages-item-inner ")]/div/div/div/div[contains(@aria-label, "{code_of_mail}")]/div/a/div/span/span[contains(@title, "Mark as read")]'))).click()
            break
        except Exception as e:
            print("Exception when mask_as_read: ", code_of_mail, e)


def get_code_by_mail(driver, email):
    code = None
    emails = None
    mail_s = email.split("@")[0]
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
            mask_mail_as_read(driver, code)
            return code
    else:
        return None


def reactive_by_phone():
    """Use when mail is locked. After login"""
    pass


def check_mail_has_phone(driver):
    pass


def add_protect_mail(driver, mail):
    pass


def verify_by_mail(driver, mail):
    pass


def reactive_mail(driver):
    pass


def setting_forward(driver, email_protect_text):

    driver.get("https://outlook.live.com/mail/0/options/mail/forwarding")
    wait_until_page_success(driver)

    try:
        WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
    except Exception as e:
        pass

    try:
        print("00000000", email_protect_text)
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).is_enabled()

        try:
            WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
        except:
            pass

        WebDriverWait(driver, 100).until(
            EC.element_to_be_clickable((By.XPATH, "//*[text()='Keep a copy of forwarded messages']"))).click()
        try:
            WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
        except:
            pass

        WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Save']"))).click()

    except:
        print("1111111", email_protect_text)
        try:
            WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
        except Exception as e:
            pass

        try:
            WebDriverWait(driver, 200).until(
                EC.element_to_be_clickable((By.XPATH, "//*[text()='Enable forwarding']")))
            driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//*[text()='Enable forwarding']"))
        except Exception as e:
            print("#EEEEEEEEEEEE", email_protect_text, e)
        print("aaaaaaaaaaaaaaaaaaaaaa", email_protect_text)
        try:
            WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
        except Exception as e:
            pass

        WebDriverWait(driver, 200).until(
            EC.element_to_be_clickable((By.XPATH, "//*[text()='Keep a copy of forwarded messages']")))
        driver.execute_script("arguments[0].click();",
                               driver.find_element(By.XPATH, "//*[text()='Keep a copy of forwarded messages']"))
        try:
            WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
        except Exception as e:
            pass
        print("bbbbbbbbbbbbbbbbbbbb", email_protect_text, email_protect_text.split("@")[0])

        browser = wait_until_page_success(driver)
        for i in range(20):
            try:
                WebDriverWait(browser, 200).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys(email_protect_text.split("@")[0])
                break
            except Exception as e:
                WebDriverWait(browser, 200).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Enable forwarding']")))
                browser.execute_script("arguments[0].click();", browser.find_element(By.XPATH, "//*[text()='Enable forwarding']"))
                WebDriverWait(browser, 200).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Keep a copy of forwarded messages']")))
                browser.execute_script("arguments[0].click();", browser.find_element(By.XPATH,"//*[text()='Keep a copy of forwarded messages']"))
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



def delete_phone(driver):
    """Use when all step is done"""
    driver.get("https://account.live.com/names/manage?mkt=en-US&refd=account.microsoft.com&refp=profile")
    wait_until_page_success(driver)
    try:
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "idRemoveAssocPhone"))).click()
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "iBtn_action"))).click()
    except:
        pass


def create_main_mail_box(driver, main_mail, main_pass):

    driver.get("https://passport.yandex.ru/auth?retpath=https%3A%2F%2Fmail.yandex.ru%2F&backpath=https%3A%2F%2Fmail.yandex.ru%2F%3Fnoretpath%3D1&from=mail&origin=hostroot_homer_auth_ru")
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "passp-field-login"))).send_keys(main_mail)
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "passp:sign-in"))).click()
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "passp-field-passwd"))).send_keys(main_pass)
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "passp:sign-in"))).click()
    wait_until_page_success(driver)
    driver.get('https://mail.yandex.ru/?uid=1130000057343225#inbox')
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Microsoft account security code']"))).click()
    print("Init Mail box browser")


def page_has_loaded(driver):
    try:
        page_state = driver.execute_script('return document.readyState;')
        return page_state == 'complete'
    except:
        return False


def wait_until_page_success(driver):
    wait = WebDriverWait(driver, 200)
    wait.until(page_has_loaded(driver))


def insert_db(mail):
    try:
        with sqlite3.connect("database_mail.db") as con:
            cur = con.cursor()
            cur.execute("INSERT INTO mail_tabel_all VALUES ('{}') ".format(mail))
            con.commit()
        print("success --- insert into database: ", mail)
    except:
        print("error --- insert into database: ", mail)


def process_after_login(driver):
    pass



def run_all_step_config_forward(mail, password, mail_protect):
    driver = create_driver()
    try:
        login_mail(driver, mail, password)

        wait_until_page_success(driver)

        process_after_login(driver)

        wait_until_page_success(driver)

        add_protect_mail(driver, mail_protect)

        wait_until_page_success(driver)

        verify_by_mail(driver, mail)

        wait_until_page_success(driver)

        setting_forward(driver)

        wait_until_page_success(driver)

        delete_phone(driver)
    except LoginFailException as e:
        pass
    except Exception as e:
        pass
    finally:
        driver.close()


def process_all_mail(list_mails: list, num_processes: int = 1, main_mail: str = "", main_password: str = ""):
    driver_mail_mail = create_driver()
    create_main_mail_box(driver_mail_mail, main_mail, main_password)

    with ThreadPoolExecutor(max_workers=num_processes) as executor:
        futures = []
        for mail, password, mail_protect in list_mails:
            future = executor.submit(run_all_step_config_forward, mail, password, mail_protect)
            futures.append(future)

        for future in futures:
            print(future.result())


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

        num_processes = self.NUM_WORKER.get()
        main_mail = self.MAIN_EMAIL.get()
        main_password = self.MAIN_EMAIL_PASSWORD.get()

        process_all_mail(self.emails, int(num_processes), main_mail, main_password)
