import sqlite3
from concurrent.futures import ThreadPoolExecutor
from tkinter import Label, NORMAL, DISABLED, Entry, Radiobutton, IntVar, Button
from tkinter.filedialog import askopenfile

import pandas as pd
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.wait import WebDriverWait

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
    pass


def buy_phone():
    """Call Api buy phone"""
    pass


def get_code_by_phone(phone_number, bill_id) -> str:
    """Get code from API buy phone"""
    pass


def mask_mail_as_read(driver, code_of_mail):
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


def get_code_by_mail(driver, email) -> str:
    pass


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


def setting_forward(driver):
    pass


def delete_phone(driver):
    """Use when all step is done"""
    pass


def create_main_mail_box(driver):
    pass


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


def load_mail_data(input_path) -> list:
    pass


def run_all_step_config_forward(mail, password, mail_protect):
    driver = create_driver()
    try:
        login_mail(driver, mail, password)

        process_after_login(driver)

        add_protect_mail(driver, mail)

        verify_by_mail(driver, mail)

        setting_forward(driver)

        delete_phone(driver)
    except LoginFailException as e:
        pass
    except Exception as e:
        pass
    finally:
        driver.close()


def process_all_mail(list_mails: list, num_processes: int = 1):
    driver_mail_mail = create_driver()
    create_main_mail_box(driver_mail_mail)

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
        input_path
        num_processes
        list_mails = load_mail_data(input_path)
        process_all_mail(list_mails, num_processes)
