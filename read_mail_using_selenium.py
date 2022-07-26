from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common import TimeoutException
from selenium.webdriver.firefox.options import Options
from concurrent.futures import ThreadPoolExecutor
from concurrent.futures import wait
import time

options = Options()
options.set_preference('intl.accept_languages', 'en-GB')
options.set_preference("permissions.default.image", 2)
options.headless = False
pool = ThreadPoolExecutor(max_workers=1)

browser = webdriver.Firefox(options=options)

browser.get(
    "https://passport.yandex.ru/auth?retpath=https%3A%2F%2Fmail.yandex.ru%2F&backpath=https%3A%2F%2Fmail.yandex.ru%2F%3Fnoretpath%3D1&from=mail&origin=hostroot_homer_auth_ru")
WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "passp-field-login"))).send_keys(
    "nhanmailao@minh.live")
WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "passp:sign-in"))).click()
WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "passp-field-passwd"))).send_keys("Team123@")
WebDriverWait(browser, 100).until(EC.element_to_be_clickable((By.ID, "passp:sign-in"))).click()
time.sleep(5)
browser.get('https://mail.yandex.ru/?uid=1130000057343225#thread/t179862510118279515')


def page_has_loaded(browser):
    page_state = browser.execute_script('return document.readyState;')
    return page_state == 'complete'


WebDriverWait(browser, 100).until(EC.element_to_be_clickable(
    (By.XPATH, '//div[contains(@aria-label, "unread, Microsoft account team, Microsoft account security code,")]')))


def get_code_by_selenium(mail_input):
    emails = browser.find_elements(By.XPATH,
                                   '//div[contains(@aria-label, "unread, Microsoft account team, Microsoft account security code,")]')
    print(len(emails))

    code = ""
    mail_s = mail_input.split("@")[0]
    mail_n = mail_s[0] + mail_s[1] + "**" + mail_s[-1]
    for mail_unread in emails:
        if mail_n in mail_unread.text:
            mail_unread_text = mail_unread.text
            security_code_str = "Security code: "
            if_you_str = "If you"
            ind1 = mail_unread_text.find(security_code_str)
            ind2 = mail_unread_text.find(if_you_str)
            code = mail_unread_text[ind1 + len(security_code_str): ind2]
            print("code: ", code, mail_input)
    return code.strip()


list_mail = ["mashekpaganig@minh.live",
             "tirybeitere@minh.live",
             "dagatawhitisc@minh.live",
             "gangwadmany@minh.live",
             "casanashimol@minh.live"]

options.headless = False
pool = ThreadPoolExecutor(max_workers=5)


future = [pool.submit(get_code_by_selenium, mail) for mail in list_mail]
print("-" * 100)


# //div[contains(@class, "item-car")]
time.sleep(1000)
