import sqlite3

import json
import selenium
from loguru import logger
from concurrent.futures import ThreadPoolExecutor
from tkinter import Label, NORMAL, DISABLED, Entry, Radiobutton, IntVar, Button
from tkinter.filedialog import askopenfile
import time

import pandas as pd
import requests
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

from tkinter import *
from tkinter.filedialog import askopenfile

PATH_FIREFOX = None


class LoginFailException(Exception):
    """Raise if can not log in to mail"""
    pass


class CantReceiveCodeFromEmailException(Exception):
    """Raise if can not receive code from email"""
    pass


class CounldntSendTheCode(Exception):
    """Raise if can not receive code from email"""
    pass


class CantReceiveCodeFromApiBuyPhone(Exception):
    """Raise if can not receive code from buy phone"""
    pass


class CantProcessBuyPhone(Exception):
    pass


def create_driver(path_firefox: str = None, headless=False) -> webdriver.Firefox:
    options = Options()
    options.set_preference('intl.accept_languages', 'en-GB')
    options.set_preference("permissions.default.image", 2)
    PATH_FIREFOX = "C:/Program Files/Mozilla Firefox/firefox.exe"
    if path_firefox is None:
        path_firefox = PATH_FIREFOX
    options.binary_location = path_firefox  # "C:/Program Files/Mozilla Firefox/firefox.exe"
    options.headless = headless
    driver = webdriver.Firefox(options=options)
    driver.maximize_window()
    return driver


def wait_until_page_success(driver):
    while not page_is_loading(driver):
        continue


def page_is_loading(driver):
    while True:
        x = driver.execute_script("return document.readyState")
        if x == "complete":
            return True
        else:
            yield False


def login_mail(driver, mail, password):
    try:
        driver.get('https://login.live.com')
        wait_until_page_success(driver)
        WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "i0116"))).send_keys(mail)
        WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
        WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "i0118"))).send_keys(password)
        WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
    except Exception as e:
        # driver.save_full_page_screenshot("html/{}.png".format(mail.split("@")[0]))
        logger.exception(mail, e)
        raise LoginFailException("Login fail")


def mask_mail_as_read(driver, code_of_mail):
    for i in range(50):
        try:
            emails = driver.find_elements(By.XPATH,
                                          f'//div/div/div/div[contains(@aria-label, "{code_of_mail}")]/div/a/div/span/span[contains(@title, "Mark as read")]')
            if len(emails):
                WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.XPATH,
                                                                           f'//div/div/div/div[contains(@aria-label, "{code_of_mail}")]/div/a/div/span/span[contains(@title, "Mark as read")]'))).click()
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
    # f'//div[contains(@class, "ns-view-messages-item-inner ")]/div/div/div/div[contains(@aria-label, "unread") and contains(@aria-label, "bo**m")]'
    sent_all = f'//div/div/div/div[contains(@aria-label, "{sent1}") and contains(@aria-label, "{sent2}")]'
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


def buy_phone(key_chothuesim):
    """Call Api buy phone"""
    url = f"https://chothuesimcode.com/api?act=number&apik={key_chothuesim}&appId=1017&carrier=Viettel"
    payload = {}
    headers = {}
    response = requests.request("GET", url, headers=headers, data=payload)
    if response.json()["Msg"] == "OK":
        num_phone = response.json()["Result"]["Number"]
        bill_id = response.json()["Result"]["Id"]
        return num_phone, bill_id
    else:
        return -1, -1


def get_code_by_phone(bill_id, key_chothuesim):
    """Get code from API buy phone"""
    url = f"https://chothuesimcode.com/api?act=code&apik={key_chothuesim}&id={bill_id}"
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


def process_code_verify(driver, bill_id, nums0=0, nums1=0, key_chothuesim=""):
    try:
        time.sleep(2)
        phone_code_verify, reponse_code = get_code_by_phone(bill_id, key_chothuesim)
        if reponse_code == 0:
            return phone_code_verify
        if nums1 == 4:
            return None
        if reponse_code == -1:
            if nums0 == 7:
                try:
                    WebDriverWait(driver, 3).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[text()='I didn’t get a code']"))).click()
                except:
                    pass
                WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.XPATH, "//input"))).clear()
                time.sleep(2)
                phone_num1, bill_id1 = buy_phone(key_chothuesim)
                WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.XPATH, "//input"))).send_keys(
                    phone_num1)
                WebDriverWait(driver, 50).until(
                    EC.element_to_be_clickable((By.XPATH, "//*[text()='Send code']"))).click()
                return process_code_verify(driver, bill_id1, 0, nums1 + 1, key_chothuesim)
            return process_code_verify(driver, bill_id, nums0 + 1, nums1, key_chothuesim)

    except Exception as e:
        print("Exception when process_code_verify11: ", e)
        return None


def process_code_verify23(browser, bill_id, nums0=0, nums1=0, key_chothuesim=""):
    try:
        time.sleep(2)
        phone_code_verify, reponse_code = get_code_by_phone(bill_id, key_chothuesim)

        if reponse_code == 0:
            return phone_code_verify
        if nums1 == 4:
            return None

        if reponse_code == -1:
            if nums0 == 7:
                browser.get("https://account.live.com/names/manage?mkt=en-US&refd=account.microsoft.com&refp=profile")
                WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.ID, "idAddPhoneAliasLink"))).click()
                WebDriverWait(browser, 50).until(
                    EC.element_to_be_clickable((By.XPATH, "//select/option[@value='VN']"))).click()
                phone_num1, bill_id1 = buy_phone(key_chothuesim)
                WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.ID, "DisplayPhoneNumber"))).send_keys(
                    phone_num1)
                WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.ID, "iBtn_action"))).click()
                return process_code_verify23(browser, bill_id1, 0, nums1 + 1, key_chothuesim)
            return process_code_verify23(browser, bill_id, nums0 + 1, nums1, key_chothuesim)
    except Exception as e:
        print("Exception when process_code_verify23: ", e)
        return None


def reactive_by_phone(driver, key_chothuesim):
    """Use when mail is locked. After login"""
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "StartAction"))).click()
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.XPATH, "//select/option[@value='VN']"))).click()
    phone_num, phone_num_id = buy_phone(key_chothuesim)
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.XPATH, "//input"))).send_keys(phone_num)
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Send code']"))).click()

    phone_code_verify = process_code_verify(driver, phone_num_id, 0, 0, key_chothuesim)
    if phone_code_verify is None:
        raise CantReceiveCodeFromApiBuyPhone("Cant receive code from api buy phone - reactive_by_phone")

    WebDriverWait(driver, 50).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@aria-label='Enter the access code']"))).send_keys(
        phone_code_verify)
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "ProofAction"))).click()

    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "FinishAction"))).click()
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "idBtn_Back"))).click()

    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "microsoft")))


def skip_popup_adv(driver):
    try:
        WebDriverWait(driver, 1).until(EC.element_to_be_clickable(
            (By.XPATH, '//div[contains(@id,"ModalFocusTrapZone")]/div[2]/div/div[2]/div/div[3]/button[2]')))
    except:
        pass


def go_to_page_forwarding(driver):
    print("go_to_page_forwarding")
    flag = 0
    try:
        driver.get("https://outlook.live.com/mail/0/options/mail/layout")
        skip_popup_adv(driver)
        WebDriverWait(driver, 50).until(EC.element_to_be_clickable(
            (By.XPATH, "//input[contains(@id, 'optionSearch')] | //a[contains(text(), 'Sign in')]")))
        if 'Try premium' in driver.page_source:
            WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.XPATH, "//a[text()='Sign in']"))).click()
        skip_popup_adv(driver)
        WebDriverWait(driver, 50).until(
            EC.element_to_be_clickable((By.XPATH, "//input[contains(@id, 'optionSearch')]")))
        skip_popup_adv(driver)
        WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.XPATH,
                                                                    "/html/body/div[4]/div[1]/div/div/div/div[2]/div[2]/div/div[2]/div/button[10]/span/span/span"))).click()
        # WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.XPATH, "//*[contains(@id, 'ModalFocusTrapZone')]/div[2]/div/div[3]/div[2]/div/div/p | //*[contains(@id, 'ModalFocusTrapZone')]/div[2]/div/div[3]/div[2])]")))
        skip_popup_adv(driver)
        WebDriverWait(driver, 50).until(
            EC.element_to_be_clickable((By.XPATH, "//*[contains(@id, 'ModalFocusTrapZone')]/div[2]/div/div[3]/div[2]")))
        driver.refresh()
        # btn Forwarding
        skip_popup_adv(driver)
        WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.XPATH,
                                                                    "/html/body/div[4]/div[1]/div/div/div/div[2]/div[2]/div/div[2]/div/button[10]/span/span/span"))).click()
        skip_popup_adv(driver)
        WebDriverWait(driver, 50).until(
            EC.element_to_be_clickable((By.XPATH, "//*[contains(@id, 'ModalFocusTrapZone')]/div[2]/div/div[3]/div[2]")))
        flag = 1
    except:
        driver.refresh()

    return flag


# WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.XPATH, '//div[contains(@id,"ModalFocusTrapZone")]/div[2]/div/div[2]/div/div[3]/button[2]')))

def enable_setting_forward_mail(driver, key_chothuesim):
    print("enable_setting_forward_mail")

    for i in range(10):
        flag = go_to_page_forwarding(driver)
        if flag:
            break

    # if "You can forward your email to another account." in driver.page_source:
    try:
        WebDriverWait(driver, 1).until(EC.element_to_be_clickable(
            (By.XPATH, "//*[contains(@id, 'ModalFocusTrapZone')]/div[2]/div/div[3]/div[2]/div/div/p")))
        print(111111111111111)
        print(222222222222222)
        pass
    except:
        # thue sdt
        print(3333333333333333)
        print(4444444444444444)
        driver.get("https://account.live.com/names/manage?mkt=en-US&refd=account.microsoft.com&refp=profile")
        for i in range(10):
            try:
                WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "idAddPhoneAliasLink"))).click()
                print("idAddPhoneAliasLink: ")
                break
            except:
                driver.get("https://account.live.com/names/manage?mkt=en-US&refd=account.microsoft.com&refp=profile")
        for i in range(10):
            try:
                WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//select/option[@value='VN']"))).click()
                break
            except:
                raise CantProcessBuyPhone("Cant Process Buy Phone")

        phone_num_th23, phone_num_id_th23 = buy_phone(key_chothuesim)
        WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "DisplayPhoneNumber"))).send_keys(
            phone_num_th23)  # nhap sdt
        WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "iBtn_action"))).click()
        phone_code_verify23 = process_code_verify23(driver, phone_num_id_th23, 0, 0, key_chothuesim)
        if phone_code_verify23 is None:
            raise CantReceiveCodeFromApiBuyPhone("Cant receive code from api buy phone - enable_setting_forward_mail")

        WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "iOttText"))).send_keys(
            phone_code_verify23)  # nhap code
        WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "iBtn_action"))).click()
        wait_until_page_success(driver)
        go_to_page_forwarding(driver)


def add_protect_mail(driver, mail, browser_):
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable(
        (By.XPATH, "//input[contains(@id, 'EmailAddress')] | //div[contains(@id, 'idDiv_SAOTCS_Proofs')]")))
    if not "Passwords can be forgotten or stolen. Just in case, add security info now to help you get back into your account if something goes wrong. " in driver.page_source:
        return

    code = None
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "EmailAddress"))).send_keys(mail)
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "iNext"))).click()
    print("MAIL::: ", mail)
    for i in range(50):
        time.sleep(2)
        try:
            code = get_code_by_mail(driver=browser_, email=mail)
            print("add_mail_protect::: ", code)
            if code is not None and code != -1:
                break
        except:
            pass
    if code is None:
        raise CantReceiveCodeFromEmailException("cant receive code from email")

    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "iOttText"))).send_keys(code)
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "iNext"))).click()
    print("add_mail_protect done ::: ", mail)


def verify_by_mail(driver, mail, browser_):
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "idDiv_SAOTCS_Proofs")))
    if not "idDiv_SAOTCS_Proofs" in driver.page_source:
        return

    code = None
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "idDiv_SAOTCS_Proofs"))).click()
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "idTxtBx_SAOTCS_ProofConfirmation"))).send_keys(
        mail)
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "idSubmit_SAOTCS_SendCode"))).click()
    WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.ID, "footer")))
    # driver.save_full_page_screenshot("html/{}.png".format("AAAAAAAA"))

    for i in range(5):
        time.sleep(1)
        if "We couldn't send the code. Please try again." in driver.page_source:
            driver.close()
            raise CounldntSendTheCode("We couldn't send the code. Please try again")

    for i in range(50):
        time.sleep(2)
        try:
            code = get_code_by_mail(driver=browser_, email=mail)
            print("process_security_email:::", code, mail)
            if code is not None:
                break
        except:
            pass
    if code is None:
        raise CantReceiveCodeFromEmailException("no have code")

    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "idTxtBx_SAOTCC_OTC"))).send_keys(code)
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "idSubmit_SAOTCC_Continue"))).click()

    wait_until_page_success(driver)

    if "The wrong code was entered" in driver.page_source:
        verify_by_mail(driver, mail, browser_)
    try:
        WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.ID, "iCancel"))).click()
    except Exception as e:
        pass
    print("process_security_email done::: ", mail)


def reactive_mail(driver, browser_, email_protect_text):
    code = None
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "iProofEmail"))).send_keys(
        email_protect_text.split("@")[0])
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "iSelectProofAction"))).click()
    for i in range(50):
        time.sleep(2)
        code = get_code_by_mail(driver=browser_, email=email_protect_text)
        print("re_active mail: ", code)
        if code is not None:
            break
    if code is None:
        raise CantReceiveCodeFromEmailException("no have code")

    wait_until_page_success(driver)
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "iOttText"))).send_keys(code)
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "iVerifyCodeAction"))).click()


def relogin(driver, mail, password):
    for i in range(50):
        try:
            login_mail(driver, mail, password)
            break
        except:
            pass
    time.sleep(2)


def setting_forward(driver, email_protect_text, mail, password, driver_main_mail, path_firefox, headless):
    go_to_page_forwarding(driver)
    try:
        WebDriverWait(driver, 1).until(EC.element_to_be_clickable(
            (By.XPATH, "//*[contains(@id, 'ModalFocusTrapZone')]/div[2]/div/div[3]/div[2]/div/div/p")))
    except:
        driver.close()
        driver = create_driver(path_firefox, headless=headless)
        relogin(driver, mail, password)
        go_to_page_profile(driver)
        verify_by_mail(driver, email_protect_text, driver_main_mail)

        go_to_page_forwarding(driver)

    try:
        WebDriverWait(driver, 1).until(EC.element_to_be_clickable(
            (By.XPATH, '//div[contains(@id,"ModalFocusTrapZone")]/div[2]/div/div[2]/div/div[3]/button[2]'))).click()
    except Exception as e:
        pass

    print("setting_forward: ", email_protect_text)
    try:
        WebDriverWait(driver, 1).until(EC.element_to_be_clickable(
            (By.XPATH, '//div[contains(@id,"ModalFocusTrapZone")]/div[2]/div/div[2]/div/div[3]/button[2]'))).click()
    except Exception as e:
        pass

    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.XPATH, '//*[contains(@id, "ModalFocusTrapZone")]/div[2]/div/div[3]/div[2]/div/div/div[1]/label/span'))).click()

    print("aaaaaaaaaaaaaaaaaaaaaa", email_protect_text)
    try:
        WebDriverWait(driver, 1).until(EC.element_to_be_clickable(
            (By.XPATH, '//div[contains(@id,"ModalFocusTrapZone")]/div[2]/div/div[2]/div/div[3]/button[2]'))).click()
    except Exception as e:
        pass

    WebDriverWait(driver, 50).until(
        EC.element_to_be_clickable((By.XPATH, '//*[contains(@id, "ModalFocusTrapZone")]/div[2]/div/div[3]/div[2]/div/div/div[2]/div[2]/label/span'))).click()

    print("bbbbbbbbbbbbbbbbbbbb", email_protect_text, email_protect_text.split("@")[0])
    try:
        WebDriverWait(driver, 1).until(EC.element_to_be_clickable(
            (By.XPATH, '//div[contains(@id,"ModalFocusTrapZone")]/div[2]/div/div[2]/div/div[3]/button[2]'))).click()
    except Exception as e:
        pass

    for i in range(20):
        try:
            WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.XPATH, '//*[contains(@id, "ModalFocusTrapZone")]/div[2]/div/div[3]/div[2]/div/div/div[2]/div/div/div/input'))).clear()
            WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.XPATH, '//*[contains(@id, "ModalFocusTrapZone")]/div[2]/div/div[3]/div[2]/div/div/div[2]/div/div/div/input'))).send_keys(email_protect_text.split("@")[0])
            break
        except Exception as e:
            WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.XPATH, '//*[contains(@id, "ModalFocusTrapZone")]/div[2]/div/div[3]/div[2]/div/div/div[1]/label/span'))).click()
            WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.XPATH, '//*[contains(@id, "ModalFocusTrapZone")]/div[2]/div/div[3]/div[2]/div/div/div[2]/div[2]/label/span'))).click()
            logger.exception(email_protect_text)

    print("cccccccccccccccccccc", email_protect_text)

    try:
        WebDriverWait(driver, 1).until(EC.element_to_be_clickable(
            (By.XPATH, '//div[contains(@id,"ModalFocusTrapZone")]/div[2]/div/div[2]/div/div[3]/button[2]'))).click()
    except Exception as e:
        pass

    WebDriverWait(driver, 50).until(
        EC.element_to_be_clickable((By.XPATH, '//*[contains(@id, "ModalFocusTrapZone")]/div[2]/div/div[3]/div[2]/div/div/div[2]/div/div/div/input'))).send_keys("@")

    print("ddddddddddddddddddddddd", email_protect_text)
    try:
        WebDriverWait(driver, 1).until(EC.element_to_be_clickable(
            (By.XPATH, '//div[contains(@id,"ModalFocusTrapZone")]/div[2]/div/div[2]/div/div[3]/button[2]'))).click()
    except Exception as e:
        pass

    WebDriverWait(driver, 50).until(
        EC.element_to_be_clickable((By.XPATH, '//*[contains(@id, "ModalFocusTrapZone")]/div[2]/div/div[3]/div[2]/div/div/div[2]/div/div/div/input'))).send_keys("m")

    print("eeeeeeeeeeeeeeeeeeee", email_protect_text)
    try:
        WebDriverWait(driver, 1).until(EC.element_to_be_clickable(
            (By.XPATH, '//div[contains(@id,"ModalFocusTrapZone")]/div[2]/div/div[2]/div/div[3]/button[2]'))).click()
    except Exception as e:
        pass

    WebDriverWait(driver, 50).until(
        EC.element_to_be_clickable((By.XPATH, '//*[contains(@id, "ModalFocusTrapZone")]/div[2]/div/div[3]/div[2]/div/div/div[2]/div/div/div/input'))).send_keys("i")

    print("ffffffffffffffffffff", email_protect_text)

    try:
        WebDriverWait(driver, 1).until(EC.element_to_be_clickable(
            (By.XPATH, '//div[contains(@id,"ModalFocusTrapZone")]/div[2]/div/div[2]/div/div[3]/button[2]'))).click()
    except Exception as e:
        pass

    WebDriverWait(driver, 50).until(
        EC.element_to_be_clickable((By.XPATH, '//*[contains(@id, "ModalFocusTrapZone")]/div[2]/div/div[3]/div[2]/div/div/div[2]/div/div/div/input'))).send_keys("n")

    print("ggggggggggggggggggggg", email_protect_text)

    try:
        WebDriverWait(driver, 1).until(EC.element_to_be_clickable(
            (By.XPATH, '//div[contains(@id,"ModalFocusTrapZone")]/div[2]/div/div[2]/div/div[3]/button[2]'))).click()
    except Exception as e:
        pass
    WebDriverWait(driver, 50).until(
        EC.element_to_be_clickable((By.XPATH, '//*[contains(@id, "ModalFocusTrapZone")]/div[2]/div/div[3]/div[2]/div/div/div[2]/div/div/div/input'))).send_keys("h")

    print("hhhhhhhhhhhhhhhhhhh", email_protect_text)

    try:
        WebDriverWait(driver, 1).until(EC.element_to_be_clickable(
            (By.XPATH, '//div[contains(@id,"ModalFocusTrapZone")]/div[2]/div/div[2]/div/div[3]/button[2]'))).click()
    except Exception as e:
        pass

    WebDriverWait(driver, 50).until(
        EC.element_to_be_clickable((By.XPATH, '//*[contains(@id, "ModalFocusTrapZone")]/div[2]/div/div[3]/div[2]/div/div/div[2]/div/div/div/input'))).send_keys(".")

    print("iiiiiiiiiiiiiiiiiiii", email_protect_text)

    try:
        WebDriverWait(driver, 1).until(EC.element_to_be_clickable(
            (By.XPATH, '//div[contains(@id,"ModalFocusTrapZone")]/div[2]/div/div[2]/div/div[3]/button[2]'))).click()
    except Exception as e:
        pass

    WebDriverWait(driver, 50).until(
        EC.element_to_be_clickable((By.XPATH, '//*[contains(@id, "ModalFocusTrapZone")]/div[2]/div/div[3]/div[2]/div/div/div[2]/div/div/div/input'))).send_keys("l")

    print("kkkkkkkkkkkkkkkkkkkkkk", email_protect_text)
    try:
        WebDriverWait(driver, 1).until(EC.element_to_be_clickable(
            (By.XPATH, '//div[contains(@id,"ModalFocusTrapZone")]/div[2]/div/div[2]/div/div[3]/button[2]'))).click()
    except Exception as e:
        pass

    WebDriverWait(driver, 50).until(
        EC.element_to_be_clickable((By.XPATH, '//*[contains(@id, "ModalFocusTrapZone")]/div[2]/div/div[3]/div[2]/div/div/div[2]/div/div/div/input'))).send_keys("i")

    print("lllllllllllllllllllll", email_protect_text)
    try:
        WebDriverWait(driver, 1).until(EC.element_to_be_clickable(
            (By.XPATH, '//div[contains(@id,"ModalFocusTrapZone")]/div[2]/div/div[2]/div/div[3]/button[2]'))).click()
    except Exception as e:
        pass

    WebDriverWait(driver, 50).until(
        EC.element_to_be_clickable((By.XPATH, '//*[contains(@id, "ModalFocusTrapZone")]/div[2]/div/div[3]/div[2]/div/div/div[2]/div/div/div/input'))).send_keys("v")

    print("mmmmmmmmmmmmmmmmmmmmmmmm", email_protect_text)
    try:
        WebDriverWait(driver, 1).until(EC.element_to_be_clickable(
            (By.XPATH, '//div[contains(@id,"ModalFocusTrapZone")]/div[2]/div/div[2]/div/div[3]/button[2]'))).click()
    except Exception as e:
        pass

    WebDriverWait(driver, 50).until(
        EC.element_to_be_clickable((By.XPATH, '//*[contains(@id, "ModalFocusTrapZone")]/div[2]/div/div[3]/div[2]/div/div/div[2]/div/div/div/input'))).send_keys("e")

    print("oooooooooooooooooooooooooo", email_protect_text)

    # todo: uncomment
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Save']"))).click()

    insert_db(mail)

    print("SUCCESS add forwarding mail: ", email_protect_text)
    return driver


def delete_phone(driver, path_firefox, mail, password, browser_):
    """Use when all step is done"""
    driver.get("https://account.live.com/names/manage?mkt=en-US&refd=account.microsoft.com&refp=profile")
    wait_until_page_success(driver)
    try:
        WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, "idRemoveAssocPhone"))).click()
        WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, "iBtn_action"))).click()
    except:
        pass
    driver.close()


def create_main_mail_box(driver, main_mail, main_pass):
    print("Start Init Mail box browser")
    driver.get(
        "https://passport.yandex.ru/auth?retpath=https%3A%2F%2Fmail.yandex.ru%2F&backpath=https%3A%2F%2Fmail.yandex.ru%2F%3Fnoretpath%3D1&from=mail&origin=hostroot_homer_auth_ru")
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "passp-field-login"))).send_keys(main_mail)
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "passp:sign-in"))).click()
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "passp-field-passwd"))).send_keys(main_pass)
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "passp:sign-in"))).click()
    WebDriverWait(driver, 100).until(
        EC.visibility_of_element_located((By.XPATH, '//a[@href="#inbox" and contains(@class, "Folder")]')))
    # WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Microsoft account security code']"))).click()

    # driver.get('https://mail.yandex.com#folder/10')
    driver.get(driver.current_url.replace("inbox", "folder/10"))
    WebDriverWait(driver, 100).until(
        EC.visibility_of_element_located((By.XPATH, '//a[@href="#inbox" and contains(@class, "Folder")]')))
    try:
        WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, '//div/div[contains(@aria-label, "Microsoft account Security code Please use the following security code for the Microsoft account ")]/div/a/div/span/div/span/span[contains(@class, "mail-ui-Arrow")]'))).click()
    except:
        print("Khong xuat hien mail tieng anh")

    try:
        WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, '//div/div[contains(@aria-label, "Tài khoản Microsoft Mã bảo mật Vui lòng sử dụng mã bảo mật sau cho tài khoản Microsoft ")]/div/a/div/span/div/span/span[contains(@class, "mail-ui-Arrow")]'))).click()
    except:
        print("Khong xuat hien mail tieng viet")

    print("End Init Mail box browser")
    return driver


def page_has_loaded(driver):
    try:
        page_state = driver.execute_script('return document.readyState;')
        return page_state == 'complete'
    except:
        return False


def check_mail_used(mail):
    with sqlite3.connect('database_mail.db') as con:
        cur = con.cursor()
        record = cur.execute('''select * from mail_tabel_all where mail = "{}" '''.format(mail))

        if len(record.fetchall()):
            return "Mail used"
        else:
            return "New mail"


def insert_db(mail):
    try:
        with sqlite3.connect("database_mail.db") as con:
            cur = con.cursor()
            cur.execute("INSERT INTO mail_tabel_all VALUES ('{}') ".format(mail))
            con.commit()
        print("success --- insert into database: ", mail)
    except Exception as e:
        print("error --- insert into database: ", mail, e)


def process_after_login(driver, driver1, mail):
    wait_until_page_success(driver)

    if "We've detected something unusual about this sign-in. For example, you might be signing in from a new location, device or app." in driver.page_source:
        WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "proofDiv0"))).click()
        reactive_mail(driver, driver1, mail)
    if "For your security and to ensure that only you have access to your account, we will ask you to verify your identity and change your password." in driver.page_source:
        pass
    # if "Stay signed in so you don't have to sign in again next time." in driver.page_source:


def process_after_login_case(driver):
    wait_until_page_success(driver)
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "footer")))

    pass_screen_update(driver)

    # driver.save_full_page_screenshot("html/{}.png".format("AAAAAAAA"))

    if "We've detected something unusual about this sign-in. For example, you might be signing in from a new location, device or app." in driver.page_source:
        # truong hop kich hoat lai mail bang cach chon mail bao ve radio button
        case = 0
    elif "For your security and to ensure that only you have access to your account, we will ask you to verify your identity and change your password." in driver.page_source:
        case = 1
        # truong hop thay doi mat khau thi tu xu ly bang tay
    # elif "Stay signed in so you don't have to sign in again next time." in driver.page_source:
    elif "Stay signed in so you don't have to sign in again next time." in driver.page_source:
        case = 2
        # truong hop login binh thuong
    elif "and we'll send a verification code to your phone. After you enter the code, you can get back into your account." in driver.page_source:
        case = 3
        # truong hop bi khoa, can thue sdt de kich hoat lai, truong hop nay ko can xoa sdt sau khi setting mail

    else:
        case = -1
    print("CASE: ", case)
    return case


def pass_screen_update(driver):
    if "As part of our effort to improve your experience across our consumer services, we're updating the Microsoft Services Agreement. We want to take this opportunity to notify you about this update." in driver.page_source:
        # truong hop login xong nhay den man hinh thong bao update
        WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "iNext"))).click()
        WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.ID, "footer")))
        # WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID, "idBtn_Back"))).click()
        # WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='home.banner.change-password-column.cta']")))
    else:
        pass

    print("Pass pass_screen_update")


def run_all_step_config_forward(mail, password, mail_protect, driver_main_mail, path_firefox, key_thuesim, headless):
    mail_used = check_mail_used(mail)
    if mail_used == "Mail used":
        print("forwarded: ", mail)
        return [mail, password, mail_protect, "forwarded"]

    driver = create_driver(path_firefox, headless=headless)
    try:
        for i in range(50):
            try:
                login_mail(driver, mail, password)
                break
            except:
                pass

        case = process_after_login_case(driver)

        wait_until_page_success(driver)
        if case == -1:
            return [mail, password, mail_protect, "manual process"]

        if case == 0:
            wait_until_page_success(driver=driver)
            WebDriverWait(driver, 100).until(
                EC.element_to_be_clickable((By.XPATH, '//input[contains(@id, "iProof")]'))).click()
            reactive_mail(driver, driver_main_mail, mail)
            go_to_page_profile(driver)
            add_protect_mail(driver, mail_protect, driver_main_mail)
            go_to_page_profile(driver)
            verify_by_mail(driver, mail_protect, driver_main_mail)
            enable_setting_forward_mail(driver, key_thuesim)
            for i in range(10):
                try:
                    driver0 = setting_forward(driver, mail_protect, mail, password, driver_main_mail, path_firefox, headless)
                    break
                except CounldntSendTheCode as counldntSendTheCode:
                    return [mail, password, mail_protect, "CounldntSendTheCode"]

                except Exception as e:
                    # driver.save_full_page_screenshot("html/{}.png".format(mail_protect.split("@")[0]))
                    logger.exception(mail_protect, e)

            try:
                delete_phone(driver0, path_firefox, mail, password, driver_main_mail)
                return [mail, password, mail_protect, "success"]

            except Exception as e:
                logger.exception(e)
                return [mail, password, mail_protect, "error delete phone"]

        if case == 1:
            return [mail, password, mail_protect, "manual process"]

        if case == 2:
            go_to_page_profile(driver)
            add_protect_mail(driver, mail_protect, driver_main_mail)
            go_to_page_profile(driver)
            verify_by_mail(driver, mail_protect, driver_main_mail)

            enable_setting_forward_mail(driver, key_thuesim)
            for i in range(10):
                try:
                    driver2 = setting_forward(driver, mail_protect, mail, password, driver_main_mail, path_firefox, headless)
                    break
                except CounldntSendTheCode as counldntSendTheCode:
                    return [mail, password, mail_protect, "CounldntSendTheCode"]
                except Exception as e:
                    logger.exception(e)
                    # driver.save_full_page_screenshot("html/{}.png".format(mail_protect.split("@")[0]))
            try:
                delete_phone(driver2, path_firefox, mail, password, driver_main_mail)
                return [mail, password, mail_protect, "success"]
            except Exception as e:
                logger.exception(e)
                return [mail, password, mail_protect, "error delete phone"]

        if case == 3:
            reactive_by_phone(driver, key_thuesim)
            go_to_page_profile(driver)
            add_protect_mail(driver, mail_protect, driver_main_mail)
            go_to_page_profile(driver)
            verify_by_mail(driver, mail, driver_main_mail)
            enable_setting_forward_mail(driver, key_thuesim)
            for i in range(10):
                try:
                    driver3 = setting_forward(driver, mail_protect, mail, password, driver_main_mail, path_firefox, headless)
                    break
                except CounldntSendTheCode as counldntSendTheCode:
                    return [mail, password, mail_protect, "CounldntSendTheCode"]
                except Exception as e:
                    # driver.save_full_page_screenshot("html/{}.png".format(mail_protect.split("@")[0]))
                    logger.info(e)
            try:
                driver3.close()
            except:
                return [mail, password, mail_protect, "Exception Close firefox"]
            return [mail, password, mail_protect, "success"]



    except LoginFailException as loginFailException:
        print("LoginFailException: ", mail)
        return [mail, password, mail_protect, "LoginFailException"]
    except CantReceiveCodeFromApiBuyPhone as cantReceiveCodeFromApiBuyPhone:
        print("CantReceiveCodeFromApiBuyPhone: ", mail)
        return [mail, password, mail_protect, "CantReceiveCodeFromApiBuyPhone"]
    except CantReceiveCodeFromEmailException as cantReceiveCodeFromApiBuyPhone:
        return [mail, password, mail_protect, "CantReceiveCodeFromEmailException"]
    except CantProcessBuyPhone as cantProcessBuyPhone:
        return [mail, password, mail_protect, "CantProcessBuyPhone"]
    except CounldntSendTheCode as counldntSendTheCode:
        return [mail, password, mail_protect, "CounldntSendTheCode"]
    except Exception as e:
        logger.info("UNKNOWN E: ", mail_protect)
        logger.exception(e)
        return [mail, password, mail_protect, "manual"]

    finally:
        try:
            driver.close()
            print("DONEEEEEEEEEE")
        except Exception as e:
            print("Exception final: ", e)

        try:
            driver0.close()
            print("DONEEEEEEEEEE0000000000")
        except Exception as e:
            print("Exception final000000000: ", e)

        try:
            driver2.close()
            print("DONEEEEEEEEEE22222222222")
        except Exception as e:
            print("Exception final22222222: ", e)

        try:
            driver3.close()
            print("DONEEEEEEEEEE3333333333")
        except Exception as e:
            print("Exception final3333333: ", e)


def go_to_page_profile(driver):
    driver.get("https://account.live.com/names/manage?mkt=en-US&refd=account.microsoft.com&refp=profile")
    wait_until_page_success(driver=driver)


def process_all_mail(list_mails: list, num_processes: int = 1, main_mail: str = "", main_password: str = "",
                     path_firefox: str = "", key_thuesim: str = "", headless=False):
    driver_mail_mail = create_driver(path_firefox, headless=headless)
    driver1 = create_main_mail_box(driver_mail_mail, main_mail, main_password)

    results_list = []
    with ThreadPoolExecutor(max_workers=num_processes) as executor:
        futures = []
        for mail, password, mail_protect in list_mails:
            future = executor.submit(run_all_step_config_forward, mail, password, mail_protect, driver1, path_firefox,
                                     key_thuesim, headless)
            futures.append(future)

        for future in futures:
            print(future.result())
            results_list.append(future.result())

    print("All task done.")
    print("results_list", results_list)
    results_df = pd.DataFrame(results_list, columns=["email", "password", "email_protect", "status"])
    results_df.to_excel("results_{}.xlsx".format(time.time()))


class MyWindow:
    def __init__(self, win):
        with open("config.json") as f:
            config = json.load(f)

        self.pool = None
        self.win = win
        self.emails = []
        self.results = []
        self.diver1 = None

        self.main_email = "nhanmailao@minh.live"
        self.main_email_password = "Team12345!"
        self.api_key = "5ec165c3b6eae475"
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
        self.MAIN_EMAIL.insert(0, config['main_mail'])

        self.lbl3 = Label(win, text='Password main mail')
        self.lbl3.place(x=50, y=150)
        self.MAIN_EMAIL_PASSWORD = Entry(bd=3, width=30)
        self.MAIN_EMAIL_PASSWORD.place(x=200, y=150)
        self.MAIN_EMAIL_PASSWORD.insert(0, config['main_password'])

        self.lbl4 = Label(win, text='Key chothuesim')
        self.lbl4.place(x=50, y=200)
        self.API_KEY = Entry(bd=3, width=30)
        self.API_KEY.place(x=200, y=200)
        self.API_KEY.insert(0, config['key_thuesim'])

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
        try:
            self.FIREFOX_PATH.insert(0, config['path_firefox'])
        except:
            pass

        # self.processBtn = Button(win, text='Process', command=self.main, state=DISABLED)
        self.processBtn = Button(win, text='Process', command=self.main)
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
        path_firefox = self.FIREFOX_PATH.get().replace("\\", "/")
        key_thuesim = self.API_KEY.get()

        print("num_processes: ", num_processes)
        print("main_mail: ", main_mail)
        print("main_password: ", main_password)
        print("path_firefox: ", path_firefox)
        print("key_thuesim: ", key_thuesim)
        print("headless: ", self.HEADLESS)

        main_mail = "nhanmailao@minh.live"
        main_password = "Team12345!"
        key_thuesim = "5ec165c3b6eae475"
        # path_firefox = "C:/Program Files/Mozilla Firefox/firefox.exe"

        process_all_mail(self.emails, int(num_processes), main_mail, main_password, path_firefox, key_thuesim, headless=self.HEADLESS)
        # process_all_mail(self.emails, 4, main_mail, main_password)


if __name__ == '__main__':
    window = Tk()
    mywin = MyWindow(window)
    window.title('Tool Forward Mail')
    window.geometry("500x500+10+10")
    window.mainloop()