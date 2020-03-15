from selenium import webdriver
import time
import datetime
from selenium.webdriver.common.keys import Keys
company_codes = []
sims = None
parola = None
folder = None
browser = None


class WINI():

    def __init__(self):
        ### Establish connection
        global browser
        options = webdriver.ChromeOptions()
        # options.add_argument('headless')
        prefs = {}
        prefs['profile.default_content_settings.popups'] = 0
        prefs['download.default_directory'] = folder
        options.add_experimental_option('prefs', prefs)
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        browser = webdriver.Chrome(executable_path=r'\\login.ds.signintra.com\DFS\RO\OTP-01\location\01_Accounts-Payable\ARRIVA SE DK\12. Other Projects\Allocation Tool\chromedriver.exe',options=options)
        browser.get('https://wini.dc.signintra.com/scr-webclient/')

    def delete_existing_value(self,param):
        global q
        if param.get_attribute('value'):
            q = len(param.get_attribute('value'))
            for a in range(len(param.get_attribute('value'))):
                param.send_keys(Keys.BACKSPACE)

    def login(self):

        user = sims
        password = parola

        user_field = browser.find_element_by_xpath(
            '/html/body/div[3]/div[3]/div/div/div/table/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr[3]/td/div/div[3]/table/tbody[2]/tr[1]/td[2]/input')
        time.sleep(1)
        self.delete_existing_value(user_field)
        time.sleep(1)
        user_field.send_keys(user)

        password_field = browser.find_element_by_xpath(
            '/html/body/div[3]/div[3]/div/div/div/table/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr[3]/td/div/div[3]/table/tbody[2]/tr[2]/td[2]/input')
        time.sleep(1)
        self.delete_existing_value(password_field)
        time.sleep(1)
        password_field.send_keys(password[0])
        time.sleep(0.5)
        password_field.send_keys(password)

        login_button = browser.find_element_by_xpath('//*[@id="z_loginButton-box"]/tbody/tr[2]/td[2]')
        login_button.click()

    def reports(self):

        reports_field = browser.find_element_by_xpath(
            '/html/body/div/div[1]/div/div/div[2]/div[1]/div/div/div/div/div/div/div/div/div[1]/div/div/div/div/table/tbody/tr[4]/td/div/span[2]')
        reports_field.click()

    def invoices(self):
        invoices_tab = browser.find_element_by_xpath(
                '/html/body/div[1]/div[1]/div/div/div[2]/div[1]/div/div/div/div/div/div/div/div/div[1]/div/div/div/div/table/tbody/tr[5]/td/div/span[5]')
        invoices_tab.click()

    def extract_report(self,entity=None,client=None,company_code=None,scan_date_lower_limit=None,scan_date_upper_limit=None,workflow_status_lower_limit=None,workflow_status_upper_limit=None,head_tickmark_check=0,line_tickmark_check=0,user_tickmark_check=0):

        clear_fields_button = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div/div[3]/div/div/div/div/div/div[2]/div[2]/div/div/div/div[1]/div/div/div/div/div/div/div/table/tbody/tr/td/table/tbody/tr[5]/td/span/table/tbody/tr[2]/td[2]')
        clear_fields_button.click()

        time.sleep(1)

        entity_field = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div/div[3]/div/div/div/div/div/div[2]/div[2]/div/div/div/div[1]/div/div/div/div/div/div/div/table/tbody/tr/td/table/tbody/tr[11]/td/div/div/div[1]/div/div/div/div/div/div/div/div/table/tbody/tr/td/table/tbody/tr/td/div/div[4]/table/tbody[2]/tr[1]/td[5]/div/i/input')
        self.delete_existing_value(entity_field)
        entity_field.send_keys(entity)

        client_field = browser.find_element_by_xpath(
            '/html/body/div[1]/div[1]/div/div/div[3]/div/div/div/div/div/div[2]/div[2]/div/div/div/div[1]/div/div/div/div/div/div/div/table/tbody/tr/td/table/tbody/tr[11]/td/div/div/div[1]/div/div/div/div/div/div/div/div/table/tbody/tr/td/table/tbody/tr/td/div/div[4]/table/tbody[2]/tr[2]/td[5]/div/i/input')
        self.delete_existing_value(client_field)
        if client:
            client_field.send_keys(client)

        company_code_field = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div/div[3]/div/div/div/div/div/div[2]/div[2]/div/div/div/div[1]/div/div/div/div/div/div/div/table/tbody/tr/td/table/tbody/tr[11]/td/div/div/div[1]/div/div/div/div/div/div/div/div/table/tbody/tr/td/table/tbody/tr/td/div/div[4]/table/tbody[2]/tr[3]/td[5]/div/i/input')
        if company_code:
            company_code_field.send_keys(company_code)

        scan_date_lower_limit_field = browser.find_element_by_xpath(
            '/html/body/div[1]/div[1]/div/div/div[3]/div/div/div/div/div/div[2]/div[2]/div/div/div/div[1]/div/div/div/div/div/div/div/table/tbody/tr/td/table/tbody/tr[11]/td/div/div/div[1]/div/div/div/div/div/div/div/div/table/tbody/tr/td/table/tbody/tr/td/div/div[4]/table/tbody[2]/tr[5]/td[5]/div/i/input')

        if scan_date_lower_limit:
            scan_date_lower_limit_field.send_keys(scan_date_lower_limit)

        scan_date_upper_limit_field = browser.find_element_by_xpath(
            '/html/body/div[1]/div[1]/div/div/div[3]/div/div/div/div/div/div[2]/div[2]/div/div/div/div[1]/div/div/div/div/div/div/div/table/tbody/tr/td/table/tbody/tr[11]/td/div/div/div[1]/div/div/div/div/div/div/div/div/table/tbody/tr/td/table/tbody/tr/td/div/div[4]/table/tbody[2]/tr[5]/td[7]/div/i/input')

        if scan_date_upper_limit:
            scan_date_upper_limit_field.send_keys(scan_date_upper_limit)

        workflow_status_lower_limit_field = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div/div[3]/div/div/div/div/div/div[2]/div[2]/div/div/div/div[1]/div/div/div/div/div/div/div/table/tbody/tr/td/table/tbody/tr[11]/td/div/div/div[1]/div/div/div/div/div/div/div/div/table/tbody/tr/td/table/tbody/tr/td/div/div[4]/table/tbody[2]/tr[7]/td[5]/div/i/input')
        if workflow_status_lower_limit:
            workflow_status_lower_limit_field.send_keys(workflow_status_lower_limit)

        workflow_status_upper_limit_field = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div/div[3]/div/div/div/div/div/div[2]/div[2]/div/div/div/div[1]/div/div/div/div/div/div/div/table/tbody/tr/td/table/tbody/tr[11]/td/div/div/div[1]/div/div/div/div/div/div/div/div/table/tbody/tr/td/table/tbody/tr/td/div/div[4]/table/tbody[2]/tr[7]/td[7]/div/i/input')
        if workflow_status_upper_limit:
            workflow_status_upper_limit_field.send_keys(workflow_status_upper_limit)

        export_button = browser.find_element_by_xpath(
            '/html/body/div[1]/div[1]/div/div/div[3]/div/div/div/div/div/div[2]/div[2]/div/div/div/div[1]/div/div/div/div/div/div/div/table/tbody/tr/td/table/tbody/tr[7]/td/span/table/tbody/tr[2]/td[2]')
        export_button.click()

        time.sleep(2)

        head_tickmark = browser.find_element_by_xpath(
            '/html/body/div[4]/div[3]/div/div/div/div/div/div/div/div/div/div[1]/div/div/div/div/div/div/div/table/tbody/tr/td/table/tbody/tr[5]/td/div/div/div[1]/div/div/div/div/div/div/div/div/table/tbody/tr/td/table/tbody/tr/td/div/div[3]/table/tbody[2]/tr[1]/td[1]/div/span/input')
        if head_tickmark_check:
            head_tickmark.click()

        line_tickmark = browser.find_element_by_xpath(
            '/html/body/div[4]/div[3]/div/div/div/div/div/div/div/div/div/div[1]/div/div/div/div/div/div/div/table/tbody/tr/td/table/tbody/tr[5]/td/div/div/div[1]/div/div/div/div/div/div/div/div/table/tbody/tr/td/table/tbody/tr/td/div/div[3]/table/tbody[2]/tr[2]/td[1]/div/span/input')
        if line_tickmark_check:
            line_tickmark.click()

        user_tickmark = browser.find_element_by_xpath(
            '/html/body/div[4]/div[3]/div/div/div/div/div/div/div/div/div/div[1]/div/div/div/div/div/div/div/table/tbody/tr/td/table/tbody/tr[5]/td/div/div/div[1]/div/div/div/div/div/div/div/div/table/tbody/tr/td/table/tbody/tr/td/div/div[3]/table/tbody[2]/tr[3]/td[1]/div/span/input')
        if user_tickmark_check:
            user_tickmark.click()

        export_button_2 = browser.find_element_by_xpath(
            '/html/body/div[4]/div[3]/div/div/div/div/div/div/div/div/div/div[1]/div/div/div/div/div/div/div/table/tbody/tr/td/table/tbody/tr[3]/td/span/table/tbody/tr[2]/td[2]')
        export_button_2.click()

        time.sleep(3)

        ok_button = browser.find_element_by_xpath(
            '/html/body/div[4]/div[3]/div/div/div/table[2]/tbody/tr/td/table/tbody/tr/td/span/table/tbody/tr[2]/td[2]')
        ok_button.click()

        time.sleep(1)

    def download_report(self):
        Report_download_icon = browser.find_element_by_xpath(
            '/html/body/div[1]/div[1]/div/div/div[2]/div[1]/div/div/div/div/div/div/div/div/div[1]/div/div/div/div/table/tbody/tr[7]/td/div/span[5]')
        Report_download_icon.click()
        Report_creation_date_lower_limit = (datetime.datetime.now() - datetime.timedelta(hours=1, minutes=5)).strftime('%b %d, %Y %I:%M:%S %p')
        time.sleep(1)
        report_creation_date_lower_limit_field = browser.find_element_by_xpath('html/body/div[1]/div[1]/div/div/div[3]/div/div/div/div/div/div[2]/div[3]/div/div/div/div[1]/div[2]/div[1]/div/div/div/div/div/div/div[2]/div/div/table/tbody/tr/td/table/tbody/tr[23]/td/i/input')
        report_creation_date_lower_limit_field.send_keys(Report_creation_date_lower_limit)
        result_button = browser.find_element_by_xpath(
            '/html/body/div[1]/div[1]/div/div/div[3]/div/div/div/div/div/div[2]/div[3]/div/div/div/div[1]/div[2]/div[1]/div/div/div/div/div/div/div[2]/div/div/table/tbody/tr/td/table/tbody/tr[11]/td/span/table/tbody/tr[2]/td[2]')
        result_button.click()
        time.sleep(2)


if __name__=='__main__':
    pass