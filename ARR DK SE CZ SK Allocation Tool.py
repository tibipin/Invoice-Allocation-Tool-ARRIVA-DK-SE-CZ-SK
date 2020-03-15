import WINI_OOP
import time
import pandas
import os
import tempfile
import datetime
from tkinter import *
from tkinter import messagebox
import numpy.random.common
import numpy.random.bounded_integers
import numpy.random.entropy
from selenium.webdriver.common.keys import Keys
WINI_OOP.folder = tempfile.mkdtemp()


class ArrivaAllocation:

    def __init__(self):

        self.countries = {
            'DK': ['00D1', '00D2', '00D3', '00D4', '00D5', '00D6', '00D7', '00D8', '00DB', '00DC', '00DD', '00DE',
                   '00DK'],
            'SE': ['00DF', '00DG', '00DH', '00DI'],
            'SK': ['00F2', '00F3', '00F4', '00F5', '00F6', '00F7', '00F8', '00F9', '00FA'],
            'CZ': ['00E4', '00E5', '00E6', '00E7', '00E8', '00E9', '00EA']}

        self.root_allocation_folder = {
            'DK': r'\\login.ds.signintra.com\DFS\RO\OTP-01\location\01_Accounts-Payable\ARRIVA SE DK\10. Trackers\Daily reports',
            'SE': r'\\login.ds.signintra.com\DFS\RO\OTP-01\location\01_Accounts-Payable\ARRIVA SE DK\10. Trackers\Daily reports',
            'SK': r'\\login.ds.signintra.com\DFS\RO\OTP-01\location\01_Accounts-Payable\SK Arriva\02. Invoice processing\01. Task Allocation',
            'CZ': r'\\login.ds.signintra.com\DFS\RO\OTP-01\location\01_Accounts-Payable\CZ Arriva\02. Invoice processing\01. Task allocation'}

        screen = Tk()
        screen.title('Arriva Invoice Allocation tool')

        sims_label = Label(screen, text='SIMS ID:')
        sims_label.grid(row=0, column=0, sticky=E)
        self.sims_entry = Entry(screen)
        self.sims_entry.grid(row=0, column=1, sticky=W)

        password_label = Label(screen, text='SIMS password:')
        password_label.grid(row=2, column=0, sticky=E)
        self.password_entry = Entry(screen)
        self.password_entry.grid(row=2, column=1, sticky=W)

        country_label = Label(screen, text='Country:')
        country_label.grid(row=4, column=0, sticky=E)
        self.country = StringVar(screen)
        self.country_menu = OptionMenu(screen, self.country, *self.countries.keys())
        self.country_menu.grid(row=4, column=1, sticky=W)
        self.country.trace('w', self.change_dropdown)

        folder_label = Label(screen, text='Root Folder:')
        folder_label.grid(row=6, column=0)
        self.folder_entry = Entry(screen, width=80)
        self.folder_entry.grid(row=6, column=1, sticky=W, columns=2)

        button_dld = Button(screen, text='Download WINI reports', bg='white', fg='black', command=self.download_rapoarte)
        button_dld.grid(row=0, column=2)
        button_aloc = Button(screen, text='Generate Allocation', bg='black', fg='white', command=self.generate_allocation)
        button_aloc.grid(row=2, column=2)

        screen.grid_columnconfigure(5, minsize=5)
        for i in [1, 3, 5]:
            screen.grid_rowconfigure(i, minsize=5)

        screen.mainloop()

    def download_rapoarte(self):

        WINI_OOP.parola = self.password_entry.get()
        WINI_OOP.sims = self.sims_entry.get()

        connection = WINI_OOP.WINI()
        time.sleep(5)
        connection.login()
        time.sleep(5)
        connection.reports()
        time.sleep(2)
        connection.invoices()
        time.sleep(2)
        for c in WINI_OOP.company_codes:
            time.sleep(1)
            connection.extract_report(entity=4, client=100, company_code=c,
                                      #workflow_status_lower_limit=2,
                                      #workflow_status_upper_limit=11,
                                      scan_date_lower_limit='Dec 1, 2019',
                                      scan_date_upper_limit='Dec 31, 2019',
                                      head_tickmark_check=1)
                                      #user_tickmark_check=1)
        time.sleep(30)
        connection.download_report()
        time.sleep(1)
        for i in range(1, (len(WINI_OOP.company_codes) + 1)):
            menu_individual_report = WINI_OOP.browser.find_element_by_xpath(
                '/html/body/div[1]/div[1]/div/div/div[3]/div/div/div/div/div/div[2]/div[3]/div/div/div/div[1]/div[1]/div/div/div/div/div/div/div/div/div[1]/div[2]/div/div/div/div[3]/table/tbody[2]/tr[' + str(
                    i) + ']/td[1]/div/table/tbody/tr/td/table/tbody/tr/td[1]/span')
            menu_individual_report.click()
            time.sleep(1)
            download_button = WINI_OOP.browser.find_element_by_xpath('/html/body/div[3]/ul/li[1]/div/div/div/a')
            download_button.click()
        messagebox.showinfo("Arriva Allocation Tool", "WINI reports successfully downloaded")

    def create_folder(self):
        self.target_folder = self.folder_entry.get()
        while True:
            director = os.listdir(self.target_folder)
            if str(datetime.datetime.today().strftime('%Y')) in director:
                self.target_folder += '\\' + str(datetime.datetime.today().strftime('%Y'))
                os.chdir(self.target_folder)
                break
            else:
                os.mkdir(self.target_folder + '\\' + str(datetime.datetime.today().strftime('%Y')))
                self.target_folder = self.target_folder + '\\' + str(datetime.datetime.today().strftime('%Y'))
                os.chdir(self.target_folder)
                break
        while True:
            director = os.listdir(self.target_folder)
            if str(datetime.datetime.today().strftime('%B')) in director:
                self.target_folder += '\\' + str(datetime.datetime.today().strftime('%B'))
                os.chdir(self.target_folder)
                break
            else:
                os.mkdir(self.target_folder + '\\' + str(datetime.datetime.today().strftime('%B')))
                self.target_folder = self.target_folder + '\\' + str(datetime.datetime.today().strftime('%B'))
                os.chdir(self.target_folder)
                break
        while True:
            director = os.listdir(self.target_folder)
            if str(datetime.datetime.today().strftime('%d.%m.%Y')) in director:
                self.target_folder += '\\' + str(datetime.datetime.today().strftime('%d.%m.%Y'))
                os.chdir(self.target_folder)
                break
            else:
                os.mkdir(self.target_folder + '\\' + str(datetime.datetime.today().strftime('%d.%m.%Y')))
                self.target_folder = self.target_folder + '\\' + str(datetime.datetime.today().strftime('%d.%m.%Y'))
                os.chdir(self.target_folder)
                break

    def change_dropdown(self, *args):
        WINI_OOP.company_codes = self.countries[str(self.country.get())]
        if self.folder_entry.get():
            self.folder_entry.delete(0, 'end')
        self.folder_entry.insert(END, str(self.root_allocation_folder[self.country.get()]))

    def generate_allocation(self):
        self.create_folder()

        # Unify the reports extracted from WINI
        self.a = pandas.DataFrame()
        self.u = pandas.DataFrame()
        self.activity_split = pandas.DataFrame()

        for file in os.listdir(WINI_OOP.folder):
            self.a = self.a.append(pandas.read_excel(WINI_OOP.folder + '\\' + file, sheet_name='Head', converters={
                'HE_Company Code': str}))
            self.u = self.u.append(pandas.read_excel(WINI_OOP.folder + '\\' + file, sheet_name='Users'))

        # Filter the information and format columns
        self.u = self.u[self.u['Type'] != 'G']
        self.u = self.u[~self.u['Groupname'].str.contains('ACC_CHECK')]

        coloane_necesare = ['HE_Transaction Number',
                            'HE_Country Code',
                            'HE_Company Code',
                            'HE_Creditor Number',
                            'HE_Creditor Name',
                            'HE_Invoice Number',
                            'HE_Invoice Type',
                            'HE_Document Subtype',
                            'HE_Invoice Date',
                            'HE_Last Change Workflow Status',
                            'HE_Active User',
                            'HE_Workflow Status']
        for coloana in self.a.columns:
            if coloana not in coloane_necesare:
                del self.a[coloana]

        self.a['HE_Transaction Number'] = self.a['HE_Transaction Number'].astype(str)
        self.a['HE_Last Change Workflow Status'] = self.a['HE_Last Change Workflow Status'].dt.date
        self.a['HE_Invoice Number'] = self.a['HE_Invoice Number'].astype(str)
        self.a['HE_Invoice Date'] = self.a['HE_Invoice Date'].dt.date
        self.a['HE_Creditor Number'] = self.a['HE_Creditor Number'].astype(str)
        self.a.replace({'HE_Creditor Number': r'\.0$'}, '', regex=True, inplace=True)
        self.u.columns = ['Type', 'HE_Active User', 'Member', 'Email']
        self.a['HE_Active User'].fillna(value='Error', inplace=True)
        self.activity_split = self.a[self.a['HE_Active User'].str.contains('ACC_CHECK')]
        pa = self.a[~self.a['HE_Active User'].str.contains('ACC_CHECK')]
        self.pending_approval = \
            pandas.merge(pa, self.u, how='inner', on='HE_Active User')

        # Calculate Inflow, Backlog and TAT

        ziua_saptamanii = datetime.date.today().isoweekday()
        Backlog = []
        Inflow = []
        SGBS = []

        for i in self.activity_split.iterrows():
            x = (datetime.date.today() - i[1]['HE_Last Change Workflow Status']).days

            ## Inflow
            if ziua_saptamanii == 1:
                if 0 < x <= 3:
                    Inflow.append('Yes')
                else:
                    Inflow.append(None)
            else:
                if x == 1:
                    Inflow.append('Yes')
                else:
                    Inflow.append(None)

            ## Backlog
            if i[1]['HE_Workflow Status'] in ['Approver 1 ok', 'Approver 2 ok']:
                if self.country.get() in ['DK', 'SE']:
                    if ziua_saptamanii in [3, 4, 5]:
                        if x >= 1:
                            Backlog.append('Yes - backlog AP2')
                        else:
                            Backlog.append(None)
                    else:
                        if x >= 3:
                            Backlog.append('Yes - backlog AP2')
                        else:
                            Backlog.append(None)
                else:
                    if ziua_saptamanii in [3, 4, 5]:
                        if x > 1:
                            Backlog.append('Yes - backlog AP2')
                        else:
                            Backlog.append(None)
                    else:
                        if x > 3:
                            Backlog.append('Yes - backlog AP2')
                        else:
                            Backlog.append(None)
            elif i[1]['HE_Workflow Status'] in ['Clarification available', 'Clarification canceled',
                                                'Asset Accountant ok']:
                if ziua_saptamanii in [3, 4, 5]:
                    if x >= 2:
                        Backlog.append('Yes')
                    else:
                        Backlog.append(None)
                else:
                    if x >= 4:
                        Backlog.append('Yes')
                    else:
                        Backlog.append(None)
            elif i[1]['HE_Workflow Status'] == 'Workflow ready':
                if i[1]['HE_Invoice Type'] == 'RMB':
                    if self.country.get() in ['DK', 'SE']:
                        if ziua_saptamanii == 1:
                            if x >= 10:
                                Backlog.append('Yes')
                            else:
                                Backlog.append(None)
                        else:
                            if x >= 8:
                                Backlog.append('Yes')
                            else:
                                Backlog.append(None)
                    else:
                        if ziua_saptamanii in [1, 2, 3]:
                            if x >= 5:
                                Backlog.append('Yes')
                            else:
                                Backlog.append(None)
                        else:
                            if x >= 3:
                                Backlog.append('Yes')
                            else:
                                Backlog.append(None)
                else:
                    if ziua_saptamanii in [1, 2, 3]:
                        if x >= 5:
                            Backlog.append('Yes')
                        else:
                            Backlog.append(None)
                    else:
                        if x >= 3:
                            Backlog.append('Yes')
                        else:
                            Backlog.append(None)
            else:
                Backlog.append('Workflow status not in SGBS scope. Please check.')

            ## TAT
            if i[1]['HE_Workflow Status'] in ['Clarification available', 'Clarification canceled',
                                              'Asset Accountant ok']:
                if ziua_saptamanii == 1:
                    if x >= 3:
                        SGBS.append('Yes')
                    else:
                        SGBS.append(None)
                else:
                    if x >= 1:
                        SGBS.append('Yes')
                    else:
                        SGBS.append(None)
            elif i[1]['HE_Workflow Status'] == 'Workflow ready':
                if i[1]['HE_Invoice Type'] == 'RMB':
                    if self.country.get() in ['DK', 'SE']:
                        if x >= 7:
                            SGBS.append('Yes')
                        else:
                            SGBS.append(None)
                    else:
                        if ziua_saptamanii in [1, 2]:
                            if x >= 4:
                                SGBS.append('Yes')
                            else:
                                SGBS.append(None)
                        else:
                            if x >= 2:
                                SGBS.append('Yes')
                            else:
                                SGBS.append(None)
                else:
                    if ziua_saptamanii in [1, 2]:
                        if x >= 4:
                            SGBS.append('Yes')
                        else:
                            SGBS.append(None)
                    else:
                        if x >= 2:
                            SGBS.append('Yes')
                        else:
                            SGBS.append(None)
            elif i[1]['HE_Workflow Status'] in ['Approver 1 ok', 'Approver 2 ok']:
                SGBS.append('Yes')
            else:
                SGBS.append('Workflow status not in SGBS scope. Please check.')

        self.activity_split['Inflow'] = Inflow
        self.activity_split['Backlog'] = Backlog
        self.activity_split['SGBS'] = SGBS

        self.activity_split['HE_Invoice Type'] = self.activity_split['HE_Invoice Type'].map(
            {'RMB': 'PO', 'ROB': 'nonPO'})

        # Return payment terms for ARR DK and SE
        if self.country.get() in ['DK', 'SE']:
            payT = pandas.read_excel(r'\\login.ds.signintra.com\DFS\RO\OTP-01\location\01_Accounts-Payable\ARRIVA SE DK\12. Other Projects\Allocation Tool\PayT ARR DK SE.xlsx')
            payT = payT.rename(columns={'Vendor': 'HE_Creditor Number', 'Company code': 'HE_Company Code'})
            payT['HE_Creditor Number'] = payT['HE_Creditor Number'].astype(str)
            tempdf = self.activity_split.loc[self.activity_split['HE_Invoice Type'] == 'nonPO']
            tempdf = tempdf.merge(payT, how='left', on=['HE_Creditor Number', 'HE_Company Code'])
            tempdf = tempdf.loc[:,['HE_Transaction Number', 'PayT']]
            self.activity_split = self.activity_split.merge(tempdf, how='left', on='HE_Transaction Number')

        # Return noPOnoPay for ARR SE
        if self.country.get() == 'SE':
            file = '\\\\login.ds.signintra.com\\DFS\\RO\\OTP-01\\location\\01_Accounts-Payable\\ARRIVA SE DK\\2. Invoice Processing\\No PO No Pay SE\\PO and no PO vendor list.xlsx'
            PO_only = pandas.read_excel(file, sheet_name='PO exceptions', converters={'Vendor': str})
            PO_only = PO_only.rename(columns={'Vendor': 'HE_Creditor Number'})
            PO_only['PO Exceptions'] = 'Yes'
            self.activity_split = self.activity_split.merge(PO_only[['HE_Creditor Number', 'PO Exceptions']], how='left', on=['HE_Creditor Number'])

        # Return noPOnoPay for ARR DK
        if self.country.get() == 'DK':
            file = '\\\\login.ds.signintra.com\\DFS\\RO\\OTP-01\\location\\01_Accounts-Payable\\ARRIVA SE DK\\2. Invoice Processing\\No PO No Pay DK\\no PO no PAY list.xlsx'
            PO_only = pandas.read_excel(file, converters={'Vendor no.': str})
            PO_only = PO_only.rename(columns={'Vendor no.': 'HE_Creditor Number'})
            self.activity_split = self.activity_split.merge(PO_only[['HE_Creditor Number', 'no PO no PAY vendor']],
                                                            how='left', on=['HE_Creditor Number'])
            # Circle K and Biofuel
            for i in self.activity_split.iterrows():
                if str(i[1]['HE_Invoice Number']).startswith('19') and len(i[1]['HE_Invoice Number']) == 8:
                    i[1]['SGBS'] = 'Yes'
                elif str(i[1]['HE_Creditor Number']) == '728822':
                    i[1]['SGBS'] = 'Yes'

        # Return special rules for ARR SK
        if self.country.get() == 'SK':
            file1 = r'\\login.ds.signintra.com\DFS\RO\OTP-01\location\01_Accounts-Payable\SK Arriva\02. Invoice processing\03. Supporting documents\\Guideline for Arriva SK vendors final version.xlsx'
            special_rules = pandas.read_excel(file1,converters={'Vendor ID': str})
            special_rules = special_rules.rename(columns={'Vendor ID': 'HE_Creditor Number'})
            self.activity_split = self.activity_split.merge(special_rules[['HE_Creditor Number', 'Special Rules']],
                                                            how='left', on=['HE_Creditor Number'])
        # Return special rules for ARR CZ
        if self.country.get() == 'CZ':
            file1 = r'\\login.ds.signintra.com\DFS\RO\OTP-01\location\01_Accounts-Payable\CZ Arriva\02. Invoice processing\03. Supporting documents\Special vendors ARR CZ.xlsx'
            special_rules = pandas.read_excel(file1, converters={'Vendor Account': str})
            special_rules = special_rules.rename(columns={'Vendor Account': 'HE_Creditor Number'})
            self.activity_split = self.activity_split.merge(special_rules[['HE_Creditor Number', 'Special Rules']],
                                                            how='left', on=['HE_Creditor Number'])

        #Inflow, Backlog and SGBS summary

        inflow_table = pandas.pivot_table(self.activity_split, index=['Inflow', 'HE_Workflow Status', 'HE_Invoice Type', 'HE_Company Code'],
                                       aggfunc='count',
                                       values=['HE_Transaction Number'],
                                       margins=True)
        backlog_table = pandas.pivot_table(self.activity_split, index=['Backlog', 'HE_Workflow Status', 'HE_Invoice Type', 'HE_Company Code'],
                                        aggfunc='count',
                                        values=['HE_Transaction Number'],
                                        margins=True)
        sgbs_table = pandas.pivot_table(self.activity_split, index=['SGBS', 'HE_Workflow Status', 'HE_Invoice Type', 'HE_Company Code'],
                                        aggfunc='count',
                                        values=['HE_Transaction Number'],
                                        margins=True)

        # Final touches
        self.activity_split['Comments'] = None
        self.activity_split = self.activity_split.drop(columns=['HE_Country Code', 'HE_Active User'])


        # Write excel file
        writer = pandas.ExcelWriter(self.target_folder + 'Activity split ARR ' + str(self.country.get()) + ' ' + str(datetime.datetime.today().strftime('%d.%m')) + '.xlsx', engine='xlsxwriter')
        vba_code = r'\\login.ds.signintra.com\DFS\RO\OTP-01\location\01_Accounts-Payable\ARRIVA SE DK\12. Other Projects\Allocation Tool\vbaproject.bin'
        with writer:
            workbook = writer.book
            frontsheet = workbook.add_worksheet('Frontsheet')
            workbook.add_vba_project(vba_code)
            frontsheet.insert_button('B2', {'macro': 'ThisWorkbook.Start_Allocation',  # add sub name
                                            'caption': 'Start Allocation',
                                            'width': 160,
                                            'height': 40})
            frontsheet.insert_button('B6', {'macro': 'ThisWorkbook.Generate_files',  # add sub name
                                            'caption': 'Generate files',
                                            'width': 160,
                                            'height': 40})
            frontsheet.insert_button('B10', {'macro': 'ThisWorkbook.Fetch',  # add sub name
                                            'caption': 'Get more invoices',
                                            'width': 160,
                                            'height': 40})
            frontsheet.hide_gridlines(2)
            self.activity_split.to_excel(writer, sheet_name='To be processed '+datetime.datetime.strftime(datetime.datetime.today(),'%d.%m.%Y'),index=False)

            inflow_table.to_excel(writer, sheet_name='Overview ' + datetime.datetime.strftime(datetime.datetime.today(),
                                                                                           '%d.%m.%Y'), startcol=0)
            backlog_table.to_excel(writer, sheet_name='Overview ' + datetime.datetime.strftime(datetime.datetime.today(),
                                                                                            '%d.%m.%Y'), startcol=5)
            sgbs_table.to_excel(writer, sheet_name='Overview ' + datetime.datetime.strftime(datetime.datetime.today(),
                                                                                            '%d.%m.%Y'), startcol=10)
            self.pending_approval.to_excel(writer, sheet_name='Pending approval ' + datetime.datetime.strftime(
                datetime.datetime.today(), '%d.%m.%Y'), index=False)

            workbook.filename = 'Activity split ARR ' + str(self.country.get()) + ' '+ datetime.datetime.strftime(datetime.datetime.today(),
                                                                                      '%d.%m.%Y') + '.xlsm'
            writer.save()
            messagebox.showinfo("Arriva Allocation Tool", 'Daily allocation macro saved @ '+str(self.target_folder))

if __name__ == '__main__':
    test = ArrivaAllocation()