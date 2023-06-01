import sys, requests, json, re, math
from functools import partial
from lib.main import MainWindow
from lib.worker import Worker, RecoveryWorker
from settings.config import *
from lib.newsletter_subscribe import NewsLetterSubscribe
from lib.create_outlook import OutlookCreate
from lib.Recovery import Recovery_Apps
from lib.gmail import Gmail_Recovery
from lib.gmail_composer import Gmail_Composer
from lib.yahoo import Yahoo_Recovery
from lib.search_emails import Search_Emails
from PySide2.QtWidgets import QApplication, QLineEdit, QLabel,QDateEdit, QCheckBox, QHeaderView, QWidget, QGridLayout
from PySide2.QtCore import *
import ipaddress, time
from faker import Faker
from datetime import datetime, timedelta
from PySide2.QtWidgets import QApplication,QMessageBox,QTableWidgetItem,QFileDialog
import PySide2


# uncomment these lines before creating an Executable version for GportlaUi Project  
if getattr(sys, 'frozen', False):
    os.chdir(sys._MEIPASS)

    
class GportalUi():
    """ GportalUi's controller class """

    def __init__(self, view):
        self._view = view
        self.is_logged = False
        self.stop_newsletter_process = False
        self.stop_create = False
        self.stop_process = False
        self.threadpool = QThreadPool()
        self._view.login_window.show()
        self._connectSignalsAndSlots()
        self.view_table  = self._view.main_window.tableWidget
        self.item  = PySide2.QtWidgets 
        self.tab_widget = self._view.main_window.tabWidget

        self._view.main_window.stop_Recovery_Btn.setEnabled(False)

        #Search Emails
        self.inbox = self._view.main_window.radioButton_Inbox_Search
        self.junk = self._view.main_window.radioButton_Junk_Search

        #Email_Composer
        self._view.main_window.Composer_hidebrowser.setEnabled(False)
        self._view.main_window.Composer_stop_Btn.setEnabled(False)

        self._view.main_window.stackedWidget.setCurrentIndex(1)
        self._view.main_window.stackedWidget_R.setCurrentIndex(1)
        self._view.main_window.Composer_stackedWidget.setCurrentIndex(1)
        self._view.main_window.stackedWidget_Search.setCurrentIndex(1)

        # Disable outlook, yahoo & firefox
        self._view.main_window.R_Switch_ISP.model().item(1).setEnabled(False)
        self._view.main_window.R_Switch_ISP.model().item(2).setEnabled(False)
        self._view.main_window.Recovery_browser.model().item(1).setEnabled(False)

    def _connectSignalsAndSlots(self):
        self._view.btn_login.clicked.connect(partial(self.loginHandler))
        self._view.main_window.news_checkall.toggled.connect(self.newsCheckAll)
        self._view.main_window.newsStartBtn.clicked.connect(self.newsStartProcess)
        self._view.main_window.newsStopBtn.clicked.connect(self.newsStopProcess)
        self._view.main_window.generate_emails_btn.clicked.connect(self.GenerateEmails_window)
        self._view.generate_window.start_gene.clicked.connect(self.Faker_generate)
        self._view.main_window.start_outlook.clicked.connect(self.O_startProcess)
        self._view.main_window.stop_outlook.clicked.connect(self.outllokStopP)
        self._view.main_window.clear_btn.clicked.connect(self.clear)
        self._view.main_window.log_output.clicked.connect(self.showlogers)
        self._view.main_window.table_output.clicked.connect(self.showtables)
        self._view.main_window.export_csv_btn.clicked.connect(self.CSV_data_Export)
        self._view.main_window.log_output.setChecked(True)
        self._view.main_window.Creatorloger.textChanged.connect(self.scrollDown)
        #recovery 
        self._view.main_window.start_Recovery_Btn.clicked.connect(self.RecoveryStartProcess)
        self._view.main_window.stop_Recovery_Btn.clicked.connect(self.Recoverstop)
        self._view.main_window.clear_btn_Recovery.clicked.connect(self.clear_Recovery)
        self._view.main_window.recovery_Log_output.clicked.connect(self.showlogers)
        self._view.main_window.table_output_recovery.clicked.connect(self.showtables)
        self._view.main_window.Recovery_file_dialog.clicked.connect(self.select_file)
        # self._view.main_window.R_Switch_ISP.currentTextChanged.connect(self.combobox_changed_value)
        self._view.main_window.Recovery_logger.textChanged.connect(self.scrollDown)
        self._view.main_window.Recovery_load_profile.clicked.connect(self.load_profiles)
        self._view.main_window.RecoveryCheck_all.clicked.connect(self.check_all_profiles)
        #Search Emails
        self._view.main_window.pushButton_Start_Search.clicked.connect(self.start_search_process)
        self._view.main_window.pushButton_Stop_Search.clicked.connect(self.stop_search_emails)
        self._view.main_window.pushButton_Clear_Search.clicked.connect(self.clear_search_emails)
        self._view.main_window.checkBox_Custom_Search.toggled.connect(self.toggle_custom_search_emails)
        self._view.main_window.radioButton_Logger_Search.clicked.connect(self.showlogers)
        self._view.main_window.radioButton_Table_Search.clicked.connect(self.showtables)
        # Compose Emails
        self._view.main_window.Composer_start_Btn.clicked.connect(self.composer_start_process)
        self._view.main_window.Composer_stop_Btn.clicked.connect(self.composer_stop)
        self._view.main_window.Composer_clear_btn.clicked.connect(self.clear_composer)
        self._view.main_window.Composer_Switch_ISP.currentTextChanged.connect(self.combobox_changed_value)
        self._view.main_window.Composer_logger_output.clicked.connect(self.showlogers)
        self._view.main_window.Composer_table_output.clicked.connect(self.showtables)
        self._view.main_window.Composer_file_dialog.clicked.connect(self.select_file)
        self._view.main_window.Composer_logger.textChanged.connect(self.scrollDown)
        self._view.main_window.Composer_load_profile.clicked.connect(self.load_profiles)
        self._view.main_window.ComposerCheck_all.clicked.connect(self.check_all_profiles)
    
    def select_file(self):
        if self.tab_widget.currentWidget().objectName() == 'emailComposer':
            scrollArea = self._view.main_window.scrollArea
            start_btn = self._view.main_window.Composer_start_Btn
            check_all = self._view.main_window.ComposerCheck_all
            profile_path = self._view.main_window.Composer_profile_path

        elif self.tab_widget.currentWidget().objectName() == 'recovery':
            scrollArea = self._view.main_window.Recovery_scrollArea
            start_btn = self._view.main_window.start_Recovery_Btn
            check_all = self._view.main_window.RecoveryCheck_all
            profile_path = self._view.main_window.Recovery_profile_path

        scrollArea.setWidgetResizable(True)
        self.scrollAreaWidgetContents = QWidget()
        self.profiles_layout = QGridLayout(self.scrollAreaWidgetContents)
        self.profiles_layout.setVerticalSpacing(1)
        self.profiles_layout.setHorizontalSpacing(1)
        scrollArea.setWidget(self.scrollAreaWidgetContents)

        start_btn.setEnabled(False)
        if check_all.isChecked(): check_all.setChecked(False)
        self._displayErrorMessage('')
        self.clearLayout(self.profiles_layout)
        profile_path.setText(QFileDialog.getExistingDirectory(caption='Select Profile\'s path', dir=profile_path.text()))
        

    def clearLayout(self, layout):
        if layout is not None:
            while layout.count():
                item = layout.takeAt(0)
                widget = item.widget()
                if widget is not None:
                    widget.deleteLater()
                else:
                    self.clearLayout(item.layout())

    def load_profiles(self):
        if self.tab_widget.currentWidget().objectName() == 'emailComposer':
            scrollArea = self._view.main_window.scrollArea
            start_btn = self._view.main_window.Composer_start_Btn
            check_all = self._view.main_window.ComposerCheck_all
            profile_path = self._view.main_window.Composer_profile_path

        elif self.tab_widget.currentWidget().objectName() == 'recovery':
            scrollArea = self._view.main_window.Recovery_scrollArea
            start_btn = self._view.main_window.start_Recovery_Btn
            check_all = self._view.main_window.RecoveryCheck_all
            profile_path = self._view.main_window.Recovery_profile_path

        scrollArea.setWidgetResizable(True)
        self.scrollAreaWidgetContents = QWidget()
        self.profiles_layout = QGridLayout(self.scrollAreaWidgetContents)
        self.profiles_layout.setVerticalSpacing(1)
        self.profiles_layout.setHorizontalSpacing(1)
        scrollArea.setWidget(self.scrollAreaWidgetContents)

        start_btn.setEnabled(False)
        if check_all.isChecked(): check_all.setChecked(False)
        self._displayErrorMessage('')
        self.clearLayout(self.profiles_layout)
        if profile_path.text().strip() == '':
            self._displayErrorMessage('<b>Profile\'s path is required!</b>')
            return
        try:
            path = os.listdir(profile_path.text())
            profiles = []
            for profile in path:
                if profile_path.text().__contains__('Chrome'):
                    if profile.__contains__('Profile') and not profile.__contains__('Guest') and not profile.__contains__('System'):
                        prf = profile
                        profiles.append(prf)
                    else: continue
                    i = profiles.index(prf)
                    checkBox = QCheckBox(prf)
                    if self.tab_widget.currentWidget().objectName() == 'emailComposer':
                        checkBox.setObjectName(f'Composer_{prf}')
                    elif self.tab_widget.currentWidget().objectName() == 'recovery':
                        checkBox.setObjectName(f'Recovery_{prf}')
                else:
                    i = path.index(profile)
                    if profile.split('.')[0] == '':
                        self._displayErrorMessage('<b>Profile\'s path is invalid!</b>')
                        break
                    try:
                        checkBox = QCheckBox(profile.split('.')[1])
                    except:
                        checkBox = QCheckBox(profile)
                    if self.tab_widget.currentWidget().objectName() == 'emailComposer':
                        checkBox.setObjectName(f'Composer_{profile}')
                    elif self.tab_widget.currentWidget().objectName() == 'recovery':
                        checkBox.setObjectName(f'Recovery_{profile}')
                checkBox.stateChanged.connect(self.check)
                row = math.floor(i/4)
                self.profiles_layout.setColumnStretch(i % 4, 1)
                self.profiles_layout.setRowStretch(row, 1)
                self.profiles_layout.addWidget(checkBox, row, i % 4)
                start_btn.setEnabled(True)
        except FileNotFoundError:
            self._displayErrorMessage('<b>Profile\'s path is invalid!</b>')
        except IndexError:
            self._displayErrorMessage('<b>Profile\'s path is invalid!</b>')
        except Exception as e:
            if str(e).__contains__('syntax is incorrect'):
                self._displayErrorMessage('<b>Profile\'s path is invalid!</b>')
            else:
                self._displayErrorMessage(f'<b>{str(e)}</b>')

    def showlogers(self) :
        if self.tab_widget.currentWidget().objectName() == 'OutlookCreator':
            self._view.main_window.stackedWidget.setCurrentIndex(1)
        if self.tab_widget.currentWidget().objectName() == 'recovery':
            self._view.main_window.stackedWidget_R.setCurrentIndex(1)
        if self.tab_widget.currentWidget().objectName() == 'emailComposer':
            self._view.main_window.Composer_stackedWidget.setCurrentIndex(1)
        if self.tab_widget.currentWidget().objectName() == 'searchEmails':
            self._view.main_window.stackedWidget_Search.setCurrentIndex(1)
    
    def showtables(self) :
        if self.tab_widget.currentWidget().objectName() == 'OutlookCreator':
            self._view.main_window.stackedWidget.setCurrentIndex(0)
        if self.tab_widget.currentWidget().objectName() == 'recovery':
            self._view.main_window.stackedWidget_R.setCurrentIndex(0)
        if self.tab_widget.currentWidget().objectName() == 'emailComposer':
            self._view.main_window.Composer_stackedWidget.setCurrentIndex(0)
        if self.tab_widget.currentWidget().objectName() == 'searchEmails':
            self._view.main_window.stackedWidget_Search.setCurrentIndex(0)
    
    def GenerateEmails_window(self) :
        #show the generate pop up window 
        self._view.generate_window.show()
        
    def _displayWelcome(self):
        self.welcome_label_message = self._view.main_window.findChild(QLabel, 'label_welcome')
        self.welcome_label_message.setText('<html><head/><body><p><span style=" font-weight:600;">Welcome, %s</span></p></body></html>'%(self.user))

        self.version_label_message = self._view.main_window.findChild(QLabel, 'label_version')
        self.version_label_message2 = self._view.main_window.findChild(QLabel, 'label_version_2')
        self.version_label_message3 = self._view.main_window.findChild(QLabel, 'label_version_3')
        self.version_label_message4 = self._view.main_window.findChild(QLabel, 'label_version_4')
        self.composer_version_label_message = self._view.main_window.findChild(QLabel, 'Composer_label_version')
        self.version_label_message.setText(GPORTALUI_VERSION)
        self.version_label_message2.setText(GPORTALUI_VERSION)
        self.version_label_message3.setText(GPORTALUI_VERSION)
        self.version_label_message4.setText(GPORTALUI_VERSION)
        self.composer_version_label_message.setText(GPORTALUI_VERSION)

    def _displayErrorMessage(self, text):
        self.error_label_message = self._view.main_window.findChild(QLabel, 'label_error_message')
        self.error_label_message.setText("")
        self.error_label_message.setText('<html><head/><body><p align="center"><span style=" font-size:8pt; color:#ff0000;">%s</span></p></body></html>'%text)
        self.error_label_message.adjustSize()

    def loginHandler(self):
        """ gportal login validation """

        username = self._view.login_window.findChild(QLineEdit, 'username')
        password = self._view.login_window.findChild(QLineEdit, 'password')
        self.user = "alaa_waladi"
        self.password = "pRtnCrm5"
        # self.user = str(username.text()).rstrip()
        # self.password = str(password.text()).rstrip()
        try:
            url_post = "http://www.gvistas.com:4808/api/v1/token/"
            payload = {"username":self.user, "password":self.password}
            headers = {'content-type': 'application/json'}
            response = requests.post(url_post, data=json.dumps(payload), headers=headers, timeout=10)
            
            status_code = response.status_code
            content_result = response._content
            res_token = str(content_result, 'utf-8').rstrip()

            username.clear()
            password.clear()
            json_data = json.loads(res_token)
            if "error" in json_data: self._view._setLoginErrorMessage(json_data["error"])
            elif status_code == 200: 
                self.is_logged = True
                self._view.login_window.close()
                self._view.main_window.show()
                self._displayWelcome()

            elif status_code == 403: self._view._setLoginErrorMessage("Access to this application is denied. Please contact administrator!")
            else: self._view._setLoginErrorMessage("Access to this application failed with status %s. Please contact administrator!"%status_code)

        except requests.exceptions.ConnectTimeout:
            self._view._setLoginErrorMessage("ConnectTimeout: Login failed! try later or contact our helpdesk")
        except Exception as e:
            self._view._setLoginErrorMessage("Exception: Login failed! try later or contact out helpdesk")

    def print_output(self, response=None):
        if response and "status" in response and not response["status"]:
            self._displayErrorMessage(response["message"])
        elif response and "status" in response and response["status"]:
            self.new_obj.logger.info(response["message"])

    def newsCheckAll(self):
        if self._view.main_window.news_checkall.isChecked():
            self._view.main_window.checkBox_contentinstitute.setChecked(True)
            self._view.main_window.checkBox_litmus.setChecked(True)
            self._view.main_window.checkBox_hubspot.setChecked(True)
            self._view.main_window.checkBox_skimm.setChecked(True)
            self._view.main_window.checkBox_generalassemb.setChecked(True)
            self._view.main_window.checkBox_thehustle.setChecked(True)
            self._view.main_window.checkBox_typography.setChecked(True)
            self._view.main_window.checkBox_buffer.setChecked(True)
            self._view.main_window.checkBox_atlasobscura.setChecked(True)
        else:
            self._view.main_window.checkBox_contentinstitute.setChecked(False)
            self._view.main_window.checkBox_litmus.setChecked(False)
            self._view.main_window.checkBox_hubspot.setChecked(False)
            self._view.main_window.checkBox_skimm.setChecked(False)
            self._view.main_window.checkBox_generalassemb.setChecked(False)
            self._view.main_window.checkBox_thehustle.setChecked(False)
            self._view.main_window.checkBox_typography.setChecked(False)
            self._view.main_window.checkBox_buffer.setChecked(False)
            self._view.main_window.checkBox_atlasobscura.setChecked(False)

    def thread_complete(self):
        #...Enable start process
        self._view.main_window.newsStartBtn.setEnabled(True)
        self._view.main_window.newsStopBtn.setEnabled(False)
        
    def print_output_o(self, response=None):
        if response and "status" in response and not response["status"]:
            self._displayErrorMessage(response["message"])
        elif response and "status" in response and response["status"]:
            self.outlook_C.logger.info(response["message"])
    
    def O_Create(self) :
        kwargs = {}
        self.outlook_C = None
        log_create = []
        # OutlookCreate(kwargs)
        if self._view.main_window.create_btn.isChecked(): log_create.append('create')
        
        if len(log_create)< 1 :
            return {"status" : False ,"message": "Error: Selecte Login Or Create To Start!" }
  
        Outlook_data = self._view.main_window.email_list_text.toPlainText().split("\n")    
        #... Validation data
        
        OutlookCreator_Data = []        
        for line in Outlook_data :
            status, message = OutlookCreate(kwargs).create_data(line)
            i = "<br>  Create format  : <b>Email, password, Proxy,First Name, Last Name, Country code, Birthday(dd/mm/YYYY)</b> " 
            if not status :
                return {"status": False, "message": "Error: invalide lines format! <b>(%s)</b>"%message+i}
            OutlookCreator_Data.append(message)
            
           
        kwargs['browser'] = self._view.main_window.outlook_browser.currentText()
        kwargs['launguage_browser'] = self._view.main_window.outlook_laung.currentText()
        kwargs['Creatorloger'] = self._view.main_window.Creatorloger
        kwargs['tableWidget'] = self._view.main_window.tableWidget
        self.outlook_C = OutlookCreate(kwargs)
        self.log_table = OutlookCreate(kwargs)
        self.outlook_C.runlogger()
        self.outlook_C.insertNewRow.connect(self.addNewRow)
        self.outlook_C.appendData.connect(self.addData)
        self.outlook_C.logger.info("Start Process ...")
        try:
            for row in OutlookCreator_Data:
                if self.stop_create: 
                    break
                for action in log_create:
                    #...check if process is stopped
                    self.outlook_C.set_Stop_Process(self.stop_create)
                    if action=="create": self.outlook_C.create_outlook(row,self.view_table,self.item)
            
            self.outlook_C.destroyDriver()

        except Exception as e:
            return {"status": False, "message": "Exception: %s"%(e)}

        return {"status": True, "message": "Finished"}
        
    def runNewsProcess(self):
        
        kwargs = {}   
        self.new_obj = None
        #...validate newsletters
        newsletters_checkbox = []
        if self._view.main_window.checkBox_contentinstitute.isChecked(): newsletters_checkbox.append("checkBox_contentinstitute")
        if self._view.main_window.checkBox_litmus.isChecked(): newsletters_checkbox.append("checkBox_litmus")
        if self._view.main_window.checkBox_hubspot.isChecked(): newsletters_checkbox.append("checkBox_hubspot")
        if self._view.main_window.checkBox_skimm.isChecked(): newsletters_checkbox.append("checkBox_skimm")
        if self._view.main_window.checkBox_generalassemb.isChecked(): newsletters_checkbox.append("checkBox_generalassemb")
        if self._view.main_window.checkBox_thehustle.isChecked(): newsletters_checkbox.append("checkBox_thehustle")
        if self._view.main_window.checkBox_typography.isChecked(): newsletters_checkbox.append("checkBox_typography")
        if self._view.main_window.checkBox_buffer.isChecked(): newsletters_checkbox.append("checkBox_buffer")
        if self._view.main_window.checkBox_atlasobscura.isChecked(): newsletters_checkbox.append("checkBox_atlasobscura")
        if len(newsletters_checkbox)<1: 
            return {"status": False, "message": "Error: at least one newsletter should be selected!"}
        
        #...validate data
        newsletters_data = self._view.main_window.news_email.toPlainText().split("\n")
        if len(newsletters_data)>50: 
            return {"status": False, "message": "Error: max allowed email list is reached!!"}
        
        news_clean_data = []
        for line in newsletters_data:
            status, message = NewsLetterSubscribe(kwargs).validate_line(line)
            i = '<br> Use : <b> Email, Proxy , Port </b> '
            if not status: 
                return {"status": False, "message": "Error: invalide lines format! <b>(%s)</b>"%message+i}
            news_clean_data.append(message)
        
        kwargs["browser"] =  self._view.main_window.news_browser.currentText()
        kwargs["hide_browser"] =  self._view.main_window.news_hidebrowser.isChecked()
        kwargs["logger_window"] = self._view.main_window.newsletter_logger
        
        self.new_obj = NewsLetterSubscribe(kwargs)
        self.new_obj.runlogger()
        self.new_obj.logger.info("Start Process")
        #...start process
        try:
            for row in news_clean_data:
                for news in newsletters_checkbox:
                    #...check if process is stopped
                    if self.stop_create:
                        self.new_obj.logger.warning("Stopped")
                        break
                    if news=="checkBox_contentinstitute": self.new_obj._contentinstitute(row)
                    if news=="checkBox_litmus": self.new_obj._litmus(row)
                    if news=="checkBox_hubspot": self.new_obj._hubspot(row)
                    if news=="checkBox_skimm": self.new_obj._skimm(row)
                    if news=="checkBox_generalassemb": self.new_obj._generalassemb(row)
                    if news=="checkBox_thehustle": self.new_obj._thehustle(row)
                    if news=="checkBox_typography": self.new_obj._typography(row)
                    if news=="checkBox_buffer": self.new_obj._buffer(row)
                    if news=="checkBox_atlasobscura": self.new_obj._atlasobscura(row)
            
            self.new_obj.destroyDriver()

        except Exception as e:
            return {"status": False, "message": "Exception: %s"%(e)}

        return {"status": True, "message": "Finished"}
    # get Isp data
    def get_isp(self) :
        if self.tab_widget.currentWidget().objectName() == 'emailComposer':
            isp = self._view.main_window.Composer_Switch_ISP.currentText()
        if self.tab_widget.currentWidget().objectName() == 'recovery':
            isp = self._view.main_window.R_Switch_ISP.currentText() 
        return isp
    
    def getSubject(self):
        if self.tab_widget.currentWidget().objectName() == 'recovery':
            search_keyword = self._view.main_window.findChild(QLineEdit, 'R_subject')
            search_keyword = str(search_keyword.text()).rstrip()
            return search_keyword
        if self.tab_widget.currentWidget().objectName() == 'emailComposer':
            subject = self._view.main_window.findChild(QLineEdit, 'Composer_subject')
            subject = str(subject.text()).rstrip()
            return subject
        
    def getBody(self):
        body = self._view.main_window.Composer_body_input.toPlainText()
        return body
    
    def combobox_changed_value(self):
        if self.tab_widget.currentWidget().objectName() == 'recovery':
            if self.get_isp() == "Outlook":
                self._view.main_window.Move_to_focused_chekbox.setEnabled(True)
            else:
                self._view.main_window.Move_to_focused_chekbox.setEnabled(False)
            if self.get_isp() == "Gmail":
                self._view.main_window.Recovery_browser.setCurrentIndex(0)
                self._view.main_window.Recovery_browser.model().item(1).setEnabled(False)
                self._view.main_window.Recovery_hidebrowser.setEnabled(False)

            else:
                self._view.main_window.Recovery_browser.model().item(1).setEnabled(True)
                self._view.main_window.Recovery_hidebrowser.setEnabled(True)


        if self.tab_widget.currentWidget().objectName() == 'emailComposer':

            if self.get_isp() == "Gmail":
                self._view.main_window.Composer_hidebrowser.setEnabled(False)

            else:
                self._view.main_window.Composer_hidebrowser.setEnabled(True)
    
    def addNewRow(self, rowNumber):
        if self.tab_widget.currentWidget().objectName() == 'emailComposer':
            self._view.main_window.Composer_table.setRowCount(rowNumber)
        elif self.tab_widget.currentWidget().objectName() == 'OutlookCreator':
            self.view_table.setRowCount(rowNumber)
        else:
            self._view.main_window.table_Recovery.setRowCount(rowNumber)
    
    def addData(self, data):
        if self.tab_widget.currentWidget().objectName() == "emailComposer":
            table = self._view.main_window.Composer_table
        elif self.tab_widget.currentWidget().objectName() == 'OutlookCreator':
            table = self.view_table
        else:
            table = self._view.main_window.table_Recovery
        for item in data:
            table.setItem(item[0], item[1], self.item.QTableWidgetItem(item[2]))

    def check(self, state):
        checkBoxes = []
        if self.tab_widget.currentWidget().objectName() == 'emailComposer':
            contains = 'Composer_'
            check_all = self._view.main_window.ComposerCheck_all
        elif self.tab_widget.currentWidget().objectName() == 'recovery':
            contains = 'Recovery_'
            check_all = self._view.main_window.RecoveryCheck_all

        for profile_checkBox in self._view.main_window.findChildren(QCheckBox):
            if profile_checkBox.objectName().__contains__(f'{contains}'):
                checkBoxes.append(profile_checkBox)
        
        if state == Qt.Checked and not check_all.isChecked() and all(checkBox.isChecked() for checkBox in checkBoxes):
            check_all.setChecked(True)
        elif state == Qt.Unchecked and check_all.isChecked():
            check_all.setChecked(False)

    def check_all_profiles(self):
        checkBoxes = []
        if self.tab_widget.currentWidget().objectName() == 'emailComposer':
            contains = 'Composer_'
            check_all = self._view.main_window.ComposerCheck_all
        elif self.tab_widget.currentWidget().objectName() == 'recovery':
            contains = 'Recovery_'
            check_all = self._view.main_window.RecoveryCheck_all

        for profile_checkBox in self._view.main_window.findChildren(QCheckBox):
            if profile_checkBox.objectName().__contains__(f'{contains}'):
                checkBoxes.append(profile_checkBox)

        if check_all.isChecked():
            for checkBox in checkBoxes:
                checkBox.setChecked(True)
        else:
            for checkBox in checkBoxes:
                checkBox.setChecked(False)

        if check_all.isChecked() and not all(checkBox.isChecked() for checkBox in checkBoxes):
            check_all.setChecked(False)
        elif not check_all.isChecked() and all(checkBox.isChecked() for checkBox in checkBoxes):
            check_all.setChecked(True)
        
    #******************* Start Recovery *******************#

    def getCheckedBoxes(self):
        start_apps = []
        if self._view.main_window.Open_inbox_chekbox.isChecked():start_apps.append('open_inbox')
        if self._view.main_window.open_spam_chekbox.isChecked():start_apps.append('open_spam')
        if self._view.main_window.not_spam_chekbox.isChecked():start_apps.append('not_spam')
        if self._view.main_window.Move_to_focused_chekbox.isChecked():start_apps.append('move_to_focused')
        return start_apps
        
    def recovery_apps_Process(self) :
        if self.stop_process:
            return

        kwargs = {}
        self.recovery = None

        kwargs["browser"] =  self._view.main_window.Recovery_browser.currentText()
        kwargs["hide_browser"] =  self._view.main_window.Recovery_hidebrowser.isChecked()
        kwargs["logger_window"] = self._view.main_window.Recovery_logger
        kwargs['launguage_browser'] = self._view.main_window.Recovery_laung.currentText()

        if self.get_isp() == 'Outlook':
            self.recovery = Recovery_Apps(kwargs)
        elif self.get_isp() == 'Gmail':
            self.recovery = Gmail_Recovery(kwargs)
        elif self.get_isp() == 'Yahoo':
            self.recovery = Yahoo_Recovery(kwargs)
        self.recovery.setStop_Recovery_Process(self.stop_process)
        self.recovery.insertNewRow.connect(self.addNewRow)
        self.recovery.appendData.connect(self.addData)
        
        # check profile path
        if self._view.main_window.Recovery_profile_path.text().strip() == '':
            return {'status':False, 'message':'<b>Profile\'s path is required!</b>'}
        
        try:
            for checkBox in self._view.main_window.findChildren(QCheckBox):
                if checkBox.objectName().__contains__('Recovery_'):
                    if checkBox.isChecked():
                        self.profiles.append(f"{self._view.main_window.Recovery_profile_path.text().strip()}/{checkBox.objectName().replace('Recovery_', '')}")
            if len(self.profiles) == 0:
                return {'status':False, 'message':'<b>Check one profile at least!</b>'}
        except FileNotFoundError:
            return {'status': False, 'message': '<b>Profile\'s path is invalid!</b>'}
        except Exception as e:
            if str(e).__contains__('syntax is incorrect'):
                return {'status': False, 'message': '<b>Profile\'s path is invalid!</b>'}
            else:
                return {'status':False, 'message':f'<b>{str(e)}</b>'}

        #check start_apps
        if len(self.getCheckedBoxes())<1:
            return {"status":False, "message":"Error You must chose at least one Action !"}

        #check keyword
        search_keyword = self.getSubject()
        if len(search_keyword) == 0:
            return {"status":False, "message":"<b>Insert Your Search Keyword</b>"}

        #check date
        date_from = self._view.main_window.findChild(QDateEdit, 'dateEdit')
        date_from = datetime.strptime(date_from.text(), "%m/%d/%Y").date()
        
        date_to = self._view.main_window.findChild(QDateEdit, 'dateTO')
        date_to = datetime.strptime(date_to.text(), "%m/%d/%Y").date()
        date_to = date_to + timedelta(days=1)

        
        #...start process
        for profile in self.profiles:
            if self.stop_process:
                break
            try:
                start_apps = str(self.getCheckedBoxes())
                rowCount = self._view.main_window.table_Recovery.rowCount()
                if self.get_isp() == 'Outlook':
                    date_to = date_to - timedelta(days=1)
                    self.recovery._outlook_acount(rowCount, start_apps, date_from, search_keyword, date_to, profile)
                elif self.get_isp() == 'Gmail':
                    self.recovery.gmail_account(rowCount, start_apps, date_from, search_keyword, date_to, profile)
                elif self.get_isp() == 'Yahoo':
                    self.recovery._Yahoo_account(rowCount, start_apps, date_from, search_keyword, date_to, profile)
            except Exception as e:
                return {"status": False, "message": f"Exception: {e}"}
        self.recovery.logger.info('Finished')
    
    def thread_complete_R(self):
        self.recovery.closeLogger()
        #...Enable start process
        self.enableRecoveryWidgets(True)
        self.stop_process = False
        self.threadpool.waitForDone()
    
    
    def print_output_R(self, response=None):
        if response and "status" in response and not response["status"]:
            self._displayErrorMessage(response["message"])
        elif response and "status" in response and response["status"]:
            self.recovery.logger.info(response["message"])

    def scrollDown(self):
        if self.tab_widget.currentWidget().objectName() == 'recovery':
            scrollBar = self._view.main_window.Recovery_logger.verticalScrollBar()

        if self.tab_widget.currentWidget().objectName() == 'emailComposer':
            scrollBar = self._view.main_window.Composer_logger.verticalScrollBar()
                
        if self.tab_widget.currentWidget().objectName() == 'OutlookCreator':
            scrollBar = self._view.main_window.Creatorloger.verticalScrollBar()

        atBottom = (scrollBar.maximum() - scrollBar.value()) < 20
        if atBottom:
            scrollBar.setValue(scrollBar.maximum())

    def enableRecoveryWidgets(self, enabled):
        self._view.main_window.start_Recovery_Btn.setEnabled(enabled)
        self._view.main_window.stop_Recovery_Btn.setEnabled(not(enabled))
        self._view.main_window.clear_btn_Recovery.setEnabled(enabled)
        self._view.main_window.Recovery_profile_path.setEnabled(enabled)
        self._view.main_window.Recovery_load_profile.setEnabled(enabled)
        self._view.main_window.Recovery_file_dialog.setEnabled(enabled)
        self._view.main_window.Recovery_scrollArea.setEnabled(enabled)
        self._view.main_window.RecoveryCheck_all.setEnabled(enabled)
        self._view.main_window.R_subject.setEnabled(enabled)
        self._view.main_window.dateEdit.setEnabled(enabled)
        self._view.main_window.dateTO.setEnabled(enabled)
        self._view.main_window.Open_inbox_chekbox.setEnabled(enabled)
        self._view.main_window.open_spam_chekbox.setEnabled(enabled)
        self._view.main_window.not_spam_chekbox.setEnabled(enabled)
        self._view.main_window.Recovery_browser.setEnabled(enabled)
        self._view.main_window.Recovery_laung.setEnabled(enabled)
        self._view.main_window.R_Switch_ISP.setEnabled(enabled)
        self._view.main_window.Recovery_hidebrowser.setEnabled(enabled)

    def RecoveryStartProcess(self):
        self.profiles = []
        self._displayErrorMessage("")
        self.enableRecoveryWidgets(False)
        self.clear_Recovery()

        # Pass the function to execute
        work = Worker(self.recovery_apps_Process)
        work.signals.result.connect(self.print_output_R)
        work.signals.finished.connect(self.thread_complete_R)
        self.threadpool.start(work)
    
    def Recoverstop(self):
        self.stop_process = True
        self.recovery.setStop_Recovery_Process(stopped=self.stop_process)
    
    def clear_Recovery(self):
        XStream.msgList.clear()
        self._view.main_window.Recovery_logger.clear()
        self._view.main_window.table_Recovery.setRowCount(0)
    
    #********************* End Recovery *******************#

    #********************* Start Search Emails *******************#

    def toggle_custom_search_emails(self):
        if self._view.main_window.checkBox_Custom_Search.isChecked():
            self._view.main_window.lineEdit_Custom_Search.setEnabled(True)
        else:
            self._view.main_window.lineEdit_Custom_Search.setEnabled(False)

    def getBox(self):
        if self.junk.isChecked():
            return "Junk"
        else:
            return "Inbox"
    
    def getResults(self):
        self.results = []
        results = self._view.main_window.findChildren(QCheckBox)
        for result in results:
            if result.isChecked():
                if result.text() == "Custom":
                    self.results.append(self._view.main_window.lineEdit_Custom_Search.text())
                else:
                    self.results.append(result.text())
        return self.results
    def clear_search_emails(self):
        self._view.main_window.Search_Logger.clear()
        self._view.main_window.Table_Search.setRowCount(0)

    def stop_search_emails(self):
        self.search_email.setStopProcess(True)

    def getSearchEmails(self, emails):
        account = emails.split(",")
        if len(account) != 2:
            return False, {"message": f"<b>Error: Invalid Email {account}<br>Format: email,password</b>"}
        email_address = account[0]
        password = account[1]
        email_match = re.search(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", email_address)
        if not(email_match):
            return False, {"message": f"Check Email {email_address}"}
        if password[0] == " " or email_address[-1] == " ":
            return False, {"message": f"Error: {account}"}
        return True, {"email": email_address, "password": password}

    def start_search_process(self):
        Search_Emails(self._view.main_window.Search_Logger, self).setStopProcess(False)
        self._view.main_window.radioButton_Logger_Search.setChecked(True)
        self._view.main_window.radioButton_Table_Search.setEnabled(False)
        self.clear_search_emails()
        self._displayErrorMessage("")
        self._view.main_window.pushButton_Start_Search.setEnabled(False)
        self._view.main_window.pushButton_Stop_Search.setEnabled(True)
        self._view.main_window.pushButton_Clear_Search.setEnabled(False)
        
        worker = Worker(self.start_search_emails)
        worker.signals.result.connect(self.print_output_Search)
        worker.signals.finished.connect(self.thread_complete_Search)
        self.threadpool.start(worker)
    
    def thread_complete_Search(self):
        #...Enable start process
        self._view.main_window.pushButton_Start_Search.setEnabled(True)
        self._view.main_window.pushButton_Stop_Search.setEnabled(False)
        self._view.main_window.pushButton_Clear_Search.setEnabled(True)
        self.threadpool.waitForDone()

    def print_output_Search(self, response=None):
        if response and "status" in response and not response["status"]:
            self._displayErrorMessage(response["message"])
        elif response and "status" in response and response["status"]:
            self.search_email.logger.info(response["message"])
    
    def start_search_emails(self):
        ### Get Date
        date_from = self._view.main_window.dateEdit_From_Search
        date_from = datetime.strptime(date_from.text(), "%m/%d/%Y").date()
        date_to = self._view.main_window.dateEdit_To_Search
        date_to = datetime.strptime(date_to.text(), "%m/%d/%Y").date()
        if date_from > date_to:
            self._view.main_window.dateEdit_From_Search.setFocus()
            return {"status":False, "message":"<b>Check Date Fields !</b>"}
        date_to = date_to + timedelta(days=1)

        ### Get Subject
        subject = self._view.main_window.lineEdit_Subject_Search.text()
        if subject == "":
            self._view.main_window.lineEdit_Subject_Search.setFocus()
            return {"status":False, "message": "<b>Insert a subject !</b>"}

        ### Get Emails
        result = self.getResults()
        if len(result) == 0:
            return {"status":False, "message": "<b>Error: You must select one result at least !<b/>"}
        box = self.getBox()
        self.search_email = Search_Emails(self._view.main_window.Search_Logger, self)
        self.search_email.runlogger()
        self.valide_line = []
        account_list = self._view.main_window.textEdit_Accounts_List_Search.toPlainText()
        emails = account_list.split("\n")
        for item in emails:
            status, message = self.getSearchEmails(item)
            if not status:
                self._view.main_window.textEdit_Accounts_List_Search.setFocus()
                return {"status":False, "message":message}
            self.valide_line.append(message)
        self.search_email.logger.info("Start Process ...")
        time.sleep(1)
        for row in self.valide_line:
            email_address = row["email"]
            password = row["password"]
            if self.search_email.getStopProcess():
                self.search_email.logger.warning("Stopped")
                break
            self.search_email.search(email_address, password, date_from, date_to, subject, result, box)
            if self.search_email.getStopProcess():
                break
        self.search_email.logger.info("Finished")
        time.sleep(1)
        self._view.main_window.radioButton_Table_Search.setEnabled(True)
        if not(self.search_email.getStopProcess()) and self.search_email.numberItems != 0:
            self._view.main_window.radioButton_Table_Search.setChecked(True)
            self._view.main_window.stackedWidget_Search.setCurrentIndex(0)
    
    def displayResult(self, results):
        NumberOfItems = len(self.getResults())
        self._view.main_window.Table_Search.setColumnCount(NumberOfItems)
        self._view.main_window.Table_Search.setHorizontalHeaderLabels(self.getResults())
        for row_number, Item in enumerate(results):
            self._view.main_window.Table_Search.insertRow(row_number)
            for col_number, key in enumerate(Item):
                self._view.main_window.Table_Search.setItem(row_number, col_number, QTableWidgetItem(str(Item[key])))
        header = self._view.main_window.Table_Search.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

    #********************** End Search Emails *******************#

    #********************** Start Compose Emails ****************#

    def enableComposerWidgets(self, boolean):
        self._view.main_window.Composer_start_Btn.setEnabled(boolean)
        self._view.main_window.Composer_stop_Btn.setEnabled(not(boolean))
        self._view.main_window.Composer_clear_btn.setEnabled(boolean)
        self._view.main_window.Composer_file_dialog.setEnabled(boolean)

        self._view.main_window.Composer_subject.setEnabled(boolean)
        self._view.main_window.Composer_recipients_input.setEnabled(boolean)
        self._view.main_window.Composer_body_input.setEnabled(boolean)
        self._view.main_window.Composer_profile_path.setEnabled(boolean)
        
        self._view.main_window.Composer_Switch_ISP.setEnabled(boolean)
        self._view.main_window.Composer_lang.setEnabled(boolean)
        self._view.main_window.Composer_send_limit.setEnabled(boolean)
        self._view.main_window.Composer_load_profile.setEnabled(boolean)

        self._view.main_window.scrollArea.setEnabled(boolean)
        self._view.main_window.ComposerCheck_all.setEnabled(boolean)
        self._view.main_window.Composer_loop.setEnabled(boolean)

    def getRecipients(self):
        if self._view.main_window.Composer_recipients_input.toPlainText().strip() == '':
            return {'status': False, 'message': '<b>Add emails to send list!</b>'}
        recipients = self._view.main_window.Composer_recipients_input.toPlainText().split('\n')
        valid_emails = []
        for email in recipients:
            match_eml = re.search(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", email)
            if match_eml:
                valid_emails.append(email)
            elif not match_eml:
                return {"status": False, "message": f"<b>Send List: invalid lines formats!</b>"}
        return valid_emails

    def clear_composer(self):
        XStream.msgList.clear()
        self._view.main_window.Composer_logger.clear()
        self._view.main_window.Composer_table.setRowCount(0)

    def composer_thread_complete(self):
        try:
            self.instance[0].logger.info("Finished")
            for item in self.instance:
                item.closeLogger()
        except:pass
        #...Enable start process
        self.enableComposerWidgets(True)
        self.stop_process = False
        self.threadpool.waitForDone()
    
    def composer_stop(self):
        if self.get_isp() == "Gmail":
            for item in self.instance:
                self.stop_process = True
                item.setStop_Composer_Process(stopped=self.stop_process)
    
    def composer_start_process(self):
        self.workers = []
        self.instance = []
        self.profiles = []
        self._displayErrorMessage("")
        self.enableComposerWidgets(False)
        self.clear_composer()

        # Pass the function to execute
        if self.get_isp() == "Gmail":
            work = Worker(self.start_gmail_composer)
            work.signals.result.connect(self.print_output_R)
            work.signals.finished.connect(self.composer_thread_complete)
            self.threadpool.start(work)
    
    def start_gmail_composer(self):
        #check subject
        self.subject = self.getSubject()
        if len(self.subject.strip()) == 0:
            return {"status":False, "message":"<b>Insert a Subject</b>"}

        # check recipients
        self.recipients = self.getRecipients()
        if str(type(self.recipients)) == "<class 'dict'>" and not self.recipients['status']:
            return {'status': False, 'message': self.recipients['message']}
        
        # check body
        self.body = self.getBody()
        self.body = self.body.strip()
        if len(self.body) < 1:
            return {"status":False, "message":"<b>Insert a message body</b>"}
        
        # check send limit
        self.send_limit = int(self._view.main_window.Composer_send_limit.value())
        if self.send_limit < 1:
            return {'status':False, 'message':'<b>The minimum value of send limit should be 1 or more!</b>'}
        
        # check loop
        self.loop = int(self._view.main_window.Composer_loop.value())
        if self.loop < 1:
            return {'status':False, 'message':'<b>The minimum value of loop should be 1 or more!</b>'}
        
        # check profile path
        if self._view.main_window.Composer_profile_path.text().strip() == '':
            return {'status':False, 'message':'<b>Profile\'s path is required!</b>'}
        
        try:
            for checkBox in self._view.main_window.findChildren(QCheckBox):
                if checkBox.objectName().__contains__('Composer_'):
                    if checkBox.isChecked():
                        self.profiles.append(os.path.join(self._view.main_window.Composer_profile_path.text().strip(), checkBox.objectName().replace('Composer_', '')))
            if len(self.profiles) == 0:
                return {'status':False, 'message':'<b>Check one profile at least!</b>'}
        except FileNotFoundError:
            return {'status': False, 'message': '<b>Profile\'s path is invalid!</b>'}
        except Exception as e:
            if str(e).__contains__('syntax is incorrect'):
                return {'status': False, 'message': '<b>Profile\'s path is invalid!</b>'}
            else:
                return {'status':False, 'message':f'<b>{str(e)}</b>'}

        for _ in range(self.loop):
            for profile in self.profiles:
                if self.stop_process or self.recipients == []:
                    break
                self.profile = profile
                self.gmail_composer_process()

    def gmail_composer_process(self):
        if self.stop_process or self.recipients == []:
            return
        
        kwargs = {}
        self.gmail_composer = None
        recipients = []
        for _ in range(self.send_limit):
            try:
                recipients.append(self.recipients[0])
                self.recipients.pop(0)
            except: pass

        kwargs["hide_browser"] =  self._view.main_window.Composer_hidebrowser.isChecked()
        kwargs["logger_window"] = self._view.main_window.Composer_logger
        kwargs['launguage_browser'] = self._view.main_window.Composer_lang.currentText()

        self.gmail_composer = Gmail_Composer(kwargs)
        self.gmail_composer.setStop_Composer_Process(stopped=self.stop_process)
        self.instance.insert(0, self.gmail_composer)
        self.instance[0].insertNewRow.connect(self.addNewRow)
        self.instance[0].appendData.connect(self.addData)
        
        try:
            rowCount = len(self.instance) - 1
            self.instance[0].gmail_account(rowCount=rowCount, item=self.item, subject=self.subject, body=self.body, recipients=recipients, profile=self.profile, send_limit=self.send_limit)
            self.stop_process = self.instance[0].getStop_Composer_Process()
        except Exception as e:
            return {"status": False, "message": f"Exception: {str(e)}"}

    #********************* End Compose Emails *******************#
    
    def Generate_data(self, line):
        params = line.split(',')
        if len(params)!=1: return False, line
        
        ip = params[0].strip()
      
        try:
            ipaddress.IPv4Address(ip)
        except: return False , "ip -> (%s)"%ip 
        #...match port
        self.ip = {"ip":ip}

        return  True , {"ip":ip}
    
    def Faker_generate(self) :
        self.new_obj = None
        domain_option = []
        
        if self._view.generate_window.domain_gen.currentText() == "outlook.com": domain_option.append("outlook.com")
        if self._view.generate_window.domain_gen.currentText() == "hotmail.com": domain_option.append("hotmail.com")
        generate_clean_data = []
        
        generate_data = self._view.generate_window.textEdit_gene.toPlainText().split("\n")
        
        if len(self._view.generate_window.textEdit_gene.toPlainText()) == 0: self.messageError("Insert IP")
        
        for line in generate_data:
            status, message = self.Generate_data(line)
            if not status: 
                return {"status": False, "message": "Error: invalide lines format! <b>(%s)</b>"%message}
            
            generate_clean_data.append(message)
        for E_domain in domain_option :pass
            
        email_list = []
        fake = Faker(['en'])
        k = generate_data
        s = len(k)
        count = 0
        for i in range(10):
            
            fname = fake.first_name()
            lname = fake.last_name()
            name = fname + '_' + lname
            num = fake.pyint(max_value = 9999)
            cnv = fake.date_between_dates(date_start=datetime(1988,1,1), date_end=datetime(2000,12,31))
            cnv = cnv.strftime('%d-%m-%Y')
            domain = timeout =  E_domain
            
            email = name + str(num) + '@' + domain + ',' + fake.pystr(max_chars = 10) + str(fake.pyint()) + ',' + k[count] + ',' + fname + ',' + lname + ',' + 'US' + ',' + str(cnv)
            email_list.append(email)
            count+=1
            if s == count :
                count = 0
        a = self.messagegenerate('Use Email List ? ')
        if a == QMessageBox.StandardButton.Yes:
            for i in email_list :
                self._view.main_window.email_list_text.appendPlainText(i)
            self._view.generate_window.close()
            self._view.generate_window.textEdit_gene.clear()
            self._view.main_window.setEnabled(True)
                
    def messagegenerate(self,msg) :
        reply = QMessageBox()
        reply.setIcon(QMessageBox.Question)
        reply.setWindowTitle("Email Generate")
        reply.setText(msg)
        reply.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        return reply.exec()
    
    def messageError(self,msgError):
        check = QMessageBox()
        check.setIcon(QMessageBox.Warning)
        check.setWindowTitle('Invalid Data')
        check.setText(msgError)
        check.setStandardButtons(QMessageBox.StandardButton.Cancel)
        return check.exec_()
    
    def newsStartProcess(self):
        
        #...disable start btn & clean logger
        self.stop_newsletter_process = False
        self._displayErrorMessage("")
        self._view.main_window.newsStartBtn.setEnabled(False)
        self._view.main_window.newsStopBtn.setEnabled(True)
        self._view.main_window.newsletter_logger.clear()

        # Pass the function to execute
        worker = Worker(self.runNewsProcess) # Any other args, kwargs are passed to the run function

        worker.signals.result.connect(self.print_output)
        worker.signals.finished.connect(self.thread_complete)
        #worker.signals.progress.connect(self.progress_fn)

        # Execute
        self.threadpool.start(worker)

    def thread_complete_o(self):
        #...Enable start process
        self.enableO_Widgets(True)
        try:
            if self._view.main_window.create_btn.isChecked():
                self.outlook_C.closeLogger()
            if self._view.main_window.login_btn.isChecked():
                self.login_log.closeLogger()
        except: pass

    def clear_outlook_creator(self):
        XStream.msgList.clear()
        self._view.main_window.Creatorloger.clear()

    def clear(self):
        XStream.msgList.clear()
        self._view.main_window.Creatorloger.clear()
        self._view.main_window.tableWidget.setRowCount(0)

    def enableO_Widgets(self, enabled:bool):
        self._view.main_window.email_list_text.setEnabled(enabled)

        self._view.main_window.create_btn.setEnabled(enabled)
        self._view.main_window.login_btn.setEnabled(enabled)
        self._view.main_window.outlook_browser.setEnabled(enabled)
        self._view.main_window.outlook_laung.setEnabled(enabled)

        self._view.main_window.generate_emails_btn.setEnabled(enabled)
        self._view.main_window.start_outlook.setEnabled(enabled)
        self._view.main_window.stop_outlook.setEnabled(not(enabled))

        self._view.main_window.clear_btn.setEnabled(enabled)
        self._view.main_window.export_csv_btn.setEnabled(enabled)
    
    def O_startProcess(self) :
        
        self.stop_create = False
        self._displayErrorMessage("")
        self.enableO_Widgets(False)
        self.clear_outlook_creator()
        # Pass the function to execute
        
        if self._view.main_window.login_btn.isChecked(): 
            worker = Worker(self.O_login)
            worker.signals.result.connect(self.print_output_login)
            worker.signals.finished.connect(self.thread_complete_o)
            self.threadpool.start(worker)
            
        else : 
        # worker = Worker(self.O_Create) # Any other args, kwargs are passed to the run function
            worker = Worker(self.O_Create)
            worker.signals.result.connect(self.print_output_o)
            worker.signals.finished.connect(self.thread_complete_o)
            # Execute
            self.threadpool.start(worker)
        
    def newsStopProcess(self):
        self.stop_newsletter_process = True
        
    def outllokStopP(self) :
        self.stop_create = True
        if self._view.main_window.create_btn.isChecked():
            self.outlook_C.set_Stop_Process(self.stop_create)
        elif self._view.main_window.login_btn.isChecked():
            self.login_log.set_Stop_Process(self.stop_create)
    
    def print_output_login(self, response=None):
        if response and "status" in response and not response["status"]:
            self._displayErrorMessage(response["message"])
        elif response and "status" in response and response["status"]:
            self.login_log.logger.info(response["message"])
        
    def O_login(self) :
        kwargs = {}
        self.new_obj_c = None
        log_create = []

        if self._view.main_window.login_btn.isChecked(): log_create.append('login')
        if len(log_create)< 1 :
            return {"status" : False ,"message": "Error: Selecte Login Or Create To Start!" }

        Outlook_data = self._view.main_window.email_list_text.toPlainText().split("\n")    
        #... Validation data
        
        OutlookLogin_Data = []        
    
        for lines in Outlook_data :
            status, message = OutlookCreate(kwargs).login_data(lines)
            i = "<br> Login Format : <b>Email, password, Proxy,First Name, Last Name,Country code, Birthday(dd/mm/YYYY) Or : Email,Password,Proxy </b>"
            if not status :
                return {"status": False, "message": "Error: invalide lines format! <b>(%s)</b>"%message+i}
            OutlookLogin_Data.append(message)
            
           
        kwargs['browser'] = self._view.main_window.outlook_browser.currentText()
        kwargs['launguage_browser'] = self._view.main_window.outlook_laung.currentText()
        kwargs['Creatorloger'] = self._view.main_window.Creatorloger
        kwargs['tableWidget'] = self._view.main_window.tableWidget
        
        self.login_log = OutlookCreate(kwargs)
        self.log_table = OutlookCreate(kwargs)
        self.login_log.runlogger()
        self.login_log.insertNewRow.connect(self.addNewRow)
        self.login_log.appendData.connect(self.addData)
        self.login_log.logger.info("Start Process ...")
        try:
            for row in OutlookLogin_Data:
                #...check if process is stopped    
                if self.stop_create: 
                    break
                self.login_log.set_Stop_Process(self.stop_create)
                for action in log_create:
                    if action=="login": self.login_log.login_outlook(row,self.view_table,self.item)
            
            self.login_log.destroyDriver()

        except Exception as e:
            return {"status": False, "message": "Exception: %s"%(e)}

        return {"status": True, "message": "Finished"}
    
    def CSV_data_Export(self) :
        user_agent = "Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:63.0) Gecko/20100101 Firefox/63.0"
        Row = self.view_table.rowCount()
        if Row > 0 :
            now = datetime.now()
            date_time = now.strftime("%Y%m%d_%H_%M_%S")
            dir = QFileDialog.getExistingDirectory()
            file_name = f"seed_hotmail_{date_time}.csv"
            path = os.path.join(dir, file_name)
            if dir : 
                r = open(path , mode='w')
                for i in range (Row) :
                    exp = f'{self.view_table.item(i, 1).text()}, {self.view_table.item(i, 2).text()}, {user_agent}, {self.view_table.item(i, 3).text()}, 3128'
                    r.write(exp)
                    if i == Row-1 :
                        pass
                    else :
                        r.write("\n")
                r.close()
                self.Export_alert("Successfully saved\n"+path)
        else :      
            self.Export_alert("No Data To Export ")
            
    def Export_alert(self,msg) :
        reply = QMessageBox()
        reply.setIcon(QMessageBox.Information)
        reply.setWindowTitle("Export Hotmail Seed")
        reply.setText(msg)
        reply.setStandardButtons(QMessageBox.StandardButton.Ok)
        
        return reply.exec()

def main():
    """ main function """
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    GportalUi(view=mainWindow)
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()