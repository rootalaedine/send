import logging
import time
from settings.config import *
from lib.webdriver_gmail import WebDriver_Gmail_Composer
from lib.html_clipboard import *
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from PyQt6.QtCore import QObject, pyqtSignal

class Gmail_Composer(WebDriver_Gmail_Composer, QObject):
    """Gmail Composer Class"""
    insertNewRow = pyqtSignal(int)
    appendData = pyqtSignal(list)
    removeItem = pyqtSignal(str)

    def __init__(self, kwargs):
        super().__init__()
        QObject.__init__(self)
        self.hide_browser = kwargs.get("hide_browser", False)
        self.launguage_browser= kwargs.get("launguage_browser", None)
        self.logger_window = kwargs.get("logger_window", None)
        self.tableWidget = kwargs.get('tableWidget',None)
    
    def getStop_Composer_Process(self):
        return self.Stop_Composer_Process

    def setStop_Composer_Process(self, stopped):
        self.Stop_Composer_Process = stopped

    def gmail_account(self, rowCount, item, subject, body, recipients, profile, send_limit) :
        while not(self.getStop_Composer_Process()):
            try:
                self.currentRowNumber = rowCount
                self.body = body
                self.item = item 
                self.subject = subject
                self.recipient = recipients
                self.profile = profile
                self.send_limit = send_limit
                self.runlogger()
                self.logger.info('Start Process')
                time.sleep(1)
                self.insertNewRow.emit(self.currentRowNumber+1)
                now = datetime.now()
                date_time = now.strftime("%m/%d/%Y, %H:%M:%S")
                data = [
                    [self.currentRowNumber, 0, date_time],
                    [self.currentRowNumber, 6, subject],
                ]
                self.appendData.emit(data)

                try:
                    if self.getStop_Composer_Process():
                        break
                    self.setUpDriver(is_hidden=self.hide_browser, launguage_browser=self.launguage_browser, profile=self.profile)
                except Exception as e: 
                    self.logger.error(str(e)) 
                    self.setStop_Composer_Process(stopped=True) 
                    break
                
                try :
                    try :
                        self.logger.info("Open URL ...")
                        self.driver.get('https://mail.google.com/mail/u/0/#inbox?compose=new')
                        time.sleep(1)
                        try:
                            WebDriverWait(self.driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//input[@name="subjectbox"]')))
                        except:
                            self.logger.error('Cannot Open URL')
                            break
                    except Exception as e:
                        self.logger.error(f'{str(e)}')
                        data = [
                            [self.currentRowNumber, 4, "Connection Error"],
                            [self.currentRowNumber, 7, "Stopped"],
                        ]
                        self.appendData.emit(data)
                        break 
                    self.compose()
                except Exception as e :
                    self.logger.error(f"Error : {str(e)}")
                    data = [
                        [self.currentRowNumber, 5, "Error"],
                        [self.currentRowNumber, 7, "Stopped"],
                    ]
                    self.appendData.emit(data)
                    self.logger.info("Ending Process")
            except Exception as e:
                data = [
                    [self.currentRowNumber, 5, "Error"],
                    [self.currentRowNumber, 7, "Stopped"],
                ]
                self.appendData.emit(data)
                self.logger.error(str(e))
            break
        if self.getStop_Composer_Process():
            self.logger.warning("Stopped")
            self.appendData.emit([[self.currentRowNumber, 7, "Stopped"]])
        time.sleep(2)
        self.destroyDriver()
    
    def compose(self):
        passed = 0
        for recipient in self.recipient.copy():
            if passed >= self.send_limit:
                break
            passed += 1
            while not(self.getStop_Composer_Process()):
                try:
                    if self.getStop_Composer_Process():
                        break
                    if len(self.recipient) > 1 and self.recipient[0] != recipient:
                        try:
                            compose_btn = WebDriverWait(self.driver, 30).until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[7]/div[3]/div/div[2]/div[1]/div[1]/div/div')))
                            if self.getStop_Composer_Process():
                                break
                            self.logger.info('Click Compose button')
                            compose_btn.click()
                        except Exception:
                            self.logger.error('Compose Button Not Found')
                            break

                    try:
                        reci = WebDriverWait(self.driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//div/input[@peoplekit-id="BbVjBd"]')))
                        if self.getStop_Composer_Process():
                            break
                        reci.click()
                        reci.send_keys(Keys.CONTROL + 'a')
                        reci.send_keys(Keys.DELETE)
                        self.logger.info('Set Email')
                        reci.send_keys(recipient)
                        time.sleep(1)
                        reci.send_keys(Keys.RETURN)
                        time.sleep(1)
                        ActionChains(self.driver, 0.5).send_keys(Keys.TAB).perform()
                    except Exception:
                        self.logger.error('"Recipient" Field not found')
                        break
                    
                    try:
                        subject = WebDriverWait(self.driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//input[@name="subjectbox"]')))
                        if self.getStop_Composer_Process():
                            break
                        subject.click()
                        subject.send_keys(Keys.CONTROL + 'a')
                        subject.send_keys(Keys.DELETE)
                        self.logger.info('Set Subject')
                        subject.send_keys(self.subject)
                    except Exception:
                        self.logger.error('"Subject" Field not found')
                        break

                    try:
                        body = WebDriverWait(self.driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//div[@role="textbox"]')))
                        if self.getStop_Composer_Process():
                            break
                        body.click()
                        body.send_keys(Keys.CONTROL + 'a')
                        body.send_keys(Keys.DELETE)
                        self.logger.info('Set message body')
                        PutHtml(self.body)
                        body.send_keys(Keys.CONTROL + 'v')
                    except Exception:
                        self.logger.error('"Body" field not found')
                        break

                    try:
                        if self.getStop_Composer_Process():
                            break
                        time.sleep(2)
                        ActionChains(self.driver, 0.5).key_down(Keys.CONTROL).send_keys(Keys.RETURN).key_up(Keys.CONTROL).perform()
                        self.logger.info('Click Send')
                    except Exception:
                        self.logger.error('Cannot click send button')
                        break

                    try:
                        WebDriverWait(self.driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="link_vsm"]')))
                        self.logger.info(f'Message sent to {recipient}')
                    except Exception:
                        self.logger.error(f'Error on sending message to {recipient}')
                        break
                except Exception as e:
                    self.logger.error(str(e))
                break
    
    def runlogger(self):
        profile_name = self.profile.split('\\')[-1]
        if not os.path.exists(LOG_PATH):
            os.makedirs(LOG_PATH)
        self.logger = logging.getLogger("GmailComposer")
        self.logger.setLevel(logging.INFO)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - [%(profile)s] - %(message)s')
        logTextBox = QTextEditLogger(self.logger_window)
        logTextBox.setFormatter(formatter)
        logTextBox.addFilter(LogFilter(profile_name))
        handler = logging.FileHandler(GMAIL_COMPOSER_LOG,'w')
        handler.setLevel(logging.INFO)
        handler.setFormatter(formatter)
        handler.addFilter(LogFilter(profile_name))
        if self.logger.hasHandlers():
            self.logger.handlers.clear()
        self.logger.addHandler(logTextBox)
        self.logger.addHandler(handler)

    def closeLogger(self):
        try:
            handlers = self.logger.handlers[:]
            for handler in handlers:
                handler.close()
                self.logger.removeHandler(handler) 
        except:
            pass
