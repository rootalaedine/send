import re, logging
import ipaddress, time, random
from settings.config import *
from lib.webdriver import WebDriver
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from PyQt6.QtCore import QObject, pyqtSignal

class Gmail_Recovery(WebDriver, QObject):
    """Gmail Recovery Class"""
    insertNewRow = pyqtSignal(int)
    appendData = pyqtSignal(list)

    def __init__(self, kwargs):
        super().__init__()
        QObject.__init__(self)
        self.browser = kwargs.get("browser", None)
        self.hide_browser = kwargs.get("hide_browser", False)
        self.launguage_browser= kwargs.get("launguage_browser", None)
        self.logger_window = kwargs.get("logger_window", None)
        self.tableWidget = kwargs.get('tableWidget',None)
        self.port = 3128
    
    def getStop_Recovery_Process(self):
        return self.Stop_Recovery_Process

    def setStop_Recovery_Process(self, stopped):
        self.Stop_Recovery_Process = stopped

    def recovery_data(self,line) :
        params = line.split(',')
        if len(params) !=3: return False, line
        
        email = params[0].strip()
        password = params[1].strip()
        ip = params[2].strip()
        #...match email
        match_eml = re.search(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", email)
        if not match_eml: return False, f"email -> {email}"
        #...match ip
        try:
            ipaddress.IPv4Address(ip)
        except: return False , f"ip -> ({ip})" 
        return True, { "email":email, "password" :password, "ip":ip}
    

    def gmail_account(self ,rowCount, start_apps, date_from, search_keyword, date_to, profile) :
        while not(self.getStop_Recovery_Process()):
            try:
                self.profile = profile
                self.runlogger()
                self.logger.info("Start Process")
                time.sleep(random.randint(1, 5))
                self.currentRowNumber = rowCount
                self.start_apps = start_apps
                self.date_from = date_from
                self.search_keyword = search_keyword
                self.date_to = date_to
                self.insertNewRow.emit(self.currentRowNumber+1)
                now = datetime.now()
                date_time = now.strftime("%m/%d/%Y, %H:%M:%S")
                data = [
                    [self.currentRowNumber, 0, date_time],
                    [self.currentRowNumber, 6, search_keyword],
                ]
                self.appendData.emit(data)
                try:
                    self.setUpDriver(browser=self.browser, is_hidden=self.hide_browser, launguage_browser=self.launguage_browser, profile=profile)
                except:
                    self.logger.error('Chrome failed to start. If chrome browser already in use, please close it.')
                    break
                try :
                    try :
                        self.logger.info("Open URL ...")
                        self.driver.get('https://mail.google.com/mail/u/0/')
                        time.sleep(random.randint(1, 5))
                    except :
                        self.logger.error(f'Failed to open URL ')
                        data = [
                            [self.currentRowNumber, 4, "Connection Error"],
                            [self.currentRowNumber, 5, "Proxy Connection Error"],
                            [self.currentRowNumber, 7, "Stopped"],
                        ]
                        self.appendData.emit(data)
                        return 
                    WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((By.XPATH, "//a[@href='https://mail.google.com/mail/u/0/#inbox']")))
                    self.checkBoxes(date_from=date_from, search_keyword=search_keyword, date_to=date_to)
                except Exception as e :
                    if self.getStop_Recovery_Process():
                        break
                    self.logger.error(f"Error : {str(e)} ")
                    data = [
                        [self.currentRowNumber, 5, "Error"],
                        [self.currentRowNumber, 7, "Stopped"],
                    ]
                    self.appendData.emit(data)
                    self.logger.info("Ending Process")
            except Exception as e:
                if self.getStop_Recovery_Process():
                    break
                data = [
                    [self.currentRowNumber, 5, "Error"],
                    [self.currentRowNumber, 7, "Stopped"],
                ]
                self.appendData.emit(data)
                self.logger.error(str(e))
            break
        if self.getStop_Recovery_Process():
            self.logger.warning("Stopped")
            self.appendData.emit([[self.currentRowNumber, 7, "Stopped"]])
        self.destroyDriver()
            
    def checkBoxes(self, date_from, search_keyword, date_to):
        if "open_inbox" in self.start_apps:
            self.Recovery_open_box("inbox", date_from, search_keyword, date_to)
                
        if "open_spam" in self.start_apps:
            self.Recovery_open_box("spam", date_from, search_keyword, date_to)

        if "not_spam" in self.start_apps:
            self.Recovery_not_spam(date_from,search_keyword,date_to)
    
    def Recovery_open_box(self, box, date_from,search_keyword,date_to):
        while not(self.getStop_Recovery_Process()):
            time.sleep(random.randint(1, 5))
            self.logger.info(f"Start Open {box.capitalize()} Process")
            try:
                try:
                    if self.getStop_Recovery_Process():
                        break
                    search_field = WebDriverWait(self.driver, 30).until(EC.visibility_of_element_located((By.XPATH, "/html/body/div[7]/div[3]/div/div[1]/div[3]/header/div[2]/div[2]/div[2]/form/div/table/tbody/tr/td/table/tbody/tr/td/div/input[1]")))
                    self.logger.info(f"Click Search Input ({box.capitalize()})")
                    time.sleep(random.randint(1, 5))
                except:
                    self.logger.error(f"Couldn't Find The Search Input ({box.capitalize()})")
                    break
                try:
                    if self.getStop_Recovery_Process():
                        break
                    search_field.click()
                    search_field.send_keys(Keys.CONTROL + "a")
                    search_field.send_keys(Keys.DELETE)
                    time.sleep(random.randint(1, 5))
                    if search_keyword != ".*":
                        search_field.send_keys(f"in:{box} subject:{search_keyword} after:{date_from} before:{date_to}")
                    else:
                        search_field.send_keys(f"in:{box} after:{date_from} before:{date_to}")
                    self.logger.info(f"Set Search Keyword [{search_keyword}] ({box.capitalize()})")
                    time.sleep(random.randint(1, 5))
                    search_field.send_keys(Keys.RETURN)
                    self.logger.info(f"Start Searching ({box.capitalize()})")
                    time.sleep(random.randint(1, 5))
                except:
                    self.logger.error(f"Error To Set Search Keyword  ({box.capitalize()})")
                    break
                try:
                    if self.getStop_Recovery_Process():
                        break
                    nothing_found = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.XPATH, "/html/body/div[7]/div[3]/div/div[2]/div[2]/div/div/div/div/div[2]/div/div[1]/div/div[2]/div[5]/div[2]/table/tbody/tr/td/span")))
                    self.logger.error(f"Nothing Found In {box.capitalize()} ")
                    break
                except:pass
                try:
                    if self.getStop_Recovery_Process():
                        break
                    result = WebDriverWait(self.driver, 30).until(EC.visibility_of_element_located((By.XPATH, "/html/body/div[7]/div[3]/div/div[2]/div[2]/div/div/div/div/div[2]/div/div[1]/div/div[2]/div[5]/div[1]/div/table/tbody/tr")))
                    number_email = len(self.driver.find_elements(By.XPATH, "/html/body/div[7]/div[3]/div/div[2]/div[2]/div/div/div/div/div[2]/div/div[1]/div/div[2]/div[5]/div[1]/div/table/tbody/tr"))
                    self.logger.info(f"Start Open Emails ({box.capitalize()})")
                    i = 1
                    while i <= number_email:
                        if self.getStop_Recovery_Process():
                            break
                        ActionChains(self.driver).send_keys(Keys.ARROW_DOWN).perform()
                        msg = self.driver.find_element(By.XPATH, f"/html/body/div[7]/div[3]/div/div[2]/div[2]/div/div/div/div/div[2]/div/div[1]/div/div[2]/div[5]/div[1]/div/table/tbody/tr[{str(i)}]/td[4]")
                        msg.click()
                        time.sleep(random.randint(1, 5))
                        found = True
                        while found:
                            try:
                                if self.getStop_Recovery_Process():
                                    break
                                ActionChains(self.driver, 0.5).key_down(Keys.SHIFT).send_keys(Keys.TAB).key_up(Keys.SHIFT).perform()
                                back_button = WebDriverWait(self.driver, 0.2).until(EC.visibility_of_element_located((By.XPATH, "/html/body/div[7]/div[3]/div/div[2]/div[2]/div/div/div/div/div[1]/div[3]/div[1]/div/div[1]/div")))
                                back_button.click()
                                found = False
                            except:pass
                        i+=1
                        time.sleep(random.randint(1, 5))
                    if self.getStop_Recovery_Process():
                        break
                    self.logger.info(f"Open Emails Succeeded ({box.capitalize()})")
                except Exception as e:
                    self.logger.info(str(e))
                    self.logger.error(f"Error To Open {box.capitalize()} Emails ")
                    break
                break
            except Exception as e :
                data = [
                    [self.currentRowNumber, 5, "Error"],
                    [self.currentRowNumber, 7, "Stopped"],
                ]
                self.appendData.emit(data)
                self.logger.error(f"Error : {str(e)} ")
                self.logger.info("Ending Process")
            break

    def Recovery_not_spam(self,date_from,search_keyword,date_to):
        while not(self.getStop_Recovery_Process()):
            time.sleep(random.randint(1, 5))
            self.logger.info(f"Start Not Spam Process")
            try:
                try:
                    if self.getStop_Recovery_Process():
                        break
                    search_field = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.XPATH, "/html/body/div[7]/div[3]/div/div[1]/div[3]/header/div[2]/div[2]/div[2]/form/div/table/tbody/tr/td/table/tbody/tr/td/div/input[1]")))
                    time.sleep(random.randint(1, 5))
                    self.logger.info("Click Search Input")
                except:
                    self.logger.error("Couldn't Find The Search Input")
                    break
                try:
                    if self.getStop_Recovery_Process():
                        break
                    search_field.click()
                    search_field.send_keys(Keys.CONTROL + "a")
                    search_field.send_keys(Keys.DELETE)
                    time.sleep(random.randint(1, 5))
                    if search_keyword != ".*":
                        search_field.send_keys(f"in:spam subject:{search_keyword} after:{date_from} before:{date_to}")
                    else:
                        search_field.send_keys(f"in:spam after:{date_from} before:{date_to}")
                    self.logger.info(f"Set Search Keyword [{search_keyword}]")
                    time.sleep(random.randint(1, 5))
                    search_field.send_keys(Keys.RETURN)
                    self.logger.info("Start Searching")
                    time.sleep(random.randint(1, 5))
                except:
                    self.logger.error(f"Error To Set Search Keyword ")
                    break
                try:
                    if self.getStop_Recovery_Process():
                        break
                    nothing_found = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.XPATH, "/html/body/div[7]/div[3]/div/div[2]/div[2]/div/div/div/div/div[2]/div/div[1]/div/div[2]/div[5]/div[2]/table/tbody/tr/td/span")))
                    self.logger.error(f"Nothing Found In Spam ")
                    break
                except:pass
                try:
                    if self.getStop_Recovery_Process():
                        break
                    result = WebDriverWait(self.driver, 30).until(EC.visibility_of_element_located((By.XPATH, "/html/body/div[7]/div[3]/div/div[2]/div[2]/div/div/div/div/div[2]/div/div[1]/div/div[2]/div[5]/div[1]/div/table/tbody/tr")))
                    checkbox = WebDriverWait(self.driver, 30).until(EC.visibility_of_element_located((By.XPATH, "/html/body/div[7]/div[3]/div/div[2]/div[2]/div/div/div/div/div[1]/div[2]/div[2]/div[1]/div/div/div[1]/div/div[1]/span")))
                    self.logger.info("Select All Emails")
                    checkbox.click()
                    time.sleep(random.randint(1, 5))
                except:
                    self.logger.error(f"Error To Select All Emails ")
                    break
                try:
                    if self.getStop_Recovery_Process():
                        break
                    not_spam_button = WebDriverWait(self.driver, 30).until(EC.visibility_of_element_located((By.XPATH, "/html/body/div[7]/div[3]/div/div[2]/div[2]/div/div/div/div/div[1]/div[2]/div[2]/div[1]/div/div/div[4]/div[1]")))
                    self.logger.info("Moving Emails From Spam To Inbox")
                    not_spam_button.click()
                    _ = WebDriverWait(self.driver, 30).until(EC.visibility_of_element_located((By.XPATH, "/html/body/div[7]/div[3]/div/div[2]/div[2]/div/div/div/div/div[2]/div/div[1]/div/div[2]/div[5]/div[2]/table/tbody/tr/td/span")))
                    self.logger.info("Move Emails From Spam To Inbox Succeeded")
                except:
                    self.logger.error(f"Error To Move Emails ")
                    break
                break
            except Exception as e :
                self.logger.error(f"Error : {str(e)} ")
                self.logger.info("Ending Process")
                data = [
                    [self.currentRowNumber, 5, "Error"],
                    [self.currentRowNumber, 7, "Stopped"],
                ]
                self.appendData.emit(data)
            break
    
    def runlogger(self):
        profile_name = self.profile.split('/')[-1]
        if not os.path.exists(LOG_PATH):
            os.makedirs(LOG_PATH)
        self.logger = logging.getLogger("GmailApps")
        self.logger.setLevel(logging.INFO)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - [%(profile)s] - %(message)s')
        logTextBox = QTextEditLogger(self.logger_window)
        logTextBox.setFormatter(formatter)
        logTextBox.addFilter(LogFilter(profile_name))
        handler = logging.FileHandler(GMAIL_RECOVERY_LOG,'w')
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

    def __del__(self):
        # close open logger
        self.closeLogger()