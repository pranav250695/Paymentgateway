import Xlutils
from selenium import webdriver
import logging
import datetime
import openpyxl
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from Config.config import TESTCLASS

driver = webdriver.Chrome(executable_path=TESTCLASS.CHROME_PATH)
driver.get('https://sandbox.isgpay.com/KotakPGRedirect/')
driver.maximize_window()
row = Xlutils.GetRowCount(TESTCLASS.DATA_PATH,"Sheet1")
logging.basicConfig(filename="C:\\Users\\pranav4013\\PycharmProjects\\pythonProject1\\PG\\PG LOG\\TEST_LOG.log",
                    format='%(asctime)s: %(levelname)s: %(message)s:',
                    datefmt='%m/%d/%Y %I:%M:%S %p'
                    )
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

#loop will read data from Excel file

for r in range(2,row+1):
    Amount = Xlutils.ReadData(TESTCLASS.DATA_PATH,"Sheet1",r,1)
    TXN_ID = Xlutils.ReadData(TESTCLASS.DATA_PATH, "Sheet1", r, 2)
    Merchant_id = Xlutils.ReadData(TESTCLASS.DATA_PATH, "Sheet1", r, 3)
    Terminal_id = Xlutils.ReadData(TESTCLASS.DATA_PATH, "Sheet1", r, 4)
    MCC = Xlutils.ReadData(TESTCLASS.DATA_PATH, "Sheet1", r, 5)
    CARD_NUM = Xlutils.ReadData(TESTCLASS.DATA_PATH,"Sheet1", r, 6)
    CARD_Exipry = Xlutils.ReadData(TESTCLASS.DATA_PATH, "Sheet1", r, 7)
    CVV = Xlutils.ReadData(TESTCLASS.DATA_PATH, "Sheet1", r, 8)
    OTP = Xlutils.ReadData(TESTCLASS.DATA_PATH, "Sheet1", r, 9)
    #first Page input fields
    print('---First Page Start---')
    logger.debug("First Page Start")
    driver.find_element_by_id('Amount').clear()
    driver.find_element_by_id('Amount').send_keys(Amount)
    driver.implicitly_wait(1000)
    driver.find_element_by_id('TxnType').clear()
    driver.find_element_by_id('TxnType').send_keys(TXN_ID)
    driver.implicitly_wait(1000)
    driver.find_element_by_id('MerchantId').clear()
    driver.find_element_by_id('MerchantId').send_keys(Merchant_id)
    driver.implicitly_wait(1000)
    driver.find_element_by_id('TerminalId').clear()
    driver.find_element_by_id('TerminalId').send_keys(Terminal_id)
    driver.implicitly_wait(1000)
    driver.find_element_by_id('MCC').clear()
    driver.find_element_by_id('MCC').send_keys(MCC)
    driver.implicitly_wait(1000)
    driver.find_element_by_id('inprocess').click()
    print('---End of First Page Start---')
    logger.info("End of First Page Start")
    # second Page input fields
    print('---Second Page Start---')
    logger.info("Second Page Start")
    driver.find_element_by_name('submit').click()
    print('---End second Page Start---')
    logger.info("End second Page Start")
    # Third Page input fields
    print('---Third Page Start---')
    logger.info("Third Page Start")
    driver.find_element_by_id('cardNum').send_keys(CARD_NUM)
    driver.find_element_by_id('cardExpiry').send_keys(CARD_Exipry)
    driver.find_element_by_id('cardSecurityCode').send_keys(CVV)
    driver.implicitly_wait(1000)
    driver.find_element_by_xpath("//div[@class='col-xs-6']/button").click()
    print('---End Third Page Start---')
    logger.info("End Third Page Start")
    # Fourth Page input fields
    print('---Fourth Page Start---')
    logger.info("Fourth Page Start")
    driver.find_element_by_name('otp').send_keys(OTP)
    driver.find_element_by_xpath("//div[@class='row']/button").click()
    Merchant_txn_number = driver.find_element_by_xpath('/html/body/table[2]/tbody/tr[2]/td[2]').text
    Merchant_id = driver.find_element_by_xpath('/html/body/table[2]/tbody/tr[4]/td[2]').text
    Terminal_id = driver.find_element_by_xpath('/html/body/table[2]/tbody/tr[5]/td[2]').text
    Amount = driver.find_element_by_xpath('/html/body/table[2]/tbody/tr[6]/td[2]').text
    Response_code = driver.find_element_by_xpath('/html/body/table[2]/tbody/tr[9]/td[2]').text
    Message = driver.find_element_by_xpath('/html/body/table[2]/tbody/tr[10]/td[2]').text
    RRN = driver.find_element_by_xpath('/html/body/table[2]/tbody/tr[11]/td[2]').text
    Authcode = driver.find_element_by_xpath('/html/body/table[2]/tbody/tr[13]/td[2]').text
    Cardnum = driver.find_element_by_xpath('/html/body/table[2]/tbody/tr[14]/td[2]').text


    # fifth Page result saved in SCREENSHOT FOLDER and LOGS generated in PGLOG folder
    # Message = driver.find_element_by_xpath("/html/body/table[2]/tbody/tr[10]/td[2]").text
    timestamp = datetime.datetime.now().strftime('%d_%m_%Y_%H_%M_%S')
    TESTCLASS.SCREENSHOT = driver.get_screenshot_as_file('SCREENSHOT\\' 'PG' + timestamp + '.png')
    print('---End Third Page Start---')
    logger.info("End Third Page Start")
    print(Message)
    if Message == "Success":
        print("transaction successfull")
        Xlutils.WriteData(TESTCLASS.DATA_PATH,'Sheet1',r,10,"Test case passed")
        driver.get("https://sandbox.isgpay.com/KotakPGRedirect/")
        logger.info("Transaction sccessfull")
    else:
        print("transaction failed")
        Xlutils.WriteData(TESTCLASS.DATA_PATH, 'Sheet1', r, 10, "Test case failed")
        logger.error("Transaction Failed")

    Xlutils.WriteData(TESTCLASS.TXN_DATA_FILE,"Sheet",r,1,Merchant_txn_number)
    Xlutils.WriteData(TESTCLASS.TXN_DATA_FILE, "Sheet", r, 2, Merchant_id)
    Xlutils.WriteData(TESTCLASS.TXN_DATA_FILE, "Sheet", r, 3, Terminal_id)
    Xlutils.WriteData(TESTCLASS.TXN_DATA_FILE, "Sheet", r, 4, Amount)
    Xlutils.WriteData(TESTCLASS.TXN_DATA_FILE, "Sheet", r, 5, Response_code)
    Xlutils.WriteData(TESTCLASS.TXN_DATA_FILE, "Sheet", r, 6, Message)
    Xlutils.WriteData(TESTCLASS.TXN_DATA_FILE, "Sheet", r, 7, RRN)
    Xlutils.WriteData(TESTCLASS.TXN_DATA_FILE, "Sheet", r, 8, Authcode)
    Xlutils.WriteData(TESTCLASS.TXN_DATA_FILE, "Sheet", r, 9, Cardnum)

