import Xlutils
from selenium import webdriver
import datetime
import logging
from Config.config import TESTCLASS


driver = webdriver.Chrome(executable_path=TESTCLASS.CHROME_PATH)
driver.get('https://sandbox.isgpay.com/KotakPGRedirect/')
driver.maximize_window()
logging.basicConfig(filename="C:\\Users\\pranav4013\\PycharmProjects\\pythonProject1\\PG\\PG LOG\\TEST_LOG.log",
                    format='%(asctime)s: %(levelname)s: %(message)s:',
                    datefmt='%m/%d/%Y %I:%M:%S %p'
                    )
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)
driver.find_element_by_link_text("Status Query").click()

row = Xlutils.GetRowCount(TESTCLASS.DATA_PATH,"Sheet2")
print(row)
for r in range(2,row+1):
    M_TXN_NO = Xlutils.ReadData(TESTCLASS.DATA_PATH,"Sheet2",r,1)
    MERCH_ID = Xlutils.ReadData(TESTCLASS.DATA_PATH, "Sheet2", r, 2)
    PASSCODE = Xlutils.ReadData(TESTCLASS.DATA_PATH, "Sheet2", r, 3)
    TERMINAL_ID = Xlutils.ReadData(TESTCLASS.DATA_PATH, "Sheet2", r, 4)
    TXN_CODE = Xlutils.ReadData(TESTCLASS.DATA_PATH, "Sheet2", r, 5)
    driver.find_element_by_name('TxnRefNo').send_keys(M_TXN_NO)
    driver.find_element_by_name('MerchantId').send_keys(MERCH_ID)
    driver.find_element_by_name('PassCode').send_keys(PASSCODE)
    driver.find_element_by_name('TerminalId').send_keys(TERMINAL_ID)
    driver.find_element_by_name('TxnType').send_keys(TXN_CODE)
    driver.find_element_by_name('SubButL').click()
    timestamp = datetime.datetime.now().strftime('%d_%m_%Y_%H_%M_%S')
    ss = driver.get_screenshot_as_file('SCREENSHOT\\' 'PG' + timestamp + '.png')
    logger.info("status query done")

    Response_code= driver.find_element_by_xpath('/html/body/table/tbody/tr[7]/td[2]').text
    print(Response_code)
    if Response_code == 'Success':
        print("Status Query success")
        Xlutils.WriteData(TESTCLASS.DATA_PATH,"Sheet2",r,6,"Status Query Paseed")
        driver.get('https://sandbox.isgpay.com/KotakPGRedirect/merchant/rac/statusQueryAPI.jsp')
        logger.info("Status Query success")
    else:
        print("Status Query Failed")
        Xlutils.WriteData(TESTCLASS.DATA_PATH, "Sheet2", r, 6,"Status Query Failed")
        driver.get('https://sandbox.isgpay.com/KotakPGRedirect/merchant/rac/statusQueryAPI.jsp')
        logger.info("Status Query Failed")



