import Xlutils
from selenium import webdriver
import logging
import datetime
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
driver.find_element_by_link_text('Capture/Refund').click()

row = Xlutils.GetRowCount(TESTCLASS.DATA_PATH,"Sheet5")

for r in range(2,row+1):
    MER_TXN_NO= Xlutils.ReadData(TESTCLASS.DATA_PATH,"Sheet5",r,1)
    AMOUNT=Xlutils.ReadData(TESTCLASS.DATA_PATH,"Sheet5",r,2)
    PASSCODE=Xlutils.ReadData(TESTCLASS.DATA_PATH,"Sheet5",r,3)
    TERMINAL_ID=Xlutils.ReadData(TESTCLASS.DATA_PATH,"Sheet5",r,4)
    MERCHANT_ID=Xlutils.ReadData(TESTCLASS.DATA_PATH,"Sheet5",r,5)
    TXN_TYPE=Xlutils.ReadData(TESTCLASS.DATA_PATH,"Sheet5",r,6)
    REF_TXN_NO=Xlutils.ReadData(TESTCLASS.DATA_PATH,"Sheet5",r,7)
    REFUND_AMOUNT=Xlutils.ReadData(TESTCLASS.DATA_PATH,"Sheet5",r,8)
    CAPTURE_AMOUNT=Xlutils.ReadData(TESTCLASS.DATA_PATH,"Sheet5",r,9)
    CURRENCY=Xlutils.ReadData(TESTCLASS.DATA_PATH,"Sheet5",r,10)
    CANCELLATION_ID=Xlutils.ReadData(TESTCLASS.DATA_PATH,"Sheet5",r,11)
    REFUND_REASON_CODE=Xlutils.ReadData(TESTCLASS.DATA_PATH,"Sheet5",r,12)

    driver.find_element_by_name('TxnRefNo').send_keys(MER_TXN_NO)
    driver.find_element_by_name('Amount').send_keys(AMOUNT)
    driver.find_element_by_name('PassCode').send_keys(PASSCODE)
    driver.find_element_by_name('TerminalId').send_keys(TERMINAL_ID)
    driver.find_element_by_name('MerchantId').send_keys(MERCHANT_ID)
    driver.find_element_by_name('TxnType').send_keys(TXN_TYPE)
    driver.find_element_by_name('RetRefNo').send_keys(REF_TXN_NO)
    driver.find_element_by_name('RefundAmount').send_keys(REFUND_AMOUNT)
    driver.find_element_by_name('CaptureAmount').send_keys(CAPTURE_AMOUNT)
    driver.find_element_by_name('Currency').send_keys(CURRENCY)
    driver.find_element_by_name('RefCancelID').send_keys(CANCELLATION_ID)
    driver.find_element_by_name('RefReasonCode').send_keys(REFUND_REASON_CODE)
    driver.find_element_by_name('SubButL').click()
    timestamp = datetime.datetime.now().strftime('%d_%m_%Y_%H_%M_%S')
    TESTCLASS.SCREENSHOT = driver.get_screenshot_as_file('SCREENSHOT\\' 'PG' + timestamp + '.png')
    message = driver.find_element_by_xpath('/html/body/table/tbody/tr[12]/td[2]').text
    print(message)
    logger.info("capture done")

    if message == 'Success':
        print('capture successfully')
        Xlutils.WriteData(TESTCLASS.DATA_PATH,"Sheet5",r,13,"test pass")
        driver.get('https://sandbox.isgpay.com/KotakPGRedirect/merchant/rac/racAPI.jsp')
        logger.info("capture done")
    else:
        print('capture failed')
        Xlutils.WriteData(TESTCLASS.DATA_PATH, "Sheet5", r, 13, "test fail")
        driver.get('https://sandbox.isgpay.com/KotakPGRedirect/merchant/rac/racAPI.jsp')
        logger.info("capture failed")

