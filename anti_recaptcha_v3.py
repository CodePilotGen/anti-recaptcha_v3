import os
import xlrd
import requests
from ibm_watson import SpeechToTextV1
from ibm_cloud_sdk_core.authenticators import IAMAuthenticator

import cv2
import numpy as np
import pyautogui
import time
import win32clipboard
import pyperclip

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

apikey = '6yqI4Ag-oUqMW_GZ02KEgwnWCPXqpUBO3Kt6ZoZhHBa5'
url = 'https://api.us-south.speech-to-text.watson.cloud.ibm.com/instances/cbe70cfa-8d0a-42fb-a31a-703be7037637'
delayTime = 2
audioToTextDelay = 10
filename = '1.mp3'

byPassUrl = 'http://scw.pjn.gov.ar/scw/home.seam'
googleIBMLink = 'https://speech-to-text-demo.ng.bluemix.net/'

pyautogui.FAILSAFE = False

def audioToText(mp3Path):
    authenticator = IAMAuthenticator(apikey)
    stt = SpeechToTextV1(authenticator=authenticator)
    stt.set_service_url(url)

    texts = []
    with open(mp3Path, 'rb') as f:
        try:
            res = stt.recognize(audio=f, content_type='audio/mp3', model='es-ES_BroadbandModel', continuous=True).get_result()
            print(res)
            alternatives = res['results'][0]['alternatives']
            for alternative in alternatives:
                transcript = alternative['transcript']
                texts.append(transcript)
        except:
            texts = []
    return texts

def saveFile(content,filename):
    with open(filename, "wb") as handle:
        for data in content.iter_content():
            handle.write(data)

def clickItem(subimage, rl='left'):
    coords = []

    pic = pyautogui.screenshot()    #taking screenshot
    pic.save('main.jpg')            #saving screenshot

    img_rgb = cv2.imread('main.jpg')
    img_gray = cv2.cvtColor(img_rgb, cv2.COLOR_BGR2GRAY)               ###beginning template matching###
    template = cv2.imread(subimage,0)                                  #
    # w, h = template.shape[::-1]                                         #

    res = cv2.matchTemplate(img_gray,template,cv2.TM_CCOEFF_NORMED)
    threshold = 0.8
    loc = np.where( res >= threshold)

    for pt in zip(*loc[::-1]):
        coords.append(pt[0])        #storing upper left coordinates of reCAPTCHA box
        coords.append(pt[1])


    if len(coords)!= 0 :
        reCAPTCHA_box_x_bias = 25
        reCAPTCHA_box_y_bias = 20

        coords[0] = coords[0] + reCAPTCHA_box_x_bias
        coords[1] = coords[1] + reCAPTCHA_box_y_bias

        pyautogui.moveTo(1, 1, duration = 0)
        pyautogui.moveTo(coords[0], coords[1], duration = 0.12)

        if(rl == 'left'):
            pyautogui.click()
        else:
            pyautogui.rightClick()
        return True

    else :
        pyautogui.moveTo(1, 1, duration = 0)
        print ('-> reCAPTCHA box not found!')
        return False

def lifecycle(jurisdiction, number, year):
    time.sleep(1)
    # pyautogui.click('subimages/submit_button.png')
    if(clickItem('subimages/ctcha_check.jpg') == False):
        # driver.refresh()
        return False
    time.sleep(1)
    if(clickItem('subimages/reshiba.jpg') == False):
        # driver.refresh()
        return False
    time.sleep(1)
    if(clickItem('subimages/download_btn.jpg', rl='right') == False):
        # driver.refresh()
        return False
    time.sleep(1)
    if(clickItem('subimages/copy_link_as.jpg') == False):
        # driver.refresh()
        return False
    time.sleep(1)
    win32clipboard.OpenClipboard()
    href = win32clipboard.GetClipboardData()
    win32clipboard.CloseClipboard()
    # print(data)
    if(href == ""):
        # driver.refresh()
        return False
    response = requests.get(href, stream=True)
    saveFile(response,filename)

    texts = audioToText(os.getcwd() + '/' + filename)

    if(len(texts) == 0):
        # driver.refresh()
        return False
    # if(clickItem('subimages/input_box.jpg') == False):
    #     pyautogui.hotkey('f5')
    #     return
    for i in range(len(texts)):
        clickItem('subimages/input_box.jpg')
        pyperclip.copy(texts[i])
        time.sleep(0.2)
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        time.sleep(1)
        iframe = driver.find_element_by_css_selector("iframe[title='Prueba reCAPTCHA']")
        driver.switch_to.frame(iframe)
        errorMsg = driver.find_elements_by_xpath('/html/body/div/div/div[1]')[0]
        # errorMsg = driver.find_element_by_xpath("//div[class='rc-audiochallenge-error-message']")[0]
        # print(errorMsg.text)
        if errorMsg.text == "Debes resolver más captchas.":
            print("wrong recognition => ", texts[i])
            if i == len(texts) - 1:
                return False
            continue
        print("exact recognition => ", texts[i])
        break

    driver.switch_to.default_content()
    sel = Select(driver.find_element_by_id('formPublica\\:camaraNumAni'))
    sel.select_by_visible_text(jurisdiction)
    driver.find_element_by_id("formPublica\\:numero").send_keys(number)
    driver.find_element_by_id("formPublica\\:anio").send_keys(year)
    driver.find_element_by_id("formPublica\\:buscarPorNumeroButton").click()
    try:
        WebDriverWait(driver, 3).until(EC.url_contains('http://scw.pjn.gov.ar/scw/expediente.seam'))
    except TimeoutException:
        # driver.refresh()
        return False
    driver.back()
    # driver.refresh()
    return True

def isBrowserClosed(driver):
    isClosed = False
    try:
        driver.get_window_size()
    except:
        isClosed = True

    return isClosed

if __name__ == "__main__":
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument('--disable-notifications')
    chrome_options.add_argument('--ignore-ssl-errors=yes')
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--incognito')
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')
    chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.164 Safari/537.36")
    
    driver = webdriver.Chrome('chromedriver.exe', options=chrome_options)
    driver.get('http://scw.pjn.gov.ar/scw/home.seam')
    
    loc_file = 'DatabaseQuery_14072021_140538239.xls'
    # loc_file = 'DatabaseQuery_test.xls'
    wb = xlrd.open_workbook(loc_file)
    ws = wb.sheet_by_index(0)
    print("{0} {1} {2}".format(ws.name, ws.nrows, ws.ncols))
    for rx in range(1, ws.nrows):
        flag = False
        jurisdiction = ws.cell(rx, 1).value
        number, year = ws.cell(rx, 2).value.split('/')
        print(rx, " row=>")
        print("{0} {1} {2}".format(jurisdiction, number, year))

        if (isBrowserClosed(driver)):
            print("Browser is closed !")
            break

        while flag == False:
            try:
                flag = lifecycle(jurisdiction, number, year)
                if (isBrowserClosed(driver)):
                    break
                driver.refresh()
            except:
                driver.refresh()
                flag = False

    driver.close()
    driver.quit()
    print("Successfully Finished !!!")