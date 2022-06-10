import os
import globals as Global
import pickle
from selenium import webdriver
from PyQt5 import QtTest
from urllib.request import urlretrieve
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class CImageInfo:
    Name= None
    Url = None

    def __init__ (self, name, url):
        self.Name= name
        self.Url = url

    def IsValid (self, file_list):
        name = self.Name + ".jpg"
        for v in file_list:
            if v == name:
                return True
        return False
    
    def LoadImg (self):
        filename = self.Name + '.jpg'
        try:
            urlretrieve (self.Url, os.path.join(Global.DF_EXPORT_IMG_PATH, filename))
            return True
        except:
            return False

class CShopeeCollector ():
    def __init__ (self, _super, _chrome):
        self.super  = _super
        self.chrome = _chrome
        self.m_AccountID = ""

        self.bEnterSellerCenter = False
        self.bEnterItemList     = False

        self.AllItemNum      = int ()
        self.AllImgInfo      = dict ()
        self.MissingItemList = dict ()
    
    def GetElm (self, ByType, _str, _limitTime):
        if self.chrome == None:return None
        locator = (ByType, _str)
        try:
            elm = WebDriverWait (self.chrome, _limitTime).until (
                EC.presence_of_element_located (locator),
                "找不到指定元素 :" + _str
            )
            return elm
        except:return None

    def GetElms (self, ByType, _str, _limitTime):
        if self.chrome == None:return None
        locator = (ByType, _str)
        try:
            elm = WebDriverWait (self.chrome, _limitTime).until (
                EC.presence_of_element_located (locator),
                "找不到指定元素 :" + _str
            )
            # fine_elements 會尋找字串中含有特定字串的所有elm
            # Ex：搜尋的字串是'abc'，含有'abc'或'xxxabcxxxx'的elm都會被搜尋到
            return self.chrome.find_elements (ByType, _str)
        except:return list ()

    def LoopGetElm (self, ByType, _str, _times = 3, _limitTime = Global.DF_GET_ITEM_ELM_TIME):
        times = 1
        elm = self.GetElm (ByType, _str, _limitTime)
        while (elm == None):
            if times >= _times:break
            elm = self.GetElm (ByType, _str, _limitTime)
            times += 1
        return elm

    def OnClickElmSafely (self, ByType, _str, _strElmName = None):
        elm = self.LoopGetElm (ByType, _str)
        if elm != None:
            try:
                elm.click ()  
                return True
            except:
                if _strElmName != None and type (_strElmName) == str:
                    self.super.PrintLog ("The " + _strElmName + " is not clickable.")
        else:
            if _strElmName != None and type (_strElmName) == str:
                self.super.PrintLog ("It cannot find out " + _strElmName + ".")
        return False


    def WaitForPage (self, _url):
        if self.chrome == None:return False
        try:
            WebDriverWait (self.chrome, float (Global.get_value ("def_waitting_for_page_loaded_secs"))).until (
                EC.url_to_be (_url),
                "找不到指定網址 :" + _url
            )
            return True
        except:return False

    def WaitForPageLoaded (self, _url, _loopTimes = 3):
        times = 1
        bLoad = self.WaitForPage (_url)
        while (bLoad == False):
            if times >= _loopTimes:
                break
            bLoad = self.WaitForPage (_url)
            times += 1
        if not bLoad:
            self.super.PrintLog ("Failed to enter the page:%s" % (_url))
        return bLoad
        
    def OpenWebdriver (self):
        driver_path = "./chromedriver"
        options = webdriver.ChromeOptions ()
        options.add_experimental_option ("useAutomationExtension", False)
        options.add_experimental_option ("excludeSwitches", ["enable-automation"])
        # 禁止彈出視窗
        prefs = {'profile.default_content_setting_values':{'notifications':2 }}
        options.add_experimental_option ('prefs', prefs)

        self.chrome = webdriver.Chrome (executable_path=driver_path, chrome_options=options)
        self.chrome.maximize_window ()
        self.GotoLoginPage ()
        QtTest.QTest.qWait (float (Global.get_value ("def_waitting_for_begining_loading_secs")) * 1000) #因為開啟新網頁時還沒仔入完整原始碼，當下會找到不到指定elm
    
    def GotoLoginPage (self):
        self.chrome.get (Global.get_value ("def_login_url"))
        # 載入成功登入時留下的cookie
        if os.path.exists ("./cookies.pkl"):
            cookies = pickle.load(open("cookies.pkl", "rb"))
            for cookie in cookies:
                if cookie['domain'] == ".shopee.tw":
                    self.chrome.add_cookie(cookie)
        self.chrome.get (Global.get_value ("def_login_url")) # 載入完cookie後要重新載入
        

    def GotoShopeeItemList (self):
        # 會以開啟新分頁的方式進入賣家後台，故切換到新分頁
        self.chrome.switch_to.window (self.chrome.window_handles[0])

        if not self.OnClickElmSafely (By.CSS_SELECTOR, "#product > ul > li:nth-child(1) > a", "Elm_AllItemList"):
            return False

        bLoad = self.WaitForPageLoaded (Global.get_value ("def_my_item_list_url"))
        if bLoad:
            # 成功進入"我的商品"後，儲存該次登入cookie
            pickle.dump(self.chrome.get_cookies(), open("cookies.pkl","wb")) # 儲存登入成功的cookies

            # 存下登入成功的帳戶ID
            elm_account = self.LoopGetElm (By.CLASS_NAME, "account-name", "Elm_Account")
            if elm_account != None:
                self.m_AccountID = elm_account.text

            return True
        else:return False

    def HandleItemImagesCollection (self):
        # 商品圖片蒐集
        # 起始頁籤顯示格式: 1 2 3 4 5 ... N (頁籤若不在第一頁格式會變動)
        # 抓取最後一頁頁碼

        # 進入"我的商品"時，可能會遇到蝦皮自家的彈出視窗，故自動關閉
        # self.OnClickElmSafely (By.CSS_SELECTOR, "#\$productlistTour > div > div > div > div.shopee-popper.shopee-popover__popper.shopee-popover__popper--light.with-arrow > div > div > div.tip-footer-wrapper > div > button")
        self.OnClickElmSafely (By.CLASS_NAME, "next-button.shopee-button.shopee-button--primary.shopee-button--normal")

        elm_group_pages = self.GetElms (By.CLASS_NAME, "shopee-pager__page", 10)
        if len (elm_group_pages) == 0:
            return False
        
        nEndPageIndex = len (elm_group_pages) - 1
        nEndPage = 0

        nCheckCnt = 0
        while True:
            try:
                nEndPage = int (elm_group_pages[nEndPageIndex].text)
                break
            except:
                nCheckCnt += 1
                if nCheckCnt >= 5:
                    return False
                QtTest.QTest.qWait (1000)

        try:
            elm_group_pages[nEndPageIndex].click ()
        except:
            self.super.PrintLog ("The Elm_EndPage is not clickable.")
            return False

        elm_item_group = self.GetElms (By.CLASS_NAME, "product-image.product-image-box", 10)
        itemNumInEndPage = len (elm_item_group)
        if itemNumInEndPage > 0:
            # 所有商品數量:每頁24個 * (頁數-1) + 最後一頁數量
            self.AllItemNum = 24 * (nEndPage - 1) + itemNumInEndPage
        else:return False

        # 回到第一頁
        elm_group_pages = self.GetElms (By.CLASS_NAME, "shopee-pager__page", 10) # 切換頁面後需要重新讀取
        try:
            elm_group_pages[0].click ()
        except:
            self.super.PrintLog ("The Elm_FirstPage is not clickable.")
            return False

        elm_item_group = list ()
        elm_item_group = self.GetElms (By.CLASS_NAME, "product-image.product-image-box", 10)
        if len (elm_item_group) == 0:
            return False

        for page in range (1, nEndPage + 1):
            # 檢查當前頁面是否正確
            elm_curr_page = self.LoopGetElm (By.CLASS_NAME, "shopee-pager__page.active")
            if elm_curr_page == None:
                return False
            while (int (elm_curr_page.text) != page):
                if self.OnClickElmSafely (By.CLASS_NAME, "shopee-button.shopee-button--small.shopee-button--frameless.shopee-button--block.shopee-pager__button-next", "Elm_PageNext"):
                    QtTest.QTest.qWait (2000)
                    elm_curr_page = self.LoopGetElm (By.CLASS_NAME, "shopee-pager__page.active")
                else:
                    return False

            error_num = 0
            elm_item_group = self.GetElms (By.CLASS_NAME, "product-image.product-image-box", 10)
            for elm in elm_item_group:
                elm_item = elm.find_element_by_tag_name ("img")
                if elm_item != None:
                    img_url = elm_item.get_attribute ('src')
                    item_name = elm_item.get_attribute ('alt')
                    img = CImageInfo (item_name, img_url)

                    count = 1
                    bLoad = img.LoadImg ()
                    while (bLoad == False):
                        error_num += 1
                        if (count >= 3): break
                        bLoad = img.LoadImg ()
                        count += 1

                    if bLoad:
                        if error_num != 0:error_num = 0
                    else:
                        self.MissingItemList[img.Name] = img
                        # 等待防爬蟲機制結束
                        self.super.PrintLog ("Waitting for crawler blocker.")
                        QtTest.QTest.qWait (float (Global.get_value ("def_waitting_for_shopee_crawler_blocker_secs")) * 1000)
                    self.AllImgInfo[img.Name] = img
        self.super.PrintLog ("All items images collection done.")

        # 商品圖片蒐集檢查
        file_list = os.listdir (Global.DF_EXPORT_IMG_PATH)
        for img_name in self.AllImgInfo:
            img = self.AllImgInfo[img_name]
            if not img.IsValid (file_list):
                self.MissingItemList[img.Name] = img
                
        checkCnt = 0
        while (len (self.MissingItemList) > 0):
            # 重新下載圖片
            error_num = 0
            file_list = os.listdir (Global.DF_EXPORT_IMG_PATH)
            for name in self.MissingItemList:
                img = self.MissingItemList[name]
                if not img.IsValid (file_list):
                    if img.LoadImg () == False:
                        self.super.PrintLog (img.Url)
                        error_num += 1
                        if error_num > 3:
                            # 等待防爬蟲機制結束
                            self.super.PrintLog ("Waitting for crawler blocker.")
                            QtTest.QTest.qWait (float (Global.get_value ("def_waitting_for_shopee_crawler_blocker_secs")) * 1000)
                    else:
                        if error_num != 0:error_num = 0
            
            # 更新缺少圖片資料
            valid_list = list ()
            file_list = os.listdir (Global.DF_EXPORT_IMG_PATH)
            for name in self.MissingItemList:
                img = self.MissingItemList[name]
                if img.IsValid (file_list):
                    valid_list.append (name)
            # 移除已經下載好的商品
            for name in valid_list:
                self.MissingItemList.pop (name, None)

            # 檢查次數上限
            if checkCnt < int (Global.get_value ("def_item_max_check_count")):
                checkCnt += 1
                self.super.PrintLog ("Checking for missing images count：%d" % (checkCnt))
            else:break

        missing_list = list ()
        file_list = os.listdir (Global.DF_EXPORT_IMG_PATH)
        for name in self.MissingItemList:
            img = self.MissingItemList[name]
            if not img.IsValid (file_list):
                missing_list.append ("圖片名稱：%s\t圖片連結：%s" % (img.Name, img.Url))

        if len (missing_list) > 0:
            self.super.PrintLog ("缺少圖片：")
            file_name = "商品圖片遺漏清單_" + self.m_AccountID + ".txt"
            with open (file_name, mode="w", encoding="utf-8") as myFile:
                for strInfo in missing_list:
                    myFile.write (strInfo)
                    self.super.PrintLog (strInfo)
        else:
            self.super.PrintLog ("Successfully collect item images.")

        self.super.PrintLog ("All items amount :" + str (self.AllItemNum))
        self.super.PrintLog ("All item images amount :" + str (len (self.AllImgInfo)))
        self.super.PrintLog ("Check done.")
        return True