import globals as Global
import time
from PyQt5 import QtWidgets
from playsound import playsound
from shopee_crawler_gui import Ui_ShopeeCrawlerGUI
from PyQt5.QtGui import QTextCursor
import ExcelCounter as Counter
import ShopeeCollector as Collector

ImgCollectionDoneInfoSound = "./sound/sound01.mp3"  # 提示音效

class MainWindowController (QtWidgets.QMainWindow):
    def __init__ (self):
        self.Log = str ()

        # 紀錄已開啟的瀏覽器狀態
        self.Collector = None

        # 紀錄來源目錄下的訂單Excel
        self.AllShopeeAccounts = dict ()

        super (MainWindowController, self).__init__ ()
        self.ui = Ui_ShopeeCrawlerGUI ()
        self.ui.setupUi (self)

    # UI物件初始化設定
    def UISetup (self):
        self.ui.button_collect_item_images.clicked.connect (self.OnHandleItemImagesCollection)                                           
        self.ui.button_process_excel_order.clicked.connect (self.OnHandleExcelOrder)                                           
        self.ui.button_save.clicked.connect (self.SaveUserData) 
        self.ui.button_cleanlog.clicked.connect (self.CleanLog)
        self.ui.lineEdit_images_collection_interval_secs.setText (str (Global.get_value ("def_get_item_elm_interval_secs")))
        self.ui.lineEdit_crawler_blocker_waitting_secs.setText (str (Global.get_value ("def_waitting_for_shopee_crawler_blocker_secs")))

    def SaveUserData (self):
        Global.set_value ("def_get_item_elm_interval_secs", float (self.ui.lineEdit_images_collection_interval_secs.text ()))
        Global.set_value ("def_waitting_for_shopee_crawler_blocker_secs", float (self.ui.lineEdit_crawler_blocker_waitting_secs.text ()))
        Global.save_user_data ()
        Global.LoadTempParameters ()

    def PrintLog (self, _message):
        self.Log += _message
        self.Log += "\n"
        self.ui.textBrowser_log.setText (self.Log)
        self.ui.textBrowser_log.repaint ()
        self.ui.textBrowser_log.moveCursor (QTextCursor.End)

    def CleanLog (self):
        self.Log = ""
        self.ui.textBrowser_log.setText (self.Log)
        self.ui.textBrowser_log.repaint ()
        self.ui.textBrowser_log.moveCursor (QTextCursor.Start)

    def FormatTimeBySecs (self, secs):
        if secs < 60:
            return "%d秒" % (secs)
        elif secs < 3600:
            return "%d分 : %d秒" % (int(secs / 60), secs % 60)
        else:
            return "%d時 : %d分 : %d秒" % (int(secs / 3600), (secs - int(int(secs / 3600) * 3600)) / 60, secs % 60)

    # 彙整商品Excel訂單 (彙整指定目錄下的所有訂單Excel)
    def OnHandleExcelOrder (self):
        self.AllShopeeAccounts.clear ()
        Counter.ReadAllShopeeOrderExcel (self, self.AllShopeeAccounts)
        Counter.ExportAllExcelData (self, self.AllShopeeAccounts)

    # 蒐集賣場商品圖片
    def OnHandleItemImagesCollection (self):
        url_login = Global.get_value ("def_login_url", "")
        if not self.Collector:
            # TODO:可加入自動偵測瀏覽器是否被關閉，若被關閉則重啟webdriver
            self.PrintLog ("Excuting...")
            self.Collector = Collector.CShopeeCollector (self, url_login)
            self.Collector.OpenWebdriver ()
        else:
            begin_time = time.time ()
            if not self.Collector.GotoShopeeItemList ():
                self.PrintLog ("Failed to go to item-list page.")
                return
            if not self.Collector.HandleItemImagesCollection ():
                self.PrintLog ("Failed to collect item images.")
                return
            end_time = time.time ()
            self.PrintLog ("Finishing item images collection with " + self.FormatTimeBySecs (end_time - begin_time) + ".")
            # playsound (ImgCollectionDoneInfoSound) #TODO:Debug
            # CauseError:utf-8 codec can't decode byte 0xb0 in position 0: invalid start byte

            