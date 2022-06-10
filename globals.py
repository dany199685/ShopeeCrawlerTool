import os

# 資料夾路徑
DF_EXPORT_EXCEL_PATH = "./export/"
DF_SOURCE_EXCEL_PATH = "./sourceExcel/"
DF_EXPORT_IMG_PATH   = "./sourceImg/"

# 使用者資料
DF_CONFIG_FILE  = "./config.txt"

def mkdir (path):
    if not os.path.exists (path):
        os.makedirs (path)

def set_value (key, value):
    _global_dict [key] = value

def get_value (key, default = None):
    try:
        return _global_dict[key]
    except:
        return default

def save_user_data ():
    strSave = str ()
    for key in _global_dict:
        strSave += key
        strSave += " "
        strSave += str (_global_dict[key])
        strSave += "\n"
    with open (DF_CONFIG_FILE, mode="w", encoding="utf-8") as myFile:
        myFile.write (strSave)
    MainWindow.PrintLog ("Save user's configuration successfully.")


def GetDataDict (file):
    if not os.path.exists (file):
        return dict ()
    
    result = dict ()
    with open (file, mode="r", encoding="utf-8") as myFile:
        kAppInfo = myFile.readlines ()  # 讀取文件內全部文字
        for strData in kAppInfo:
            strData = strData[0:(len(strData) - 1)] # 排除換行符號
            tempStr = strData.split()
            result [tempStr[0]] = tempStr[1]
    return result

def LoadParameters ():
    config_dict = GetDataDict (DF_CONFIG_FILE)

    #【Url】
    # 起始登入畫面
    set_value ("def_login_url", config_dict.get ("def_login_url", ""))

    # 賣家後台-我的商品網址
    set_value ("def_my_item_list_url", config_dict.get ("def_my_item_list_url", ""))
    
    #【Shopee Parameters】
    set_value ("def_item_max_num_in_a_page", int (config_dict.get ("def_item_max_num_in_a_page", 24)))

    set_value ("def_waitting_for_shopee_crawler_blocker_secs", float (config_dict.get ("def_waitting_for_shopee_crawler_blocker_secs", 10.0)))

    set_value ("def_item_max_check_count", int (config_dict.get ("def_item_max_check_count", 10)))

    set_value ("def_default_get_elm_waitting_secs", float (config_dict.get ("def_default_get_elm_waitting_secs", 10.0)))
    
    set_value ("def_waitting_for_page_loaded_secs", float (config_dict.get ("def_waitting_for_page_loaded_secs", 10.0)))
    
    set_value ("def_get_item_elm_interval_secs", float (config_dict.get ("def_get_item_elm_interval_secs", 0.3)))

    # 剛開啟瀏覽器等待網頁載入時間
    set_value ("def_waitting_for_begining_loading_secs", float (config_dict.get ("def_waitting_for_begining_loading_secs", 3.0)))

    #【Excel】
    set_value ("def_worksheet_default_name", config_dict.get ("def_worksheet_default_name", "orders"))
    
    set_value ("def_export_excel_name", config_dict.get ("def_export_excel_name", "蝦皮賣場資料彙整"))

    set_value ("def_itemMainParams_num", int (config_dict.get ("def_itemMainParams_num", 2)))

    set_value ("def_itemId_col", int (config_dict.get ("def_itemId_col", 0)))

    set_value ("def_buyerId_col", int (config_dict.get ("def_buyerId_col", 4)))

    set_value ("def_itemName_col", int (config_dict.get ("def_itemName_col", 23)))

    set_value ("def_itemMainId_col", int (config_dict.get ("def_itemMainId_col", 27)))

    set_value ("def_itemOptionId_col", int (config_dict.get ("def_itemOptionId_col", 28)))

    set_value ("def_itemNum_col", int (config_dict.get ("def_itemNum_col", 29)))

    set_value ("def_itemImg_col_index", config_dict.get ("def_itemImg_col_index", "A"))

    set_value ("def_shopId_col_index", config_dict.get ("def_shopId_col_index", "B"))
    
    set_value ("def_itemId_col_index", config_dict.get ("def_itemId_col_index", "C"))
    
    set_value ("def_itemName_col_index", config_dict.get ("def_itemName_col_index", "D"))
    
    set_value ("def_item_subId_col_index", config_dict.get ("def_item_subId_col_index", "E"))
    
    set_value ("def_itemNum_col_index", config_dict.get ("def_itemNum_col_index", "F"))
    
    set_value ("def_itemDetail_col_index", config_dict.get ("def_itemDetail_col_index", "G"))

    set_value ("def_grid_autoSetHeight_by_minOrderSize", int (config_dict.get ("def_grid_autoSetHeight_by_minOrderSize", 10)))

    set_value ("def_itemImg_grid_min_height", int (config_dict.get ("def_itemImg_grid_min_height", 180)))

    set_value ("def_itemImg_grid_width", float (config_dict.get ("def_itemImg_grid_width", 28.5)))

    set_value ("def_item_shopId_grid_width", int (config_dict.get ("def_item_shopId_grid_width", 15)))

    set_value ("def_itemId_grid_width", int (config_dict.get ("def_itemId_grid_width", 15)))

    set_value ("def_itemName_grid_width", int (config_dict.get ("def_itemName_grid_width", 50)))

    set_value ("def_itemSubId_grid_width", int (config_dict.get ("def_itemSubId_grid_width", 20)))

    set_value ("def_itemNum_grid_width", int (config_dict.get ("def_itemNum_grid_width", 15)))

    set_value ("def_itemDetail_grid_width", int (config_dict.get ("def_itemDetail_grid_width", 40)))

    set_value ("def_itemImg_width", int (config_dict.get ("def_itemImg_width", 200)))

    set_value ("def_itemImg_height", int (config_dict.get ("def_itemImg_height", 200)))

    set_value ("def_itemDetail_spliter", config_dict.get ("def_itemDetail_spliter", " | "))

    strTitle = config_dict.get ("def_excel_titles", "商品參考圖,賣場編號,商品主要貨號,商品名稱,商品選項貨號,數量,訂單詳細資訊")
    set_value ("def_excel_titles", strTitle.split (','))

def LoadTempParameters ():
    global DF_GET_ITEM_ELM_TIME
    DF_GET_ITEM_ELM_TIME = float (get_value ("def_get_item_elm_interval_secs"))

    global DF_GET_ELM_DEFAULT_TIME
    DF_GET_ELM_DEFAULT_TIME = float (get_value ("def_default_get_elm_waitting_secs"))

def _init ():
    global _global_dict 
    _global_dict = {}
    LoadParameters ()
    LoadTempParameters ()

    from Controller import MainWindowController
    global MainWindow
    MainWindow = MainWindowController ()
    MainWindow.UISetup ()

    mkdir (DF_EXPORT_EXCEL_PATH)
    mkdir (DF_EXPORT_IMG_PATH)
    mkdir (DF_SOURCE_EXCEL_PATH)


