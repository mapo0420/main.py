import xlwings as xw
import pyautogui as gui
import time


from selenium import webdriver
from selenium.webdriver.common.by import By
import time


#参考書名から確認テストデータへのpathを取得できるよう辞書を定義
ref_book_to_path={"【5訂版】システム英単語":r"D:\①英語\②日大レベル\システム英単語\5訂版\★[5訂版]システム英単語_20210804本部.xlsm",
                  "［4thEd］NextStage":r"D:\①英語\②日大レベル\Next Stage\問題文付き_推奨4thNextStage_20221029本部.xlsm",
                  "[六] 英単語ターゲット1900":r"D:\①英語\②日大レベル\ターゲット1900\6訂版\【臨時】6訂版ターゲット1900_20220213本部.xlsm",
                  "必携英単語LEAP":r"D:\①英語\②日大レベル\必携英単語 LEAP\★必携英単語 LEAP_20210811本部.xlsm",
                  "DATABASE":r"D:\①英語\②日大レベル\Datebase4500\Datebase4500[五訂版]_20220313本部 (1).xlsm",
                  "英文法ポラリス1":r"D:\①英語\②日大レベル\関正生の英文法ポラリス1 標準レベル\★関正生の英文法ポラリス1 標準レベル_20220410本部 (1).xlsm"}

pdf_page={"参考書":["ページ番号"],
          "英語長文ポラリスレベル1":["1-2","3-6","7-9","10-11","12-13","13-15","16-17","18-19","20-22","23-24","25-27","28-29"]
          }


exceltest=["【5訂版】システム英単語","［4thEd］NextStage","[六] 英単語ターゲット1900","必携英単語LEAP","DATABASE","英文法ポラリス1"]
pdftest=[]
irregular=["DATABASE"]



#参考書名、範囲、問題数を受け取って確認テストを印刷する関数を定義
def print_test(text,range,num):

    #excelの場合
    if text in exceltest:
        print("参考書:"+text+" "+"範囲:"+range+" "+"問題数:"+num+"で確認テストを作製します。")
        
        #参考書名から参考書のpathを取得
        path=ref_book_to_path[text]
        
        #範囲の始めと終わりの問題番号を取得
        s=range.split("-")[0]
        e=range.split("-")[1]
        print("出題範囲は"+s+"から"+e)

        #確認テストのexcelファイルを開いて操作シートに移動
        wb=excel.books.open(path)
        sh=wb.sheets[1]
        sh.activate()
        print(text+"のexcelファイルを開いて操作シートに移動")

        time.sleep(0.5)
        
        #「まずはここをクリック」の座標をクリック
        gui.click(x=281,y=320)


        time.sleep(0.5)

        #出題範囲を入力
        gui.write(s)
        gui.press("enter")
        gui.write(e)
        gui.press("enter")
        print("出題範囲を入力")

        #例外的な処理が必要ない参考書について、問題数を入力してそのまま印刷
        if not text in irregular:
            gui.write(num)
            gui.press(["enter","enter"])
        
        #例外処理が必要な参考書について、必要な処理を加える
        else:
            #自動で印刷されない参考書について、解答とテストのシートを取得
            answer_sh=wb.sheets[3]
            test_sh=wb.sheets[4]
            
            time.sleep(0.5)

            #古文単語は大抵50問で出題されるので、2ページ印刷するよう設定
            if text=="古文単語315":
                sp=1
                ep=2

            #その他の参考書は1ページのままで設定
            else:
                sp=1
                ep=1
            #解答とテストを指定されたページ数で印刷
            answer_sh.api.PrintOut(From=sp,To=ep,Copies=1,ActivePrinter="Apeos C2060")
            test_sh.api.PrintOut(From=sp,To=ep,ActivePrinter="Apeos C2060")
            
        wb.close()  
    
    #PDFの場合    
    elif text in pdftest:
        page_list=pdf_page[text]

        #問題番号を指定してある場合("$"から始めるなどして書き方を統一して欲しい)
        if num.startswith("$"):
            num=num[1:] #"$を削除"
            num=num.split(",")
            for n in num:
                page=page_list[int(n)-1]
                #pdfの印刷すべきページ数を「page:str」に格納
                #ここからpdfを印刷するコードを書く

        #範囲と問題数のみ指定してある場合
        else:
            s=range.split("-")[0]
            e=range.split("-")[1]
            start_page=page_list[int(s)-1].split("-")[0]
            end_page=page_list[int(e)-1].split("-")[1]
            page=start_page+end_page
            #pdfの印刷すべきページ数を「page:str」に格納
            #ここからpdfを印刷するコードを書く
    else:
        print(text+"には未対応です。")


browser = webdriver.Chrome(executable_path="./chromedriver") 
USER = "T2103801"
PASS = "hiroki1107"
# GoogleChromeを起動
time.sleep(3)
# browser.implicitly_wait(3)
# ログインするサイトへアクセス
url_login = "https://www.takedajuku-system.com/user-auths/login/lecturer/"
browser.get(url=url_login)
time.sleep(3)
print('ログイン完了')
# USER,PASSを入力
em_user = browser.find_element(By.ID,'loginid')
em_user.clear()
em_user.send_keys(USER)
em_pass = browser.find_element(By.ID,'password')
em_pass.clear()
em_pass.send_keys(PASS)
print('入力完了')
# ログインボタンをクリック
btn_login = browser.find_element(By.CLASS_NAME,'btn-primary')
time.sleep(3)
btn_login.click()
print('ログイン完了')



# メニューボタン → 特訓スケジュール
time.sleep(2)
btn_menu = browser.find_element(By.ID,'sidr-menu-button')
time.sleep(3)
btn_menu.click()
print('メニューボタンクリック')
li_btn = browser.find_elements(By.ID,'sidr-id-buttonloading')
time.sleep(10)
btn_sch = li_btn[0]
btn_sch.click()
time.sleep(10)
print('特訓スケジュールボタンクリック')



# 週ボタンをクリック
btn_week = browser.find_element(By.XPATH,'//*[@id="calendar"]/div[1]/div[2]/div/button[2]')
time.sleep(1)
btn_week.click()
time.sleep(2)
print('週ボタンをクリック')


# 矢印を押して先週分へ
btn_arrow = browser.find_element(By.XPATH,'/html/body/div[1]/div/div[3]/div[2]/div[1]/div[3]/div[1]/div[1]/button[2]')
time.sleep(1)
btn_arrow.click()
time.sleep(2)
print('矢印ボタンをクリック')


# カレンダーの要素をリストで取得
elm_all = browser.find_elements(By.CLASS_NAME,'js-calendar-column')
time.sleep(3)
print('リストを取得')
length = len(elm_all)
print(length)



def saved():
    pass

from selenium.common.exceptions import ElementNotInteractableException

# elm_all[3].click()

temp = browser.find_element(By.XPATH,'/html/body/div[1]/div/div[3]/div[2]/div[1]/div[3]/div[2]/div/table/tbody/tr/td/div/div/div/div[2]/table/tbody/tr[4]/td[1]/a/div/span/div')
temp.click()
btn_repo = browser.find_element(By.ID,'report_view_page')

# try:
time.sleep(2)
# レポートを見るボタン
btn_repo.click() 
# except ElementNotInteractableException:
    # unsaved()
# else:
    # saved()

time.sleep(3)
# text = browser.('/html/body/div[1]/div/div[3]/div[2]/div[4]/div/div[2]/div/div[1]/div[2]/div/div')
hw = browser.find_elements(By.CLASS_NAME,'homework-item')
hw_num = len(hw)


#Chromeのウィンドウを右に移動
browser.set_window_position(800,200)

#excelアプリを起動して全画面表示
excel=xw.App(visible=True) 
time.sleep(0.5)
gui.doubleClick("C:\RupiaBot\確認テスト自動作成\excel.header.png")

for i in range(hw_num):
    text_xpath = f'/html/body/div[1]/div/div[3]/div[2]/div[4]/div/div[2]/div[{i+1}]/div[1]/div[2]/div/div'
    range_xpath = f'/html/body/div[1]/div/div[3]/div[2]/div[4]/div/div[2]/div[{i+1}]/div[1]/div[3]/div/div[1]'
    num_xpath = f'/html/body/div[1]/div/div[3]/div[2]/div[4]/div/div[2]/div[{i+1}]/div[1]/div[4]/div/div[1]'
    text = browser.find_element(By.XPATH,text_xpath)
    time.sleep(2)
    print(text.text.split(' ')[0])
    text=text.text.split(' ')[0]
    range = browser.find_element(By.XPATH,range_xpath)
    time.sleep(2)
    print(len(range.text))
    print(range.text.split("\n")[0])

    range=range.text.split('\n')[0]
    num = browser.find_element(By.XPATH,num_xpath)
    time.sleep(2)
    print(len(num.text))
    print(num.text)
    num=num.text

    #Chromeのウィンドウを右に移動
    browser.set_window_position(800,200)

    #取得したデータを引数として関数に渡して確認テストを印刷
    print_test(text,range,num)



#excelアプリを終了
excel.quit()

