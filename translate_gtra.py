import win32com.client
import os
import pathlib
import time
import sys
import re
import urllib.parse
import chromedriver_binary
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
from time import sleep


class Translator:       #seleniumによる翻訳を定義するクラス
    def __init__(self):
        self.options = Options()
        self.options.binary_location = "C:\\Program Files\\chrome-win\\chrome.exe"
        #self.options.add_argument("--headless")

        self.browser = webdriver.Chrome(options=self.options)
        self.browser.minimize_window()
        self.browser.implicitly_wait(2)

    def gtrans(self, txt , lg1 , lg2):  # lg1からlg2にGoogle翻訳する関数
        buf = txt
        if buf.replace(" ","") == "":                   # 入力が空ならそのまま返す
            return txt
        
        #if re.sub(""):正規表現でいろいろする
            
        if txt.replace(",", "").replace(".", "").replace("-", "").replace("_", "").replace("（", "").replace("）", "").replace("(", "").replace(")", "").replace("*", "").isnumeric() == True:
            # 2020.09.30のように入力が数字＋記号のみならそのまま返す
            return txt

        # 翻訳したい文をURLに埋め込んでからアクセス
        text_for_url = urllib.parse.quote_plus(txt, safe='')
        url = "https://translate.google.co.jp/#{1}/{2}/{0}".format(text_for_url , lg1 , lg2)
        self.browser.get(url)

        # 少し待つ
        wait_time = len(txt) / 1000
        if wait_time < 0.5:
            wait_time = 0.5
        time.sleep(wait_time)

        # 翻訳結果を抽出
        soup = BeautifulSoup(self.browser.page_source, "html.parser")
        ret =  soup.find(class_="tlid-translation translation")

        try:
            return ret.text
        except AttributeError:
            print (txt)
            return


def main():
    Application = win32com.client.Dispatch("Word.Application")
    Application.Visible = True
    global lg
    lg = ["ja" , "en"]

    os.chdir(os.path.dirname(__file__))
    path = os.getcwd()
    folder = path + "\\Translate"
    WordFolder = pathlib.Path(folder)
    Filepath=[ str(p) for p in WordFolder.iterdir()]
    Filename=[p.name for p in WordFolder.iterdir() if p.is_file()]

    for k in range(len(Filepath)):
        GetTranslation(Application, Filepath[k] , Filename[k] , folder)


def GetTranslation(Application, Filepath, Filename, folder):
    doc = Application.Documents.Open(Filepath)
    cmax = doc.paragraphs.count

    array = []

    for i in reversed(range(1, cmax+1)):
        #rtext = doc.paragraphs(i).Range.Text
        #rtext = rtext.replace("\r","")
        #ori_text = rtext.replace("\x07","")
        ori_text = doc.paragraphs(i).Range.Text
        print(i)
        if str(ori_text) == "":
            z = 1
        else:
            z = 0
            translation = translator.gtrans(ori_text, *lg)
            array.extend([ori_text, translation])
        if z == 0:
            try:
                #doc.Paragraphs(i).Range.InsertAfter(translation + "\n")     #挿入
                doc.Paragraphs(i).Range = translation
            except:
                print(i, "error")
    
    newFilePath = folder + "\\translated_" + Filename
    print(Filename)
    doc.SaveAs2(newFilePath)
    doc.Close()

if __name__ == "__main__":
    translator = Translator()
    main()