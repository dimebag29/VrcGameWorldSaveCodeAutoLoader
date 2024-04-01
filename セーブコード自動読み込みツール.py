# ==============================================================================================================
# 作成者:dimebag29 作成日:2024年4月1日 バージョン:v0.2
# (Author:dimebag29 Creation date:April 1, 2024 Version:v0.2)
#
# このプログラムのライセンスはLGPLv3です。pynputライブラリのライセンスを継承しています。
# (This program is licensed to LGPLv3. Inherits the license of the pynput library.)
# https://www.gnu.org/licenses/lgpl-3.0.html.en
#
# 開発環境 (Development environment)
# ･python 3.7.5
# ･auto-py-to-exe 2.36.0 (used to create the exe file)
#
# exe化時のauto-py-to-exeの設定
# ･ひとつのファイルにまとめる (--onefile)
# ･ウィンドウベース (--windowed)
# ･exeアイコン設定 (--icon) (VrcGameWorldSaveCodeAutoLoaderIcon.ico)
# ･追加ファイルでGUI用アイコンと処理完了音ファイル、処理警告音ファイルを追加 (--add-data)
#   (VrcGameWorldSaveCodeAutoLoaderIcon.ico, Normal.wav, Error.wav)
#
# ToN (TerrorsOfNowhere)でのセーブコードの自動保存設定
AutoSaveInTerrorsOfNowhere = True
# ==============================================================================================================

# python 3.7.5の標準ライブラリ (Libraries included as standard in python 3.7.5)
import subprocess
import os
import sys
import time
import threading
import socket
import json
import glob
from itertools import islice
import re
import datetime
from tkinter import *
import winsound

# 外部ライブラリ (External libraries)
from pynput import keyboard                                                     # Version:1.7.6
import psutil                                                                   # Version:5.9.6
import pyperclip                                                                # Version:1.8.2


# ================================================= 関数定義 ====================================================
# GUI操作関係 ------------------------------------------------------------------
# GUIの保存場所ボタンが押された時に呼ばれる関数
def Push_OpenSaveFolderButton():
    # エクスプローラーで、プレイヤー別フォルダが保存されるフォルダを開く
    subprocess.Popen(["explorer", SaveFileDir], shell=True)


# GUIの手動読み込みボタンが押された時に呼ばれる関数
def Push_LoadButton():
     # 現在のワールド用のセーブコード保存ファイルのパスを取得してみる。無かったら""が返ってくる
    NowEnteringWorldSaveCodeFilePath = TryGetSaveCodeFilePathOfNowEnteringWorld(False)
    GetNewestSaveCodeFromFilePathOnClipboard(NowEnteringWorldSaveCodeFilePath)


# GUIの保存ボタンが押された時、XSOverlayの手首メニューの「次の曲」ボタンが押された時、ToNのセーブコードがログに確認された際に呼ばれる関数
def Push_SaveButton(InputSaveCode = None):

    NowDateTime = datetime.datetime.now()                                       # 現在の日時を取得しておく

    # 文字列の引数があったら、その文字列を保存。
    # 引数が無かったら、現在のクリップボードの中身を保存。
    if str == type(InputSaveCode):
        SaveCodeValue = InputSaveCode                                           # 文字列の引数の中身を保存
    else:
        SaveCodeValue = pyperclip.paste()                                       # 現在のクリップボードの中身を保存

    NowEnteringWorldSaveCodeFilePath = TryGetSaveCodeFilePathOfNowEnteringWorld(True) # セーブコードを保存するファイルパスを取得

    # すでにセーブコードファイルが存在したら、前回保存時とセーブコードがちゃんと変わってるか確認
    if True == os.path.exists(NowEnteringWorldSaveCodeFilePath):
        # セーブコードファイル内の最新のセーブコードを取得
        with open(NowEnteringWorldSaveCodeFilePath, 'r', encoding="utf-8") as f:
            SaveCodeList = [s.rstrip() for s in f.readlines()]                  # ファイル全体をリストとして読み込み (内包表記で末尾の改行コードを削除)
            SaveCodeList = [a for a in SaveCodeList if a != '']                 # 内包表記で空の要素を駆逐する https://qiita.com/github-nakasho/items/a08e21e80cbc9761db2f#%E5%86%85%E5%8C%85%E8%A1%A8%E8%A8%98%E3%81%A7%E7%A9%BA%E3%81%AE%E8%A6%81%E7%B4%A0%E3%82%92%E9%A7%86%E9%80%90%E3%81%99%E3%82%8B
            #print(*SaveCodeList, sep="\n")
            TempLoadSaveCode_Time  = SaveCodeList[-1].split(SplitWordInSaveCode)[0] # 最終行から日時を取得
            TempLoadSaveCode_Value = SaveCodeList[-1].split(SplitWordInSaveCode)[1] # 最終行からセーブコード部分を取得
        
        # 前回保存時とセーブコードが同じだったら
        if TempLoadSaveCode_Value == SaveCodeValue:
            BeforeSaveTime = datetime.datetime.strptime(TempLoadSaveCode_Time, "%Y-%m-%d %H:%M:%S")         # セーブコードが保存されていた日時
            HowLongAgo =  ReturnJapaneseAboutElapsedTimeFromTimedelta(NowDateTime - BeforeSaveTime) + "前"  # セーブコードがどのくらい前のものか計算

            # XSOverlay通知
            ViewXSOverlayNotification("保存中止 : " + HowLongAgo + "に保存されたセーブコードと同じです")

            # GUI通知
            UpdateSystemMessage("⚠ 保存中止 ⚠\n" + HowLongAgo + "に保存されたセーブコードと同じです")

            # 音を鳴らす
            time.sleep(0.2)
            winsound.PlaySound(SoundPath_Error, winsound.SND_FILENAME)

            # 前回保存時とセーブコードが同じだったので、保存処理をせず、ここで関数を終える
            return

    # ファイルに追記モードで書き込み ---------------------------------------------
    with open(NowEnteringWorldSaveCodeFilePath, 'a', encoding="utf-8") as f:
        f.write("\n\n" + NowDateTime.strftime("%Y-%m-%d %H:%M:%S") + SplitWordInSaveCode + SaveCodeValue)

    ViewXSOverlayNotification("セーブコードを保存しました")                     # セーブコード保存完了をXSOverlay通知
    UpdateSystemMessage("セーブコードを保存しました")                           # セーブコード保存完了をGUI通知
    button_Load["state"] = "normal"                                             # GUIの読み込みボタンを押せるようにする

    # 音を鳴らす
    time.sleep(0.2)
    winsound.PlaySound(SoundPath_Normal, winsound.SND_FILENAME)


# GUIのシステムメッセージのクリアボタンが押された時に呼ばれる関数
def Push_ClearButton():
    TextView_SystemMessage["text"] = ""


# 引数でもらったメッセージを日時付きでGUIのシステムメッセージに表示する関数
def UpdateSystemMessage(Message):
    NowTime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")             # 現在の日時をstrで取得
    TextView_SystemMessage["text"] = NowTime + "\n" + Message                   # システムメッセージに表示


# XSOverlayからの操作関係 -------------------------------------------------------
# メディアキー入力監視スレッド
def StartMediakeyLoggingThread():
    global listener                                                             # 関数内で変数の値を変更したい場合はglobalにする

    with keyboard.Listener(on_press=None, on_release=None, win32_event_filter=win32_event_filter, suppress=False) as listener:
        listener.join()
    # https://stackoverflow.com/questions/54394219/pynput-capture-keys-prevent-sending-them-to-other-applications
    # https://github.com/moses-palmer/pynput/issues/170#issuecomment-602743287


# メディアキー(次の曲)が押された時の処理
def win32_event_filter(msg, data):
    global BeforeMediaKeyPushedTime

    # VRモードで「次の曲」ボタンが押されたら実行
    if True == PlayerInfoDict["VrMode"] and 0xB0 == data.vkCode:
        if 256 == msg:                                                          # キー上げ
            NowTime = time.time()                                               # 現在の日時を保存しておく
            if 1.0 > (NowTime - BeforeMediaKeyPushedTime):                      # 前回「次の曲」ボタンが押されてから経過した時間が1.0秒未満だったら実行
                # クリップボードの内容を保存する関数をスレッドとして実行
                # (処理に時間がかかるとsuppress_event()が効かない場合があったのでスレッドとして実行する)
                Push_SaveButtonThread = threading.Thread(target=Push_SaveButton, daemon=True)
                Push_SaveButtonThread.start()
                NowTime = 0.0                                                   # トリプルクリック時に2回Push_SaveButtonThreadが実行されるのを防止
            BeforeMediaKeyPushedTime = NowTime
        listener.suppress_event()                                               # メディアキー(次の曲)入力が他のプログラムに伝わらないようにここで抑止
    # https://docs.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes
    #   data.vkCode >> 再生/一時停止:0xB3、次の曲:0xB0、前の曲:0xB1
    # https://github.com/moses-palmer/pynput/issues/170#issuecomment-602743287
    #   msg >> キー下げ:257、キー上げ:256


# XSOverlayに通知を出す処理
def ViewXSOverlayNotification(NotificationInput):
    global Message                                                              # 関数内で変数の値を変更したい場合はglobalにする
    
    # VRじゃなかったらこの関数を終える
    if True != PlayerInfoDict["VrMode"]:
        return
    
    print("XSOverlay通知文 : " + NotificationInput)
    Message["content"] = NotificationInput                                      # XSOverlay通知文を更新

    SendData = json.dumps(Message).encode("utf-8")                              # 通知文字列をJSON形式にエンコード
    MySocket = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)                 # ソケット通信用インスタンスを生成
    MySocket.sendto(SendData, ("127.0.0.1", 42069))                             # XSOverlayに通知依頼を送信
    MySocket.close()                                                            # ソケット通信終了
    # https://zenn.dev/eeharumt/scraps/95f49a62dd809a
    # https://gist.github.com/nekochanfood/fc8017d8247b358154062368d854be9c


# その他 -----------------------------------------------------------------------
# VRChatのログフォルダ内から、名前の日時が一番新しいログファイルのパスを返す関数
def GetNewestLogFilePath():
    # LogFileDir内でファイル名検索文字列LogFileSearchWordにヒットしたファイルのパスリスト
    LogFilePathList = list(glob.glob(LogFileDir + LogFileSearchWord))

    # ひとつもログファイルを取得できなかったら空を返す
    if 0 >= len(LogFilePathList):
        return ""

    LogFilePathList.sort(reverse=True)                                          # 名前でソート
    return LogFilePathList[0]                                                   # 日時が一番新しいログファイルのパスを返す


# ファイルの行数を取得する関数
def GetLastLineNumFromLogFile(InputFilePath):
    # ファイルの行数を取得 https://stackoverflow.com/questions/845058/how-to-get-the-line-count-of-a-large-file-cheaply-in-python/1019572#1019572
    with open(InputFilePath, "rb") as f:
        Line = sum([1 for Temp in f])                                           # ファイルの行数分1が格納されたリストを内包表記で生成しSumすることで行数取得
    return Line


# VRChatが起動しているか確認する関数
def CheckVRChatRunning():
    IsVRChatRunning = False
    
    SendCommand = 'tasklist /nh /fi "IMAGENAME eq VRChat.exe"'
    output = subprocess.run(SendCommand, startupinfo=StartupInfo, stdout=subprocess.PIPE, text=True)
    #print(output.stdout.split('\n')[1])

    if "VRChat.exe" in output.stdout.split('\n')[1]:
        IsVRChatRunning = True

    return IsVRChatRunning


# ログファイルからプレイヤー情報(VRかどうか、ユーザー名、ユーザーID)を取得する関数
def TryGetUserInfo(InputLogFilePath):
    global PlayerInfoDict

    # 初期化
    PlayerInfoDict["VrMode"] = None                                             # VRかどうか
    PlayerInfoDict["Name"]   = ""                                               # ユーザー名
    PlayerInfoDict["Id"]     = ""                                               # ユーザーID

    # ログファイルを開く
    with open(InputLogFilePath, 'r', encoding="utf-8") as f:
        for Line in f:
            # VRかどうか 取得
            if "XR Device: " in Line:                                           # 現在の行内に「XR Device: 」という文字列が含まれてるか
                if "None" in Line:                                              # 現在の行内に「None」という文字列が含まれてるか
                    PlayerInfoDict["VrMode"] = False                            # VRではない
                else:
                    PlayerInfoDict["VrMode"] = True                             # VRである
                
            # ユーザー名, ユーザーID 取得
            if "User Authenticated: " in Line:                                  # 現在の行内に「User Authenticated: 」という文字列が含まれてるか
                SplitStr = Line[54:].split(" (usr_")                            # 現在の行を「 (usr_」で区切ってリスト化
                PlayerInfoDict["Name"] = CheckAndFix_DirAndFileName(SplitStr[0])# ユーザー名
                PlayerInfoDict["Id"]   = "usr_" + SplitStr[1].rstrip()[:-1]     # ユーザーID
            
            # 全てのプレイヤー情報がそろったらTrueを返す
            if None != PlayerInfoDict["VrMode"] and "" != PlayerInfoDict["Name"] and "" != PlayerInfoDict["Id"]:
                return True
    
    # まだそろってなかったらFalseを返す
    return False


# 現在のプレイヤー用のフォルダを作る。必ずTureを返す
def CreateNowAccountDir():
    global PlayerDir

    SearchDir = os.path.join(SaveFileDir, "*" + PlayerInfoDict["Id"] + "*")     # フォルダ名にプレイヤーIDを含むフォルダ検索用パス
    HitSearchDirList = list(glob.glob(SearchDir))                               # フォルダ名にプレイヤーIDを含むフォルダのパスがまとめられたリスト

    # 存在してほしいフォルダのパス
    PlayerDir = os.path.join(SaveFileDir, PlayerInfoDict["Name"] + "_" + PlayerInfoDict["Id"])
    #print("PlayerDir : " + PlayerDir)

    # プレイヤーIDを含むフォルダが無かったら作る
    if 0 >= len(HitSearchDirList):
        os.makedirs(PlayerDir, exist_ok=True)                                   # フォルダを生成。ないはずだけどすでにあったらそのまま
        print("Create PlayerDir")
        return True
    
    # 一番新しい、フォルダ名にプレイヤーIDを含むフォルダのパスを取得
    HitSearchDirList.sort(key=os.path.getmtime, reverse=True)    
    HitSearchDir = HitSearchDirList[0]

    # 存在してほしいフォルダと完全一致するフォルダがあったらそのまま
    if PlayerDir == HitSearchDir:
        print("Already Created PlayerDir")
        return True
    
    # プレイヤーIDが部分一致するフォルダがあったらリネームする (ユーザー名が変更されたということ)
    else:
        print("Rename PlayerDir")
        os.rename(HitSearchDir, PlayerDir)
        return True


# 現在のワールド用のセーブコード保存ファイルのパスを取得してみる関数
# 引数Falseの場合、ファイルが無かったら""を返し、引数Trueの場合、存在してほしかったファイルパスを返す
def TryGetSaveCodeFilePathOfNowEnteringWorld(ReturnIdealPath):

    # 存在してほしいファイルパス
    IdealPath = os.path.join(PlayerDir, NowEnteringWorldDict["Name"] + "_" + NowEnteringWorldDict["Id"] + ".txt")

    SearchPath = os.path.join(PlayerDir, "*" + NowEnteringWorldDict["Id"] + "*.txt")    # ファイル名にワールドIDを含むファイル検索用パス
    HitSearchPathList = list(glob.glob(SearchPath))                                     # ファイル名にワールドIDを含むファイルのパスがまとめられたリスト

    # ワールドIDを含むファイルが無かったら""を返す
    if 0 >= len(HitSearchPathList) and False == ReturnIdealPath:
        print("Not Created SaveCodeFile of Now Entering World")
        return ""
    
    # 引数Trueの場合、ワールドIDを含むファイルがなかった場合でも、存在してほしかったファイルパスを返す
    # Push_SaveButton()でセーブコード保存ファイルを新規作成する際のパス取得に使用
    if 0 >= len(HitSearchPathList) and True == ReturnIdealPath:
        print("Return ideal SaveFile path (File not created)")
        return IdealPath
    
    # 一番新しい、ファイル名にワールドIDを含むファイルのパスを取得
    HitSearchPathList.sort(key=os.path.getmtime, reverse=True)    
    HitSearchPath = HitSearchPathList[0]

    # 存在してほしいファイルパスと完全一致するファイルがあったらそのまま
    if IdealPath == HitSearchPath:
        print("Already Created SaveCodeFile of Now Entering World")
    # ワールドIDが部分一致するファイルがあったらリネームする (ワールド名が変更されたということ)
    else:
        print("Rename SaveCodeFile of Now Entering World")
        os.rename(HitSearchPath, IdealPath)
    return IdealPath
    

# 引数で指定されたセーブコードファイルのパスから最新のセーブコードを取得し、クリップボードに保存する
def GetNewestSaveCodeFromFilePathOnClipboard(InputSaveCodeFilePath):

    # セーブコードDict初期化
    LoadedSaveCodeDict = {"Time":"", "Value":""}

    # ファイルからセーブコードDictを更新
    with open(InputSaveCodeFilePath, 'r', encoding="utf-8") as f:
        SaveCodeList = [s.rstrip() for s in f.readlines()]                      # ファイル全体をリストとして読み込み (内包表記で末尾の改行コードを削除)
        SaveCodeList = [a for a in SaveCodeList if a != '']                     # 内包表記で空の要素を駆逐する https://qiita.com/github-nakasho/items/a08e21e80cbc9761db2f#%E5%86%85%E5%8C%85%E8%A1%A8%E8%A8%98%E3%81%A7%E7%A9%BA%E3%81%AE%E8%A6%81%E7%B4%A0%E3%82%92%E9%A7%86%E9%80%90%E3%81%99%E3%82%8B
        #print(*SaveCodeList, sep="\n")
        LoadedSaveCodeDict["Time"]  = SaveCodeList[-1].split(SplitWordInSaveCode)[0] # 最終行から日時部分を取得
        LoadedSaveCodeDict["Value"] = SaveCodeList[-1].split(SplitWordInSaveCode)[1] # 最終行からセーブコード部分を取得
        #print("GetSaveCode!! Length = " + str(len(LoadedSaveCodeDict["Value"] )))

    # セーブコードをクリップボードにコピー
    pyperclip.copy(LoadedSaveCodeDict["Value"])

    NowDateTime = datetime.datetime.now()                                                           # 現在の日時を取得しておく
    BeforeSaveTime = datetime.datetime.strptime(LoadedSaveCodeDict["Time"], "%Y-%m-%d %H:%M:%S")    # セーブコードが保存されていた日時
    HowLongAgo =  ReturnJapaneseAboutElapsedTimeFromTimedelta(NowDateTime - BeforeSaveTime) + "前"  # セーブコードがどのくらい前のものか計算

    # このワールドのセーブコードをクリップボードに読み込んだことを通知
    ViewXSOverlayNotification(HowLongAgo + "のセーブコードをクリップボードに読み込み完了")  # XSOverlay
    UpdateSystemMessage(HowLongAgo + "のセーブコードを\nクリップボードに読み込み完了")      # GUI

    # 音を鳴らす
    time.sleep(0.2)
    winsound.PlaySound(SoundPath_Normal, winsound.SND_FILENAME)


# フォルダ名、ファイル名に使えない文字列をハイフンに置換する
def CheckAndFix_DirAndFileName(InputStr):
    # http://iori084.blog.fc2.com/blog-entry-16.html
    return re.sub(r'[\\|/|:|?|"|<|>|\|]', '-', InputStr)


# Timedelta型からおおよその経過時間を日本語strで返す
def ReturnJapaneseAboutElapsedTimeFromTimedelta(Td):
    Days    = Td.days
    Hours   = Td.seconds // 3600
    Minutes = (Td.seconds // 60) % 60
    Seconds = Td.seconds % 60

    if 0 < Days:
        return str(Days) + "日"
    elif 0 < Hours:
        return str(Hours) + "時間"
    elif 0 < Minutes:
        return str(Minutes) + "分"
    elif 0 < Seconds:
        return str(Seconds) + "秒"



# メインループスレッド -----------------------------------------------------------
def MainLoopThread():
    FirstLoop = True

    LogFilePath =  ""                                                           # 最新のログファイルのパス
    LogFilePath_Old = ""                                                        # 前回ループのログファイルパス

    LastLineNum = 0                                                             # 何行目まで読み込む必要があるか
    LastLineNum_Old = 0                                                         # 前回ループのログファイル読み込み済み行数

    EnteringWorldId_Old = ""                                                    # 前回ループのワールドID

    LogFileNewCreatedFlag = False                                               # ログファイルが新しく生成されたかフラグ
    GetUserInfo = False                                                         # プレイヤー情報(VrMode, Name, Id)を取得できたらTrue
    AlreadyCreatedAccountDir = False                                            # 現在のプレイヤー用のセーブコード保存フォルダの確認･作成が完了していたらTrue

    FindedTerrorsOfNowhereSaveCode = ""                                         # ログ内に見つけたTerrors of Nowhereのセーブコード


    # 最新のログファイルパスを取得
    LogFilePath = GetNewestLogFilePath()

    # VRChatよりも後にこのツールが起動していたら、Logの新規生成確認をスキップしてプレイヤー情報を取得できるようにしておく
    if True == CheckVRChatRunning():
        LogFileNewCreatedFlag = True
    
    # ログファイル監視ループ
    while True:
        time.sleep(3)                                                           # 少し待つ

        # 前回ループの情報を保存しておく -----------------------------------------
        LogFilePath_Old = LogFilePath                                           # 前回ループのログファイルパスを保存しておく
        LastLineNum_Old = LastLineNum                                           # 前回ループのログファイル読み込み済み行数を保存しておく
        EnteringWorldId_Old = NowEnteringWorldDict["Id"]                        # 前回ループのワールドIDを保存しておく
        FindedTerrorsOfNowhereSaveCode = ""                                     # ログ内に見つけたTerrors of Nowhereのセーブコード初期化

        # VRChatが起動してなかったら色々リセットしてスキップ------------------------
        if False == CheckVRChatRunning():
            LogFileNewCreatedFlag = False                                       # ログファイルが新しく生成されたかフラグを折る
            GetUserInfo = False                                                 # プレイヤー情報を取得できたフラグを折る
            AlreadyCreatedAccountDir = False                                    # 現在のプレイヤー用のセーブコード保存フォルダの確認･作成が完了していたフラグを折る
            NowEnteringWorldDict["Name"] = ""                                   # 現在のワールド名を空にする
            NowEnteringWorldDict["Id"]   = ""                                   # 現在のワールドIDを空にする
            TextView_WorldName["text"] = "-"                                    # GUIのワールド名表示を初期化
            button_Load["state"] = "disable"                                    # GUIの読み込みボタンを押せないようにする
            button_Save["state"] = "disable"                                    # GUIの保存ボタンを押せないようにする
            TextView_SystemMessage["text"] = ""                                 # GUIのシステムメッセージをクリアする
            FindedTerrorsOfNowhereSaveCode = ""                                 # ログ内に見つけたTerrors of Nowhereのセーブコード初期化
            #print("VRChat is Not Running")
            continue                                                            # ログファイル監視処理をここでスキップ
        #print("VRChat is Running")
        
        # Logが新規生成されてた -------------------------------------------------
        LogFilePath = GetNewestLogFilePath()                                    # 最新のログファイルパスを取得

        # ログファイルがひとつも生成されてなかったらスキップ (キャッシュ消去後やVRC初回起動時など想定)
        if "" == LogFilePath:
            continue

        # 前回ループからログファイルパスが変わってたら実行 (Logが新規生成されたということ)
        if LogFilePath_Old != LogFilePath:
            print("LogFile NewCreated")
            LogFileNewCreatedFlag = True                                        # ログファイルが新しく生成されたフラグを立てる
        
        # ログファイルが新しく生成されたフラグがたってなったらログファイル監視処理をここでスキップ
        if False == LogFileNewCreatedFlag:
            continue
        
        # プレイヤー情報(VrMode, Name, Id)を取得 --------------------------------
        if False == GetUserInfo:                                                # プレイヤー情報をまだ取得出来てなかったら実行
            print("TryGetPlayerInfo")
            GetUserInfo = TryGetUserInfo(LogFilePath)                           # プレイヤー情報取得をトライ。成功したらTrue、失敗したらFalseが返ってくる
            print(PlayerInfoDict)

            # VRChat起動直後で、プレイヤー情報をまだ取得できなかったらログファイル監視処理をここでスキップ
            if False == GetUserInfo:
                continue
                
        # 現在のプレイヤー用のフォルダを作る --------------------------------------
        if False == AlreadyCreatedAccountDir:                                   # 現在のプレイヤー用のセーブコード保存フォルダが作られてなかったら実行
            print("CreateNowAccountDir")
            AlreadyCreatedAccountDir = CreateNowAccountDir()                     # 現在のプレイヤー用のセーブコード保存フォルダを作る。必ずTrueが返ってくる

        # ログファイル追記分をチェックして、ワールドが変わってるか確認 --------------
        LastLineNum = GetLastLineNumFromLogFile(LogFilePath)                    # ログファイルの行数を取得
        #print(str(LastLineNum_Old), str(LastLineNum))

        if LastLineNum_Old != LastLineNum:
            with open(LogFilePath, 'r', encoding="utf-8") as f:                 # 読み込み専用で開く
                LineList_Iterator = islice(f, LastLineNum_Old, LastLineNum)     # 行数範囲指定して読込
                LineList = [str(t).rstrip() for t in LineList_Iterator]         # イテレータリストからstrリストに変換。文末の改行も消去

                # ワールド名とワールドIDが書かれてれば取得する
                for Line in LineList:
                    # ワールド名取得
                    if "Entering Room: " in Line:
                        # ディレクトリ名、ファイル名に使えない文字はハイフンに変換してもらってから取得
                        NowEnteringWorldDict["Name"] = CheckAndFix_DirAndFileName(Line[61:])
                        # GUIのワールド名を表示
                        TextView_WorldName["text"] = NowEnteringWorldDict["Name"]
                    
                    # ワールドID取得
                    if "Joining wrld_" in Line:
                        NowEnteringWorldDict["Id"] = Line[54:95]
                    
                    # Terrors of Nowhereのセーブコードがあったら実行
                    if "-  [START]" in Line:
                        # [START] から [END] までの文字列を抽出
                        FindedTerrorsOfNowhereSaveCode = Line.split("[START]")[1].split("[END]")[0]
        
        # ワールド移動していたら実行 (同じワールドへのリジョインやインスタンス移動時は実行しない)
        if EnteringWorldId_Old != NowEnteringWorldDict["Id"]:
            button_Save["state"] = "normal"                                     # GUIの保存ボタンを押せるようにする

            # 現在のワールド用のセーブコード保存ファイルのパスを取得してみる。引数がFalseの場合、無かったら""が返ってくる
            NowEnteringWorldSaveCodeFilePath = TryGetSaveCodeFilePathOfNowEnteringWorld(False)
            if "" != NowEnteringWorldSaveCodeFilePath:
                button_Load["state"] = "normal"                                 # GUIの読み込みボタンを押せるようにする
                # セーブコードファイルから一番新しいセーブコードをクリップボードに読み込み
                GetNewestSaveCodeFromFilePathOnClipboard(NowEnteringWorldSaveCodeFilePath)
            else:
                button_Load["state"] = "disable"                                # GUIの読み込みボタンを押せないようにする
                TextView_SystemMessage["text"] = ""                             # GUIのシステムメッセージをクリアする
        
        # 初回ループ時ではなく、Terrors of Nowhereのワールド内にいて自動保存が有効で、セーブコードがログに出力されていたら実行
        if False == FirstLoop and TerrorsOfNowhereWorldId == NowEnteringWorldDict["Id"] and True == AutoSaveInTerrorsOfNowhere and "" != FindedTerrorsOfNowhereSaveCode:
            Push_SaveButton(FindedTerrorsOfNowhereSaveCode)                     # セーブコード保存
            pyperclip.copy(FindedTerrorsOfNowhereSaveCode)                      # セーブコードをクリップボードにコピー
        
        FirstLoop = False
        # LoopEnd---------------------------------------------------------------


# ================================================== 初期化 ====================================================
SoftwareName = u"セーブコード自動読み込みツール"                                  # このソフトの名前
SoftwareVersion = u"v0.1"                                                       # このソフトのバージョン

ExeDir = os.path.dirname(sys.argv[0])                                           # このプログラム(exe)が置かれているディレクトリ取得

LogFileDir = os.path.expanduser("~/AppData/LocalLow/VRChat/VRChat/")            # VRChatのログファイルが保存されるディレクトリ
LogFileSearchWord = "output_log*.txt"                                           # VRChatのログファイル名 検索文字列

SaveFileDir = ""                                                                # プレイヤー別フォルダが保存されるフォルダ
PlayerDir = ""                                                                  # 現在のプレイヤーのワールド別セーブコードファイルが保存されるフォルダ

PlayerInfoDict       = {"VrMode":False, "Name":"", "Id":""}                     # 現在のプレイヤー情報 (VRかどうか、ユーザー名、ユーザーID)
NowEnteringWorldDict = {"Name":"", "Id":""}                                     # 現在のワールド情報 (ワールド名、ワールドID)

SplitWordInSaveCode = "<MySplitWord>"                                           # セーブコードファイル内で「保存日時」と「セーブコード」を区切る文字列

BeforeMediaKeyPushedTime = time.time()                                          # XSOverlay手首メニューの「次の曲」ボタンが前回押された日時

TerrorsOfNowhereWorldId = "wrld_a61cdabe-1218-4287-9ffc-2a4d1414e5bd"           # ゲームワールド TerrorsOfNowhere のワールドID

# サウンドファイルのパスを取得 https://stackoverflow.com/questions/31836104/pyinstaller-and-onefile-how-to-include-an-image-in-the-exe-file
try:
    TempFileDir = sys._MEIPASS
except Exception:
    TempFileDir = os.path.abspath(".")
    
SoundPath_Normal = os.path.join(TempFileDir, "Normal.wav")                      # 処理完了音
SoundPath_Error  = os.path.join(TempFileDir, "Error.wav")                       # 処理警告音
# Download the audio file from Springin’ Sound Stock (https://www.springin.org/sound-stock/)
# ・Normal.wav : Sボタン・システム カテゴリ, 「正解4」
# ・Error.wav  : Sボタン・システム カテゴリ, 「確認1」

# Iconのパスを取得
IconPath = os.path.join(TempFileDir, "VrcGameWorldSaveCodeAutoLoaderIcon.ico")

# XSOverlay共通通知文 https://xiexe.github.io/XSOverlayDocumentation/#/NotificationsAPI?id=xsoverlay-message-object
Message = {
    "messageType" : 1,          # 1 = Notification Popup, 2 = MediaPlayer Information, will be extended later on.
    "index" : 0,                # Only used for Media Player, changes the icon on the wrist.
    "timeout" : 5.0,            # How long the notification will stay on screen for in seconds
    "height" : 100.0,           # Height notification will expand to if it has content other than a title. Default is 175
    "opacity" : 1.0,            # Opacity of the notification, to make it less intrusive. Setting to 0 will set to 1.
    "volume" : 0.0,             # Notification sound volume.
    "audioPath" : "",           # File path to .ogg audio file. Can be "default", "error", or "warning". Notification will be silent if left empty.
    "title" : SoftwareName,     # Notification title, supports Rich Text Formatting
    "useBase64Icon" : False,    # Set to true if using Base64 for the icon image
    "icon" : "default",         # Base64 Encoded image, or file path to image. Can also be "default", "error", or "warning"
    "sourceApp" : "TEST_App"    # Somewhere to put your app name for debugging purposes
    }

# subprocessでコマンド実行したときにコマンドプロンプトのウインドウが表示されないようにする設定 (https://chichimotsu.hateblo.jp/entry/20140712/1405147421)
StartupInfo = subprocess.STARTUPINFO()
StartupInfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
StartupInfo.wShowWindow = subprocess.SW_HIDE

# ============================================= 多重起動してたら終了 =============================================
MyExeName = os.path.basename(sys.argv[0])                                       # 自分のexe名を取得 (拡張子付き)

ProcessHitCount = 0                                                             # 自分を同じ名前のexeがプロセス内に何個あるかカウントする用
for MyProcess in psutil.process_iter():                                         # プロセス一覧取得
    try:
        if MyExeName == os.path.basename(MyProcess.exe()):                      # 自分を同じ名前のexeだったら
            ProcessHitCount = ProcessHitCount + 1                               # カウントアップ
    except:
        pass

# 単一起動時はexeが2つある(なぜかはわからない)。それを超えていたら多重起動しているということなのでここで終了
if 2 < ProcessHitCount:
    sys.exit(0)


# ========================= セーブコードファイルを保存するフォルダを作る。すでにあったらそのまま =====================
SaveFileDir = os.path.join(ExeDir, "SaveCodeFiles")
os.makedirs(SaveFileDir, exist_ok=True)                                         # フォルダを生成。すでにあったらそのまま


# =============================================== スレッド開始 ==================================================
# メディアキー入力監視開始 (無限ループさせてる関数なので、daemon=Trueでデーモン化しないとメインスレッドが終了しても生き残り続けてしまう)
MediakeyLoggingThread = threading.Thread(target=StartMediakeyLoggingThread, daemon=True)
MediakeyLoggingThread.start()

# VRChatが実行されているか監視開始 (無限ループさせてる関数なので、daemon=Trueでデーモン化しないとメインスレッドが終了しても生き残り続けてしまう)
VRChatLoggingThread = threading.Thread(target=MainLoopThread, daemon=True)
VRChatLoggingThread.start()


# ================================================= GUI生成 ====================================================
# ウィンドウ設定 ----------------------------------------------------------------
root = Tk()                                                                     # Tkクラスのインスタンス化
root.title(SoftwareName + " " + SoftwareVersion)                                # タイトル設定
root.geometry("360x222")                                                        # 画面サイズ設定
root.resizable(False, False)                                                    # リサイズ不可に設定
root.configure(bg="#333333")                                                    # 背景色

# ワールド名表示 ----------------------------------------------------------------
TextView_WorldTitle = Label(text=u"現在のワールド", foreground="#CCCCCC", background="#333333", font=("TkDefaultFont", 12, "bold", "underline"))
TextView_WorldTitle.place(x=0, y=5, width=360, height=20)

TextView_WorldName = Label(text=u"-", foreground="#CCCCCC", background="#333333", font=("TkDefaultFont", 18, "bold"))
TextView_WorldName.place(x=0, y=30, width=360, height=20)

# ボタン設定 --------------------------------------------------------------------
button_Load = Button(text=u"手動\n読込", font=("Helvetica", 12, "bold"), state="disable", command=Push_LoadButton)
button_Load.place(x=0, y=60, width=60, height=56)

button_Open = Button(text=u"保存\n場所", font=("Helvetica", 12, "bold"), command=Push_OpenSaveFolderButton)
button_Open.place(x=60, y=60, width=60, height=56)

button_Save = Button(text=u"現在のクリップボードの\n中身を保存", font=("Helvetica", 12, "bold"), state="disable", command=Push_SaveButton)
button_Save.place(x=125, y=60, width=235, height=56)

# 区切り線 ---------------------------------------------------------------------
Separator = Label(text=u"", foreground="#CCCCCC", background="#555555")
Separator.place(x=0, y=122, width=360, height=1)

# システムメッセージ ------------------------------------------------------------
TextView_SystemTitle = Label(text=u"システムメッセージ:", foreground="#CCCCCC", background="#333333", font=("TkDefaultFont", 12, "bold"), anchor=NW)
TextView_SystemTitle.place(x=0, y=125, width=360, height=22)

button_Clear = Button(text=u"クリア", font=("Helvetica", 8, "bold"), command=Push_ClearButton)
button_Clear.place(x=295, y=126.5, width=60, height=20)

TextView_SystemMessage = Label(text=u"", foreground="#CCCCCC", background="#555555", font=("TkDefaultFont", 12, "bold"), anchor=NW, justify="left")
TextView_SystemMessage.place(x=5, y=150, width=350, height=67)

# GUI生成 ----------------------------------------------------------------------
root.iconbitmap(IconPath)                                                       # アイコン設定
root.mainloop()                                                                 # 画面を表示し続ける
