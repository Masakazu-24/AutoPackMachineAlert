# %%
import tkinter as tk
from tkinter import ttk
import tkinter.font as font
import tkinter.messagebox as tkm
import pandas as pd
from pathlib import Path
import pymsteams
import sys
import time
import socket
import os
import datetime
import threading

## 生成したTeamsのWebhookURLを変数に格納
TEAMS_WEB_HOOK_URL = "https://kyoceragp.webhook.office.com/webhookb2/1637c1a5-b0e2-4cb4-a46e-0e22f48d27d0@82cc187e-25d5-45e4-8c34-8434bf6075fe/IncomingWebhook/7be33a3145c9405f9da5e014577b85f4/4c26940d-3caf-4724-8f3a-f28d4cf57f1d"

log_directory = "./log/"
log_file_name = time.strftime("%Y%m%d", time.localtime()) + "_log.txt"
log_path = os.path.join(log_directory, log_file_name)
# 稼働灯確認用マイコンIP/PORT
ip = '192.168.98.4'
port = 50000
socket1 = None


#エラー処理
def failer(e):
    exc_type, exc_obj, tb = sys.exc_info()
    lineno = tb.tb_lineno
    global Err
    # Teamsに投稿
    myTeamsMessage = pymsteams.connectorcard(TEAMS_WEB_HOOK_URL)
    myTeamsMessage.title("Error")
    myTeamsMessage.text("  (line" + str(lineno) + "):" + str(e.args) + "\n")
    myTeamsMessage.send()
    if not os.path.exists(log_directory):
        os.makedirs(log_directory)
    if not os.path.exists(log_path):
        with open(log_path, 'w') as log_file:
            log_file.write(str(datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S"))+" Logfile_created")
    #ログ出力
    with open(log_path,"a",encoding="utf-8") as f:
        f.write("\n"+str(datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")) + "  (line" + str(lineno) + "):" + str(e.args) )

def send_message(Name,Mail,Message): #Nameはローマ字の頭大文字で名+姓の順　間半角スペース'Masakazu Nishi'の形式
    global sv_name,sv_mail
    if sv_name.get() and sv_mail.get():
        teams_message = pymsteams.connectorcard(TEAMS_WEB_HOOK_URL)
        teams_message.payload = {
            "type": "message",
            "attachments": [
                {
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": {
                    "type": "AdaptiveCard",
                    "body": [
                        {
                            "type": "TextBlock",
                            "text": "<at>"+str(Name)+"</at>\r \n"+str(Message)
                        }
                    ],
                    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                    "version": "1.0",
                    "msteams": {
                        "entities": [
                            {
                                "type": "mention",
                                "text": "<at>"+ str(Name)+ "</at>",
                                "mentioned": {
                                    "id": str(Mail),
                                    "name": str(Name)
                                }
                            }
                        ]
                    }
                }
            }]
        }
        teams_message.send()
    else:
        print('Teams messege canceled')

def M5_connect():
    global socket1
    server = (ip, port)
    try:
        # M5に接続
        socket1 = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        socket1.connect(server)
        print('Connection Success')
        sv_connect_status.set('M5 接続成功')
    except Exception as e:
        print('M5 Connection Fail')
        failer(e)
        # Teamsに投稿
        myTeamsMessage = pymsteams.connectorcard(TEAMS_WEB_HOOK_URL)
        myTeamsMessage.title("Error")
        myTeamsMessage.text("M5 Connection Fail")
        myTeamsMessage.send()
        sys.exit(1)

def send_receve():
    global flg_red,flg_yellow,flg_green,socket1
    while True:
        print('---send_receve---\n')
        # サーバにコマンド送信
        command = 'GET\r'
        socket1.send(command.encode("UTF-8"))
        dt = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        # サーバから受信
        recieve = socket1.recv(4096).decode()
        recieve = recieve.replace('\r\n','').split(',')
        # print('recieve:\n',recieve)
        flg_red = int(recieve[3])
        flg_yellow = int(recieve[6])
        flg_green = int(recieve[9])

        print('From M5 Red/Yellow/Green: ',flg_red,flg_yellow,flg_green)
        time.sleep(3)


def toggle_button_color():
    global flg_red,flg_green,flg_green,socket1,sv_combined_name,sv_mail
    global sv_status_1,sv_status_1_bk,sv_status_2,sv_status_2_bk,flg_status_1,flg_status_2
    while True:
        print('---toggle_button_color---\n')
        # Red
        if flg_red == 0:
            red_button["bg"] = "gray"
        elif flg_red == 1:
            red_button["bg"] = "red"
        elif flg_red == 2:
            if red_button["bg"] == "red":
                red_button["bg"] = "gray"
            else:
                red_button["bg"] = "red"
        # Yellow
        if flg_yellow == 0:
            yellow_button["bg"] = "gray"
        elif flg_yellow == 1:
            yellow_button["bg"] = "yellow"
        elif flg_yellow == 2:
            if yellow_button["bg"] == "yellow":
                yellow_button["bg"] = "gray"
            else:
                yellow_button["bg"] = "yellow"
        # Green
        if flg_green == 0:
            green_button["bg"] = "gray"
        elif flg_green == 1:
            green_button["bg"] = "green"
        elif flg_green == 2:
            if green_button["bg"] == "green":
                green_button["bg"] = "gray"
            else:
                green_button["bg"] = "green"
        # フラグの値を文字列に変換
        flg_red_str = int(flg_red)
        flg_yellow_str = int(flg_yellow)
        flg_green_str = int(flg_green)
        # ステータスを取得
        status_1 = str(df_status_pattern[(df_status_pattern['flg_red'] == flg_red_str) &
                                (df_status_pattern['flg_yellow'] == flg_yellow_str) &
                                (df_status_pattern['flg_green'] == flg_green_str)]['status_1'].values[0])

        status_2 = str(df_status_pattern[(df_status_pattern['flg_red'] == flg_red_str) &
                                (df_status_pattern['flg_yellow'] == flg_yellow_str) &
                                (df_status_pattern['flg_green'] == flg_green_str)]['status_2'].values[0])
        sv_status_1.set(status_1)
        sv_status_2.set(status_2)
        print('status_1:',status_1)
        print('status_2:',status_2)
        if str(sv_status_1.get()) == 'nan':
            sv_status_1.set('')
        if str(sv_status_2.get()) == 'nan':
            sv_status_2.set('')
        if str(sv_status_1.get()) == str(sv_status_1_bk.get()):
            flg_status_1 = 0
        else:
            flg_status_1 = 1
            sv_status_1_bk.set(sv_status_1.get())
        if str(sv_status_2.get()) == str(sv_status_2_bk.get()):
            flg_status_1 = 0
        else:
            flg_status_1 = 1
            sv_status_2_bk.set(sv_status_2.get())
        
        if flg_status_1 == 1 or flg_status_2 == 1:
            print('flg_status_1:',flg_status_1)
            print('flg_status_2:',flg_status_2)

            Message = ''
            if sv_status_1.get():
                Message += str(sv_status_1.get())
            if sv_status_2.get():
                if sv_status_1.get():
                    Message += ' & '
                Message += str(sv_status_2.get())
            if len(Message) > 0:
                Name = sv_name.get()
                Mail = sv_combined_name.get()
                send_message(Name,Mail,Message)
        time.sleep(1)

def focus_out_format(event):
    global sv_name
    global sv_family_name
    global sv_mail
    global sv_combined_name
    print('focus_out_format')
    name = None
    family_name = None
    mark = None
    try:
        name, family_name, mark = event.widget.get().strip().split('.') # 入力された文字列に空白が含まれる場合は削除
    except:
        tkm.showerror(message='メールアドレスに不正な文字列が入力されています',title='Error')
    if name:#空白場合はスルー
        result = name[0].upper() + name[1:].lower()
        sv_name.set(result)
        print('name:',sv_name.get())
    if family_name:
        result = family_name[0].upper() + family_name[1:].lower()
        sv_family_name.set(result)
        print('family_name:',sv_family_name.get())
    if name and family_name:
        sv_combined_name.set(sv_name.get()+" "+sv_family_name.get())
        sv_mail.set(str(event.widget.get().strip())+"@kyocera.jp")

def send_test_message():
    global sv_combined_name
    global sv_name
    global sv_family_name
    global sv_mail
    Name = sv_combined_name.get()
    if Name == "":
        sv_mail.set(txt_mail_account.get().strip()+'@kyocera.jp')
        first_name, family_name, mark = txt_mail_account.get().strip().split('.')
        first_name = first_name[0].upper() + first_name[1:].lower()
        family_name = family_name[0].upper() + family_name[1:].lower()
        sv_name.set(str(first_name))
        sv_family_name.set(str(family_name))
        sv_combined_name.set(str(first_name)+' '+str(family_name))
        Name = sv_combined_name.get()
    Mail = sv_mail.get()
    print('--send_test_message--')
    print('Name:',Name)
    print('Mail:',Mail)
    Message = 'テスト送信'
    send_message(Name,Mail,Message)
    tkm.showinfo(title='確認',message='Teamsでメッセージを送信しました。\n受信できない場合は入力したメールアドレスが正しいか確認してください。')

def clear_all():
    global sv_combined_name
    global sv_mail
    global sv_name
    global sv_family_name
    sv_combined_name.set('')
    sv_mail.set('')
    sv_name.set('')
    sv_family_name.set('')
    txt_mail_account.delete(0,'end')

#常に最前面に表示
root = tk.Tk()
root.geometry("1280x800+0+0")
root.title('自動梱包機')
frame_settings = tk.LabelFrame(root, height=400, width=300, pady=0, padx=0, borderwidth=2, relief="groove",text='設定')
frame_settings.place(x=20, y=20)
frame_status = tk.LabelFrame(root, height=600, width=800, pady=0, padx=0, borderwidth=2, relief="groove",text='マシン状態')
frame_status.place(x=400, y=20)

## フラグ類-----------------------------------------------------
# メンション先の名前とメールアドレス
sv_name = tk.StringVar() #ローマ字　'Masakazu'の形式
sv_family_name = tk.StringVar() #ローマ字　'Nishi'の形式
sv_combined_name = tk.StringVar()
sv_mail = tk.StringVar() #office365メールアドレス
sv_connect_status = tk.StringVar() #M5の接続状態

# 稼働灯フラグ
flg_green = 0
flg_yellow = 0
flg_red = 0
buzzer_1 = 0
buzzer_2 = 0

# マシン状態
sv_status_1 = tk.StringVar() #手動中/原点復帰中/自動運転中/自動停止中
sv_status_2 = tk.StringVar() #異常中/異常解除待/不足‐満杯
sv_status_1_bk = tk.StringVar() #変化確認
sv_status_2_bk = tk.StringVar() #変化確認
flg_status_1 = 0 #変化確認
flg_status_2 = 0 #変化確認
df_status_pattern = pd.read_csv('Status_pattern.csv',encoding='utf-8')
# ----------------------------------------------------------------


lbl_mail = tk.Label(frame_settings,text="ﾒｰﾙｱﾄﾞﾚｽ",font=("",10),anchor=tk.E,width=10)
lbl_mail.place(x=10, y=0)
txt_mail_account = tk.Entry(frame_settings,width=20)
txt_mail_account.place(x=90, y=0)
txt_mail_account.bind("<FocusOut>",focus_out_format)
lbl_domain = tk.Label(frame_settings,text="@kyocera.jp",font=("",10),anchor=tk.E,width=10)
lbl_domain.place(x=200, y=0)
lbl_name = tk.Label(frame_settings,text="名前",font=("",10),anchor=tk.E,width=10)
lbl_name.place(x=10, y=40)
lbl_first_name = tk.Label(frame_settings,width=10,textvariable=sv_name,relief="sunken")
lbl_first_name.place(x=90, y=40)
lbl_family_name = tk.Label(frame_settings,width=10,textvariable=sv_family_name,relief="sunken")
lbl_family_name.place(x=180, y=40)

test_button = tk.Button(frame_settings, text="テスト送信",command=send_test_message)
test_button.place(x=100, y=80)
clear_button = tk.Button(frame_settings, text="クリア",command=clear_all)
clear_button.place(x=20, y=80)


# マシン状態配置
sv_connect_status.set('M5 stack 未接続')
lbl_connect_status = tk.Label(frame_status,width=20,textvariable=sv_connect_status,relief="groove")
lbl_connect_status.place(x=20, y=20)

lbl_machine_status1 = tk.Label(frame_status,width=20,textvariable=sv_status_1,relief="groove")
lbl_machine_status1.place(x=20, y=40)
lbl_machine_status2 = tk.Label(frame_status,width=20,textvariable=sv_status_2,relief="groove")
lbl_machine_status2.place(x=20, y=60)

red_button = tk.Button(frame_status, text="R", bg="gray",width=20, height=10)
red_button.place(x=20, y=80)
yellow_button = tk.Button(frame_status, text="Y", bg="gray",width=20, height=10)
yellow_button.place(x=20, y=240)
green_button = tk.Button(frame_status, text="G", bg="gray",width=20, height=10)
green_button.place(x=20, y=400)

# toggle_button_red = tk.Button(frame_status, text="監視スタート", command=toggle_button_color)
# toggle_button_red.place(x=200, y=80)

M5_connect()
# スレッドを作成し、送受信とボタンの色変更を並列で実行
send_receive_thread = threading.Thread(target=send_receve)
color_change_thread = threading.Thread(target=toggle_button_color)

# スレッドを開始
send_receive_thread.start()
color_change_thread.start()


root.mainloop()



# %%
