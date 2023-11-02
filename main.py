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


## 生成したTeamsのWebhookURLを変数に格納
TEAMS_WEB_HOOK_URL = "https://kyoceragp.webhook.office.com/webhookb2/1637c1a5-b0e2-4cb4-a46e-0e22f48d27d0@82cc187e-25d5-45e4-8c34-8434bf6075fe/IncomingWebhook/7be33a3145c9405f9da5e014577b85f4/4c26940d-3caf-4724-8f3a-f28d4cf57f1d"

log_directory = "./log/"
log_file_name = time.strftime("%Y%m%d", time.localtime()) + "_log.txt"
log_path = os.path.join(log_directory, log_file_name)
# 稼働灯確認用マイコンIP/PORT
ip = '192.168.98.4'
port = 50000


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

def M5_connect():
    server = (ip, port)
    try:
        # M5に接続
        socket1 = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        socket1.connect(server)
        print('Connection Success')
        sv_connect_status.set('M5 接続成功')
    except Exception as e:
        print('Connection Fail')
        failer(e)
        # Teamsに投稿
        myTeamsMessage = pymsteams.connectorcard(TEAMS_WEB_HOOK_URL)
        myTeamsMessage.title("Error")
        myTeamsMessage.text("Connection Fail")
        myTeamsMessage.send()
        sys.exit(1)


def toggle_button_color(color):
    interval = 500
    global flg_red,flg_green,flg_green
    print('Red/green/Blue: ',flg_red,flg_yellow,flg_green)
    #赤色
    if color == 'red':
        if flg_red == 0:
            red_button["bg"] = "gray"
            return
        elif flg_red == 1:
            red_button["bg"] = "red"
            return
        elif flg_red == 2:
            if red_button["bg"] == "red":
                red_button["bg"] = "gray"
                root.after(interval, lambda:toggle_button_color(color))
                return
            else:
                red_button["bg"] = "red"
                root.after(interval, lambda:toggle_button_color(color))
                return
    #黄色
    if color == 'yellow':
        if flg_yellow == 0:
            yellow_button["bg"] = "gray"
            return
        elif flg_yellow == 1:
            yellow_button["bg"] = "yellow"
            return
        elif flg_yellow == 2:
            if yellow_button["bg"] == "yellow":
                yellow_button["bg"] = "gray"
                root.after(interval, lambda:toggle_button_color(color))
                return
            else:
                yellow_button["bg"] = "yellow"
                root.after(interval, lambda:toggle_button_color(color))
                return
    #緑色
    if color == 'green':
        if flg_green == 0:
            green_button["bg"] = "gray"
            return
        elif flg_green == 1:
            green_button["bg"] = "green"
            return
        elif flg_green == 2:
            if green_button["bg"] == "green":
                green_button["bg"] = "gray"
                root.after(interval, lambda:toggle_button_color(color))
                return
            else:
                green_button["bg"] = "green"
                root.after(interval, lambda:toggle_button_color(color))
                return
    

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


def switch_mode(color):
    global flg_red
    global flg_yellow
    global flg_green
    global sv_combined_name
    global sv_mail
    
    Name = sv_combined_name.get()
    Mail = sv_mail.get()
    print('--switch_mode--')
    print('Red/Yellow/Green: ',flg_red,flg_yellow,flg_green)
    if color == 'red':
        if flg_red == 0:
            toggle_button_red["text"] = "点灯"
            flg_red = 1
            toggle_button_color('red')
            Message = '赤色点灯'
            # send_message(Name,Mail,Message)
            status_check(color)
            return
        elif flg_red == 1:
            toggle_button_red["text"] = "点滅"
            flg_red = 2
            toggle_button_color('red')    
            Message = '赤色点滅'
            # send_message(Name,Mail,Message)
            status_check(color)
            return
        elif flg_red == 2:
            toggle_button_red["text"] = "消灯"
            flg_red = 0
            toggle_button_color('red')
            Message = '赤色消灯'
            # send_message(Name,Mail,Message)
            status_check(color)
            return
    if color == 'yellow':
        if flg_yellow == 0:
            toggle_button_yellow["text"] = "点灯"
            flg_yellow = 1
            toggle_button_color('yellow')
            Message = '黄色点灯'
            # send_message(Name,Mail,Message)
            status_check(color)
            return
        elif flg_yellow == 1:
            toggle_button_yellow["text"] = "点滅"
            flg_yellow = 2
            toggle_button_color('yellow')    
            Message = '黄色点滅'
            # send_message(Name,Mail,Message)
            status_check(color)
            return
        elif flg_yellow == 2:
            toggle_button_yellow["text"] = "消灯"
            flg_yellow = 0
            toggle_button_color('yellow')
            Message = '黄色消灯'
            # send_message(Name,Mail,Message)
            status_check(color)
            return
    if color == 'green':
        if flg_green == 0:
            toggle_button_green["text"] = "点灯"
            flg_green = 1
            toggle_button_color('green')
            Message = '緑色点灯'
            # send_message(Name,Mail,Message)
            status_check(color)
            return
        elif flg_green == 1:
            toggle_button_green["text"] = "点滅"
            flg_green = 2
            toggle_button_color('green')    
            Message = '緑色点滅'
            # send_message(Name,Mail,Message)
            status_check(color)
            return
        elif flg_green == 2:
            toggle_button_green["text"] = "消灯"
            flg_green = 0
            toggle_button_color('green')
            Message = '緑色消灯'
            # send_message(Name,Mail,Message)
            status_check(color)
            return

def status_check(color):
    global flg_red,flg_yellow,flg_green,sv_status_1,sv_status_2
    ## sv_status_1
    if color == 'yellow':
        if flg_yellow == 1:
            sv_status_1.set('手動中')
        elif flg_yellow == 2 and flg_green == 0:
            sv_status_1.set('原点復帰中')
        if flg_green == 1 and flg_yellow == 2:
            sv_status_1.set('')
            sv_status_2.set('不足/満杯')    
        if flg_yellow == 0:
            sv_status_1.set('')
    elif color == 'green':
        if flg_green == 1:
            sv_status_1.set('自動運転中')
        elif flg_green == 2:
            sv_status_1.set('自動停止中')
        if flg_green == 0:
            sv_status_1.set('')
    elif color == 'red':
        ## sv_status_2
        if flg_red == 1:
            sv_status_2.set('異常')
        elif flg_red == 2:
            sv_status_2.set('異常解除待')
        if flg_red == 0:
            sv_status_2.set('')

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

toggle_button_red = tk.Button(frame_status, text="R", command=lambda: switch_mode('red'))
toggle_button_red.place(x=200, y=80)
toggle_button_yellow = tk.Button(frame_status, text="Y", command=lambda: switch_mode('yellow'))
toggle_button_yellow.place(x=200, y=240)
toggle_button_green = tk.Button(frame_status, text="G", command=lambda: switch_mode('green'))
toggle_button_green.place(x=200, y=400)



root.mainloop()



# %%
