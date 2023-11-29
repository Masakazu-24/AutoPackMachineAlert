# %%
import tkinter as tk
from tkinter import ttk
import tkinter.messagebox as tkm
import pandas as pd
import pymsteams
import sys
import time
import socket
import os
import datetime
import threading
import subprocess

## 生成したTeamsのWebhookURLを変数に格納
# 西テスト用
# TEAMS_WEB_HOOK_URL = "https://kyoceragp.webhook.office.com/webhookb2/1637c1a5-b0e2-4cb4-a46e-0e22f48d27d0@82cc187e-25d5-45e4-8c34-8434bf6075fe/IncomingWebhook/7be33a3145c9405f9da5e014577b85f4/4c26940d-3caf-4724-8f3a-f28d4cf57f1d"
# 本番用
TEAMS_WEB_HOOK_URL = "https://kyoceragp.webhook.office.com/webhookb2/9ec90966-965d-483a-85a1-eb37475e942a@82cc187e-25d5-45e4-8c34-8434bf6075fe/IncomingWebhook/3aafa31a7d844978acde270e561dedb1/4c26940d-3caf-4724-8f3a-f28d4cf57f1d"
log_directory = "./log/"
log_file_name = time.strftime("%Y%m%d", time.localtime()) + "_log.txt"
log_path = os.path.join(log_directory, log_file_name)
# 稼働灯確認用マイコンIP/PORT
ip = '192.168.98.4'
port = 50000
socket1 = None
# Wi-Fi接続用batファイル
bat_file_path = r'wifi_connect.bat'


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
    sys.exit()

def send_message(Name,Mail,Message): #Nameはローマ字の頭大文字で名+姓の順　間半角スペース'Masakazu Nishi'の形式
    global sv_mail
    if sv_mail.get():
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
    max_retries = 10
    try:
        # M5に接続
        socket1 = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        socket1.connect(server)
        print('Connection Success')
        sv_connect_status.set('M5 接続成功')
    except ConnectionRefusedError as e:
        print('ConnectionRefusedError')
        tkm.showerror('ConnectionRefusedError','M5接続失敗')
        failer(e)
    except Exception as e:
        print('M5 Connection Fail')
        result = subprocess.run([bat_file_path],capture_output=True,text=True)
        print(f"標準出力：\n{result.stdout}")
        print(f"標準エラー：\n{result.stderr}")
        print(f"終了コード：\n{result.returncode}")
        for _ in range(max_retries):
            with socket.socket(socket.AF_INET,socket.SOCK_STREAM) as s:
                result = s.connect_ex((ip,port))
                if result == 0:
                    print('Connection successfully!!')
                    break
                else:
                    print('Connection failed. Retrying...')
                    time.sleep(1)
        else:
            print('Exceeded maximum retries. Connection failed')
        try:
            # M5に接続
            socket1 = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            socket1.connect(server)
            print('Connection Success at 2nd try')
            sv_connect_status.set('M5 接続成功')
        except Exception as e:
            tkm.showerror('Error','M5接続失敗\n\n稼働灯監視用マイコンへの接続に失敗しました。\n以下の設定を確認してください：\n\n1.Wi-Fi接続設定を開き、Wi-Fi 2がAutoPackingMachineに接続されていることを確認してください。接続されていない場合は接続し、再度プログラムを立ち上げてください。\n\n2.ブラウザを開き、192.168.98.4にアクセスします。アクセスができない場合はマイコンの再起動を行います。稼働灯下にある本体のディスプレイとサイドボタンを長押しし、画面が一度消えて再度表示されたことを確認したら、1のWi-Fi接続を確認後に再度プログラムを立ち上げてください。')
            # Teamsに投稿
            myTeamsMessage = pymsteams.connectorcard(TEAMS_WEB_HOOK_URL)
            myTeamsMessage.title("Error")
            myTeamsMessage.text("M5 Connection Fail")
            myTeamsMessage.send()
            failer(e)
            sys.exit(1)

def send_receve(exit_signal):
    global flg_red,flg_yellow,flg_green,socket1,sv_connect_status
    if sv_connect_status.get() == 'M5 接続成功':
        while not exit_signal.is_set():
            try:
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
                time.sleep(10)
            except ConnectionAbortedError:
                print('ConnectionAbortedError')
            except Exception as e:
                failer(e)
    else:
        print('Cannot send to M5\n')


def toggle_button_color(exit_signal):
    global flg_red,flg_green,flg_green,socket1,sv_combined_name,sv_mail
    global sv_status_1,sv_status_1_bk,sv_status_2,sv_status_2_bk,flg_status_1,flg_status_2,sv_connect_status
    if sv_connect_status.get() == 'M5 接続成功':
        while not exit_signal.is_set():
            try:
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
                alert_status = int(df_status_pattern[(df_status_pattern['flg_red'] == flg_red_str) &
                                        (df_status_pattern['flg_yellow'] == flg_yellow_str) &
                                        (df_status_pattern['flg_green'] == flg_green_str)]['alert_flg'].values[0])
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
                    flg_status_2 = 0
                else:
                    flg_status_2 = 1
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
                    if len(Message) > 0 and alert_status ==1:
                        Name = sv_combined_name.get()
                        Mail = sv_mail.get()
                        send_message(Name,Mail,Message)
            except Exception as e:
                failer(e)
            time.sleep(1)
    else:
        print('Cannot run toggle_button_color\n')

def send_test_message():
    global sv_name_JP,sv_combined_name,sv_mail
    if sv_name_JP.get() != "":
        Mail = sv_mail.get()
        Name = sv_combined_name.get()
        print('--send_test_message--')
        print('Name:',Name)
        print('Mail:',Mail)
        Message = 'テスト送信'
        send_message(Name,Mail,Message)
        tkm.showinfo(title='確認',message='Teamsでメッセージを送信しました。\n受信できない場合は入力したメールアドレスが正しいか確認してください。')

def clear_all():
    global sv_combined_name,sv_mail,sv_mail_head,sv_name,sv_name_JP,sv_family_name
    sv_combined_name.set('')
    sv_mail.set('')
    sv_mail_head.set('')
    sv_name.set('')
    sv_name_JP.set('')
    sv_family_name.set('')
  
exit_signal = threading.Event()
def close_window():
    print('---close_window---\n')
    sub_win = tk.Toplevel()
    sub_win.geometry('100x100')
    sub_win.attributes("-topmost",True)
    label = tk.Label(sub_win, text="停止中です・・・")
    label.place(x=10,y=10)
    sub_win.update()
    exit_signal.set() #並列処理停止フラグ
    try:
        socket1.close() #M5接続解除
    except:
        print('Fail to socket1 close()')
    send_receive_thread.join()
    print('send_recieve joined')
    sub_win.destroy()
    color_change_thread.join()
    print('color_change joined')
    root.destroy()
    root.quit()

# チェックボックスの状態を更新してStatus_pattern.csvを保存
def update_status_pattern():
    # print(status_checkboxes)
    for i in range(len(status_checkboxes)):
        # print('i in update_status_patten:: ',i)
        row = df_status_pattern.iloc[i]
        alert_flg = status_checkboxes[i].get()
        df_status_pattern.at[i, 'alert_flg'] = alert_flg
    df_status_pattern.to_csv('Status_pattern.csv', index=False)
    tkm.showinfo("Info",'設定が保存されました')
    open_sub_window()

# サブウィンドウを作成
def open_sub_window():
    global status_checkboxes, sub_window_exist, sub_window 
    if not sub_window_exist:
        sub_window_exist = True
        sub_window = tk.Toplevel(root)
        sub_window.title("Teams通知を行うタイミングにチェックを入れてください")
        sub_window.geometry("800x800")
        def on_close():
            global sub_window_exist
            sub_window_exist = False
            sub_window.destroy()
        sub_window.protocol("WM_DELETE_WINDOW", on_close)

        status_checkboxes = []
        status_vars = []

        # チェックボックスを作成し、Status_pattern.csvからalert_flgを読み込む
        for i, row in df_status_pattern.iterrows():
            flg_red = row['flg_red']
            flg_yellow = row['flg_yellow']
            flg_green = row['flg_green']
            status_1 = row['status_1']
            status_2 = row['status_2']
            tmp_txt = f'Pattern {i+1} - (Red: {flg_red}, Yellow: {flg_yellow}, Green: {flg_green}, Status1: {status_1}, Status2: {status_2})'
            exec(f"var_{i}=tk.IntVar()")
            alert_flg = df_status_pattern.at[i,'alert_flg']
            status_vars.append(alert_flg)
            eval(f"var_{i}.set({alert_flg})")
            eval(f"status_checkboxes.append(var_{i})")
            exec(f"checkbox_{i} = tk.Checkbutton(sub_window, text='{tmp_txt}', variable=var_{i}, anchor='w', justify='left')")
            eval(f"checkbox_{i}.pack(fill='x')")
            if alert_flg == 1:
                exec(f"checkbox_{i}.select()")
                print(f"checkbox_{i} selected")
        print('\nstatus_checkboxes in def open_sub_window:\n',status_checkboxes)
        print('\nstatus_vars in def open_sub_window:\n',status_vars)
        save_button = tk.Button(sub_window, text="Save", command=update_status_pattern)
        save_button.pack()
    else:
        sub_window.lift()

def on_name_selected(event):
    selected_name = sv_name_JP.get()
    selected_email = name_email_df.loc[name_email_df['Name'] == selected_name, 'mail_head'].values[0]
    print(selected_email)
    sv_mail_head.set(selected_email)
    sv_mail.set(sv_mail_head.get()+str('@kyocera.jp'))
    first,family,tag = selected_email.split('.')
    sv_name.set(first.capitalize())
    sv_family_name.set(family.capitalize())
    sv_combined_name.set(sv_name.get()+' '+sv_family_name.get())

#常に最前面に表示
root = tk.Tk()
root.geometry("1280x800+0+0")
root.title('自動梱包機')
frame_settings = tk.LabelFrame(root, height=400, width=350, pady=0, padx=0, borderwidth=2, relief="groove",text='設定')
frame_settings.place(x=20, y=20)
frame_status = tk.LabelFrame(root, height=600, width=800, pady=0, padx=0, borderwidth=2, relief="groove",text='マシン状態')
frame_status.place(x=400, y=20)

## フラグ類-----------------------------------------------------
# メンション先の名前とメールアドレス
sv_name_JP = tk.StringVar() #漢字氏名
sv_name = tk.StringVar() #ローマ字　'Masakazu'の形式
sv_family_name = tk.StringVar() #ローマ字　'Nishi'の形式
sv_combined_name = tk.StringVar()
sv_mail_head = tk.StringVar() #メールユーザー名（@の前）
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

# Teams通知タイミング設定のフラグ類
status_checkboxes = []
status_vars = []
sub_window_exist = False
# ----------------------------------------------------------------


# 名前とメールアドレスを読み込む
name_email_df = pd.read_csv('Name_email.csv',encoding='utf-8')

# 名前のドロップダウン
name_label = tk.Label(frame_settings, text="氏名:")
name_label.place(x=10, y=0)
name_dropdown = ttk.Combobox(frame_settings, textvariable=sv_name_JP, values=name_email_df['Name'].tolist())
name_dropdown.place(x=100, y=0)
name_dropdown.bind("<<ComboboxSelected>>", on_name_selected)

# 選択された名前に対応するメールアドレスの表示
email_label = tk.Label(frame_settings, text="メールアドレス:")
email_label.place(x=10, y=40)
email_entry = tk.Entry(frame_settings, textvariable=sv_mail_head, state="readonly")
email_entry.place(x=100, y=40)
lbl_domain = tk.Label(frame_settings,text="@kyocera.jp",font=("",10),anchor=tk.E,width=10)
lbl_domain.place(x=220, y=40)

lbl_first_name = tk.Label(frame_settings,width=10,textvariable=sv_name,relief="sunken")
lbl_first_name.place(x=100, y=80)
lbl_family_name = tk.Label(frame_settings,width=10,textvariable=sv_family_name,relief="sunken")
lbl_family_name.place(x=180, y=80)
test_button = tk.Button(frame_settings, text="テスト送信",command=send_test_message, bg='#00ff7f')
test_button.place(x=150, y=120)
clear_button = tk.Button(frame_settings, text="クリア",command=clear_all)
clear_button.place(x=100, y=120)
# サブウィンドウを開くボタンを作成
open_sub_window_button = tk.Button(frame_settings, text="Teams通知設定", command=open_sub_window)
open_sub_window_button.place(x=100, y=200)


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

M5_connect()
# スレッドを作成し、送受信とボタンの色変更を並列で実行
send_receive_thread = threading.Thread(target=send_receve,args=(exit_signal,))
color_change_thread = threading.Thread(target=toggle_button_color,args=(exit_signal,))

# スレッドを開始
send_receive_thread.start()
color_change_thread.start()

# ウィンドウを閉じるイベントの設定
root.protocol("WM_DELETE_WINDOW", close_window)
root.mainloop()



# %%
