import threading
import time
import re
import asyncio
import discord  #discord_bot
from discord.ext import commands #discord_bot
from datetime import datetime  # 時刻
import customtkinter as ctk  # UI
import pandas as pd  # Excel保存
from playsound import playsound  # MP3再生
from tkinter import messagebox, filedialog
import serial  # シリアル通信ライブラリ

# Matplotlib関連
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# ----- Discord Bot の設定 -----
# Bot のプレフィックスは "!" に設定（任意）
intents = discord.Intents.default()
discord_bot = commands.Bot(command_prefix="!", intents=intents)
bot_loop = None

# 通知を送信するチャンネルID
# 自身が使用するbotを入れたサーバーのメッセージチャットのチャンネルID
CHANNEL_ID = 1234567890123456789


# グローバル変数
data = []  # Excel保存用の測定データリスト
EXCEL_FILE = None  # 保存ファイル名

def send_command(ser, command):
    """
    シリアル通信でコマンドを送信し、応答を取得する。
    """
    ser.write((command + "\r").encode("ascii"))
    print(f"送信: {command}")
    time.sleep(0.5)
    response = ser.readline().decode("ascii").strip()
    print(f"受信: {response}")
    return response

def alarm_sound(n):
    """アラーム音をn回鳴らす"""
    for _ in range(n):
        playsound("alarm.mp3")
        time.sleep(0.5)

@discord_bot.event
async def on_ready():
    global bot_loop
    bot_loop = discord_bot.loop
    print(f"Discord Bot: Logged in as {discord_bot.user}")

async def send_message(message):
    channel = discord_bot.get_channel(CHANNEL_ID)
    if channel:
        await channel.send(message)
    else:
        print("Discord Bot: 指定したチャンネルが見つかりません")

def start_discord_bot():
    # Bot を実行
    discord_bot.run("MTM0NDI4NDcyNjc3MjQ5ODQ5NA.GrZpSQ.MInm6ETryrm9Pcl8kN3GYvflED2-oGhFd_PPZU") # 自身が作成したbotのトークンを記入

def send_discord_notification(message):
    try:
        # メッセージをUTF-8エンコード→デコードして、エラーがあれば置換する
        message = message.encode('utf-8', errors='replace').decode('utf-8')
        future = asyncio.run_coroutine_threadsafe(send_message(message), bot_loop)
        future.result(timeout=5)
    except Exception as e:
        print("Discord通知エラー:", e)



class MeasurementApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("測定データ収集システム")
        self.geometry("1200x700")
        
        # 各センサのCOMポート設定
        self.com_port1 = "COM15"    # ピラニ1[Pa]
        self.com_port2 = "COM13"    # ピラニ2[Pa]
        self.com_port3 = "COM16"    # 電離真空計[Pa]
        self.com_port4 = "COM14"   # 熱電対
        self.com_port5 = "COM12"   # ヒーター電圧

        self.record_interval = 10  # 記録間隔（秒）
        self.original_record_interval = 10
        self.room_temperature = 26.7  # 初期室温（℃）
        self.measurement_running = False

        # 温度通知閾値（設定可能）
        self.target_temperature = 200  # 目標温度
        self.danger_temperature = 250  # 危険温度
        self.limit_temperature  = 300  # 限界温度

        # 測定データ保持用リスト（測定開始からの経過秒）
        self.time_data = []         # [s]
        self.pirani1_data = []      # ピラニ1 [Pa]
        self.pirani2_data = []      # ピラニ2 [Pa]
        self.ion_data = []          # 電離真空計 [Pa]
        self.thermocouple_data = [] # 熱電対（温度, ℃）
        self.heater_data = []       # ヒーター電圧 [V]
        self.ion_data2 = []         # 基板温度測定開始語の電離真空計 [Pa]

        # イベントマーカーリスト（タプル: (記録時刻, "イベント名", "色")）
        self.event_markers = []

        # 各センサ用シリアル接続
        self.pirani1_ser = None
        self.pirani2_ser = None
        self.ion_ser = None
        self.thermocouple_ser = None
        self.heater_ser = None

        # ログウィンドウ用
        self.log_window = None
        self.log_textbox = None

        # 内部フラグ
        self.show_substrate_graphs = False  # 基板温度測定開始後に表示するグラフ群
        self.vapor_events = []              # 蒸着開始／終了の目印用 (time, color)

        # アラーム通知用状態
        self.notif_200_triggered = False
        self.notif_250_triggered = False
        self.last_300_notif_time = None
        self.ion_notify_heater = False  # ヒーター起動可能通知
        self.ion_notify_vapor = False   # 蒸着可能通知
        self.volt_ten = False

        self.heater_increase_flag = {10: False, 20: False, 30: False, 40: False}
        self.heater_increase_timestamp = {10: None, 20: None, 30: None, 40: None}
        # ヒーター電圧下げ通知用：基板温度測定終了後に、当日の最大値からの下がりを管理
        self.basis_ended = False
        self.last_decrease_notif_voltage = None
        self.last_decrease_notif_timestamp = None

        self.create_layout()
        self.update_current_time()
        self.update_graphs()

    def create_layout(self):
        # ヘッダー領域（上部）
        self.header_frame = ctk.CTkFrame(self, height=20)
        self.header_frame.pack(side="top", fill="x")
        self.current_time_label = ctk.CTkLabel(self.header_frame, text="", font=("Arial", 14))
        self.current_time_label.pack(side="right", padx=10, pady=5)
        # 通知テスト用ボタン
        ctk.CTkButton(self.header_frame, text="通知テスト", command=lambda: send_discord_notification("テスト通知です"), fg_color="#7289DA", hover_color="#304FB5").pack(side="left",pady=5,padx=10)

        # コンテンツ領域
        self.content_frame = ctk.CTkFrame(self)
        self.content_frame.pack(side="top", fill="both", expand=True)

        # 左側：操作パネル
        self.left_frame = ctk.CTkFrame(self.content_frame, width=200)
        self.left_frame.pack(side="left", fill="y", padx=10, pady=10)

        ctk.CTkButton(self.left_frame, text="測定開始", command=self.start_measurement, fg_color="#008000", hover_color="#003300").pack(pady=10)
        ctk.CTkButton(self.left_frame, text="基板温度測定開始", command=self.start_basis, fg_color="#008000", hover_color="#003300").pack(pady=10)
        ctk.CTkButton(self.left_frame, text="蒸着開始", command=self.start_vapor_deposition, fg_color="#008000", hover_color="#003300").pack(pady=10)
        ctk.CTkButton(self.left_frame, text="基板温度測定終了", command=self.end_basis, fg_color="#DB0016", hover_color="#A30013").pack(pady=10)
        ctk.CTkButton(self.left_frame, text="蒸着終了", command=self.end_vapor_deposition, fg_color="#DB0016", hover_color="#A30013").pack(pady=10)
        ctk.CTkButton(self.left_frame, text="測定終了", command=self.end_measurement, fg_color="#DB0016", hover_color="#A30013").pack(pady=10)
        ctk.CTkButton(self.left_frame, text="Excel保存", command=self.date_keep, fg_color="#FFDD00", hover_color="#CCB000", text_color="black").pack(pady=10)
        ctk.CTkButton(self.left_frame, text="グラフ保存", command=self.save_graph_images, fg_color="#FFDD00", hover_color="#CCB000", text_color="black").pack(pady=10)
        ctk.CTkButton(self.left_frame, text="ログ表示", command=self.open_log_window).pack(pady=10)
        ctk.CTkButton(self.left_frame, text="設定", command=self.open_settings_window, fg_color="#464646", hover_color="#2B2B2B").pack(pady=10)
        # 終了ボタン（赤色）を一番下に配置
        exit_button = ctk.CTkButton(self.left_frame, text="終了", command=self.quit, fg_color="red", hover_color="darkred")
        exit_button.pack(side="bottom", pady=10)

        self.room_temp_label = ctk.CTkLabel(self.left_frame, text=f"現在の室温: {self.room_temperature}℃", font=("Arial", 12))
        self.room_temp_label.pack(pady=10)

        # 右側：グラフ表示エリア（タブ付き、３タブ）
        self.right_frame = ctk.CTkFrame(self.content_frame)
        self.right_frame.pack(side="right", expand=True, fill="both", padx=10, pady=10)
        self.graph_tabview = ctk.CTkTabview(self.right_frame, width=800, height=300)
        self.graph_tabview.pack(fill="both", expand=True)
        self.graph_tabview.add("真空度")        
        self.graph_tabview.add("電離真空計")
        self.graph_tabview.add("温度＆電圧")

        # 真空度グラフ（常に電離真空計のデータを表示）
        self.vac_fig, self.vac_ax = plt.subplots(figsize=(5,3))
        self.vac_canvas = FigureCanvasTkAgg(self.vac_fig, master=self.graph_tabview.tab("真空度"))
        self.vac_canvas.get_tk_widget().pack(fill="both", expand=True)

        # 電離真空計グラフ（対数表示）
        self.ion_fig, self.ion_ax = plt.subplots(figsize=(5,3))
        self.ion_canvas = FigureCanvasTkAgg(self.ion_fig, master=self.graph_tabview.tab("電離真空計"))
        self.ion_canvas.get_tk_widget().pack(fill="both", expand=True)

        # 温度＆電圧グラフ（双軸）
        self.temp_fig, self.temp_ax = plt.subplots(figsize=(5,3))
        self.temp_ax2 = self.temp_ax.twinx()
        self.temp_canvas = FigureCanvasTkAgg(self.temp_fig, master=self.graph_tabview.tab("温度＆電圧"))
        self.temp_canvas.get_tk_widget().pack(fill="both", expand=True)

    def update_current_time(self):
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.current_time_label.configure(text=current_time)
        self.after(1000, self.update_current_time)

    def append_log_line(self, message):
        if self.log_textbox is not None:
            self.log_textbox.insert("end", message + "\n")
            self.log_textbox.see("end")

    def update_graphs(self):
        # --- 真空度タブ（常に電離真空計のデータを表示） ---
        self.vac_ax.clear()
        if self.time_data and self.ion_data:
            self.vac_ax.semilogy(self.time_data, self.ion_data, marker='o', linestyle='-', label="Ion Gauge")
        self.vac_ax.set_xlabel("Time [s]")
        self.vac_ax.set_ylabel("Vacuum [Pa]")
        self.vac_ax.set_title("degree of vacuum")
        # プロットしたグラフにイベントマーカーを追加
        for t, label, color in self.event_markers:
            if t <= self.time_data[-1]:
                self.vac_ax.plot(t, self.ion_data[self.time_data.index(t)] if t in self.time_data else 1e-6,
                                 marker='D', color=color, markersize=8)
                self.vac_ax.annotate(label, (t, self.ion_data[self.time_data.index(t)] if t in self.time_data else 1e-6),
                                     textcoords="offset points", xytext=(0,10), ha='center')
        self.vac_ax.legend()
        self.vac_fig.tight_layout()
        self.vac_canvas.draw()

        # 電離真空計グラフ（基板温度測定開始後のみ）
        if self.show_substrate_graphs:
            self.ion_ax.clear()
            if self.ion_data2:
                # 直近の self.ion_data の個数に合わせた x 軸部分を使用
                x = self.time_data[-len(self.ion_data2):]
                self.ion_ax.semilogy(x, self.ion_data2, marker='o', linestyle='-', label="ionization vacuum gauge")
        self.ion_ax.set_xlabel("Time [s]")
        self.ion_ax.set_ylabel("Vacuum [Pa]")
        self.ion_ax.set_title("ionization vacuum gauge")
        self.ion_ax.legend()
        self.ion_fig.tight_layout()
        self.ion_canvas.draw()

        # 温度＆電圧グラフ（基板温度測定開始後のみ）
        if self.show_substrate_graphs:
            self.temp_ax.clear()
            self.temp_ax2.clear()
            if self.thermocouple_data:
                x_temp = self.time_data[-len(self.thermocouple_data):]
                self.temp_ax.plot(x_temp, self.thermocouple_data, 'r-', marker='o', label="temperature [℃]")
                self.temp_ax.axhline(self.target_temperature, color='red', linestyle='--', linewidth=2, label=f"target temperature {self.target_temperature}℃")
            if  self.heater_data:
                x_voltage = self.time_data[-len(self.heater_data):]
                self.temp_ax2.plot(x_voltage, self.heater_data, 'b-', marker='x', label="Voltage [V]")
        self.temp_ax.set_xlabel("Time [s]")
        self.temp_ax.set_ylabel("Temperature [℃]", color='r')
        self.temp_ax.set_ylim(0, 400)
        self.temp_ax.set_yticks(range(0, 401, 50))
        self.temp_ax2.set_ylabel("Voltage [V]", color='b', labelpad=10)
        self.temp_ax2.set_ylim(0, 50)
        self.temp_ax2.set_yticks(range(0, 51, 5))
        # 右軸ラベルの位置調整
        self.temp_ax2.yaxis.set_label_position("right")
        self.temp_ax2.yaxis.tick_right()
        self.temp_ax.legend(loc="upper left")
        self.temp_ax2.legend(loc="upper right")
        self.temp_ax.set_title("Temperature and Voltage")
        self.temp_fig.tight_layout()
        self.temp_canvas.draw()

        self.after(1000, self.update_graphs)

    # --- 各センサの測定関数 ---
    def get_pirani1_measurement(self):
        if self.pirani1_ser is None:
            try:
                self.pirani1_ser = serial.Serial(self.com_port1, 9600, timeout=1,
                                                  bytesize=8, stopbits=1, parity=serial.PARITY_NONE)
                time.sleep(2)
                if send_command(self.pirani1_ser, "CO") != "OK":
                    print("ピラニ1: CO失敗")
            except Exception as e:
                print(f"ピラニ1 シリアル接続エラー: {e}")
                return "Error"
        try:
            return send_command(self.pirani1_ser, "P0")
        except Exception as e:
            print(f"ピラニ1 測定エラー: {e}")
            return "Error"

    def get_pirani2_measurement(self):
        if self.pirani2_ser is None:
            try:
                self.pirani2_ser = serial.Serial(self.com_port2, 9600, timeout=1,
                                                  bytesize=8, stopbits=1, parity=serial.PARITY_NONE)
                time.sleep(2)
                if send_command(self.pirani2_ser, "CO") != "OK":
                    print("ピラニ2: CO失敗")
            except Exception as e:
                print(f"ピラニ2 シリアル接続エラー: {e}")
                return "Error"
        try:
            return send_command(self.pirani2_ser, "P0")
        except Exception as e:
            print(f"ピラニ2 測定エラー: {e}")
            return "Error"

    def get_ion_gauge_measurement(self):
        if self.ion_ser is None:
            try:
                self.ion_ser = serial.Serial(self.com_port3, 9600, timeout=1,
                                             bytesize=8, stopbits=1, parity=serial.PARITY_NONE)
                time.sleep(2)
                if send_command(self.ion_ser, "RE") != "OK":
                    print("電離真空計: RE失敗")
                if send_command(self.ion_ser, "F1") != "OK":
                    print("電離真空計: F1失敗")
            except Exception as e:
                print(f"電離真空計 シリアル接続エラー: {e}")
                return "Error"
        try:
            return send_command(self.ion_ser, "RP")
        except Exception as e:
            print(f"電離真空計 測定エラー: {e}")
            return "Error"

    def get_thermocouple_measurement(self):
        if self.thermocouple_ser is None:
            try:
                self.thermocouple_ser = serial.Serial(self.com_port4, 9600, timeout=1,
                                                      bytesize=8, stopbits=1, parity=serial.PARITY_NONE)
                time.sleep(0.1)
                self.thermocouple_ser.write(b'MAIN:FUNC DCV\r\n')
                time.sleep(0.1)
            except Exception as e:
                print(f"熱電対 シリアル接続エラー: {e}")
                return "Error"
        try:
            self.thermocouple_ser.write(b'MAIN:MEAS? XNOW\r\n')
            time.sleep(0.1)
            response = self.thermocouple_ser.readline().decode().strip()
            pattern = r'[-+]\d+\.\d+E[-+]\d+'
            match = re.search(pattern, response)
            if match:
                measurement = match.group(0)
                voltage_V = float(measurement)
                voltage_uV = voltage_V * 1_000_000
                try:
                    temp = self.room_temperature + (voltage_uV / 40.6125)
                    return str(temp)
                except Exception as e:
                    print(f"熱電対 温度変換エラー: {e}")
                    return "Error"
            else:
                print("熱電対: 測定値が見つかりません。")
                return "Error"
        except Exception as e:
            print(f"熱電対 測定エラー: {e}")
            return "Error"

    def get_heater_voltage_measurement(self):
        if self.heater_ser is None:
            try:
                self.heater_ser = serial.Serial(self.com_port5, 9600, timeout=1,
                                                bytesize=8, stopbits=1, parity=serial.PARITY_NONE)
                time.sleep(0.1)
                self.heater_ser.write(b'MAIN:FUNC ACV\r\n')
                time.sleep(0.1)
            except Exception as e:
                print(f"ヒーター電圧 シリアル接続エラー: {e}")
                return "Error"
        try:
            self.heater_ser.write(b'MAIN:MEAS? XNOW\r\n')
            time.sleep(0.1)
            response = self.heater_ser.readline().decode().strip()
            pattern = r'[-+]\d+\.\d+E[-+]\d+'
            match = re.search(pattern, response)
            if match:
                measurement = match.group(0)
                return measurement
            else:
                print("ヒーター電圧: 測定値が見つかりません。")
                return "Error"
        except Exception as e:
            print(f"ヒーター電圧 測定エラー: {e}")
            return "Error"

    def measure_data(self):
        while self.measurement_running:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            pirani1_value = self.get_pirani1_measurement()
            pirani2_value = self.get_pirani2_measurement()
            ion_value = self.get_ion_gauge_measurement()
            thermo_value = self.get_thermocouple_measurement()
            heater_value = self.get_heater_voltage_measurement()

            try: pirani1_val = float(pirani1_value)
            except: pirani1_val = float('nan')
            try: pirani2_val = float(pirani2_value)
            except: pirani2_val = float('nan')
            try: ion_val = float(ion_value)
            except: ion_val = float('nan')
            try: thermo_val = float(thermo_value)
            except: thermo_val = float('nan')
            try: heater_val = float(heater_value)
            except: heater_val = float('nan')

            current_time_sec = time.time() - self.start_time
            self.time_data.append(current_time_sec)
            self.pirani1_data.append(pirani1_val)
            self.pirani2_data.append(pirani2_val)
            self.ion_data.append(ion_val)
            if self.show_substrate_graphs:
                self.ion_data2.append(ion_val)
                self.thermocouple_data.append(thermo_val)
                self.heater_data.append(heater_val)

            log_line = (f"{timestamp} | ピラニ1: {pirani1_value} Pa | ピラニ2: {pirani2_value} Pa | "
                        f"電離: {ion_value} Pa | 温度: {thermo_value} ℃ | 電圧: {heater_value} V")
            self.append_log_line(log_line)
            data.append([timestamp, pirani1_value, pirani2_value, ion_value, thermo_value, heater_value])

            # ----- アラーム・通知処理 -----
            if self.show_substrate_graphs:
                # 目標温度(200℃)の通知（1回だけ）
                try:
                    if thermo_val >= self.target_temperature and not self.notif_200_triggered:
                        alarm_sound(1)
                        send_discord_notification("基盤温度が200℃になりました")
                        self.notif_200_triggered = True
                except Exception as e:
                    print("通知処理エラー（200℃）:", e)

                # 危険温度(250℃)の通知（1回だけ）
                try:
                    if thermo_val >= self.danger_temperature and not self.notif_250_triggered:
                        alarm_sound(5)
                        send_discord_notification("基盤温度が250℃になりました")
                        self.notif_250_triggered = True
                except Exception as e:
                    print("通知処理エラー（250℃）:", e)

                # 限界温度(300℃)の通知：温度が300℃以上の場合、5分毎に通知する
                try:
                    if thermo_val >= self.limit_temperature:
                        now = time.time()
                        if self.last_300_notif_time is None or (now - self.last_300_notif_time >= 300):
                            alarm_sound(5)
                            send_discord_notification("基盤温度が300℃を超えました")
                            self.last_300_notif_time = now
                    else:
                        # 温度が300℃以下になったら、通知用タイマーをリセット
                        self.last_300_notif_time = None
                except Exception as e:
                    print("通知処理エラー（300℃）:", e)

                    # 電離真空計関連の通知：各条件とも1回のみ送信
                try:
                    if ion_val <= 5e-4 and not self.ion_notify_heater:
                        send_discord_notification("ヒーター起動可能")
                        self.ion_notify_heater = True
                    if ion_val <= 2.67e-4 and thermo_val >= 200 and not self.ion_notify_vapor:
                        send_discord_notification("蒸着可能")
                        self.ion_notify_vapor = True
                except Exception as e:
                    print("通知処理エラー電離真空計関連:", e)

                # ヒーター電圧上げ通知（各しきい値：10,20,30,40 V）
                try:
                    for threshold in [10, 20, 30, 40]:
                        if heater_val >= threshold:
                            # しきい値以上の場合
                            if self.heater_increase_timestamp[threshold] is None:
                                # 初めてしきい値に達した時刻を記録
                                self.heater_increase_timestamp[threshold] = time.time()
                            elif not self.heater_increase_flag[threshold] and time.time() - self.heater_increase_timestamp[threshold] >= 20*60:
                                # しきい値に達してから20分以上経過している場合に通知
                                send_discord_notification(f"ヒーター電圧が{threshold}Vになってから20分経過しました (温度 {thermo_val} ℃, 電圧 {heater_val} V)")
                                self.heater_increase_flag[threshold] = True
                        else:
                            # しきい値を下回った場合はタイマーと通知フラグをリセット
                            self.heater_increase_timestamp[threshold] = None
                            self.heater_increase_flag[threshold] = False
                except Exception as e:
                    print("通知処理エラーヒーター:", e)

                # ヒーター電圧下げ通知（基板温度測定終了後）
                try:
                    if self.basis_ended:
                        # 初回に最大電圧からの下がりを監視（20V下がったタイミング）
                        if self.last_decrease_notif_voltage is not None and heater_val <= self.last_decrease_notif_voltage - 20:
                            if self.last_decrease_notif_timestamp is None:
                                self.last_decrease_notif_timestamp = time.time()
                            elif time.time() - self.last_decrease_notif_timestamp >= 20*60:
                                send_discord_notification(f"ヒーター電圧を下げてください (温度 {thermo_val} ℃, 電圧 {heater_val} V)")
                                # 更新：次の通知の基準を現在の電圧にする
                                self.last_decrease_notif_voltage = heater_val
                                self.last_decrease_notif_timestamp = None
                        # ヒーター電圧が2V以下になったら下げ通知を停止
                        if heater_val < 2:
                            self.basis_ended = False
                except Exception as e:
                    print("通知処理エラー終了中:", e)


            time.sleep(self.record_interval)


    def end_measurement(self):
        self.measurement_running = False
        if self.ion_ser is not None:
            try:
                send_command(self.ion_ser, "LO")
                send_command(self.ion_ser, "MAIN:LOC")
            except Exception as e:
                print(f"電離真空計コマンド送信エラー: {e}")
        if self.pirani1_ser is not None:
            try:
                send_command(self.pirani1_ser, "CF")
            except Exception as e:
                print(f"ピラニ1 CF送信エラー: {e}")
        if self.pirani2_ser is not None:
            try:
                send_command(self.pirani2_ser, "CF")
            except Exception as e:
                print(f"ピラニ2 CF送信エラー: {e}")
        if self.thermocouple_ser is not None:
            try:
                send_command(self.thermocouple_ser, "MAIN:LOC")
            except Exception as e:
                print(f"熱電対 MAIN:LOC送信エラー: {e}")

        for ser in [self.ion_ser, self.pirani1_ser, self.pirani2_ser, self.thermocouple_ser, self.heater_ser]:
            if ser is not None:
                try:
                    ser.close()
                except Exception as e:
                    print(f"シリアルポートクローズエラー: {e}")
        self.ion_ser = self.pirani1_ser = self.pirani2_ser = self.thermocouple_ser = self.heater_ser = None
        print("測定を停止し、各機器をローカルモードに戻しました。")

    def start_measurement(self):
        try:
            self.ion_ser = serial.Serial(self.com_port3, 9600, timeout=1,
                                         bytesize=8, stopbits=1, parity=serial.PARITY_NONE)
            time.sleep(2)
            if send_command(self.ion_ser, "RE") != "OK":
                print("電離真空計: RE失敗")
            if send_command(self.ion_ser, "F1") != "OK":
                print("電離真空計: F1失敗")
        except Exception as e:
            print(f"電離真空計 シリアル接続オープンエラー: {e}")
            return

        self.show_substrate_graphs = False
        self.baseline_heater_voltage = None
        self.notif_200_triggered = False
        self.notif_250_triggered = False
        self.last_300_notif_time = None
        self.heater_increase_notif = {10: None, 20: None, 30: None, 40: None}

        self.start_time = time.time()
        self.time_data = []
        self.pirani1_data = []
        self.pirani2_data = []
        self.ion_data = []
        self.thermocouple_data = []
        self.heater_data = []

        self.measurement_running = True
        self.measurement_thread = threading.Thread(target=self.measure_data, daemon=True)
        self.measurement_thread.start()

    def start_basis(self):  #基板温度測定開始
        self.show_substrate_graphs = True
        self.ion_data2 = []
        t = self.time_data[-1] if self.time_data else 0
        self.event_markers.append((t, "start temperature", "green")) #self.event_markers.append((t, "イベント名", "色"))
        if self.heater_data:
            self.baseline_heater_voltage = self.heater_data[-1]
        else:
            self.baseline_heater_voltage = 0
        send_discord_notification("基板温度測定を開始しました")

    def end_basis(self):
        self.show_substrate_graphs = False
        self.basis_ended = True
        t = self.time_data[-1] if self.time_data else 0
        self.event_markers.append((t, "end temperature", "orange"))
        if self.heater_data:
            self.last_decrease_notif_voltage = max(self.heater_data)
        else:
            self.last_decrease_notif_voltage = 0
        self.last_decrease_notif_timestamp = None
        send_discord_notification("基板温度測定を終了しました")

    def start_vapor_deposition(self):
        self.record_interval = 1
        self.vapor_events.append((self.time_data[-1] if self.time_data else 0, 'blue'))
        t = self.time_data[-1] if self.time_data else 0
        self.event_markers.append((t, "start vapor deposition", "blue"))
        send_discord_notification("蒸着開始")

    def end_vapor_deposition(self):
        self.record_interval = self.original_record_interval
        self.vapor_events.append((self.time_data[-1] if self.time_data else 0, 'red'))
        t = self.time_data[-1] if self.time_data else 0
        self.event_markers.append((t, "end vapor deposition", "red"))
        send_discord_notification("蒸着終了")

    def save_graph_images(self):
        """
        ファイル保存ダイアログで名前を指定し、
        各タブのグラフを画像ファイルとして保存する。
        例：ユーザーが「graph.png」と指定した場合、
        「graph_pirani.png」「graph_ion.png」「graph_temp.png」として保存する。
        """
        filename = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Files", "*.png")],
            title="グラフ画像の保存ファイル名を指定してください"
        )
        if filename:
            import os
            base, ext = os.path.splitext(filename)
            try:
                self.vac_fig.savefig(base + "_all_ion" + ext)
                self.ion_fig.savefig(base + "_ion" + ext)
                self.temp_fig.savefig(base + "_temp" + ext)
                messagebox.showinfo("保存完了", "グラフ画像が保存されました。")
            except Exception as e:
                messagebox.showerror("保存エラー", f"グラフ画像の保存中にエラーが発生しました: {e}")

    def date_keep(self):
        global EXCEL_FILE
        if not data:
            messagebox.showerror("エラー", "保存するデータがありません")
            return
        if not EXCEL_FILE:
            EXCEL_FILE = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excelファイル", "*.xlsx")],
                title="保存ファイル名を指定してください"
            )
        if not EXCEL_FILE:
            messagebox.showwarning("警告", "ファイルが選択されませんでした。\n保存を中止します。")
            return
        df = pd.DataFrame(data, columns=["Timestamp", "ピラニ1", "ピラニ2", "電離真空計", "熱電対", "ヒーター電圧"])
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="w") as writer:
            df.to_excel(writer, index=False)
        messagebox.showinfo("保存完了", f"データを{EXCEL_FILE}に保存しました。")

    def open_settings_window(self):
        if hasattr(self, "settings_window") and self.settings_window is not None and self.settings_window.winfo_exists():
            self.settings_window.lift()
            return
        self.settings_window = ctk.CTkToplevel(self)
        self.settings_window.title("設定")
        self.settings_window.geometry("400x500")
        self.settings_window.attributes('-topmost', True)
        self.settings_window.after(100, lambda: self.settings_window.attributes('-topmost', False))

        com_frame = ctk.CTkFrame(self.settings_window)
        com_frame.pack(fill="x", pady=5)

        frame1 = ctk.CTkFrame(com_frame)
        frame1.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ctk.CTkLabel(frame1, text="ピラニ1:", anchor="w").pack(side="left", padx=(0,5))
        self.com_port1_entry = ctk.CTkEntry(frame1, width=100)
        self.com_port1_entry.insert(0, self.com_port1)
        self.com_port1_entry.pack(side="left")

        frame2 = ctk.CTkFrame(com_frame)
        frame2.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        ctk.CTkLabel(frame2, text="ピラニ2:", anchor="w").pack(side="left", padx=(0,5))
        self.com_port2_entry = ctk.CTkEntry(frame2, width=100)
        self.com_port2_entry.insert(0, self.com_port2)
        self.com_port2_entry.pack(side="left")

        frame3 = ctk.CTkFrame(com_frame)
        frame3.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        ctk.CTkLabel(frame3, text="電離真空計:", anchor="w").pack(side="left", padx=(0,5))
        self.com_port3_entry = ctk.CTkEntry(frame3, width=100)
        self.com_port3_entry.insert(0, self.com_port3)
        self.com_port3_entry.pack(side="left")

        frame4 = ctk.CTkFrame(com_frame)
        frame4.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        ctk.CTkLabel(frame4, text="熱電対:", anchor="w").pack(side="left", padx=(0,5))
        self.com_port4_entry = ctk.CTkEntry(frame4, width=100)
        self.com_port4_entry.insert(0, self.com_port4)
        self.com_port4_entry.pack(side="left")

        frame5 = ctk.CTkFrame(com_frame)
        frame5.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        ctk.CTkLabel(frame5, text="ヒーター電圧:", anchor="w").pack(side="left", padx=(0,5))
        self.com_port5_entry = ctk.CTkEntry(frame5, width=100)
        self.com_port5_entry.insert(0, self.com_port5)
        self.com_port5_entry.pack(side="left")

        # 追加：温度閾値設定
        threshold_frame = ctk.CTkFrame(self.settings_window)
        threshold_frame.pack(fill="x", pady=5)
        ctk.CTkLabel(threshold_frame, text="目標温度 (℃):", anchor="w").grid(row=0, column=0, padx=5, pady=5)
        self.target_temp_entry = ctk.CTkEntry(threshold_frame, width=50)
        self.target_temp_entry.insert(0, str(self.target_temperature))
        self.target_temp_entry.grid(row=0, column=1, padx=5, pady=5)

        ctk.CTkLabel(threshold_frame, text="危険温度 (℃):", anchor="w").grid(row=1, column=0, padx=5, pady=5)
        self.danger_temp_entry = ctk.CTkEntry(threshold_frame, width=50)
        self.danger_temp_entry.insert(0, str(self.danger_temperature))
        self.danger_temp_entry.grid(row=1, column=1, padx=5, pady=5)

        ctk.CTkLabel(threshold_frame, text="限界温度 (℃):", anchor="w").grid(row=2, column=0, padx=5, pady=5)
        self.limit_temp_entry = ctk.CTkEntry(threshold_frame, width=50)
        self.limit_temp_entry.insert(0, str(self.limit_temperature))
        self.limit_temp_entry.grid(row=2, column=1, padx=5, pady=5)

        ctk.CTkLabel(self.settings_window, text="記録間隔（秒）:", anchor="w").pack(fill="x", pady=5)
        self.record_interval_entry = ctk.CTkEntry(self.settings_window)
        self.record_interval_entry.insert(0, str(self.record_interval))
        self.record_interval_entry.pack(fill="x", pady=5)

        ctk.CTkLabel(self.settings_window, text="現在の室温（℃）:", anchor="w").pack(fill="x", pady=5)
        self.room_temp_entry = ctk.CTkEntry(self.settings_window)
        self.room_temp_entry.insert(0, str(self.room_temperature))
        self.room_temp_entry.pack(fill="x", pady=5)

        save_button = ctk.CTkButton(self.settings_window, text="保存", command=self.save_settings)
        save_button.pack(pady=10)

    def save_settings(self):
        try:
            self.com_port1 = self.com_port1_entry.get()
            self.com_port2 = self.com_port2_entry.get()
            self.com_port3 = self.com_port3_entry.get()
            self.com_port4 = self.com_port4_entry.get()
            self.com_port5 = self.com_port5_entry.get()
            self.record_interval = int(self.record_interval_entry.get())
            self.room_temperature = float(self.room_temp_entry.get())
            self.target_temperature = float(self.target_temp_entry.get())
            self.danger_temperature = float(self.danger_temp_entry.get())
            self.limit_temperature  = float(self.limit_temp_entry.get())
            self.room_temp_label.configure(text=f"現在の室温: {self.room_temperature}℃")
            print(f"設定を保存しました: ピラニ1={self.com_port1}, ピラニ2={self.com_port2}, "
                  f"電離真空計={self.com_port3}, 熱電対={self.com_port4}, ヒーター電圧={self.com_port5}, "
                  f"記録間隔={self.record_interval}秒, 目標温度={self.target_temperature}℃, 危険温度={self.danger_temperature}℃, 限界温度={self.limit_temperature}℃")
            self.settings_window.destroy()
        except ValueError:
            print("設定にエラーがあります。正しい値を入力してください。")

    def open_log_window(self):
        if self.log_window is None or not self.log_window.winfo_exists():
            self.log_window = ctk.CTkToplevel(self)
            self.log_window.title("リアルタイムログ")
            self.log_window.geometry("600x400")
            self.log_window.attributes('-topmost', True)
            self.log_window.after(5000, lambda: self.log_window.attributes('-topmost', False))
            self.log_textbox = ctk.CTkTextbox(self.log_window)
            self.log_textbox.pack(fill="both", expand=True)
        else:
            self.log_window.lift()

def play_alarm():
    try:
        playsound("alarm.mp3")
    except Exception as e:
        print(f"アラーム音再生中にエラーが発生しました: {e}")

if __name__ == "__main__":
    # Discord Bot を別スレッドで起動
    discord_thread = threading.Thread(target=start_discord_bot, daemon=True)
    discord_thread.start()
    # メインアプリケーションを起動
    app = MeasurementApp()
    app.mainloop()
