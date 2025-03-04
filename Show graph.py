import customtkinter as ctk
from tkinter import filedialog, messagebox, Toplevel, Label, Entry
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from datetime import datetime

class ExcelGraphViewer(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Excel Graph Viewer")
        self.geometry("1200x700")
        # タイムスタンプ設定（既存）
        self.timestamp_settings = {
            "start temperature": {"time": None, "color": "green"},
            "end temperature": {"time": None, "color": "orange"},
            "start vapor deposition": {"time": None, "color": "blue"},
            "end vapor deposition": {"time": None, "color": "red"}
        }
        # 新たにオフセット（開始秒数）の設定
        self.vacuum_offset = 0.0  # 真空度 (指定秒から) 用オフセット
        self.temp_offset = 0.0    # 温度＆電圧 (指定秒から) 用オフセット
        self.create_layout()

    def create_layout(self):
        # 上部：操作パネル（PNG保存、タイムスタンプ設定、開始秒数設定、終了ボタン）
        self.header_frame = ctk.CTkFrame(self, height=40)
        self.header_frame.pack(side="top", fill="x")
        self.current_time_label = ctk.CTkLabel(self.header_frame, text="", font=("Arial", 14))
        self.current_time_label.pack(side="right", padx=10, pady=5)

        self.left_frame = ctk.CTkFrame(self, width=200)
        self.left_frame.pack(side="left", fill="y", padx=10, pady=10)
        ctk.CTkButton(self.left_frame, text="PNG保存", command=self.save_png_images,
                      fg_color="#4CAF50", hover_color="#388E3C").pack(padx=10, pady=5)
        ctk.CTkButton(self.left_frame, text="タイムスタンプ設定", command=self.open_timestamp_window,
                      fg_color="#008000", hover_color="#003300").pack(padx=10, pady=5)
        ctk.CTkButton(self.left_frame, text="タイムスタンプの追加", command=self.open_add_timestamp_window,
                      fg_color="#008000", hover_color="#003300").pack(padx=10, pady=5)
        ctk.CTkButton(self.left_frame, text="開始秒数設定", command=self.open_offset_window,
                      fg_color="#FFA000", hover_color="#FF8F00").pack(padx=10, pady=5)
        ctk.CTkButton(self.left_frame, text="終了", command=self.quit,
                      fg_color="red", hover_color="darkred").pack(padx=10, pady=5)
        ctk.CTkButton(self.left_frame, text="Excelファイルを開く", command=self.load_data).pack(padx=10, pady=5)
        
        
        # 右側：タブ付きグラフ表示エリア（5タブ）
        self.tabview = ctk.CTkTabview(self, width=900, height=600)
        self.tabview.pack(side="right", expand=True, fill="both", padx=10, pady=10)
        # 既存タブ
        self.tabview.add("真空度")
        self.tabview.add("温度＆電圧")
        # 追加タブ
        self.tabview.add("真空度 (指定秒から)")
        self.tabview.add("温度＆電圧 (指定秒から)")

        # 真空度タブ用グラフ（常に全データを表示）
        self.vac_fig, self.vac_ax = plt.subplots(figsize=(5,3))
        self.vac_canvas = FigureCanvasTkAgg(self.vac_fig, master=self.tabview.tab("真空度"))
        self.vac_canvas.get_tk_widget().pack(fill="both", expand=True)

        # 温度＆電圧タブ用グラフ（双軸）
        self.temp_fig, self.temp_ax = plt.subplots(figsize=(5,3))
        self.temp_ax2 = self.temp_ax.twinx()
        self.temp_canvas = FigureCanvasTkAgg(self.temp_fig, master=self.tabview.tab("温度＆電圧"))
        self.temp_canvas.get_tk_widget().pack(fill="both", expand=True)

        # 追加タブ：真空度 (指定秒から)
        self.vac_off_fig, self.vac_off_ax = plt.subplots(figsize=(5,3))
        self.vac_off_canvas = FigureCanvasTkAgg(self.vac_off_fig, master=self.tabview.tab("真空度 (指定秒から)"))
        self.vac_off_canvas.get_tk_widget().pack(fill="both", expand=True)

        # 追加タブ：温度＆電圧 (指定秒から)
        self.temp_off_fig, self.temp_off_ax = plt.subplots(figsize=(5,3))
        self.temp_off_ax2 = self.temp_off_ax.twinx()
        self.temp_off_canvas = FigureCanvasTkAgg(self.temp_off_fig, master=self.tabview.tab("温度＆電圧 (指定秒から)"))
        self.temp_off_canvas.get_tk_widget().pack(fill="both", expand=True)

    def open_offset_window(self):
        """開始秒数設定ウィンドウ"""
        off_win = Toplevel(self)
        off_win.title("開始秒数設定")
        off_win.geometry("300x150")
        Label(off_win, text="真空度開始秒数:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        vacuum_entry = Entry(off_win)
        vacuum_entry.insert(0, str(self.vacuum_offset))
        vacuum_entry.grid(row=0, column=1, padx=10, pady=5)
        Label(off_win, text="温度＆電圧開始秒数:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        temp_entry = Entry(off_win)
        temp_entry.insert(0, str(self.temp_offset))
        temp_entry.grid(row=1, column=1, padx=10, pady=5)
        def save_offsets():
            try:
                self.vacuum_offset = float(vacuum_entry.get())
                self.temp_offset = float(temp_entry.get())
                off_win.destroy()
                # 再描画
                if hasattr(self, "data"):
                    self.plot_graphs()
            except ValueError:
                messagebox.showerror("入力エラー", "開始秒数は数値で入力してください。")
        ctk.CTkButton(off_win, text="保存", command=save_offsets,
                      fg_color="#4CAF50", hover_color="#388E3C").grid(row=2, column=0, columnspan=2, pady=10)

    def open_timestamp_window(self):
        ts_window = Toplevel(self)
        ts_window.title("タイムスタンプ設定")
        ts_window.geometry("400x250")
        entries = {}
        row = 0
        for key, val in self.timestamp_settings.items():
            Label(ts_window, text=key).grid(row=row, column=0, padx=10, pady=5, sticky="w")
            initial_val = "" if val["time"] is None else str(val["time"])
            entry = Entry(ts_window)
            entry.insert(0, initial_val)
            entry.grid(row=row, column=1, padx=10, pady=5)
            Label(ts_window, text=val["color"], fg=val["color"]).grid(row=row, column=2, padx=10, pady=5)
            entries[key] = entry
            row += 1
        def save_timestamps():
            for key, entry in entries.items():
                val = entry.get().strip()
                if val:
                    try:
                        t_val = float(val)
                        self.timestamp_settings[key]["time"] = t_val
                    except ValueError:
                        messagebox.showerror("入力エラー", f"{key} の値が数値ではありません: {val}")
                        return
                else:
                    self.timestamp_settings[key]["time"] = None
            ts_window.destroy()
            if hasattr(self, "data"):
                self.plot_graphs()
        ctk.CTkButton(ts_window, text="保存", command=save_timestamps,
                      fg_color="#4CAF50", hover_color="#388E3C").grid(row=row, column=0, columnspan=3, pady=10)

    def open_add_timestamp_window(self):
        win = Toplevel(self)
        win.title("タイムスタンプ追加")
        win.geometry("400x200")

        Label(win, text="グラフ名:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        graph_entry = Entry(win)
        graph_entry.insert(0, "温度＆電圧")  # 例としてデフォルト
        graph_entry.grid(row=0, column=1, padx=10, pady=5)

        Label(win, text="時間 (秒):").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        time_entry = Entry(win)
        time_entry.insert(0, "100")  # 例として100秒
        time_entry.grid(row=1, column=1, padx=10, pady=5)

        Label(win, text="ラベル:").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        label_entry = Entry(win)
        label_entry.insert(0, "任意のタイムスタンプ")
        label_entry.grid(row=2, column=1, padx=10, pady=5)

        Label(win, text="色:").grid(row=3, column=0, padx=10, pady=5, sticky="w")
        color_entry = Entry(win)
        color_entry.insert(0, "purple")
        color_entry.grid(row=3, column=1, padx=10, pady=5)

        def save_and_add():
            try:
                g_name = graph_entry.get().strip()
                t_val = float(time_entry.get().strip())
                lab = label_entry.get().strip()
                col = color_entry.get().strip()
                self.add_timestamp_to_graph(g_name, t_val, lab, col)
                win.destroy()
            except ValueError:
                messagebox.showerror("入力エラー", "時間は数値で入力してください。")
        ctk.CTkButton(win, text="追加", command=save_and_add,
                    fg_color="#4CAF50", hover_color="#388E3C").grid(row=4, column=0, columnspan=2, pady=10)


    def add_timestamp_to_graph(self, graph_name, timestamp, label, color):
        """
        指定されたグラフにタイムスタンプのマーカーを追加する。
        
        Parameters:
            graph_name (str): "真空度", "温度＆電圧", "真空度 (指定秒から)", "温度＆電圧 (指定秒から)" など
            timestamp (float): x軸（秒）の値
            label (str): 表示するラベル
            color (str): ラインとラベルの色
        """
        # 対象のグラフ軸を選択
        if graph_name == "真空度":
            ax = self.vac_ax
            canvas = self.vac_canvas
        elif graph_name == "温度＆電圧":
            ax = self.temp_ax
            canvas = self.temp_canvas
        elif graph_name == "真空度 (指定秒から)":
            ax = self.vac_off_ax
            canvas = self.vac_off_canvas
        elif graph_name == "温度＆電圧 (指定秒から)":
            ax = self.temp_off_ax
            canvas = self.temp_off_canvas
        else:
            print("不明なグラフ名:", graph_name)
            return

        # 垂直線を追加
        ax.axvline(x=timestamp, color=color, linestyle='--', alpha=0.7)
        # y座標はグラフの上端の95%の位置にラベルを表示（適宜調整可能）
        y_min, y_max = ax.get_ylim()
        y_text = y_max * 0.95
        ax.text(timestamp, y_text, label, rotation=45, fontsize=9, color=color)
        canvas.draw()

    def load_data(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return
        try:
            self.data = pd.read_excel(file_path)
        except Exception as e:
            messagebox.showerror("エラー", f"Excelファイルの読み込みに失敗しました: {e}")
            return
        try:
            self.data["Timestamp"] = pd.to_datetime(self.data["Timestamp"])
            self.data["Elapsed"] = (self.data["Timestamp"] - self.data["Timestamp"].iloc[0]).dt.total_seconds()
        except Exception as e:
            messagebox.showerror("エラー", f"Timestamp列の変換に失敗しました: {e}")
            return
        self.plot_graphs()

    def plot_graphs(self):
        # --- 真空度タブ（常に全データを表示） ---
        vac_ax = self.vac_ax
        vac_ax.clear()
        if "Elapsed" in self.data.columns and "電離真空計" in self.data.columns:
            vac_ax.semilogy(self.data["Elapsed"], self.data["電離真空計"],
                            marker='o', linestyle='-', label="Ion Gauge")
            # すべてのタイムスタンプ（4つ）を表示
            for label_text, setting in self.timestamp_settings.items():
                t_val = setting["time"]
                color = setting["color"]
                if t_val is not None:
                    vac_ax.axvline(x=t_val, color=color, linestyle='--', alpha=0.7)
                    y_max = self.data["電離真空計"].max() if not self.data["電離真空計"].empty else 1
                    vac_ax.text(t_val, y_max, label_text, rotation=45, fontsize=9, color=color)
        vac_ax.set_xlabel("Time [s]")
        vac_ax.set_xlim(left=0)
        vac_ax.set_ylabel("Vacuum [Pa]")
        vac_ax.set_title("degree of vacuum")
        vac_ax.legend()
        vac_ax.figure.tight_layout()
        self.vac_canvas.draw()

        # --- 温度＆電圧タブ（基礎データ全体、全4タイムスタンプ表示） ---
        temp_ax = self.temp_ax
        temp_ax2 = self.temp_ax2
        temp_ax.clear()
        temp_ax2.clear()
        if "Elapsed" in self.data.columns and "熱電対" in self.data.columns:
            temp_ax.plot(self.data["Elapsed"], self.data["熱電対"], 'r-', marker='o', label="Temperature [℃]")
            temp_ax.axhline(200, color='red', linestyle='--', linewidth=2, label="Target 200℃")
        if "Elapsed" in self.data.columns and "ヒーター電圧" in self.data.columns:
            temp_ax2.plot(self.data["Elapsed"], self.data["ヒーター電圧"], 'b-', marker='x', label="Voltage [V]")
        # すべてのタイムスタンプ表示
        for label_text, setting in self.timestamp_settings.items():
            t_val = setting["time"]
            color = setting["color"]
            if t_val is not None:
                temp_ax.axvline(x=t_val, color=color, linestyle='--', alpha=0.7)
                # y座標は温度軸の上部（例:380℃）
                temp_ax.text(t_val, 380, label_text, rotation=45, fontsize=9, color=color)
        temp_ax.set_xlabel("Time [s]")
        temp_ax.set_xlim(left=0)
        temp_ax.set_ylabel("Temperature [℃]", color='r')
        temp_ax.set_ylim(0, 400)
        temp_ax.set_yticks(range(0, 401, 50))
        temp_ax2.set_ylabel("Voltage [V]", color='b', labelpad=10)
        temp_ax2.set_ylim(0, 50)
        temp_ax2.set_yticks(range(0, 51, 5))
        temp_ax2.yaxis.set_label_position("right")
        temp_ax2.yaxis.tick_right()
        temp_ax.legend(loc="upper left")
        temp_ax2.legend(loc="upper right")
        temp_ax.set_title("Temperature and Voltage")
        temp_ax.figure.tight_layout()
        self.temp_canvas.draw()

        # --- 追加タブ：真空度 (指定秒から)  ---
        vac_off_ax = self.vac_off_ax
        vac_off_ax.clear()
        if "Elapsed" in self.data.columns and "電離真空計" in self.data.columns:
            # 指定秒以降のデータを抽出し、オフセットを引いて調整
            df_vac = self.data[self.data["Elapsed"] >= self.vacuum_offset].copy()
            if not df_vac.empty:
                df_vac["Adjusted"] = df_vac["Elapsed"] - self.vacuum_offset
                vac_off_ax.semilogy(df_vac["Adjusted"], df_vac["電離真空計"],
                                    marker='o', linestyle='-', label="ionization vacuum gauge")
                # タイムスタンプは "start vapor deposition" と "end vapor deposition" のみ表示
                for label_text, setting in self.timestamp_settings.items():
                    if label_text in ["start vapor deposition", "end vapor deposition"]:
                        t_val = setting["time"]
                        color = setting["color"]
                        if t_val is not None and t_val >= self.vacuum_offset:
                            # 調整したx値
                            vac_off_ax.axvline(x=t_val - self.vacuum_offset, color=color, linestyle='--', alpha=0.7)
                            y_max = df_vac["電離真空計"].max() if not df_vac["電離真空計"].empty else 1
                            vac_off_ax.text(t_val - self.vacuum_offset, y_max, label_text, rotation=45, fontsize=9, color=color)
        vac_off_ax.set_xlabel("Time [s]")
        vac_off_ax.set_xlim(left=0)
        vac_off_ax.set_ylabel("Vacuum [Pa]")
        vac_off_ax.set_title("degree of vacuum")
        vac_off_ax.legend()
        vac_off_ax.figure.tight_layout()
        self.vac_off_canvas.draw()

        # --- 追加タブ：温度＆電圧 (指定秒から) ---
        temp_off_ax = self.temp_off_ax
        temp_off_ax2 = self.temp_off_ax2
        temp_off_ax.clear()
        temp_off_ax2.clear()
        if "Elapsed" in self.data.columns and "熱電対" in self.data.columns:
            df_temp = self.data[self.data["Elapsed"] >= self.temp_offset].copy()
            if not df_temp.empty:
                df_temp["Adjusted"] = df_temp["Elapsed"] - self.temp_offset
                temp_off_ax.plot(df_temp["Adjusted"], df_temp["熱電対"],
                                'r-', marker='o', label="Temperature [℃]")
                temp_off_ax.axhline(200, color='red', linestyle='--', linewidth=2, label="Target 200℃")
                # タイムスタンプは "start vapor deposition" と "end vapor deposition" のみ表示
                for label_text, setting in self.timestamp_settings.items():
                    if label_text in ["start vapor deposition", "end vapor deposition"]:
                        t_val = setting["time"]
                        color = setting["color"]
                        if t_val is not None and t_val >= self.temp_offset:
                            temp_off_ax.axvline(x=t_val - self.temp_offset, color=color, linestyle='--', alpha=0.7)
                            temp_off_ax.text(t_val - self.temp_offset, 380, label_text, rotation=45, fontsize=9, color=color)
        if "Elapsed" in self.data.columns and "ヒーター電圧" in self.data.columns:
            df_volt = self.data[self.data["Elapsed"] >= self.temp_offset].copy()
            if not df_volt.empty:
                df_volt["Adjusted"] = df_volt["Elapsed"] - self.temp_offset
                temp_off_ax2.plot(df_volt["Adjusted"], df_volt["ヒーター電圧"],
                                'b-', marker='x', label="Voltage [V]")
        temp_off_ax.set_xlabel("Time [s]")
        temp_off_ax.set_xlim(left=0)
        temp_off_ax.set_ylabel("Temperature [℃]", color='r')
        temp_off_ax.set_ylim(0, 400)
        temp_off_ax.set_yticks(range(0, 401, 50))
        temp_off_ax2.set_ylabel("Voltage [V]", color='b', labelpad=10)
        temp_off_ax2.set_ylim(0, 50)
        temp_off_ax2.set_yticks(range(0, 51, 5))
        temp_off_ax2.yaxis.set_label_position("right")
        temp_off_ax2.yaxis.tick_right()
        temp_off_ax.legend(loc="upper left")
        temp_off_ax2.legend(loc="upper right")
        temp_off_ax.set_title("Temperature and Voltage")
        temp_off_ax.figure.tight_layout()
        self.temp_off_canvas.draw()

    def save_png_images(self):
        filename = filedialog.asksaveasfilename(defaultextension=".png",
                                                filetypes=[("PNG Files", "*.png")],
                                                title="グラフ画像の保存ファイル名を指定してください")
        if filename:
            import os
            base, ext = os.path.splitext(filename)
            try:
                self.vac_fig.savefig(base + "_vac" + ext)
                self.temp_fig.savefig(base + "_temp" + ext)
                self.vac_off_fig.savefig(base + "_vac_offset" + ext)
                self.temp_off_fig.savefig(base + "_temp_offset" + ext)
                messagebox.showinfo("保存完了", "グラフ画像が保存されました。")
            except Exception as e:
                messagebox.showerror("保存エラー", f"グラフ画像の保存中にエラーが発生しました: {e}")

if __name__ == "__main__":
    app = ExcelGraphViewer()
    app.mainloop()
