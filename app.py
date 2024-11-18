# pyinstaller -F --paths=D:\Task\Ota\robosell_app\app\.venv\Lib\site-packages app.py --hide-console hide-early --icon=D:\Task\Ota\robosell_app\app\icon1.ico --add-data "D:\Task\Ota\robosell_app\app\icon1.ico;."
# Standard library imports
import datetime
import json
import os
import sys
import threading
import time
import urllib.parse
import winreg as wrg
import requests

# Third-party library imports
import pandas as pd
import requests
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter.ttk import Progressbar

# Custom module imports
import find_contact
import send_contact

api_url = "https://autofill.robosell.jp/"
version = 3.8


def set_regidit_key(value):
    # Store location of HKEY_CURRENT_USER
    location = wrg.HKEY_CURRENT_USER

    # Store path in soft
    soft = wrg.OpenKeyEx(location, r"SOFTWARE\\")
    key_1 = wrg.CreateKey(soft, "Autofill")

    # Creating values in Geeks
    wrg.SetValueEx(key_1, "api_key", 0, wrg.REG_SZ, value)

    if key_1:
        wrg.CloseKey(key_1)


def get_regidit_key():
    # Store location of HKEY_CURRENT_USER
    location = wrg.HKEY_CURRENT_USER

    # Storing path in soft
    try:
        soft = wrg.OpenKeyEx(location, r"SOFTWARE\\Autofill\\")
    except FileNotFoundError:
        set_regidit_key("")
        return ""
    else:
        # reading values in value_1 and value_2
        api_key = str(wrg.QueryValueEx(soft, "api_key"))
        api_key = api_key[2:-5]
        # Closing folder
        if soft:
            wrg.CloseKey(soft)
        return api_key


def resource_path(relative_path):
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def update_progress(current_value, total_value, auto_cotact):
    auto_contact.total_value_var.set(f"登録数: {total_value}")
    auto_contact.current_value_var.set(f"進行数: {current_value}")
    auto_contact.progress["value"] = (current_value / total_value) * 100
    auto_contact.update_idletasks()


class Auto_contact(tk.Tk):
    def __init__(self):
        super().__init__()
        style = ttk.Style()
        self.title(f"RoboSell {version}")

        style.configure(
            "Custom.TButton",
            foreground="#15903d",
            background="#15903d",
            font=("Yu Mincho", 11),
            relief="raised",
        )

        self.mana_id = ""
        self.info_data = ""
        self.url_item_array = []
        self.api_key = ""
        self.chat_api_key = ""
        self.user_data = ""
        self.start_time = ""
        self.end_time = ""
        # Set window size and padding
        self.geometry("570x550")
        self.configure(padx=20, pady=20)
        icon_path = resource_path("icon1.ico")
        self.iconbitmap(icon_path)
        self.save_data = []
        self.url_count = 0
        self.total_url_cnt = 0

        # Set row and column padding
        self.rowconfigure(0, pad=10)
        self.rowconfigure(1, pad=10)
        self.rowconfigure(2, pad=20)
        self.rowconfigure(3, pad=20)
        self.rowconfigure(4, pad=20)
        self.rowconfigure(5, pad=20)
        self.columnconfigure(0, pad=20)
        self.columnconfigure(1, pad=36)
        self.columnconfigure(2, pad=20)
        self.columnconfigure(3, pad=12)
        self.columnconfigure(4, pad=12)

        self.menubar = tk.Menu(self)

        # Create a File menu
        self.file_menu = tk.Menu(self.menubar, tearoff=0)
        self.file_menu.add_command(
            label="CSV入力 (O)", font=("MS Gothic", 9), command=self.open_csv_file
        )
        self.file_menu.add_command(
            label="CSV出力 (S)", font=("MS Gothic", 9), command=self.save_csv
        )
        self.file_menu.add_separator()
        self.file_menu.add_command(label="ヘルプ  (H)", font=("MS Gothic", 9))
        self.file_menu.add_separator()
        self.file_menu.add_command(
            label="完了", command=self.quit, font=("MS Gothic", 9)
        )

        # Create an Edit menu
        self.edit_menu = tk.Menu(self.menubar, tearoff=0)
        self.edit_menu.add_command(
            label="開始 (A)", font=("MS Gothic", 9), command=self.start_work
        )
        self.edit_menu.add_command(
            label="停止 (P)", font=("MS Gothic", 9), command=self.stop_work
        )
        self.edit_menu.add_separator()
        self.edit_menu.add_command(
            label="APIキーの登録", font=("MS Gothic", 9), command=self.open_api_frame
        )

        # Add the menus to the self.menubar
        self.menubar.add_cascade(
            label="ファイル", menu=self.file_menu, font=("MS Gothic", 9)
        )
        self.menubar.add_cascade(
            label="アクション", menu=self.edit_menu, font=("MS Gothic", 9)
        )

        # create hotkey
        self.bind("<Control-Key-o>", self.open_csv_file)
        self.bind("<Control-Key-s>", self.save_csv)
        self.bind("<Control-Key-a>", self.start_work)
        self.bind("<Control-Key-p>", self.stop_work)

        # Configure the self window to use the self.menubar
        self.config(menu=self.menubar)

        # Create buttons for the first row
        self.register_api_btn = ttk.Button(
            self,
            text="APIキー登録",
            command=self.open_api_frame,
            style="Custom.TButton",
        )
        self.register_api_btn.grid(row=0, column=0)

        self.api_key_txt = tk.Label(
            self, text="APIキー:", justify="right", anchor="nw", font=("Yu Mincho", 10)
        )
        self.api_key_txt.grid(row=0, column=1, sticky="nw", columnspan=3) 

        self.register_mana_btn = ttk.Button(
            self,
            text="管理ID登録",
            command=self.open_mana_frame,
            style="Custom.TButton",
        )
        self.register_mana_btn.grid(row=1, column=0)

        self.mana_id_txt = tk.Label(
            self, text="管理ID:", justify="right", anchor="nw", font=("Yu Mincho", 10)
        )
        self.mana_id_txt.grid(row=1, column=1, sticky="nw", columnspan=3) 

        self.open_button = ttk.Button(
            self,
            text="顧客データ登録",
            command=self.open_csv_file,
            style="Custom.TButton",
        )
        self.open_button.grid(row=2, column=0)

        # self.mana_id_label = tk.Label(
        #     self, text="管理ID:", justify="left", anchor="nw", font=("Yu Mincho", 10)
        # )

        # self.mana_id_label.grid(row=2, column=1, sticky="nw", columnspan=3)

        # self.mana_id_input =  tk.Entry(self, font=("Yu Mincho", 12), width=8)
        # self.mana_id_input.grid(row=2, column=2)

        # Create buttons for the second row
        self.start_btn = ttk.Button(
            self, text="開始", command=self.start_work, style="Custom.TButton"
        )
        self.start_btn.grid(row=2, column=3)

        self.top_btn = ttk.Button(
            self, text="停止", command=self.stop_work, style="Custom.TButton"
        )
        self.top_btn.grid(row=2, column=4)

        # Create StringVar objects
        self.current_value_var = tk.StringVar(self)
        self.total_value_var = tk.StringVar(self)

        # Create a progress bar and labels for current/total values
        progress_label = tk.Label(
            self, textvariable=self.current_value_var, font=("Yu Mincho", 12)
        )
        progress_label.grid(row=3, column=1)
        total_label = tk.Label(
            self, textvariable=self.total_value_var, font=("Yu Mincho", 12)
        )
        total_label.grid(row=3, column=2)
        self.progress = Progressbar(
            self, orient=tk.HORIZONTAL, length=500, mode="determinate"
        )
        self.progress.grid(row=4, column=0, columnspan=5)

        # Create the frame group for the third row
        self.list_frame = tk.Frame(self)
        # self.list_frame.grid(row=2, column=0, columnspan=4, sticky="nsew")

        # Create a Canvas widget
        canvas = tk.Canvas(self.list_frame)
        canvas.grid(row=0, column=0, sticky="nsew")

        # Create a Scrollbar widget
        scrollbar = ttk.Scrollbar(
            self.list_frame, orient="vertical", command=canvas.yview
        )
        scrollbar.grid(row=0, column=1, sticky="ns")

        # Configure the Canvas widget to use the Scrollbar
        canvas.configure(yscrollcommand=scrollbar.set)

        # Create a Frame inside the Canvas
        self.inner_frame = ttk.Frame(canvas)

        canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")

        # Add content to the inner Frame

        # Configure grid weights for resizing
        self.list_frame.grid_rowconfigure(0, weight=1)
        self.list_frame.grid_columnconfigure(0, weight=1)

        # Configure the Canvas to resize with the Frame
        canvas.bind(
            "<Configure>",
            lambda event: canvas.configure(scrollregion=canvas.bbox("all")),
        )

        # Enable scrolling with the mouse wheel
        canvas.bind_all(
            "<MouseWheel>",
            lambda event: canvas.yview_scroll(int(-1 * (event.delta / 120)), "units"),
        )

        # Create a button for the fourth row
        show_msg = """
            進行中は画面を閉じないでください。  
            途中終了する場合、停止してから画面を閉じてください。  
            すべての処理が終了後、自動で画面が閉じられます。  
        """
        self.alarmtxt = tk.Label(
            self,
            width=46,
            justify="left",
            anchor="sw",
            font=("Yu Mincho", 10),
            text=show_msg,
        )
        self.alarmtxt.configure(borderwidth=1, padx=2)
        self.alarmtxt.configure(relief="raised")
        self.save_btn = ttk.Button(
            self, text="閉じる", command=self.quit_action, style="Custom.TButton"
        )
        self.save_btn.grid(row=5, column=3)

    def quit_action(self):
        if self.url_count != 0:
            show_msg = """
                進行中は画面を閉じないでください。\n
                途中終了する場合、停止してから画面を閉じてください。\n
                すべての処理が終了後、自動で画面が閉じられます。
            """
            messagebox.showinfo("警告", show_msg)
            return
        self.quit()

    def open_csv_file(self, event=None):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        self.file_path = file_path
        if file_path:
            try:
                self.info_data = pd.read_csv(file_path, encoding="SJIS")
            except UnicodeDecodeError:
                try:
                    self.info_data = pd.read_csv(file_path, encoding="UTF8")
                except UnicodeDecodeError:
                    messagebox.showinfo(
                        "警告",
                        "Failed to read the CSV file with SJIS and UTF8 encodings.",
                    )
                    return
            self.file_name = file_path.split("/")[-1]

            if "本文" in self.info_data.columns:
                self.info_data.rename(
                    columns={"本文": "お問い合わせ詳細"}, inplace=True
                )

            if "企業URL" in self.info_data.columns:
                self.info_data.rename(columns={"企業URL": "サイトurl"}, inplace=True)

            header = list(self.info_data.columns)
            required_columns = ["顧客ID", "サイトurl", "企業名", "お問い合わせ詳細"]

            # Check if all required columns are present
            missing_columns = [col for col in required_columns if col not in header]
            if missing_columns:
                show_msg = (
                    """必須項目がありません。詳しくはマニュアルをご確認ください。"""
                )
                messagebox.showinfo("警告", show_msg)
                return
        for i in range(0, len(self.info_data)):
            item_data = self.info_data.iloc[i]
            if "nan" == str(item_data.iloc[1]):
                continue
            list_item = List_item(self.inner_frame, item_data, i)
            self.url_item_array.append(list_item)
            list_item.grid(row=i, column=0, columnspan=4, sticky="nw")
        if len(self.url_item_array) > 500:
            show_msg = """一度に登録できる件数は最大500です。"""
            messagebox.showinfo("警告", show_msg)
            return
        update_progress(0, len(self.url_item_array), self)

    def open_api_frame(self):
        self.register_api_btn.config(
            state=tk.DISABLED
        )  # Disable the button after clicking
        new_window = tk.Toplevel(self, borderwidth=2, relief="ridge")
        api_register_frame = Register_api(self, new_window)
        api_register_frame.grid(row=0, column=0)

    def open_mana_frame(self):
        self.register_mana_btn.config(
            state=tk.DISABLED
        )  # Disable the button after clicking
        new_window = tk.Toplevel(self, borderwidth=2, relief="ridge")
        mana_register_frame = Register_mana(self, new_window)
        mana_register_frame.grid(row=0, column=0)

    def start_work(self, event=None):
        if self.api_key == "":
            show_msg = (
                """APIキーがありません。APIキーを入力して続行してください。"""
            )
            messagebox.showinfo("警告", show_msg)
            return
        if self.mana_id == "":
            show_msg = (
                """管理IDがありません。管理IDを入力して続行してください。"""
            )
            messagebox.showinfo("警告", show_msg)
            return
        url = api_url + "api/get_contact_data"
        data = {"api_key": self.api_key, "domain": "test", "mana_id": self.mana_id}
        plan_type = ""
        response = requests.get(url, params=data)
        if response.status_code == 200:
            self.start_time = response.json()["start_time"]
            self.end_time = response.json()["end_time"]
            plan_type = response.json()["plan_type"]
            cur_mana_id = self.mana_id
            # messagebox.showinfo("警告", response.json()["mana_info"])
        current_time = datetime.datetime.now().time()
        start_time = datetime.datetime.strptime(self.start_time, "%H:%M:%S").time()
        end_time = datetime.datetime.strptime(self.end_time, "%H:%M:%S").time()
        # if current_time >= start_time and current_time <= end_time:
        #     pass
        # else:
        #     show_msg = "設定した時間帯ではありません!\n設定した時間帯は" + self.start_time[:-3] + "～" + self.end_time[:-3] + "です。"
        #     messagebox.showinfo("警告", show_msg)
        #     return

        start_flg = True

        if plan_type == "無料プラン":

            for item in self.info_data["お問い合わせ詳細"]:

                if ".doc1.jp/" not in item or str(item) == "":
                    start_flg = False
                    show_msg = """
                        別サービス「DocTrack」の契約（無料）が必須となります。
                        https://doctrack.jp/
                    """
                    messagebox.showinfo("アラート", show_msg)
                    break
                else:
                    start_flg = True

        if self.api_key != "" and start_flg:
            url = api_url + "api/send_file"
            file_data = self.info_data.to_json(orient="records")

            files = {"file_data": (self.file_name, open(self.file_path, "rb"))}
            # Include other parameters in the 'data' parameter
            data = {"api_key": self.api_key, "file_name": self.file_name}

            response = requests.post(url, data=data, files=files)

            if response.status_code == 200:
                self.file_id = response.json()["file_id"]
            else:
                print(response.status_code, response.json()["error"])
                messagebox.showinfo(
                    "アラート", "ファイルアップロードが成功しませんでした。"
                )
                return
            self.start_btn.config(text="再開")
            self.start_btn.config(state=tk.DISABLED)
            self.alarmtxt.place(relx=0.05, rely=0.8)
            self.start_thread = threading.Thread(target=self.start_process)
            self.start_thread.start()
            self.start_btn.config(state=tk.NORMAL)
            self.save_btn.configure(state=tk.DISABLED)

    def start_process(self):
        for item in self.url_item_array:
            item.start_process()

    def start_process(self):
        chunk_size = 5
        self.url_count = len(self.url_item_array)
        self.total_url_cnt = len(self.url_item_array)
        num_chunks = (len(self.url_item_array) + chunk_size - 1) // chunk_size

        threads = []
        for i in range(num_chunks):
            start_index = i * chunk_size
            end_index = (i + 1) * chunk_size
            chunk_items = self.url_item_array[start_index:end_index]

            for item in chunk_items:
                thread = threading.Thread(target=item.process)
                thread.start()
                threads.append(thread)
                time.sleep(10)

        for thread in threads:
            thread.join()

    def stop_work(self, event=None):
        self.stop_thread = threading.Thread(target=self.stop_process)
        self.stop_thread.start()

    def stop_process(self):
        self.top_btn.config(state=tk.DISABLED)
        for item in self.url_item_array:
            item.stop_process()
            if item.state_txt.cget("text") == "進行中":
                item.state_txt.config(text="停止")
        self.top_btn.config(state=tk.NORMAL)

    def save_csv(self, event=None):
        csv_data = pd.DataFrame(
            self.save_data, columns=["顧客ID", "企業名", "結果", "完了日時", "メモ"]
        )
        csv_data.to_csv("save_data.csv", index=False, encoding="SJIS")
        messagebox.showinfo("成功", "データ保存成功!")


class Register_api(tk.Frame):
    def __init__(self, parent, frame):
        super().__init__(parent)
        self.configure(borderwidth=2, relief="solid")
        self.parent = parent
        self.frame = frame
        self.frame.title("APIキー登録")
        self.frame.overrideredirect(True)
        self.frame.attributes("-topmost", True)
        x_pos = parent.winfo_x() + 80
        y_pos = parent.winfo_y() + 200
        self.frame.geometry(f"+{x_pos}+{y_pos}")
        frame.rowconfigure(0)
        frame.rowconfigure(1, pad=30)
        frame.rowconfigure(2, pad=20)

        frame.columnconfigure(0, pad=20)
        frame.columnconfigure(1, pad=20)

        self.move_bar = tk.Label(
            self.frame,
            bg="#22c55e",
            width=40,
            height=2,
            text="APIキー登録",
            font=("Yu Mincho", 12),
            fg="white",
        )
        self.move_bar.grid(row=0, column=0, columnspan=2, padx=0)

        self.register_btn = ttk.Button(
            self.frame,
            text="APIキー登録",
            command=self.register_api_key,
            style="Custom.TButton",
        )
        self.register_btn.grid(row=1, column=0)

        self.api_text = tk.Entry(self.frame, font=("Yu Mincho", 12), width=28)
        self.api_text.grid(row=1, column=1)
        self.api_text.insert(0, "APIキーを貼り付けてください")
        self.api_text.focus()

        self.close_button = ttk.Button(
            self.frame, text="取消", command=self.close_frame, style="Custom.TButton"
        )
        self.close_button.grid(row=2, column=0, columnspan=2)
        self.frame.protocol("WM_DELETE_WINDOW", self.close_frame)

        self.move_bar.bind("<ButtonPress-1>", self.on_drag_start)
        self.move_bar.bind("<B1-Motion>", self.on_drag_motion)

    def close_frame(self):
        self.parent.register_api_btn.config(
            state=tk.NORMAL
        )  # Enable the button in the main frame
        self.frame.destroy()

    def register_api_key(self):
        url = api_url + "api/get_contact_data"
        apikey = self.api_text.get()
        data = {"api_key": apikey, "version": version, "domain": "test"}

        response = requests.get(url, params=data)
        if response.status_code == 200:
            data = response.json()["message"]
            user_data = response.json()["user_data"]
            chat_api_key = response.json()["chat_api_key"]
            self.parent.user_data = json.loads(user_data)
            messagebox.showinfo("お知らせ", "APIキーが正確に登録されました。")
            auto_contact.api_key = apikey
            auto_contact.api_key_txt.config(text="APIキー:" + apikey)
            auto_contact.chat_api_key = chat_api_key
            auto_contact.start_time = response.json()["start_time"]
            auto_contact.end_time = response.json()["end_time"]
            self.close_frame()

        elif response.status_code == 401:
            messagebox.showerror("お知らせ", response.json()["message"])
        elif response.status_code == 505:
            messagebox.showerror("お知らせ", response.json()["message"])
            auto_contact.register_api_btn.config(state=tk.DISABLED)
            auto_contact.open_button.config(state=tk.DISABLED)
            auto_contact.start_btn.config(state=tk.DISABLED)
            auto_contact.top_btn.config(state=tk.DISABLED)
            auto_contact.save_btn.config(state=tk.DISABLED)
        else:
            messagebox.showerror("お知らせ", "サーバーから応答がありません。")

    def on_drag_start(self, event):
        self.frame.startX = event.x
        self.frame.startY = event.y

    def on_drag_motion(self, event):
        deltax = event.x - self.frame.startX
        deltay = event.y - self.frame.startY
        x = self.frame.winfo_x() + deltax
        y = self.frame.winfo_y() + deltay
        self.frame.geometry(f"+{x}+{y}")

    def __del__(self):
        self.parent.register_api_btn.config(state=tk.NORMAL)

class Register_mana(tk.Frame):
    def __init__(self, parent, frame):
        super().__init__(parent)
        self.configure(borderwidth=2, relief="solid")
        self.parent = parent
        self.frame = frame
        self.frame.title("管理ID登録")
        self.frame.overrideredirect(True)
        self.frame.attributes("-topmost", True)
        x_pos = parent.winfo_x() + 80
        y_pos = parent.winfo_y() + 200
        self.frame.geometry(f"+{x_pos}+{y_pos}")
        frame.rowconfigure(0)
        frame.rowconfigure(1, pad=30)
        frame.rowconfigure(2, pad=20)

        frame.columnconfigure(0, pad=20)
        frame.columnconfigure(1, pad=20)

        self.move_bar = tk.Label(
            self.frame,
            bg="#22c55e",
            width=40,
            height=2,
            text="管理ID登録",
            font=("Yu Mincho", 12),
            fg="white",
        )
        self.move_bar.grid(row=0, column=0, columnspan=2, padx=0)

        self.register_btn = ttk.Button(
            self.frame,
            text="管理ID登録",
            command=self.register_mana_key,
            style="Custom.TButton",
        )
        self.register_btn.grid(row=1, column=0)

        self.mana_id = tk.Entry(self.frame, font=("Yu Mincho", 12), width=28)
        self.mana_id.grid(row=1, column=1)
        self.mana_id.insert(0, "管理IDを貼り付けてください")
        self.mana_id.focus()

        self.close_button = ttk.Button(
            self.frame, text="取消", command=self.close_frame, style="Custom.TButton"
        )
        self.close_button.grid(row=2, column=0, columnspan=2)
        self.frame.protocol("WM_DELETE_WINDOW", self.close_frame)

        self.move_bar.bind("<ButtonPress-1>", self.on_drag_start)
        self.move_bar.bind("<B1-Motion>", self.on_drag_motion)

    def close_frame(self):
        self.parent.register_mana_btn.config(
            state=tk.NORMAL
        )  # Enable the button in the main frame
        self.frame.destroy()

    def register_mana_key(self):
        url = api_url + "api/get_mana_data"
        mana_id = self.mana_id.get()
        data = {"mana_id": mana_id, "version": version, "domain": "test"}

        response = requests.get(url, params=data)
        if response.status_code == 200:
            messagebox.showinfo("お知らせ", "管理IDが正確に登録されました。")
            auto_contact.mana_id = self.mana_id.get()
            auto_contact.mana_id_txt.config(text="管理ID:" + mana_id)
            self.close_frame()

        elif response.status_code == 401:
            messagebox.showerror("お知らせ", response.json()["message"])
        elif response.status_code == 505:
            messagebox.showerror("お知らせ", response.json()["message"])
        else:
            messagebox.showerror("お知らせ", "サーバーから応答がありません。")

    def on_drag_start(self, event):
        self.frame.startX = event.x
        self.frame.startY = event.y

    def on_drag_motion(self, event):
        deltax = event.x - self.frame.startX
        deltay = event.y - self.frame.startY
        x = self.frame.winfo_x() + deltax
        y = self.frame.winfo_y() + deltay
        self.frame.geometry(f"+{x}+{y}")

    def __del__(self):
        self.parent.register_mana_btn.config(state=tk.NORMAL)


class List_item(tk.Frame):
    def __init__(self, parent, data, i):
        super().__init__(parent)
        self.data = data
        self.i = i + 1
        self.error_array = []
        self.configure(pady=3)

        self.search_data = None

        # Create label fields in the small frame group
        self.id_txt = tk.Label(
            self, text=str(self.i) + ".", width=5, height=1, font=("Yu Mincho", 12)
        )
        self.id_txt.grid(row=0, column=0)
        self.url_txt = tk.Label(
            self,
            text=self.data.iloc[1],
            justify="left",
            width=20,
            anchor="nw",
            height=1,
            font=("Yu Mincho", 12),
        )
        self.url_txt.grid(row=0, column=1, sticky="nw")

        self.state_txt = tk.Label(
            self, text="ステータス", width=10, height=1, font=("Yu Mincho", 12)
        )
        self.state_txt.grid(row=0, column=2)
        self.driver = ""
        self.options = ""
        self.error_label = tk.Label(
            self, text="  ", width=10, height=1, font=("Yu Mincho", 12)
        )
        self.error_label.grid(row=0, column=3)

    def process(self):
        domain = self.data.iloc[1]
        protocol = domain.split("://")[0]
        domain = domain.split("://")[1]
        protocol = protocol + "://"
        profile_data = self.data
        for key, key_value in auto_contact.user_data.items():
            profile_data[key] = key_value

        domain_parts = domain.split("/")
        domain = domain_parts[0]

        url = api_url + "api/get_contact_data"
        contact_page_url = False
        req = {
            "api_key": auto_contact.api_key,
            "domain": domain,
            "cnt": auto_contact.total_url_cnt,
        }
        response = requests.get(url, params=req)
        if "type" in response.json() and response.json()["type"] == "CantUseApiKey":
            self.state_txt.config(text="失敗")
            self.error_array.append({"field": "API", "msg": "No found APIKey"})
            result = "自動送信【失敗】"
        elif "type" in response.json() and response.json()["type"] == "noForm":
            result = "フォーム無し"
        else:
            result = "自動送信【失敗】"
            contact_finder = find_contact.FindContact()
            contact_page_url = contact_finder.find_contact_page_url(self.data.iloc[1])
            contact_form = {}
            if contact_page_url != False:
                contact_form = contact_finder.find_contact_form(contact_page_url)

            if contact_form != {}:
                send_process = send_contact.SendContact()
                send_data = send_process.send_data(
                    contact_page_url, contact_form, profile_data
                )
                if send_data == "success":
                    result = "自動送信【成功】"
            else:
                result = "フォーム未発見"
        auto_contact.url_count -= 1
        update_progress(
            auto_contact.total_url_cnt - auto_contact.url_count,
            auto_contact.total_url_cnt,
            auto_contact,
        )

        if auto_contact.url_count < 1:
            auto_contact.start_btn.config(state=tk.DISABLED)
            auto_contact.top_btn.config(state=tk.DISABLED)
            auto_contact.save_btn.config(state=tk.DISABLED)
            print("--------- end ---------")

        client_id = self.data.iloc[0]
        now_time = datetime.datetime.now()

        url = api_url + "api/send_result"
        parts = profile_data["お問い合わせ詳細"].splitlines()  # Split string into lines
        encoded_parts = []

        for part in parts:
            part = part.replace("\t", " ")
            if part.startswith("http://") or part.startswith("https://"):
                encoded_part = urllib.parse.quote(part, safe=":/?=&")
                encoded_parts.append(encoded_part)
            else:
                encoded_parts.append(part)

        # Reconstruct the memo with preserved formatting
        memo = "\n".join(encoded_parts)

        send_data = {
            "customer_id": client_id,
            "company_name": profile_data["企業名"],
            "result": result,
            "contact_url": contact_page_url,
            "now_time": now_time.strftime("%Y-%m-%d %H:%M:%S"),
            "memo": memo,
            "file_id": auto_contact.file_id,
        }
        req = {
            "api_key": auto_contact.api_key,
            "data": json.dumps(send_data).replace("'", '"'),
        }
        response = requests.get(url, params=req)

        if auto_contact.url_count == 0:
            auto_contact.quit()

    def start_process(self):
        self.service = ChromeService(executable_path=ChromeDriverManager().install())
        self.options = webdriver.ChromeOptions()
        self.options.add_argument("--ignore-certificate-error")
        self.options.add_argument("--ignore-ssl-errors")
        self.options.add_argument("--headless")
        self.driver = webdriver.Chrome(options=self.options)
        self.driver.maximize_window()
        self.process_thread = threading.Thread(target=self.process)
        self.process_thread.start()

    def send_error(self):
        domain = self.data.iloc[1]
        domain = domain.split("://")[1]
        if domain[-1] == "/":
            domain = domain[:-1]
        new_data = ""
        error = json.dumps(self.error_array)
        if self.search_data == None:
            new_data = "register_domain"
        else:
            new_data = json.dumps({"data": self.search_data})
            new_data = new_data.replace("'", "'")
        url = api_url + "api/send_error_data"
        req = {
            "api_key": auto_contact.api_key,
            "domain": domain,
            "message": error,
            "new_data": new_data,
        }
        response = requests.get(url, params=req)

    def stop_process(self):
        self.process_thread.join()
        self.driver.quit()
        del self.process_thread

    def __del__(self):
        self.process_thread.join()
        del self.process_thread


if __name__ == "__main__":
    auto_contact = Auto_contact()
    auto_contact.mainloop()
