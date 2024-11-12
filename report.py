import tkinter as tk
from tkinter import ttk
import socket, threading
import json
from openpyxl import load_workbook, Workbook
import os, sys

class ITReportForm(tk.Tk):
    def __init__(self, local_host, port, mode, client_socket=None):
        super().__init__()
        self.host = local_host
        self.port = port
        self.mode = mode
        self.client_socket = client_socket
        base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
        icon_path = os.path.join(base_path, "sign.ico")
        self.title("IT Report Form")
        self.iconbitmap(icon_path)
        if self.mode == "client":
            self.geometry(self._center_window(400, 400))
        else:
            self.geometry(self._center_window(400, 250))


        self.configure_gui(self, 3, 1)

        self.create_widget()

    def _center_window(self, width, height):
        x = (self.winfo_screenwidth() - width) // 2
        y = (self.winfo_screenheight() - height) // 2
        return f"{width}x{height}+{x}+{y}"
    
    def configure_gui(self, frame, rows, cols):
        # if 
        for row in range(rows):
            frame.rowconfigure(row, weight=1)
        for col in range(cols):
            frame.columnconfigure(col, weight=1)

    def create_widget(self):
        self.font = ("helvetica", 10)
        btn_frm = self.widget(parent=self, name="frame", row=2, col=0, sticky="n")
        self.configure_gui(btn_frm, 1, 2)
        
        if self.mode == "server":
            self.user_widget()
            # self.start_server_thread()
            submit_btn = self.widget(parent=btn_frm, name="button", text="Submit", row=0, col=1, command=lambda: self.handle_client_request(self.client_socket))

        else:
            self.user_widget()
            self.it_widget()
            send_btn = self.widget(parent=btn_frm, name="button", text="Request", row=0, col=0, command=self.start_client_thread)
            submit_btn = self.widget(parent=btn_frm, name="button", text="Submit", row=0, col=1, command=self.on_data_submit)
        

    def widget(self, **kwargs):
        widget = None
        parent = kwargs.get("parent")
        font = kwargs.get("font")
        widget_name = kwargs.get("name")
        text = kwargs.get("text")
        btn_command = kwargs.get("command")
        row = kwargs.get("row")
        col = kwargs.get("col")
        padx = kwargs.get("padx")
        pady = kwargs.get("pady")
        sticky = kwargs.get("sticky")
        colspan = kwargs.get("colspan")
        vals = kwargs.get("values")

        if widget_name == "combo":
            widget = ttk.Combobox(parent, font=font, width=30, values=vals)
        if widget_name == "entry":
            widget = ttk.Entry(parent, font=font, width=32)
        if widget_name == "frame":
            widget = ttk.Frame(parent)
        if widget_name == "lframe":
            widget = ttk.LabelFrame(parent, text=text)
        if widget_name == "button":
            widget = ttk.Button(parent, text=text, command=btn_command)
        if widget_name == "label":
            widget = ttk.Label(parent, text=text, font=(*font, "bold"))
        widget.grid(row=row, column=col, padx=5, pady=5, sticky=sticky, columnspan=colspan)
        return widget

    def user_widget(self):
        # workbook = load_workbook("output.xlsx")
        # sheet = workbook["Sheet1"]
        # data_opt = {
        #     "usernames": [cell.value for cell in sheet["C"]],
        #     "names": [cell.value for cell in sheet["D"]]        
        # }
        # usernames = []
        # names = []
        # if self.mode == "client":
        # data_opt = json.loads(self.client_socket.recv(1024).decode("utf-8"))
            # usernames = data_opt.get("usernames", [])
            # names = data_opt.get("names", [])
        user_lbl_frm = self.widget(parent=self, name="lframe", text="User Details", row=0, col=0, sticky="s")
        self.configure_gui(user_lbl_frm, 1, 1)
        username_lbl = self.widget(parent=user_lbl_frm, name="label", text="Username:", font=self.font, row=0, col=0, sticky="e")
        self.username = self.widget(parent=user_lbl_frm, name="combo", font=self.font, row=0, col=1, sticky="w")

        name_lbl = self.widget(parent=user_lbl_frm, name="label", text="Name:", font=self.font, row=1, col=0, sticky="e")
        self.name = self.widget(parent=user_lbl_frm, name="combo", font=self.font, row=1, col=1, sticky="w")

        des_lbl = self.widget(parent=user_lbl_frm, name="label", text="Description:", font=self.font, row=2, col=0, sticky="e")
        self.des = self.widget(parent=user_lbl_frm, name="entry", font=self.font, row=2, col=1, sticky="w")
        self.des.delete(0, tk.END)
        self.des.insert(tk.END, "CCR")

        time_lbl = self.widget(parent=user_lbl_frm, name="label", text="Time:", font=self.font, row=3, col=0, sticky="e")
        self.time = self.widget(parent=user_lbl_frm, name="entry", font=self.font, row=3, col=1, sticky="w")

    def it_widget(self):
        it_lbl_frm = self.widget(parent=self, name="lframe", text="IT Information", row=1, col=0, sticky="n")

        ip_lbl = self.widget(parent=it_lbl_frm, name="label", text="IP address:", font=self.font, row=0, col=0, sticky="e")
        self.ip_address_widget = self.widget(parent=it_lbl_frm, name="combo", font=self.font, row=0, col=1, sticky="w")

        it_lbl = self.widget(parent=it_lbl_frm, name="label", text="IT Personel:", font=self.font, row=1, col=0, sticky="e")
        self.it = self.widget(parent=it_lbl_frm, name="combo", font=self.font, row=1, col=1, sticky="w")

        duration_lbl = self.widget(parent=it_lbl_frm, name="label", text="Duration:", font=self.font, row=2, col=0, sticky="e")
        self.duration = self.widget(parent=it_lbl_frm, name="entry", font=self.font, row=2, col=1, sticky="w")

        issue_lbl = self.widget(parent=it_lbl_frm, name="label", text="Issue:", font=self.font, row=3, col=0, sticky="e")
        self.issue = self.widget(parent=it_lbl_frm, name="entry", font=self.font, row=3, col=1, sticky="w")

    def start_client_thread(self):
        threading.Thread(target=self.connect_server, daemon=True).start()


    def on_data_submit(self):
        file_path = "output.xlsx"

        if os.path.exists(file_path):

            workbook = load_workbook("output.xlsx")
            sheet = workbook["Sheet1"]
        else:
            workbook = Workbook()
            workbook.active.title = "Sheet1"
            sheet = workbook["Sheet1"]
            
        it_name = self.it.get().upper()
        username = self.username.get()
        name = self.name.get().upper()
        des = self.des.get()
        ip_address = self.ip_address_widget.get()
        time = self.time.get()
        duration = self.duration.get()
        issue = self.issue.get().capitalize()

        data = [it_name, username, name, des, ip_address, time, duration, issue]


        next_row = sheet.max_row + 1
        for col, value in enumerate(data, start=2):
            sheet.cell(row=next_row, column=col, value=value)
        workbook.save("output.xlsx")
        self.destroy()


    def connect_server(self):
        self.data = None
        workbook = load_workbook("output.xlsx")
        sheet = workbook["Sheet1"]
        data_opt = {
            "usernames": [cell.value for cell in sheet["C"]],
            "names": [cell.value for cell in sheet["D"]]        
        }
        
        remote_host = self.ip_address_widget.get()
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as client_socket:
                client_socket.connect((remote_host, self.port))
                print("Connection Accepted by server")
                # client_socket.sendall(json.dumps(data_opt).encode("utf-8"))
                data = client_socket.recv(1024).decode("utf-8")
                self.data = json.loads(data)
                self.update_user_data(mode=self.mode)
        except socket.error as e:
            print(f"Connection error: {e}")

    def handle_client_request(self, client_socket):
        with client_socket:
            data = self.update_user_data(self.mode)
            user_data = json.dumps(data).encode("utf-8")
            client_socket.sendall(user_data)
            self.destroy()

    def update_user_data(self, mode):
        if mode == "server":
            return {
                "username": self.username.get(),
                "name": self.name.get(),
                "des": self.des.get(),
                "time": self.time.get()
            }
        else:
            if self.data:
                self.username.configure(values=self.data.get("username"))
                self.username.set(self.data.get("username"))

                self.name.configure(values=self.data.get("name"))
                self.name.set(self.data.get("name"))

                self.des.delete(0, tk.END)
                self.des.insert(tk.END, self.data.get("des"))

                self.time.delete(0, tk.END)
                self.time.insert(tk.END, self.data.get("time"))


def server_start():
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as server_socket:
            server_socket.bind((host, port))
            server_socket.listen()
            print("server started and listening")


            while True:
                client_socket, _ = server_socket.accept()
                display_gui(client_socket)
                print("Client connected")
                # threading.Thread(target=server_start(), daemon=True)

    except socket.error as e:
        print(f"Connection error: {e}")

def display_gui(client_socket):
    with client_socket:
        app = ITReportForm(host, port, mode, client_socket)
        app.mainloop()


def start_server_thread():
    threading.Thread(target=server_start, daemon=True).start()


if __name__ == "__main__":
    host = socket.gethostbyname(socket.gethostname())
    port = 12345
    while True:
        mode = input("Enter mode(client/server): ").strip().lower()
        if mode == "client" or mode == "server":
            if mode == "client":
                app = ITReportForm(host, port, mode)
                app.mainloop()
            else:
                server_start()
            break
        print("Please Enter Correct Mode")
    