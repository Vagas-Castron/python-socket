import tkinter as tk
from tkinter import ttk
import socket, threading
import json

class ITReportForm(tk.Tk):
    def __init__(self, host, port, mode):
        super().__init__()
        self.host = host
        self.port = port
        self.mode = mode

        self.title("IT Report Form")
        self.geometry(self._center_window(400, 500))

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
            self.start_server_thread()
            submit_btn = self.widget(parent=btn_frm, name="button", text="Submit", row=0, col=1, command=lambda: self.handle_client_request(self.client_socket))

        else:
            self.user_widget()
            self.it_widget()
            send_btn = self.widget(parent=btn_frm, name="button", text="Request", row=0, col=0, command=self.start_client_thread)
            submit_btn = self.widget(parent=btn_frm, name="button", text="Submit", row=0, col=1)
        


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

        if widget_name == "combo":
            widget = ttk.Combobox(parent, font=font, width=30)
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
        self.ip_address = self.widget(parent=it_lbl_frm, name="combo", font=self.font, row=0, col=1, sticky="w")

        it_lbl = self.widget(parent=it_lbl_frm, name="label", text="IT Personel:", font=self.font, row=1, col=0, sticky="e")
        self.it = self.widget(parent=it_lbl_frm, name="combo", font=self.font, row=1, col=1, sticky="w")

        duration_lbl = self.widget(parent=it_lbl_frm, name="label", text="Duration:", font=self.font, row=2, col=0, sticky="e")
        self.duration = self.widget(parent=it_lbl_frm, name="entry", font=self.font, row=2, col=1, sticky="w")

        issue_lbl = self.widget(parent=it_lbl_frm, name="label", text="Issue:", font=self.font, row=3, col=0, sticky="e")
        self.issue = self.widget(parent=it_lbl_frm, name="entry", font=self.font, row=3, col=1, sticky="w")

    def start_client_thread(self):
        threading.Thread(target=self.connect_server, daemon=True).start()


    def start_server_thread(self):
        threading.Thread(target=self.server_start, daemon=True).start()


    def connect_server(self):
        self.data = None
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as client_socket:
                client_socket.connect((self.host, self.port))
                print("Connection Accepted by server")
                data = client_socket.recv(1024).decode("utf-8")
                self.data = json.loads(data)
                self.update_user_data(mode=self.mode)
        except socket.error as e:
            print(f"Connection error: {e}")

    def server_start(self):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as server_socket:
                server_socket.bind((self.host, self.port))
                server_socket.listen()

                while True:
                    self.client_socket, _ = server_socket.accept()
                    print("Client connected")
                    threading.Thread(target=self.server_start(), daemon=True)

        except socket.error as e:
            print(f"Connection error: {e}")

    def handle_client_request(self, client_socket):
        with client_socket:
            data = self.update_user_data(self.mode)
            user_data = json.dumps(data).encode("utf-8")
            client_socket.sendall(user_data)

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




    # def data_update(self):
        

    # def widget_label(self, parent, text, font=None):
    #     return ttk.Label(parent, text=text, font=font)
    
    # def widget_position(self, widg, row, col):
    #     widg.grid(row=row, column=col, padx=5, pady=5, sticky=sticky)


if __name__ == "__main__":
    host = '192.168.0.250'
    port = 12345
    while True:
        mode = input("Enter mode(client/server): ").strip().lower()
        if mode == "client" or mode == "server":
            break
        print("Please Enter Correct Mode")
    app = ITReportForm(host, port, mode)
    app.mainloop()