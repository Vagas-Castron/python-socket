import tkinter as tk
from tkinter import messagebox as mb, filedialog
from tkinter import ttk
import socket, threading
import json
from pystray import Icon, MenuItem, Menu
from PIL import Image, ImageTk
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import NamedStyle
import os, sys, re
from pathlib import Path
import queue
import time
from datetime import datetime, timedelta



class ITReportForm(tk.Tk):
    def __init__(self, mode, port=None, client_socket=None):
        super().__init__()
        self.mode = mode
        self.port = port
        self.client_socket = client_socket
        self.data = None
        self.tray_icon = None
        self.issue_option = "user-specific"
        self.data_ready = threading.Event()
        
        self.set_config_file()
        self.file_path = self.get_file_path()


        base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
        icon_path = os.path.join(base_path, "sign.ico")
        self.title("IT Report Form")
        self.iconbitmap(icon_path)
        if self.mode == "client":
            self.geometry(self._center_window(400, 250))
        else:
            self.geometry(self._center_window(400, 450))


        self.configure_gui(self, 4, 1)

        self.create_widget()
        # if self.mode == "server":
        #     self.start_server_thread()

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
        btn_frm = self.widget(parent=self, name="frame", row=3, col=0, sticky="n")
        self.configure_gui(btn_frm, 1, 3)
        
        if self.mode == "client":
            self.client_widgets()
            self.user_lbl_frm.grid(row=0, column=0, sticky="s")
            submit_btn = self.widget(parent=btn_frm, name="button", text="Submit", row=0, col=1, command=self.start_client_thread)

        else:
            self.minimize_to_tray()
            self.client_widgets()
            self.server_widgets()
            self.dropdown_frame = ttk.LabelFrame(self, text="General Problem")
            self.add_dropdown_content()
            self.config(menu=self.create_menu())
            self.specific_issue()
        
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

    def client_widgets(self):

        self.user_lbl_frm = ttk.LabelFrame(self, text="User Details")
        self.configure_gui(self.user_lbl_frm, 1, 1)

        usernames = self.user_options().get("usernames")
        username_lbl = self.widget(parent=self.user_lbl_frm, name="label", text="Username:", font=self.font, row=0, col=0, sticky="e")
        self.username = self.widget(parent=self.user_lbl_frm, name="combo", font=self.font, row=0, col=1, sticky="w", values=usernames)
        usernames = self.user_options().get("usernames")
        self.username.bind("<KeyRelease>", lambda event: self.combobox_filter(event, self.username, usernames))

        names = self.user_options().get("names")
        name_lbl = self.widget(parent=self.user_lbl_frm, name="label", text="Name:", font=self.font, row=1, col=0, sticky="e")
        self.name = self.widget(parent=self.user_lbl_frm, name="combo", font=self.font, row=1, col=1, sticky="w", values=names)
        names = self.user_options().get("names")
        self.name.bind("<KeyRelease>", lambda event: self.combobox_filter(event, self.name, names))

        des_lbl = self.widget(parent=self.user_lbl_frm, name="label", text="Description:", font=self.font, row=2, col=0, sticky="e")
        self.des = self.widget(parent=self.user_lbl_frm, name="entry", font=self.font, row=2, col=1, sticky="w")


        # Get current time in 24-hour format
        current_time = time.strftime("%H:%M")
        # print("Current time:", current_time)
        time_lbl = self.widget(parent=self.user_lbl_frm, name="label", text="Time:", font=self.font, row=3, col=0, sticky="e")
        self.time = self.widget(parent=self.user_lbl_frm, name="entry", font=self.font, row=3, col=1, sticky="w")
        
        if self.mode == "client":
            send_to_lbl = self.widget(parent=self.user_lbl_frm, name="label", text="Send to:", font=self.font, row=4, col=0, sticky="e")
            self.send_to = self.widget(parent=self.user_lbl_frm, name="combo", font=self.font, row=4, col=1, sticky="w")
        
        if self.mode == "client":
            self.time.delete(0, tk.END)
            self.time.insert(tk.END, current_time)
            self.des.delete(0, tk.END)
            self.des.insert(tk.END, "CC")

        if self.mode == "server":
            self.name["state"] = "disabled"
            self.username["state"] = "disabled"
            # self.des.config(state='readonly')
            # self.time.config(state='readonly')

    def server_widgets(self):
        self.it_lbl_frm = ttk.LabelFrame(self, text="IT Details")

        ip_lbl = self.widget(parent=self.it_lbl_frm, name="label", text="IP address:", font=self.font, row=0, col=0, sticky="e")
        self.ip_address_widget = self.widget(parent=self.it_lbl_frm, name="combo", font=self.font, row=0, col=1, sticky="w")
        self.ip_address_widget["state"] = "disabled"

        it_lbl = self.widget(parent=self.it_lbl_frm, name="label", text="IT Personel:", font=self.font, row=1, col=0, sticky="e")
        it_names = self.user_options().get("it")
        self.it = self.widget(parent=self.it_lbl_frm, name="combo", font=self.font, row=1, col=1, sticky="w", values=it_names)
        self.it.bind("<KeyRelease>", lambda event: self.combobox_filter(event, self.it, it_names))

        duration_lbl = self.widget(parent=self.it_lbl_frm, name="label", text="Duration:", font=self.font, row=2, col=0, sticky="e")
        self.duration = self.widget(parent=self.it_lbl_frm, name="entry", font=self.font, row=2, col=1, sticky="w")

        issue_lbl = self.widget(parent=self.it_lbl_frm, name="label", text="Issue:", font=self.font, row=3, col=0, sticky="e")
        self.issue = self.widget(parent=self.it_lbl_frm, name="entry", font=self.font, row=3, col=1, sticky="w")

    def create_menu(self):

        menu = tk.Menu(self)

        option_menu = tk.Menu(menu, tearoff=0)
        option_menu.add_command(label="Agent Specific", command=self.specific_issue)
        option_menu.add_command(label="Mass Issue", command=self.general_issue)

        preference_menu = tk.Menu(menu, tearoff=0)
        preference_menu.add_command(label="Save To", command=self.file_selector)

        menu.add_cascade(label="Options", menu=option_menu)
        menu.add_cascade(label="Settings", menu=preference_menu)
        return menu

    def get_app_folder(self):
        # Cross-platform AppData path resolution
        if os.name == 'nt':  # For Windows
            app_data_path = os.getenv('LOCALAPPDATA')
        else:  # For Linux/macOS
            app_data_path = os.path.expanduser("~/.local/share")

        # Validate if app_data_path is resolved
        if not app_data_path:
            raise EnvironmentError("Unable to locate AppData or equivalent directory.")
        
        # Define the path to your application's folder in AppData
        app_folder = os.path.join(app_data_path, "ReportApp")  # Path to your custom folder
        # Create the folder if it doesn't exist
        os.makedirs(app_folder, exist_ok=True)
        
        return app_folder
    
    def file_selector(self):
        # Open the file dialog to select a file
        file_path = filedialog.askopenfilename(
            title="Select a file", 
            # initialdir=self.file_path, 
            filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*"))
        )
        self.set_config_file(file_path)
        self.file_path = file_path


    def set_config_file(self, file_path=None):

        app_folder = self.get_app_folder()
        config_file = os.path.join(app_folder, "config.json")

        if not os.path.exists(config_file):
            # If the file doesn't exist, create it and write the default data
            # Get the path to the Desktop folder
            desktop_path = Path.home() / 'Desktop'
            config_data = {
                "file_path": os.path.join(desktop_path, "output.xlsx"),  # Default path or data
            }
            with open(config_file, "w") as file:
                json.dump(config_data, file, indent=4)
        else:
            if file_path:
                try:
                    with open(config_file, "r") as f:
                        config_data = json.load(f)  # Load existing config data
                        # Update the config with the new file path
                        config_data["file_path"] = file_path

                    with open(config_file, "w") as file:
                        json.dump(config_data, file, indent=4)

                except FileNotFoundError:
                    config_file = {}  # If config.json doesn't exist, create a new dictionary
        self.config_file = config_file


    def get_file_path(self):
        try:
            with open(self.config_file, "r") as file:
                config_file = json.load(file)
            return config_file.get("file_path")
        except Exception:
            print(Exception)
        

    def general_issue(self):
        self.dropdown_frame.grid(row=2, column=0, sticky="n", padx=5, pady=5)
        self.user_lbl_frm.grid_remove()
        self.it_lbl_frm.grid_remove()
        self.issue_option = "general"


    def specific_issue(self):
        self.user_lbl_frm.grid(row=0, column=0, sticky="s")
        self.it_lbl_frm.grid(row=1, column=0, sticky="n")
        # self.general_issue_txt.delete("1.0", tk.END)
        self.dropdown_frame.grid_remove()  # Hide the frame
        self.issue_option = "user-specific"



    def add_dropdown_content(self):

        # self.issue_option = "general"
        issue_label = ttk.Label(self.dropdown_frame, text="NOTE: ", font=("helvetica", 10, "bold"))
        issue_label.grid(row=0, column=0, padx=5, pady=5, sticky="ne")
        self.general_issue_txt = tk.Text(self.dropdown_frame, wrap="word", height=5, width=40, font=("helvetica", 10, "bold"))
        self.general_issue_txt.grid(row=0, column=1, padx=5, pady=5)
    # def create_receive_widget(self):

    def start_client_thread(self):
        threading.Thread(target=self.connect_server, args=(self.client_socket,), daemon=True).start()

    def minimize_to_tray(self):
    # Hide the window
        self.withdraw()

        # Create an icon for the tray
        image = Image.open("sign.ico")  # Replace with your icon file
        menu = Menu(MenuItem("Show", self.show_window), MenuItem("Exit", self.exit_app))
        self.tray_icon = Icon("name", image, "Report App", menu)
        self.tray_icon.run()

    # Function to handle showing the Tkinter window
    def show_window(self, icon=None, item=None):
        if self.tray_icon:
            self.tray_icon.stop()
        self.tray_icon = None
        self.deiconify()
        print("window shown?")

    def exit_app(self, icon=None, item=None):
        if self.tray_icon:
            self.tray_icon.stop()
        self.destroy()

    def get_month_range_name(self):
        """
        Generates a string representing the current and next month in the format:
        'CurrentMonth-NextMonth YYYY'
        """
        curr_date = datetime.now()
        # Get the current and next month
        current_month = curr_date.strftime("%B")  # e.g., "October"
        prev_month_date = curr_date - timedelta(days=31)  # Jump to the previous month
        prev_month = prev_month_date.strftime("%B")  # e.g., "November"
        
        # Get the year
        year = prev_month_date.strftime("%Y")  # Year corresponds to the previous month
        
        # Create the sheet name
        sheet_name = f"{current_month}-{prev_month} {year}"
        print(sheet_name)
        return sheet_name


    def creat_new_worksheet(self, workbook):
        today = datetime.now().date()
        file_name = self.get_file_path() # Path to your existing workbook

        # Format the sheet name as YYYY-MM-DD
        sheet_name = self.get_month_range_name()
        if today.day == 22:

            # Check if the sheet already exists to avoid duplication
            if sheet_name not in workbook.sheetnames:
                # Add a new sheet with the specific name
                worksheet = workbook.create_sheet(title=sheet_name)
                workbook.active = workbook.sheetnames.index(worksheet.title)
                print(f"Worksheet '{sheet_name}' added to the workbook.")
            else:
                print(f"Worksheet '{sheet_name}' already exists in the workbook.")

            # Save the workbook
            workbook.save(file_name)
            print(f"Workbook saved as: {file_name}")
            return workbook.active
        else:
            # Check if the workbook has only one sheet named "Sheet1"
            if len(workbook.sheetnames) == 1 and workbook.active.title == "Sheet1":
                # Remove "Sheet1"
                sheet_to_remove = workbook["Sheet1"]
                workbook.remove(sheet_to_remove)
                
                # Create a new custom-named sheet
                new_sheet_name = sheet_name
                workbook.create_sheet(new_sheet_name)
                workbook.active = workbook.sheetnames.index(new_sheet_name)
                workbook.save(file_name)
                return workbook.active
            return None

    

    def date_updater(self, sheet):
        dates = [cell.value for cell in sheet["A"] if cell.value is not None]
        dates = [datetime.strptime(str(date), "%Y-%m-%d").date() if isinstance(date, str) else date for date in dates]
        current_date = datetime.now().date()
        date_style = NamedStyle(name="custom_date_style", number_format="DD-MMM-YYYY")

        if dates:
            previous_date = max(dates)  # Assuming the dates are sorted or sparse
            print(f"Previous recorded date: {previous_date}")
        else:
            print("No previous date found.")
            previous_date = None
        
        if previous_date:
            if current_date > previous_date.date():
                next_row = sheet.max_row + 2
                sheet[f"A{next_row}"] = current_date
                sheet[f"A{next_row}"].number_format = "DD-MMM-YYYY"
                return next_row
            else:
                return sheet.max_row + 1
        else:
            next_row = sheet.max_row + 1
            sheet[f"A{next_row}"] = current_date
            sheet[f"A{next_row}"].number_format = "DD-MMM-YYYY"
            return next_row

    def data_reset(self):
        if self.issue_option == "user-specific":
            self.username.set("")
            self.name.set("")
            self.des.delete(0, tk.END)
            self.time.delete(0, tk.END)
            self.ip_address_widget.set("")
            self.duration.delete(0, tk.END)
            self.issue.delete(0, tk.END)
            self.it.set("")
        else:
            self.general_issue_txt.delete("1.0", tk.END)

    # def data_received(self):
        

    def on_data_submit(self):
        file_path = self.file_path

        try:
            if file_path.split(".")[-1] != "xlsx":
                raise ValueError("Invalid file selected \nMake sure you have selected valid file for saving data")
            workbook = load_workbook(file_path)
            sheet = self.creat_new_worksheet(workbook) or workbook.active
            print(f"this is the working sheet {sheet.title}")
        except Exception as e:
            mb.showerror("File Error", e)
            return

        it_name = self.it.get().upper()
        username = self.username.get()
        name = self.name.get().upper()
        des = self.des.get()
        ip_address = self.ip_address_widget.get()
        time = self.time.get()
        duration = self.duration.get()
        issue = self.issue.get().capitalize()
        general_issue = self.general_issue_txt.get("1.0", tk.END).strip()

        value_error = "Seems like you did not fill all the data \nSorry you can not submit with bank field"
        try:
            if self.issue_option == "user-specific":
                if it_name == "" or issue == "" or duration == "":
                    raise ValueError(value_error)

                data = [it_name, username, name, des, ip_address, time, duration, issue]

                next_row = self.date_updater(sheet) #Update date and return (next_row) which next row should the data be added to, depending on date position
                print("testing if it reaches her moment before writing")
                for col, value in enumerate(data, start=2):
                    sheet.cell(row=next_row, column=col, value=value)
            else:
                if general_issue == "":
                    raise ValueError(value_error)
                next_row = sheet.max_row + 1
                start_col = 2
                end_col =  9
                data = general_issue
                merge_range = f"{get_column_letter(start_col)}{next_row}:{get_column_letter(end_col)}{next_row}"
                sheet.merge_cells(merge_range)
                sheet.cell(row=next_row, column=start_col, value=data)
            print("testing if it reaches her moment before saving")
            workbook.save(file_path)
            print(file_path)
            self.data_reset()
            self.specific_issue()
            self.data_ready.set()
            self.minimize_to_tray()
        except Exception as e:
            mb.showerror("Submission Error", e)
        # finally:

    
    # a mothod for sending data to the server 
    def send_data(self, client_socket):
        # data['ip_address'] = client_socket.laddr[0]
        data = json.dumps(self.update_user_data()).encode("utf-8")
        try:
            client_socket.sendall(data)
            client_socket.close()
            self.destroy()
        except Exception as e:
            print(f"Error: {e}")

    def combobox_filter(self, event, combobox, option_values):
        input_value = combobox.get()
        threading.Thread(target=self.filter_options, args=(input_value, option_values, combobox), daemon=True).start()


    def filter_options(self, typed_text, option_values, combobox):
        # option_values = self.user_options().get("usernames")
        filtered_values = []
        if typed_text == "":
            combobox["values"] = option_values
        else:
            for option in option_values:
                if typed_text.lower() in option.lower():
                    filtered_values.append(option)
            self.after(0, lambda: self.update_combobox(typed_text, filtered_values, combobox))

    def update_combobox(self, typed_text, values, combobox):
        combobox["values"] = values
        combobox.delete(0, tk.END)            # Clear the entry
        combobox.insert(0, typed_text)        # Restore typed text
        combobox.icursor(len(typed_text)) 
        combobox.focus_set()
        # self.after(50, lambda: self.username.event_generate('<Down>'))

    def user_options(self):
        # self.data = None
        workbook = load_workbook("it_agent.xlsx")
        sheet = workbook["Sheet1"]
        return {
            "usernames": [cell.value for cell in sheet["B"] if cell.value is not None],
            "names": [cell.value for cell in sheet["A"] if cell.value is not None] ,
            "it": [cell.value for cell in sheet["C"] if cell.value is not None] 
        }

    def connect_server(self, client_socket):
      
        addr = self.send_to.get()
        port = self.port
        ipv4_pattern = re.compile(r'^((25[0-5]|2[0-4][0-9]|[0-1]?[0-9][0-9]?)\.){3}(25[0-5]|2[0-4][0-9]|[0-1]?[0-9][0-9]?)$')
                                  
        data = self.update_user_data()
        # self.send_data(client_socket, data)
        username = data.get("username")
        name = data.get("name")
        time = data.get("time")
        if username == "" or name == "" or time == "":
            error_des = "Please Make Sure The Required Fields Are Correctly Filled With Correct Information"
            mb.showerror("Error", error_des)
            return 
           
        if bool(ipv4_pattern.match(addr)):
            try:
                if client_socket is None:
                    client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                    client_socket.connect((addr, port))
                    self.client_socket = client_socket
                    print("Connection Accepted by server")
                    self.send_data(client_socket)
                else:
                    print(client_socket)
                    self.send_data(client_socket)
                    
            except socket.error as e:
                print(f"Connection error: {e}")
                mb.showinfo("Connection", f"{e}")
                return
        else:
            error_des = "Invalid IP address \nIP Range => 0.0.0.0 - 255.255.255.255 \nExample (172.168.1.10)"
            mb.showerror("IP Address Error", error_des)


    def update_user_data(self, data=None):
        self.data = data
        print("user data update function")
        if self.mode == "client":
            return {
                "username": self.username.get(),
                "name": self.name.get(),
                "des": self.des.get(),
                "time": self.time.get()
            }
        else:
            print(self.data)
            if self.data:
                self.username.configure(values=self.data.get("username"))
                self.username.set(self.data.get("username"))

                self.name.configure(values=self.data.get("name"))
                self.name.set(self.data.get("name"))

                self.des.delete(0, tk.END)
                self.des.insert(tk.END, self.data.get("des"))

                self.time.delete(0, tk.END)
                self.time.insert(tk.END, self.data.get("time"))

                self.ip_address_widget.delete(0, tk.END)
                self.ip_address_widget.set(self.data.get("client_ip_addr"))
                # self.deiconify()
                self.show_window()
                self.data_ready.clear()


def server_start(server_addr, server_port, data_queue):
    try:
        server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        server_socket.bind((server_addr, server_port))
        server_socket.listen()
        print("server started and listening")
        # display_gui(data_queue)
        threading.Thread(target=server_listening, args=(server_socket, data_queue,), daemon=True).start()

    except socket.error as e:
        print(f"Connection error: {e}")


def server_listening(server_socket, data_queue):
    while True:
        try:
            client_socket, addr = server_socket.accept()
            print(f"Client {addr} connected")
            threading.Thread(target=handle_clients, args=(client_socket ,data_queue), daemon=True).start()
        except socket.error as e:
            print(f"Error accepting connection: {e}")


def handle_clients(client_socket, data_queue):
    try:
        data = client_socket.recv(1024).decode("utf-8")
        data = json.loads(data)
        print(f"Received: {data}")
        data["client_ip_addr"] = client_socket.getpeername()[0]
        data_queue.put(data)
        
        # client_socket.close()
    except socket.error as e:
        print(f"Socket Error: {e}")
    finally:
        client_socket.close()


def display_gui(queue):
    print("gui processing display")
    mode = "server"
    app = ITReportForm(mode=mode)
    threading.Thread(target=process_gui_queue, args=(app, queue), daemon=True).start()
    app.protocol("WM_DELETE_WINDOW", app.minimize_to_tray)
    app.mainloop()

def process_gui_queue(app, queue):
    while True:
        try:
            data = queue.get()  # Get data sent from the server
            print("gui processing queue")
            if data == "STOP":
                break
            # Process data and update GUI
            app.update_user_data(data)
            app.data_ready.wait()

        except Exception as e:
            print(f"GUI update error: {e}")


if __name__ == "__main__":
    host = socket.gethostbyname(socket.gethostname())
    port = 12345
    data_queue = queue.Queue()

    # Get the installation directory of the executable
    if getattr(sys, 'frozen', False):
        # When running from the bundled executable, get the installation directory
        app_dir = os.path.dirname(sys.executable)
    else:
        # If running as a script (before packaging), use the current directory
        app_dir = os.path.dirname(os.path.abspath(__file__))
        # Define the path to the config.json file in the installation folder
    config_path = os.path.join(app_dir, 'config.json')

    print(app_dir)
    # Open the config file
    try:
        with open(config_path, 'r') as file:
            config_data = json.load(file)
            print("Config loaded successfully:", config_data)
    except FileNotFoundError:
            print(f"Error: {config_path} not found.")


    while True:
        mode = config_data.get("mode")
        if mode == "client" or mode == "server":
            if mode == "client":
                app = ITReportForm(mode, port)
                app.mainloop()
            else:
                threading.Thread(target=server_start, args=(host, port, data_queue), daemon=True).start()
                display_gui(data_queue)
            break