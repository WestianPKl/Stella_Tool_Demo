import pandas as pd
import tkinter as tk
import tkinter.font as tkFont
import ttkbootstrap as btk
from ttkbootstrap.constants import *
from tkinter import messagebox
import subprocess
import os
from .config.config import ConfigInformations
from .climatic_chambers.climatic_chambers import ClimaticTests
from .thermal_shock.thermal_shock import ThermalShockTests
from .other_chambers.other_chambers import OtherTests
from .loggers.loggers import Loggers


class App:
    def __init__(self, root):
        # Current version with additional update informations:
        current_version = ConfigInformations.current_version
        current_location = os.getcwd()
        link = os.path.join(current_location, ConfigInformations.link)

        # Main GUI:
        root.title("Data Processing Tool")
        width = 304
        height = 300
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = "%dx%d+%d+%d" % (
            width,
            height,
            (screenwidth - width) / 2,
            (screenheight - height) / 2,
        )
        root.geometry(alignstr)
        root.resizable(width=False, height=False)

        menubar = tk.Menu(root)
        file = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file)
        file.add_command(
            label="Update",
            command=lambda: self.check_update(current_version, link, root),
        )
        file.add_separator()
        file.add_command(label="Exit", command=root.destroy)
        root.config(menu=menubar)

        # Main GUI buttons:
        button_combined_climatic = btk.Button(
            root,
            text="Chamber",
            command=lambda: [ClimaticTests(root, 0)],
            bootstyle=DARK,
        )
        button_combined_climatic.place(x=50, y=40, width=92, height=30)

        button_combined_ts = btk.Button(
            root,
            text="TS Chamber",
            command=lambda: [ThermalShockTests(root, 0)],
            bootstyle=DARK,
        )
        button_combined_ts.place(x=170, y=40, width=92, height=30)

        button_salt = btk.Button(
            root,
            text="Salt Chamber",
            command=lambda: [OtherTests(root, 0)],
            bootstyle=DARK,
        )
        button_salt.place(x=40, y=110, width=102, height=30)

        button_specialised = btk.Button(
            root,
            text="Specialized",
            command=lambda: [OtherTests(root, 1)],
            bootstyle=DARK,
        )
        button_specialised.place(x=160, y=110, width=102, height=30)

        button_rotronic = btk.Button(
            root,
            text="Rotronic",
            command=lambda: [Loggers(root, 0)],
            bootstyle=DARK,
        )
        button_rotronic.place(x=10, y=180, width=82, height=30)

        button_vaisala = btk.Button(
            root,
            text="Vaisala",
            command=lambda: [Loggers(root, 1)],
            bootstyle=DARK,
        )
        button_vaisala.place(x=110, y=180, width=82, height=30)

        button_keithley = btk.Button(
            root,
            text="Keithley",
            command=lambda: [Loggers(root, 2)],
            bootstyle=DARK,
        )
        button_keithley.place(x=210, y=180, width=82, height=30)

        button_grafana = btk.Button(
            root,
            text="Grafana",
            command=lambda: [Loggers(root, 3)],
            bootstyle=DARK,
        )
        button_grafana.place(x=50, y=220, width=82, height=30)

        button_agilent = btk.Button(
            root,
            text="Agilent",
            command=lambda: [Loggers(root, 4)],
            bootstyle=DARK,
        )
        button_agilent.place(x=170, y=220, width=82, height=30)

        # Main GUI labels:
        label_chamber = tk.Label(root)
        ft = tkFont.Font(family="Helvetica", size=15)
        label_chamber["font"] = ft
        label_chamber["fg"] = "#333333"
        label_chamber["justify"] = "center"
        label_chamber["text"] = "Chamber Data"
        label_chamber.place(x=10, y=0, width=282, height=40)

        label_other_chambers = tk.Label(root)
        ft = tkFont.Font(family="Helvetica", size=15)
        label_other_chambers["font"] = ft
        label_other_chambers["fg"] = "#333333"
        label_other_chambers["justify"] = "center"
        label_other_chambers["text"] = "Other Test Chambers"
        label_other_chambers.place(x=10, y=70, width=282, height=40)

        label_loggers = tk.Label(root)
        ft = tkFont.Font(family="Helvetica", size=15)
        label_loggers["font"] = ft
        label_loggers["fg"] = "#333333"
        label_loggers["justify"] = "center"
        label_loggers["text"] = "Logger Data"
        label_loggers.place(x=10, y=140, width=282, height=40)

        creator_message = tk.Label(root)
        ft = tkFont.Font(family="Helvetica", size=10)
        creator_message["font"] = ft
        creator_message["fg"] = "#333333"
        creator_message["justify"] = "right"
        creator_message["text"] = "Stella " + current_version + " by PKl"
        creator_message.place(x=10, y=260, width=282, height=20)

    def check_update(self, current_version, link, root):
        # Auto-update
        try:
            update = pd.read_csv(link + "Stella_version.txt")
            next_version = update.iloc[0]["Version"]

            if current_version < next_version:
                popup_win_0 = tk.Toplevel(root)
                width = 250
                height = 100
                screenwidth = popup_win_0.winfo_screenwidth()
                screenheight = popup_win_0.winfo_screenheight()
                alignstr = "%dx%d+%d+%d" % (
                    width,
                    height,
                    (screenwidth - width) / 2,
                    (screenheight - height) / 2,
                )
                popup_win_0.geometry(alignstr)
                popup_win_0.resizable(width=False, height=False)
                popup_win_0.attributes("-topmost", True)

                button_update = btk.Button(
                    popup_win_0,
                    text="Yes",
                    command=lambda: [popup_win_0.destroy(), update_app(link)],
                    bootstyle=SUCCESS,
                )
                button_update.place(x=30, y=50, width=70, height=30)

                button_dont_update = btk.Button(
                    popup_win_0,
                    text="No",
                    command=lambda: [popup_win_0.destroy(), without_upd()],
                    bootstyle=DANGER,
                )
                button_dont_update.place(x=140, y=50, width=70, height=30)

                label_update = tk.Label(popup_win_0)
                ft = tkFont.Font(family="Helvetica", size=9)
                label_update["font"] = ft
                label_update["fg"] = "#333333"
                label_update["justify"] = "center"
                label_update["text"] = (
                    "New version is now available: "
                    + next_version
                    + "\nDo you want to download it?"
                )
                label_update.place(x=20, y=0, width=200, height=40)

                def update_app(link):
                    subprocess.call(link + "Stella.bat")

                def without_upd():
                    messagebox.showinfo(
                        title="Update", message="Update me soon, please!"
                    )
                    pass

            else:
                messagebox.showinfo(title="Update", message="No updates are available.")

        except:
            messagebox.showwarning(
                title="Update",
                message="Update check couldn't be completed."
                "\nCheck your connection to local update folder.",
            )
            pass
