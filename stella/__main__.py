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
        currentVersion = ConfigInformations.currentVersion
        currentLocation = os.getcwd()
        link = os.path.join(currentLocation, ConfigInformations.link)

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
            label="Update", command=lambda: self.CheckUpdate(currentVersion, link, root)
        )
        file.add_separator()
        file.add_command(label="Exit", command=root.destroy)
        root.config(menu=menubar)

        # Main GUI buttons:
        buttonCombinedClimatic = btk.Button(
            root,
            text="Chamber",
            command=lambda: [ClimaticTests(root, 0)],
            bootstyle=DARK,
        )
        buttonCombinedClimatic.place(x=50, y=40, width=92, height=30)

        buttonCombinedTS = btk.Button(
            root,
            text="TS Chamber",
            command=lambda: [ThermalShockTests(root, 0)],
            bootstyle=DARK,
        )
        buttonCombinedTS.place(x=170, y=40, width=92, height=30)

        buttonSalt = btk.Button(
            root,
            text="Salt Chamber",
            command=lambda: [OtherTests(root, 0)],
            bootstyle=DARK,
        )
        buttonSalt.place(x=40, y=110, width=102, height=30)

        buttonSalt = btk.Button(
            root,
            text="Specialized",
            command=lambda: [OtherTests(root, 1)],
            bootstyle=DARK,
        )
        buttonSalt.place(x=160, y=110, width=102, height=30)

        buttonRotronic = btk.Button(
            root,
            text="Rotronic",
            command=lambda: [Loggers(root, 0)],
            bootstyle=DARK,
        )
        buttonRotronic.place(x=10, y=180, width=82, height=30)

        buttonVaisala = btk.Button(
            root,
            text="Vaisala",
            command=lambda: [Loggers(root, 1)],
            bootstyle=DARK,
        )
        buttonVaisala.place(x=110, y=180, width=82, height=30)

        buttonKeithley = btk.Button(
            root,
            text="Keithley",
            command=lambda: [Loggers(root, 2)],
            bootstyle=DARK,
        )
        buttonKeithley.place(x=210, y=180, width=82, height=30)

        buttonGrafana = btk.Button(
            root,
            text="Grafana",
            command=lambda: [Loggers(root, 3)],
            bootstyle=DARK,
        )
        buttonGrafana.place(x=50, y=220, width=82, height=30)

        buttonAgilent = btk.Button(
            root,
            text="Agilent",
            command=lambda: [Loggers(root, 4)],
            bootstyle=DARK,
        )
        buttonAgilent.place(x=170, y=220, width=82, height=30)

        # Main GUI labels:
        labelChamber = tk.Label(root)
        ft = tkFont.Font(family="Helvetica", size=15)
        labelChamber["font"] = ft
        labelChamber["fg"] = "#333333"
        labelChamber["justify"] = "center"
        labelChamber["text"] = "Chamber Data"
        labelChamber.place(x=10, y=0, width=282, height=40)

        labelOtherChambers = tk.Label(root)
        ft = tkFont.Font(family="Helvetica", size=15)
        labelOtherChambers["font"] = ft
        labelOtherChambers["fg"] = "#333333"
        labelOtherChambers["justify"] = "center"
        labelOtherChambers["text"] = "Other Test Chambers"
        labelOtherChambers.place(x=10, y=70, width=282, height=40)

        labelLoggers = tk.Label(root)
        ft = tkFont.Font(family="Helvetica", size=15)
        labelLoggers["font"] = ft
        labelLoggers["fg"] = "#333333"
        labelLoggers["justify"] = "center"
        labelLoggers["text"] = "Logger Data"
        labelLoggers.place(x=10, y=140, width=282, height=40)

        creatorMessage = tk.Label(root)
        ft = tkFont.Font(family="Helvetica", size=10)
        creatorMessage["font"] = ft
        creatorMessage["fg"] = "#333333"
        creatorMessage["justify"] = "right"
        creatorMessage["text"] = "Stella " + currentVersion + " by PKl"
        creatorMessage.place(x=10, y=260, width=282, height=20)

    def CheckUpdate(self, currentVersion, link, root):
        # Auto-update
        try:
            update = pd.read_csv(link + "Stella_version.txt")
            nextVersion = update.iloc[0]["Version"]

            if currentVersion < nextVersion:
                popupWin0 = tk.Toplevel(root)
                width = 250
                height = 100
                screenwidth = popupWin0.winfo_screenwidth()
                screenheight = popupWin0.winfo_screenheight()
                alignstr = "%dx%d+%d+%d" % (
                    width,
                    height,
                    (screenwidth - width) / 2,
                    (screenheight - height) / 2,
                )
                popupWin0.geometry(alignstr)
                popupWin0.resizable(width=False, height=False)
                popupWin0.attributes("-topmost", True)

                buttonUpdate = btk.Button(
                    popupWin0,
                    text="Yes",
                    command=lambda: [popupWin0.destroy(), UpdateApp(link)],
                    bootstyle=SUCCESS,
                )
                buttonUpdate.place(x=30, y=50, width=70, height=30)

                buttondontUpdate = btk.Button(
                    popupWin0,
                    text="No",
                    command=lambda: [popupWin0.destroy(), WithoutUpd()],
                    bootstyle=DANGER,
                )
                buttondontUpdate.place(x=140, y=50, width=70, height=30)

                labelUpdate = tk.Label(popupWin0)
                ft = tkFont.Font(family="Helvetica", size=9)
                labelUpdate["font"] = ft
                labelUpdate["fg"] = "#333333"
                labelUpdate["justify"] = "center"
                labelUpdate["text"] = (
                    "New version is now available: "
                    + nextVersion
                    + "\nDo you want to download it?"
                )
                labelUpdate.place(x=20, y=0, width=200, height=40)

                def UpdateApp(link):
                    subprocess.call(link + "Stella.bat")

                def WithoutUpd():
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
