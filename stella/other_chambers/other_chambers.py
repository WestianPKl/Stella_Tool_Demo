import pandas as pd
import tkinter as tk
import os
from tkinter import filedialog
import tkinter.font as tkFont
import ttkbootstrap as btk
from ttkbootstrap.constants import *
from ..utility.message import Message
from ..data.dataFormating import DataFormating


class OtherTests:
    def __init__(self, root, option):
        self.root = root
        if option == 0:
            self.ReadSalt()
        elif option == 1:
            self.ReadZUTMichalin()

    def ReadSalt(self):
        # Salt data processing:
        filepaths = filedialog.askopenfilenames(
            parent=self.root,
            filetypes=(
                ("SIMPATI files", "*.x*"),
                ("all files", ".*"),
            ),
            initialdir=os.getcwd(),
        )
        filepaths2 = list(filepaths)
        listData = []
        for filepath in filepaths2:
            try:
                self.data = pd.read_csv(
                    filepath,
                    header=3,
                    sep=";",
                    low_memory=False,
                    encoding_errors="ignore",
                )
                newColumnNames = {"Unnamed: 0": "Date and Time"}
                self.data.rename(columns=newColumnNames, inplace=True)
                self.data.drop(self.data.index[[0]], axis=0, inplace=True)
                self.data = pd.DataFrame(self.data)
                self.data = self.data.rename(
                    columns={
                        i: "Temperature"
                        for i in self.data.columns
                        if i.startswith("A:Temperat")
                    }
                )
                self.data = self.data.rename(
                    columns={
                        i: "Temperature Set"
                        for i in self.data.columns
                        if i.startswith("N:Temperat")
                    }
                )
                newColumnNames = {
                    "N:Humidifier  ": "Humidifier Set",
                    "A:Humidifier  ": "Humidifier",
                    "N:Temp.Humidif": "Humidifier Set",
                    "A:Temp.Humidif": "Humidifier",
                    "N:Humidity    ": "Humidity Set",
                    "A:Humidity    ": "Humidity",
                }
                self.data.rename(columns=newColumnNames, inplace=True)
                self.data["Temperature"] = self.data["Temperature"].str.replace(
                    ",", ".", regex=True
                )
                self.data["Temperature Set"] = self.data["Temperature Set"].str.replace(
                    ",", ".", regex=True
                )
                self.data["Humidifier"] = self.data["Humidifier"].str.replace(
                    ",", ".", regex=True
                )
                self.data["Humidifier Set"] = self.data["Humidifier Set"].str.replace(
                    ",", ".", regex=True
                )
                try:
                    self.data["Humidity Set"] = self.data["Humidity Set"].str.replace(
                        ",", ".", regex=True
                    )
                    self.data["Humidity"] = self.data["Humidity"].str.replace(
                        ",", ".", regex=True
                    )
                except:
                    pass
                self.data["Date and Time"] = pd.to_datetime(
                    self.data["Date and Time"], dayfirst=True, errors="coerce"
                )
                self.data["Temperature Set"] = pd.to_numeric(
                    self.data["Temperature Set"], errors="coerce"
                )
                self.data["Temperature"] = pd.to_numeric(
                    self.data["Temperature"], errors="coerce"
                )
                self.data["Humidifier Set"] = pd.to_numeric(
                    self.data["Humidifier Set"], errors="coerce"
                )
                self.data["Humidifier"] = pd.to_numeric(
                    self.data["Humidifier"], errors="coerce"
                )
                listData.append(self.data)
            except:
                return (
                    "error",
                    "Wrong file format."
                    "\nPlease contact with: piotr.klys92@gmail.com",
                    0,
                )
        try:
            self.data = pd.concat(listData, axis=0, ignore_index=False)
            self.data = self.data.sort_values(by="Date and Time", ascending=True)
        except:
            return (
                "error",
                "Wrong file format." "\nPlease contact with: piotr.klys92@gmail.com",
                0,
            )
        try:
            # Data processing variants - popup:
            popupWin1 = tk.Toplevel(self.root)
            width = 233
            height = 147
            screenwidth = popupWin1.winfo_screenwidth()
            screenheight = popupWin1.winfo_screenheight()
            alignstr = "%dx%d+%d+%d" % (
                width,
                height,
                (screenwidth - width) / 2,
                (screenheight - height) / 2,
            )
            popupWin1.geometry(alignstr)
            popupWin1.resizable(width=False, height=False)
            popupWin1.attributes("-topmost", True)
            popupWin1.minsize(233, 147)
            popupWin1.maxsize(233, 147)

            buttonTemperature = btk.Button(
                popupWin1,
                text="Salt test",
                command=lambda: [popupWin1.destroy(), saltTest(self)],
                bootstyle=DARK,
            )
            buttonTemperature.place(x=10, y=50, width=102, height=30)

            buttonHumidity = btk.Button(
                popupWin1,
                text="Salt Humidity",
                command=lambda: [popupWin1.destroy(), humidityTest(self)],
                bootstyle=DARK,
            )
            buttonHumidity.place(x=120, y=50, width=102, height=30)

            labelData = tk.Label(popupWin1)
            labelData["anchor"] = "n"
            ft = tkFont.Font(family="Helvetica", size=15)
            labelData["font"] = ft
            labelData["fg"] = "#333333"
            labelData["justify"] = "center"
            labelData["text"] = "Data type"
            labelData.place(x=25, y=10, width=182, height=37)

            buttonQuit = btk.Button(
                popupWin1,
                text="Close",
                command=lambda: [
                    popupWin1.destroy(),
                ],
                bootstyle=DANGER,
            )
            buttonQuit.place(x=70, y=100, width=92, height=30)
        except:
            popupWin1.destroy()

        def saltTest(self):
            ##Salt test.
            data2 = self.data.filter(
                [
                    "Date and Time",
                    "Temperature Set",
                    "Temperature",
                    "Humidifier Set",
                    "Humidifier",
                ],
                axis=1,
            )
            optionTest = {
                "dewing": 0,
                "secasi": False,
                "angel": False,
                "samples": False,
                "timeabsOne": False,
                "timeabsAll": False,
                "indigo": False,
                "insight": False,
                "grafana": False,
                "agilent": False,
                "rotronic": False,
                "keithleymanager": False,
                "gasChamber": False,
                "humBath": True,
                "indexZUT": False,
                "coolingSys": False,
                "saltHum": False,
                "hum": False,
            }
            DataFormating(self.root, data2, optionTest, "Salt")

        def humidityTest(self):
            # Salt with humidity.
            try:
                data2 = self.data.filter(
                    [
                        "Date and Time",
                        "Temperature Set",
                        "Temperature",
                        "Humidifier Set",
                        "Humidifier",
                        "Humidity Set",
                        "Humidity",
                    ],
                    axis=1,
                )
                data2["Humidity Set"] = pd.to_numeric(
                    data2["Humidity Set"], errors="coerce"
                )
                data2["Humidity"] = pd.to_numeric(data2["Humidity"], errors="coerce")
                optionTest = {
                    "dewing": 0,
                    "secasi": False,
                    "angel": False,
                    "samples": False,
                    "timeabsOne": False,
                    "timeabsAll": False,
                    "indigo": False,
                    "insight": False,
                    "grafana": False,
                    "agilent": False,
                    "rotronic": False,
                    "keithleymanager": False,
                    "gasChamber": False,
                    "humBath": False,
                    "indexZUT": False,
                    "coolingSys": False,
                    "saltHum": True,
                    "hum": True,
                }
                DataFormating(self.root, data2, optionTest, "Salt")

            except KeyError:
                return Message(
                    "warning",
                    "Test option unavailable." "\nPlease select different test type.",
                    0,
                )

    def ReadZUTMichalin(self):
        # GasCorrosion data processing:
        try:
            # Data processing variants - popup:
            popupWin1 = tk.Toplevel(self.root)
            width = 490
            height = 148
            screenwidth = popupWin1.winfo_screenwidth()
            screenheight = popupWin1.winfo_screenheight()
            alignstr = "%dx%d+%d+%d" % (
                width,
                height,
                (screenwidth - width) / 2,
                (screenheight - height) / 2,
            )
            popupWin1.geometry(alignstr)
            popupWin1.resizable(width=False, height=False)
            popupWin1.attributes("-topmost", True)
            popupWin1.minsize(490, 148)
            popupWin1.maxsize(490, 148)

            GButton_51 = btk.Button(
                popupWin1,
                text="Gas Corrosion",
                command=lambda: [popupWin1.destroy(), GasChamber(self)],
                bootstyle=DARK,
            )
            GButton_51.place(x=10, y=50, width=110, height=30)

            GButton_89 = btk.Button(
                popupWin1,
                text="Splash Chamber",
                command=lambda: [popupWin1.destroy(), SplashChamber(self)],
                bootstyle=DARK,
            )
            GButton_89.place(x=130, y=50, width=110, height=30)

            GButton_991 = btk.Button(
                popupWin1,
                text="Solar Chamber",
                command=lambda: [popupWin1.destroy(), SolarChamber(self)],
                bootstyle=DARK,
            )
            GButton_991.place(x=250, y=50, width=110, height=30)

            GButton_581 = btk.Button(
                popupWin1,
                text="Cooling System",
                command=lambda: [popupWin1.destroy(), CoolingSystem(self)],
                bootstyle=DARK,
            )
            GButton_581.place(x=370, y=50, width=110, height=30)

            GLabel_861 = tk.Label(popupWin1)
            ft = tkFont.Font(family="Helvetica", size=15)
            GLabel_861["font"] = ft
            GLabel_861["fg"] = "#333333"
            GLabel_861["justify"] = "center"
            GLabel_861["text"] = "Test type:"
            GLabel_861.place(x=190, y=10, width=110, height=25)

            GButton_728 = btk.Button(
                popupWin1,
                text="Close",
                command=lambda: [
                    popupWin1.destroy(),
                ],
                bootstyle=DANGER,
            )
            GButton_728.place(x=190, y=100, width=110, height=30)
        except:
            popupWin1.destroy()

        def GasChamber(self):
            # GasChamber data processing:
            filepaths = filedialog.askopenfilenames(
                parent=self.root,
                filetypes=(
                    ("CSV Files", "*.csv"),
                    ("all files", ".*"),
                ),
                initialdir=os.getcwd(),
            )
            filepaths2 = list(filepaths)

            # Multiple data function:
            listData = []
            for filepath in filepaths2:
                try:
                    self.data = pd.read_csv(
                        filepath, sep=",", low_memory=False, encoding_errors="ignore"
                    )
                    self.data["Date"] = self.data["Date"].str.replace("-", ".")
                    self.data["Date and Time"] = self.data[["Date", "Time"]].apply(
                        " ".join, axis=1
                    )
                    self.data = self.data.drop(columns=["Date"], axis=1)
                    self.data = self.data.drop(columns=["Time"], axis=1)
                    self.data = self.data.drop(columns=["Millisecond"], axis=1)
                    self.data["Date and Time"] = pd.to_datetime(
                        self.data["Date and Time"], errors="coerce"
                    )
                    newColumnNames = {
                        "Temp set": "Temperature Set",
                        "Temp act": "Temperature IN",
                        "Humi set": "Humidity Set",
                        "Humi in act": "Humidity IN",
                        "Humi out act": "Humidity OUT",
                        "Temp out act": "Temperature OUT",
                    }
                    self.data.rename(columns=newColumnNames, inplace=True)
                    self.data["Temperature Set"] = pd.to_numeric(
                        self.data["Temperature Set"], errors="coerce"
                    )
                    self.data["Temperature IN"] = pd.to_numeric(
                        self.data["Temperature IN"], errors="coerce"
                    )
                    self.data["Temperature OUT"] = pd.to_numeric(
                        self.data["Temperature OUT"], errors="coerce"
                    )
                    self.data["Humidity Set"] = pd.to_numeric(
                        self.data["Humidity Set"], errors="coerce"
                    )
                    self.data["Humidity IN"] = pd.to_numeric(
                        self.data["Humidity IN"], errors="coerce"
                    )
                    self.data["Humidity OUT"] = pd.to_numeric(
                        self.data["Humidity OUT"], errors="coerce"
                    )
                    listData.append(self.data)
                except:
                    return Message(
                        "error",
                        "Wrong file format."
                        "\nPlease contact with: piotr.klys92@gmail.com",
                        0,
                    )

            try:
                self.data = pd.concat(listData, axis=0, ignore_index=False)
            except:
                return Message("warning", "File was not selected", 0)

            self.data = self.data.sort_values(by="Date and Time", ascending=True)
            data2 = self.data.filter(
                [
                    "Date and Time",
                    "Temperature Set",
                    "Temperature IN",
                    "Temperature OUT",
                    "Humidity Set",
                    "Humidity IN",
                    "Humidity OUT",
                ],
                axis=1,
            )
            optionTest = {
                "dewing": 0,
                "secasi": False,
                "angel": False,
                "samples": False,
                "timeabsOne": False,
                "timeabsAll": False,
                "indigo": False,
                "insight": False,
                "grafana": False,
                "agilent": False,
                "rotronic": False,
                "keithleymanager": False,
                "gasChamber": True,
                "humBath": False,
                "indexZUT": False,
                "coolingSys": False,
                "saltHum": False,
                "hum": True,
            }
            DataFormating(self.root, data2, optionTest, "GasChamber")

        def SplashChamber(self):
            # Splash data processing:
            filepaths = filedialog.askopenfilenames(
                parent=self.root,
                filetypes=(
                    ("CSV Files", "*.csv"),
                    ("all files", ".*"),
                ),
                initialdir=os.getcwd(),
            )
            filepaths2 = list(filepaths)

            # Multiple data function:
            listData = []
            for filepath in filepaths2:
                try:
                    self.data = pd.read_csv(
                        filepath, sep=",", low_memory=False, encoding_errors="ignore"
                    )
                    self.data["Date"] = self.data["Date"].str.replace("-", ".")
                    self.data["Date and Time"] = self.data[["Date", "Time"]].apply(
                        " ".join, axis=1
                    )
                    self.data = self.data.drop(columns=["Date"], axis=1)
                    self.data = self.data.drop(columns=["Time"], axis=1)
                    self.data = self.data.drop(columns=["Millisecond"], axis=1)
                    self.data["Date and Time"] = pd.to_datetime(
                        self.data["Date and Time"], errors="coerce"
                    )
                    newColumnNames = {
                        "Chamber set": "Temperature Set",
                        "Chamber act": "Temperature",
                        "Bath set": "Bath Temp Set",
                        "Bath act": "Bath Temp",
                        "on/off (1/0)": "On/Off",
                    }
                    self.data.rename(columns=newColumnNames, inplace=True)
                    self.data["Temperature Set"] = pd.to_numeric(
                        self.data["Temperature Set"], errors="coerce"
                    )
                    self.data["Temperature"] = pd.to_numeric(
                        self.data["Temperature"], errors="coerce"
                    )
                    self.data["Bath Temp Set"] = pd.to_numeric(
                        self.data["Bath Temp Set"], errors="coerce"
                    )
                    self.data["Bath Temp"] = pd.to_numeric(
                        self.data["Bath Temp"], errors="coerce"
                    )
                    self.data["On/Off"] = pd.to_numeric(
                        self.data["On/Off"], errors="coerce"
                    )
                    listData.append(self.data)
                except:
                    return Message(
                        "error",
                        "Wrong file format."
                        "\nPlease contact with: piotr.klys92@gmail.com",
                        0,
                    )
            try:
                self.data = pd.concat(listData, axis=0, ignore_index=False)
            except:
                return Message("warning", "File was not selected", 0)

            self.data = self.data.sort_values(by="Date and Time", ascending=True)
            data2 = self.data.filter(
                [
                    "Date and Time",
                    "Temperature Set",
                    "Temperature",
                    "Bath Temp Set",
                    "Bath Temp",
                    "On/Off",
                ],
                axis=1,
            )
            optionTest = {
                "dewing": 0,
                "secasi": False,
                "angel": False,
                "samples": False,
                "timeabsOne": False,
                "timeabsAll": False,
                "indigo": False,
                "insight": False,
                "grafana": False,
                "agilent": False,
                "rotronic": False,
                "keithleymanager": False,
                "gasChamber": False,
                "humBath": True,
                "indexZUT": True,
                "coolingSys": False,
                "saltHum": False,
                "hum": False,
            }
            DataFormating(self.root, data2, optionTest, "SplashChamber")

        def SolarChamber(self):
            # Solar data processing:
            filepaths = filedialog.askopenfilenames(
                parent=self.root,
                filetypes=(
                    ("CSV Files", "*.csv"),
                    ("all files", ".*"),
                ),
                initialdir=os.getcwd(),
            )
            filepaths2 = list(filepaths)

            # Multiple data function:
            listData = []
            for filepath in filepaths2:
                try:
                    self.data = pd.read_csv(
                        filepath, sep=",", low_memory=False, encoding_errors="ignore"
                    )
                    self.data["Date"] = self.data["Date"].str.replace("-", ".")
                    self.data["Date and Time"] = self.data[["Date", "Time"]].apply(
                        " ".join, axis=1
                    )
                    self.data = self.data.drop(columns=["Date"], axis=1)
                    self.data = self.data.drop(columns=["Time"], axis=1)
                    self.data = self.data.drop(columns=["Millisecond"], axis=1)
                    self.data["Date and Time"] = pd.to_datetime(
                        self.data["Date and Time"], errors="coerce"
                    )
                    newColumnNames = {
                        "Temp set": "Temperature Set",
                        "Temp act": "Temperature",
                        "Humi set": "Humidity Set",
                        "Humi act": "Humidity",
                    }
                    self.data.rename(columns=newColumnNames, inplace=True)
                    self.data["Temperature Set"] = pd.to_numeric(
                        self.data["Temperature Set"], errors="coerce"
                    )
                    self.data["Temperature"] = pd.to_numeric(
                        self.data["Temperature"], errors="coerce"
                    )
                    self.data["Humidity Set"] = pd.to_numeric(
                        self.data["Humidity Set"], errors="coerce"
                    )
                    self.data["Humidity"] = pd.to_numeric(
                        self.data["Humidity"], errors="coerce"
                    )
                    self.data["UV"] = pd.to_numeric(self.data["UV"], errors="coerce")
                    self.data["On/Off"] = pd.to_numeric(
                        self.data["On/Off"], errors="coerce"
                    )
                    listData.append(self.data)
                except:
                    return Message(
                        "error",
                        "Wrong file format."
                        "\nPlease contact with: piotr.klys92@gmail.com",
                        0,
                    )

            try:
                self.data = pd.concat(listData, axis=0, ignore_index=False)
            except:
                return Message("warning", "File was not selected", 0)

            self.data = self.data.sort_values(by="Date and Time", ascending=True)
            self.data["Temperature"] = pd.to_numeric(
                self.data["Temperature"], errors="coerce"
            )
            self.data["Humidity"] = pd.to_numeric(
                self.data["Humidity"], errors="coerce"
            )
            data2 = self.data.filter(
                [
                    "Date and Time",
                    "Temperature Set",
                    "Temperature",
                    "Humidity Set",
                    "Humidity",
                    "UV",
                    "On/Off",
                ],
                axis=1,
            )
            optionTest = {
                "dewing": 0,
                "secasi": False,
                "angel": False,
                "samples": False,
                "timeabsOne": False,
                "timeabsAll": False,
                "indigo": False,
                "insight": False,
                "grafana": False,
                "agilent": False,
                "rotronic": False,
                "keithleymanager": False,
                "gasChamber": False,
                "humBath": False,
                "indexZUT": True,
                "coolingSys": False,
                "saltHum": False,
                "hum": True,
            }
            DataFormating(self.root, data2, optionTest, "SolarChamber")

        def CoolingSystem(self):
            filepaths = filedialog.askopenfilenames(
                parent=self.root,
                filetypes=(
                    ("CSV Files", "*.csv"),
                    ("all files", ".*"),
                ),
                initialdir=os.getcwd(),
            )
            filepaths2 = list(filepaths)

            # Multiple data function:
            listData = []
            for filepath in filepaths2:
                try:
                    self.data = pd.read_csv(
                        filepath, sep=",", low_memory=False, encoding_errors="ignore"
                    )
                    self.data["Date"] = self.data["Date"].str.replace("-", ".")
                    self.data["Date and Time"] = self.data[["Date", "Time"]].apply(
                        " ".join, axis=1
                    )
                    self.data = self.data.drop(columns=["Date"], axis=1)
                    self.data = self.data.drop(columns=["Time"], axis=1)
                    self.data = self.data.drop(columns=["Millisecond"], axis=1)
                    self.data["Date and Time"] = pd.to_datetime(
                        self.data["Date and Time"], errors="coerce"
                    )
                    newColumnNames = {
                        "Temp IN": "Temperature In",
                        "Temp OUT": "Temperature Out",
                        "Temp Cham": "Temperature Chamber",
                    }
                    self.data.rename(columns=newColumnNames, inplace=True)
                    self.data["Temperature In"] = pd.to_numeric(
                        self.data["Temperature In"], errors="coerce"
                    )
                    self.data["Temperature Out"] = pd.to_numeric(
                        self.data["Temperature Out"], errors="coerce"
                    )
                    self.data["Temperature Chamber"] = pd.to_numeric(
                        self.data["Temperature Chamber"], errors="coerce"
                    )
                    self.data["Pressure"] = pd.to_numeric(
                        self.data["Pressure"], errors="coerce"
                    )
                    self.data["Flow"] = pd.to_numeric(
                        self.data["Flow"], errors="coerce"
                    )
                    listData.append(self.data)
                except:
                    return Message(
                        "error",
                        "Wrong file format."
                        "\nPlease contact with: piotr.klys92@gmail.com",
                        0,
                    )

            try:
                self.data = pd.concat(listData, axis=0, ignore_index=False)
            except:
                return Message("warning", "File was not selected", 0)

            self.data = self.data.sort_values(by="Date and Time", ascending=True)
            self.data = pd.DataFrame(self.data)
            data2 = self.data.filter(
                [
                    "Date and Time",
                    "Temperature In",
                    "Temperature Out",
                    "Temperature Chamber",
                    "Pressure",
                    "Flow",
                ],
                axis=1,
            )
            optionTest = {
                "dewing": 0,
                "secasi": False,
                "angel": False,
                "samples": False,
                "timeabsOne": False,
                "timeabsAll": False,
                "indigo": False,
                "insight": False,
                "grafana": False,
                "agilent": False,
                "rotronic": False,
                "keithleymanager": False,
                "gasChamber": False,
                "humBath": False,
                "indexZUT": False,
                "coolingSys": True,
                "saltHum": False,
                "hum": False,
            }
            DataFormating(self.root, data2, optionTest, "CoolingSystem")