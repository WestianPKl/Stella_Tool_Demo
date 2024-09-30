import pandas as pd
import tkinter as tk
import os
from tkinter import filedialog
import tkinter.font as tkFont
import ttkbootstrap as btk
from ttkbootstrap.constants import *
from ..utility.message import Message
from ..data.dataFormating import DataFormating


class ClimaticTests:
    def __init__(self, root, option):
        self.root = root
        if option == 0:
            self.Combined()

    def Combined(self):
        filepaths = filedialog.askopenfilenames(
            parent=self.root, filetypes=(("all files", ".*"),), initialdir=os.getcwd()
        )
        filepaths2 = list(filepaths)
        listData = []
        for filepath in filepaths2:
            try:
                if filepath.endswith(".csv"):
                    self.data = pd.read_csv(
                        filepath, sep=";", low_memory=False, encoding_errors="ignore"
                    )
                    self.data = self.data.rename(
                        columns={
                            i: "Temperature"
                            for i in self.data.columns
                            if i.startswith("A:Temper")
                        }
                    )
                    newColumnNames = {
                        "Date and time       ": "Date and Time",
                        "Temperature": "Temperature Set",
                        "Temperature.1": "Temperature",
                        "Humidity": "Humidity Set",
                        "Humidity.1": "Humidity",
                        "Bath temp.": "Bath Temp Set",
                        "Bath temp..1": "Bath Temp",
                        "Temper": "Temperature Set",
                        "Temper.1": "Temperature",
                        "Feuchte": "Humidity Set",
                        "Feuchte.1": "Humidity",
                        "Bath temp": "Bath Temp Set",
                        "Bath temp.1": "Bath Temp",
                        "BathTemp": "Bath Temp Set",
                        "BathTemp.1": "Bath Temp",
                    }
                    self.data.rename(columns=newColumnNames, inplace=True)
                    self.data.drop(
                        self.data.index[
                            [
                                0,
                                1,
                            ]
                        ],
                        axis=0,
                        inplace=True,
                    )
                elif filepath.endswith(".asc"):
                    self.data = pd.read_csv(
                        filepath,
                        header=2,
                        sep="	",
                        low_memory=False,
                        encoding_errors="ignore",
                    )
                    newColumnNames = {
                        "Date Time": "Date and Time",
                        "T setpoint": "Temperature Set",
                        "T Dry bulb": "Temperature",
                        "Relative humidity setpoint": "Humidity Set",
                        "Relative humidity": "Humidity",
                    }
                    self.data.rename(columns=newColumnNames, inplace=True)
                    self.data = self.data.filter(
                        [
                            "Date and Time",
                            "Temperature Set",
                            "Temperature",
                            "Humidity Set",
                            "Humidity",
                        ],
                        axis=1,
                    )
                    self.data["Date and Time"] = pd.to_datetime(
                        self.data["Date and Time"], dayfirst=True, errors="coerce"
                    )
                    try:
                        self.data["Temperature Set"] = self.data[
                            "Temperature Set"
                        ].str.replace("+", "", regex=True)
                        self.data["Temperature"] = self.data["Temperature"].str.replace(
                            "+", "", regex=True
                        )
                        self.data["Humidity Set"] = self.data[
                            "Humidity Set"
                        ].str.replace("+", "", regex=True)
                        self.data["Humidity"] = self.data["Humidity"].str.replace(
                            "+", "", regex=True
                        )
                    except:
                        pass
                elif filepath.endswith(".txt"):
                    self.data = pd.read_csv(
                        filepath,
                        header=2,
                        sep="\t",
                        low_memory=False,
                        encoding_errors="ignore",
                    )
                    newColumnNames = {
                        "Date / Time": "Date and Time",
                        "Setpoint C": "Temperature Set",
                        "Value C": "Temperature",
                        "Setpoint Humidity rH %": "Humidity Set",
                        "Value Humidity rH %": "Humidity",
                    }
                    self.data.rename(columns=newColumnNames, inplace=True)
                    self.data["Date and Time"] = pd.to_datetime(
                        self.data["Date and Time"], dayfirst=True, errors="coerce"
                    )
                elif filepath.endswith(".xls"):
                    self.data = pd.read_csv(
                        filepath,
                        skiprows=2,
                        header=None,
                        sep=";",
                        low_memory=False,
                        encoding_errors="ignore",
                    )
                    newColumnNames = {0: "Time", 1: "Temperature", 2: "Temperature Set"}
                    self.data.rename(columns=newColumnNames, inplace=True)
                else:
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
                    self.data = self.data.rename(
                        columns={
                            i: "Temperature"
                            for i in self.data.columns
                            if i.startswith("A:Temper")
                        }
                    )
                    self.data = self.data.rename(
                        columns={
                            i: "Temperature Set"
                            for i in self.data.columns
                            if i.startswith("N:Temper")
                        }
                    )
                    self.data = self.data.rename(
                        columns={
                            i: "Humidity"
                            for i in self.data.columns
                            if i.startswith("A:Hum")
                        }
                    )
                    self.data = self.data.rename(
                        columns={
                            i: "Humidity Set"
                            for i in self.data.columns
                            if i.startswith("N:Hum")
                        }
                    )
                    self.data = self.data.rename(
                        columns={
                            i: "Bath Temp"
                            for i in self.data.columns
                            if i.startswith("T.bath     ")
                        }
                    )
                    self.data = self.data.rename(
                        columns={
                            i: "Humidity"
                            for i in self.data.columns
                            if i.startswith("A:rel.")
                        }
                    )
                    self.data = self.data.rename(
                        columns={
                            i: "Humidity Set"
                            for i in self.data.columns
                            if i.startswith("N:rel.")
                        }
                    )
                    self.data = self.data.rename(
                        columns={
                            i: "Bath Temp"
                            for i in self.data.columns
                            if i.startswith("A:T.bath   ")
                        }
                    )
                    newColumnNames = {
                        "A:Wilgotnosc  ": "Humidity",
                        "N:Wilgotnosc  ": "Humidity Set",
                        "A:Temp. kapiel": "Bath Temp",
                        "A:temper.  ": "Temperature",
                        "N:temper.  ": "Temperature Set",
                    }
                    self.data.rename(columns=newColumnNames, inplace=True)
                    self.data["Date and Time"] = pd.to_datetime(
                        self.data["Date and Time"], dayfirst=True, errors="coerce"
                    )
                try:
                    self.data["Date and Time"] = pd.to_datetime(
                        self.data["Date and Time"], errors="coerce"
                    )
                    self.data = self.data.filter(
                        [
                            "Date and Time",
                            "Temperature Set",
                            "Temperature",
                            "Humidity Set",
                            "Humidity",
                            "Bath Temp",
                        ],
                        axis=1,
                    )
                except:
                    pass
                listData.append(self.data)
            except ValueError:
                return Message(
                    "error",
                    "Wrong file format."
                    "\nPlease contact with: piotr.klys92@gmail.com",
                    0,
                )
        try:
            self.data = pd.concat(listData, axis=0, ignore_index=False)
            try:
                self.data = self.data.sort_values(by="Date and Time", ascending=True)
                optionTest = {
                    "dewing": 1,
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
                    "saltHum": False,
                }
            except:
                self.data = self.data.sort_values(by="Time", ascending=True)
                optionTest = {
                    "dewing": 0,
                    "secasi": True,
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
                    "saltHum": False,
                }
            self.TestType(optionTest, "Climatic")
        except:
            return Message(
                "error",
                "Wrong file format." "\nPlease contact with: piotr.klys92@gmail.com",
                0,
            )

    def TestType(self, optionTest, name):
        # Popup GUI for Climatic Chambers:
        try:
            self.data["Temperature"] = self.data["Temperature"].str.replace(
                ",", ".", regex=True
            )
            self.data["Temperature Set"] = self.data["Temperature Set"].str.replace(
                ",", ".", regex=True
            )
        except:
            pass
        try:
            self.data["Humidity Set"] = self.data["Humidity Set"].str.replace(
                ",", ".", regex=True
            )
            self.data["Humidity"] = self.data["Humidity"].str.replace(
                ",", ".", regex=True
            )
        except:
            pass
        try:
            self.data["Bath Temp"] = self.data["Bath Temp"].str.replace(
                ",", ".", regex=True
            )
        except:
            pass
        try:
            self.data["Date and Time"] = pd.to_datetime(
                self.data["Date and Time"], errors="coerce"
            )
        except:
            pass
        self.data["Temperature Set"] = pd.to_numeric(
            self.data["Temperature Set"], errors="coerce"
        )
        self.data["Temperature"] = pd.to_numeric(
            self.data["Temperature"], errors="coerce"
        )
        try:
            self.data["Humidity Set"] = pd.to_numeric(
                self.data["Humidity Set"], errors="coerce"
            )
            self.data["Humidity"] = pd.to_numeric(
                self.data["Humidity"], errors="coerce"
            )
        except:
            pass
        try:
            if optionTest["dewing"] == 1:
                self.popupWin1 = tk.Toplevel(self.root)
                width = 335
                height = 148
                screenwidth = self.popupWin1.winfo_screenwidth()
                screenheight = self.popupWin1.winfo_screenheight()
                alignstr = "%dx%d+%d+%d" % (
                    width,
                    height,
                    (screenwidth - width) / 2,
                    (screenheight - height) / 2,
                )
                self.popupWin1.geometry(alignstr)
                self.popupWin1.resizable(width=False, height=False)
                self.popupWin1.attributes("-topmost", True)
                self.popupWin1.minsize(335, 148)
                self.popupWin1.maxsize(335, 148)

                buttonTemperature = btk.Button(
                    self.popupWin1,
                    text="Temperature",
                    command=lambda: [
                        self.popupWin1.destroy(),
                        Temperature(self, self.data, optionTest, name),
                    ],
                    bootstyle=DARK,
                )
                buttonTemperature.place(x=20, y=50, width=92, height=30)

                buttonHumidity = btk.Button(
                    self.popupWin1,
                    text="Humidity",
                    command=lambda: [
                        self.popupWin1.destroy(),
                        Humidity(self, self.data, optionTest, name),
                    ],
                    bootstyle=DARK,
                )
                buttonHumidity.place(x=120, y=50, width=92, height=30)

                buttonDewing = btk.Button(
                    self.popupWin1,
                    text="Dewing",
                    command=lambda: [
                        self.popupWin1.destroy(),
                        Dewing(self, self.data, optionTest, name),
                    ],
                    bootstyle=DARK,
                )
                buttonDewing.place(x=220, y=50, width=92, height=30)

                labelData = tk.Label(self.popupWin1)
                labelData["anchor"] = "n"
                ft = tkFont.Font(family="Helvetica", size=15)
                labelData["font"] = ft
                labelData["fg"] = "#333333"
                labelData["justify"] = "center"
                labelData["text"] = "Test type"
                labelData.place(x=25, y=10, width=280, height=36)

                buttonQuit = btk.Button(
                    self.popupWin1,
                    text="Close",
                    command=lambda: [self.popupWin1.destroy()],
                    bootstyle=DANGER,
                )
                buttonQuit.place(x=120, y=100, width=92, height=30)

            elif optionTest["dewing"] == 0:
                self.popupWin1 = tk.Toplevel(self.root)
                width = 243
                height = 147
                screenwidth = self.popupWin1.winfo_screenwidth()
                screenheight = self.popupWin1.winfo_screenheight()
                alignstr = "%dx%d+%d+%d" % (
                    width,
                    height,
                    (screenwidth - width) / 2,
                    (screenheight - height) / 2,
                )
                self.popupWin1.geometry(alignstr)
                self.popupWin1.resizable(width=False, height=False)
                self.popupWin1.attributes("-topmost", True)
                self.popupWin1.minsize(243, 147)
                self.popupWin1.maxsize(243, 147)

                buttonTemperature = btk.Button(
                    self.popupWin1,
                    text="Temperature",
                    command=lambda: [
                        self.popupWin1.destroy(),
                        Temperature(self, self.data, optionTest, name),
                    ],
                    bootstyle=DARK,
                )
                buttonTemperature.place(x=10, y=50, width=102, height=30)

                buttonHumidity = btk.Button(
                    self.popupWin1,
                    text="Humidity",
                    command=lambda: [
                        self.popupWin1.destroy(),
                        Humidity(self, self.data, optionTest, name),
                    ],
                    bootstyle=DARK,
                )
                buttonHumidity.place(x=130, y=50, width=102, height=30)

                labelData = tk.Label(self.popupWin1)
                labelData["anchor"] = "n"
                ft = tkFont.Font(family="Helvetica", size=15)
                labelData["font"] = ft
                labelData["fg"] = "#333333"
                labelData["justify"] = "center"
                labelData["text"] = "Test type"
                labelData.place(x=30, y=10, width=182, height=37)

                buttonQuit = btk.Button(
                    self.popupWin1,
                    text="Close",
                    command=lambda: [self.popupWin1.destroy()],
                    bootstyle=DANGER,
                )
                buttonQuit.place(x=80, y=100, width=82, height=30)
        except:
            self.popupWin1.destroy()

        def Temperature(self, data, optionTest, name):
            # For temperature tests - only temperature values will be exported to excel.
            data2 = data.filter(
                ["Time", "Date and Time", "Temperature Set", "Temperature"], axis=1
            )
            optionTest["hum"] = False
            DataFormating(self.root, data2, optionTest, name)

        def Humidity(self, data, optionTest, name):
            # For humidity tests - only temperature and humidity values will be exported to excel.
            try:
                data2 = data.filter(
                    [
                        "Time",
                        "Date and Time",
                        "Temperature Set",
                        "Temperature",
                        "Humidity Set",
                        "Humidity",
                    ],
                    axis=1,
                )
                optionTest["hum"] = True
                if (
                    not "Humidity" in data2.columns
                    and not "Humidity Set" in data2.columns
                ):
                    optionTest["hum"] = False
                DataFormating(self.root, data2, optionTest, name)

            except KeyError:
                Message(
                    "warning",
                    "Test option unavailable." "\nPlease select different test type.",
                    0,
                )

        def Dewing(self, data, optionTest, name):
            # For dewing tests - temperature, humidity and water bath temperature values will be exported to excel.
            try:
                data2 = data.filter(
                    [
                        "Date and Time",
                        "Temperature Set",
                        "Temperature",
                        "Humidity Set",
                        "Humidity",
                        "Bath Temp",
                    ],
                    axis=1,
                )
                data2["Bath Temp"] = pd.to_numeric(data2["Bath Temp"], errors="coerce")
                optionTest["hum"] = True
                if (
                    not "Humidity" in data2.columns
                    and not "Humidity Set" in data2.columns
                ):
                    optionTest["hum"] = False
                if not "Bath Temp" in data2.columns:
                    optionTest["dewing"] = 0
                DataFormating(self.root, data2, optionTest, name)

            except KeyError:
                Message(
                    "warning",
                    "Test option unavailable." "\nPlease select different test type.",
                    0,
                )
