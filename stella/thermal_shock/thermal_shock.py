import pandas as pd
import tkinter as tk
from tkinter import filedialog
import tkinter.font as tkFont
import ttkbootstrap as btk
from ttkbootstrap.constants import *
from ..utility.message import Message
from ..data.dataFormating import DataFormating


class ThermalShockTests:
    def __init__(self, root, option):
        self.root = root
        if option == 0:
            self.combined()

    def combined(self):
        # Thermal Shocks data processing:
        filepaths = filedialog.askopenfilenames(
            parent=self.root,
            filetypes=(("all files", ".*"),),
        )
        filepaths = list(filepaths)
        list_data = []
        for filepath in filepaths:
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
                    new_column_names = {
                        "Date and time       ": "Date and Time",
                        "Temp hot": "Temperature Hot Set",
                        "Temp hot.1": "Temperature Hot",
                        "TempHotChamb": "Temperature Hot Set",
                        "TempHotChamb.1": "Temperature Hot",
                        "Temp cold": "Temperature Cold Set",
                        "Temp cold.1": "Temperature Cold",
                        "TempColdChamber": "Temperature Cold Set",
                        "TempColdChamber.1": "Temperature Cold",
                        "Temp cage.1": "Temperature Basket",
                        "Temp. basket.1": "Temperature Basket",
                        "Lift up   ": "Basket Position",
                        "Basket up ": "Basket Position",
                    }
                    self.data.rename(columns=new_column_names, inplace=True)
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
                else:
                    self.data = pd.read_csv(
                        filepath,
                        header=3,
                        sep=";",
                        low_memory=False,
                        encoding_errors="ignore",
                    )
                    new_column_names = {"Unnamed: 0": "Date and Time"}
                    self.data.rename(columns=new_column_names, inplace=True)
                    self.data.drop(self.data.index[[0]], axis=0, inplace=True)
                    self.data = pd.DataFrame(self.data)
                    new_column_names = {
                        "A:Komora gorac": "Temperature Hot",
                        "N:Komora gorac": "Temperature Hot Set",
                        "A:Komora zimna": "Temperature Cold",
                        "N:Komora zimna": "Temperature Cold Set",
                        "A:Temp. windy ": "Temperature Basket",
                        "A:Temp. Basket": "Temperature Basket",
                        "A:TempLift ": "Temperature Basket",
                        "A:Basket": "Temperature Basket",
                        "A:Poloz. windy": "Basket Position",
                        "A:Basketpositi": "Basket Position",
                        "Basket Up/Down": "Basket Position",
                        "Lift Up  ": "Basket Position",
                    }
                    self.data.rename(columns=new_column_names, inplace=True)
                    self.data = self.data.rename(
                        columns={
                            i: "Temperature Hot"
                            for i in self.data.columns
                            if i.startswith("A:Hot")
                        }
                    )
                    self.data = self.data.rename(
                        columns={
                            i: "Temperature Hot Set"
                            for i in self.data.columns
                            if i.startswith("N:Hot")
                        }
                    )
                    self.data = self.data.rename(
                        columns={
                            i: "Temperature Cold"
                            for i in self.data.columns
                            if i.startswith("A:Cold")
                        }
                    )
                    self.data = self.data.rename(
                        columns={
                            i: "Temperature Cold Set"
                            for i in self.data.columns
                            if i.startswith("N:Cold")
                        }
                    )
                    self.data["Date and Time"] = pd.to_datetime(
                        self.data["Date and Time"], dayfirst=True, errors="coerce"
                    )
                self.data["Date and Time"] = pd.to_datetime(
                    self.data["Date and Time"], errors="coerce"
                )
                self.data = self.data.filter(
                    [
                        "Date and Time",
                        "Temperature Hot Set",
                        "Temperature Hot",
                        "Temperature Cold Set",
                        "Temperature Cold",
                        "Temperature Basket",
                        "Basket Position",
                    ],
                    axis=1,
                )
                list_data.append(self.data)
            except ValueError:
                return Message(
                    "error",
                    "Wrong file format."
                    "\nPlease contact with: piotr.klys92@gmail.com",
                    0,
                )

        try:
            self.data = pd.concat(list_data, axis=0, ignore_index=False)
            self.data = self.data.sort_values(by="Date and Time", ascending=True)
            option_test = {
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
                "saltHum": False,
                "hum": False,
            }
            self.TestType(option_test, "ThermalShock")
        except:
            return Message(
                "error",
                "Wrong file format." "\nPlease contact with: piotr.klys92@gmail.com",
                0,
            )

    def TestType(self, option_test, name):
        try:
            self.data["Temperature Hot"] = self.data["Temperature Hot"].str.replace(
                ",", ".", regex=True
            )
            self.data["Temperature Hot Set"] = self.data[
                "Temperature Hot Set"
            ].str.replace(",", ".", regex=True)
            self.data["Temperature Cold"] = self.data["Temperature Cold"].str.replace(
                ",", ".", regex=True
            )
            self.data["Temperature Cold Set"] = self.data[
                "Temperature Cold Set"
            ].str.replace(",", ".", regex=True)
        except:
            pass
        self.data["Date and Time"] = pd.to_datetime(
            self.data["Date and Time"], errors="coerce"
        )
        self.data["Temperature Hot Set"] = pd.to_numeric(
            self.data["Temperature Hot Set"], errors="coerce"
        )
        self.data["Temperature Hot"] = pd.to_numeric(
            self.data["Temperature Hot"], errors="coerce"
        )
        self.data["Temperature Cold Set"] = pd.to_numeric(
            self.data["Temperature Cold Set"], errors="coerce"
        )
        self.data["Temperature Cold"] = pd.to_numeric(
            self.data["Temperature Cold"], errors="coerce"
        )
        try:
            self.data["Temperature Basket"] = pd.to_numeric(
                self.data["Temperature Basket"], errors="coerce"
            )
            self.data["Basket Position"] = pd.to_numeric(
                self.data["Basket Position"], errors="coerce"
            )
        except:
            pass
        try:
            # Data processing variants - popup:
            popup_win_1 = tk.Toplevel(self.root)
            width = 325
            height = 148
            screenwidth = popup_win_1.winfo_screenwidth()
            screenheight = popup_win_1.winfo_screenheight()
            alignstr = "%dx%d+%d+%d" % (
                width,
                height,
                (screenwidth - width) / 2,
                (screenheight - height) / 2,
            )
            popup_win_1.geometry(alignstr)
            popup_win_1.resizable(width=False, height=False)
            popup_win_1.attributes("-topmost", True)
            popup_win_1.minsize(325, 148)
            popup_win_1.maxsize(325, 148)

            button_temperature = btk.Button(
                popup_win_1,
                text="Temperature",
                command=lambda: [
                    popup_win_1.destroy(),
                    temperature(self, self.data, option_test, name),
                ],
                bootstyle=DARK,
            )
            button_temperature.place(x=15, y=50, width=92, height=30)

            button_basket_position = btk.Button(
                popup_win_1,
                text="Basket Pos.",
                command=lambda: [
                    popup_win_1.destroy(),
                    basket_position(self, self.data, option_test, name),
                ],
                bootstyle=DARK,
            )
            button_basket_position.place(x=115, y=50, width=92, height=30)

            button_basket = btk.Button(
                popup_win_1,
                text="Basket Temp.",
                command=lambda: [
                    popup_win_1.destroy(),
                    basket(self, self.data, option_test, name),
                ],
                bootstyle=DARK,
            )
            button_basket.place(x=215, y=50, width=92, height=30)

            label_data = tk.Label(popup_win_1)
            label_data["anchor"] = "n"
            ft = tkFont.Font(family="Helvetica", size=15)
            label_data["font"] = ft
            label_data["fg"] = "#333333"
            label_data["justify"] = "center"
            label_data["text"] = "Data type"
            label_data.place(x=20, y=10, width=280, height=36)

            button_close = btk.Button(
                popup_win_1,
                text="Close",
                command=lambda: [
                    popup_win_1.destroy(),
                ],
                bootstyle=DANGER,
            )
            button_close.place(x=115, y=100, width=92, height=30)
        except:
            popup_win_1.destroy()

        def temperature(self, data, option_test, name):
            # Thermal Shock chamber temperature only.
            data2 = data.filter(
                [
                    "Date and Time",
                    "Temperature Hot Set",
                    "Temperature Hot",
                    "Temperature Cold Set",
                    "Temperature Cold",
                ],
                axis=1,
            )
            DataFormating(self.root, data2, option_test, name)

        def basket_position(self, data, option_test, name):
            # Thermal Shock chamber temperature with basket position digital output.
            data = data.filter(
                [
                    "Date and Time",
                    "Temperature Hot Set",
                    "Temperature Hot",
                    "Temperature Cold Set",
                    "Temperature Cold",
                    "Basket Position",
                ],
                axis=1,
            )
            data2 = pd.DataFrame(data)
            DataFormating(self.root, data2, option_test, name)

        def basket(self, data, option_test, name):
            # Thermal Shock chamber temperature with basket position digital output and basket temperature.
            data = data.filter(
                [
                    "Date and Time",
                    "Temperature Hot Set",
                    "Temperature Hot",
                    "Temperature Cold Set",
                    "Temperature Cold",
                    "Temperature Basket",
                    "Basket Position",
                ],
                axis=1,
            )
            data2 = pd.DataFrame(data)
            DataFormating(self.root, data2, option_test, name)
