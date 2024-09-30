import pandas as pd
import tkinter as tk
import os
from tkinter import filedialog
import tkinter.font as tkFont
import ttkbootstrap as btk
from ttkbootstrap.constants import *
from ..utility.message import Message
from ..data.dataFormating import DataFormating


class Loggers:
    def __init__(self, root, option):
        self.root = root
        if option == 0:
            self.ReadRotronic()
        elif option == 1:
            self.ReadVaisala()
        elif option == 2:
            self.ReadKeithley()
        elif option == 3:
            self.ReadGrafana()
        elif option == 4:
            self.ReadAgilent()

    def ReadRotronic(self):
        # Rotronic data processing:
        filepaths = filedialog.askopenfilenames(
            parent=self.root,
            filetypes=(
                ("Excel Files", "*.xls"),
                ("all files", ".*"),
            ),
            initialdir=os.getcwd(),
        )
        filepaths2 = list(filepaths)
        # try:
        # Multiple data function:
        listData = []
        for filepath in filepaths2:
            indexNumber = str(filepaths2.index(filepath))
            DT = "Date and Time Probe "
            T = "Temperature Probe "
            H = "Humidity Probe "
            try:
                self.data = pd.read_csv(
                    filepath,
                    skiprows=25,
                    sep="	",
                    parse_dates=[["Date    ", "Time    "]],
                    low_memory=False,
                    encoding_errors="ignore",
                )
                self.data = self.data.drop(self.data.index[0])
                try:
                    self.data = self.data.drop(columns=["none"], axis=1)
                except:
                    pass
                newColumnNames = {
                    "Date    _Time    ": DT + indexNumber,
                    "Temperature": T + indexNumber,
                    "Humidity": H + indexNumber,
                }
                self.data.rename(columns=newColumnNames, inplace=True)
            except:
                try:
                    self.data = pd.read_csv(
                        filepath,
                        skiprows=41,
                        sep="	",
                        parse_dates=[["Date    ", "Time    "]],
                        low_memory=False,
                        encoding_errors="ignore",
                    )
                    self.data = self.data.drop(self.data.index[0])
                    try:
                        self.data = self.data.drop(columns=["none"], axis=1)
                    except:
                        pass
                    newColumnNames = {
                        "Date    _Time    ": DT + indexNumber,
                        "Temperature": T + indexNumber,
                        "Humidity": H + indexNumber,
                    }
                    self.data.rename(columns=newColumnNames, inplace=True)
                except:
                    self.data = pd.read_csv(
                        filepath,
                        skiprows=30,
                        sep="	",
                        parse_dates=[["Date    ", "Time    "]],
                        low_memory=False,
                        encoding_errors="ignore",
                    )
                    self.data = self.data.drop(self.data.index[0])
                    try:
                        self.data = self.data.drop(columns=["none"], axis=1)
                    except:
                        pass
                    newColumnNames = {
                        "Date    _Time    ": DT + indexNumber,
                        "Temperature": T + indexNumber,
                        "Humidity": H + indexNumber,
                    }
                    self.data.rename(columns=newColumnNames, inplace=True)
            self.data[T + indexNumber] = pd.to_numeric(
                self.data[T + indexNumber], errors="coerce"
            )
            self.data[H + indexNumber] = pd.to_numeric(
                self.data[H + indexNumber], errors="coerce"
            )
            self.data = self.data.dropna(subset=[H + indexNumber, T + indexNumber])
            listData.append(self.data)
        try:
            self.data = pd.concat(listData, axis=1, ignore_index=False)
        except:
            return Message("warning", "File was not selected", 0)
        try:
            # Data processing variants - popup:
            popupWin1 = tk.Toplevel(self.root)
            width = 223
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
            popupWin1.minsize(223, 147)
            popupWin1.maxsize(223, 147)

            buttonTemperature = btk.Button(
                popupWin1,
                text="dd.mm.yyyy",
                command=lambda: [popupWin1.destroy(), DayFirst(self)],
                bootstyle=DARK,
            )
            buttonTemperature.place(x=15, y=50, width=92, height=30)

            buttonHumidity = btk.Button(
                popupWin1,
                text="yyyy.mm.dd",
                command=lambda: [popupWin1.destroy(), YearFirst(self)],
                bootstyle=DARK,
            )
            buttonHumidity.place(x=115, y=50, width=92, height=30)

            labelData = tk.Label(popupWin1)
            labelData["anchor"] = "n"
            ft = tkFont.Font(family="Helvetica", size=15)
            labelData["font"] = ft
            labelData["fg"] = "#333333"
            labelData["justify"] = "center"
            labelData["text"] = "Select date format"
            labelData.place(x=20, y=10, width=182, height=37)

            buttonQuit = btk.Button(
                popupWin1,
                text="Close",
                command=lambda: [
                    popupWin1.destroy(),
                ],
                bootstyle=DANGER,
            )
            buttonQuit.place(x=65, y=100, width=92, height=30)
        except:
            popupWin1.destroy()

        def DayFirst(self):
            maxCol = len(self.data.columns)
            index = 0
            for col in range(0, maxCol):
                try:
                    self.data["Date and Time Probe " + str(index)] = pd.to_datetime(
                        self.data["Date and Time Probe " + str(index)],
                        format="%d.%m.%y %H:%M:%S",
                        errors="coerce",
                    )
                    index += 1
                except:
                    pass
            RotronicSave(self)

        def YearFirst(self):
            maxCol = len(self.data.columns)
            index = 0
            for col in range(0, maxCol):
                try:
                    self.data["Date and Time Probe " + str(index)] = pd.to_datetime(
                        self.data["Date and Time Probe " + str(index)],
                        format="%y.%m.%d %H:%M:%S",
                        errors="coerce",
                    )
                    index += 1
                except:
                    pass
            RotronicSave(self)

        def RotronicSave(self):
            # File save:
            data2 = pd.DataFrame(self.data)
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
                "rotronic": True,
                "keithleymanager": False,
                "gasChamber": False,
                "humBath": False,
                "indexZUT": False,
                "coolingSys": False,
                "saltHum": False,
                "hum": True,
            }
            DataFormating(self.root, data2, optionTest, "Rotronic")

    def ReadVaisala(self):
        # Vaisala data processing:
        try:
            # Data processing variants - popup:
            popupWin1 = tk.Toplevel(self.root)
            width = 223
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
            popupWin1.minsize(223, 147)
            popupWin1.maxsize(223, 147)

            buttonTemperature = btk.Button(
                popupWin1,
                text="Indigo520",
                command=lambda: [popupWin1.destroy(), IndigoLogger(self)],
                bootstyle=DARK,
            )
            buttonTemperature.place(x=20, y=50, width=82, height=30)

            buttonHumidity = btk.Button(
                popupWin1,
                text="Insight",
                command=lambda: [popupWin1.destroy(), InsightSW(self)],
                bootstyle=DARK,
            )
            buttonHumidity.place(x=120, y=50, width=82, height=30)

            labelData = tk.Label(popupWin1)
            labelData["anchor"] = "n"
            ft = tkFont.Font(family="Helvetica", size=15)
            labelData["font"] = ft
            labelData["fg"] = "#333333"
            labelData["justify"] = "center"
            labelData["text"] = "Used logger"
            labelData.place(x=20, y=10, width=182, height=37)

            buttonQuit = btk.Button(
                popupWin1,
                text="Close",
                command=lambda: [
                    popupWin1.destroy(),
                ],
                bootstyle=DANGER,
            )
            buttonQuit.place(x=70, y=100, width=82, height=30)
        except:
            popupWin1.destroy()

        def IndigoLogger(self):
            # INDIGO520 logger data processing:
            filepaths = filedialog.askopenfilenames(
                parent=self.root,
                filetypes=(
                    ("CSV Files", "*.csv"),
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
                        sep=",",
                        names=[
                            "Date and Time",
                            "Probe",
                            "Parameter",
                            "Measurement",
                            "Unit",
                        ],
                        low_memory=False,
                        encoding_errors="ignore",
                    )
                    if self.data["Date and Time"].str.contains("Z").any() == True:
                        self.data.drop("Unit", inplace=True, axis=1)
                        self.data["Date and Time"] = self.data[
                            "Date and Time"
                        ].str.replace("T", " ", regex=True)
                        self.data["Date and Time"] = self.data[
                            "Date and Time"
                        ].str.replace("Z", "", regex=True)
                        self.data["Date and Time"] = pd.to_datetime(
                            self.data["Date and Time"], errors="coerce"
                        )
                        # Probe 1
                        self.dataProbe1 = self.data.query("Probe == 'Probe 1'")

                        # Probe 1 Temperature
                        self.dataProbe1T = self.dataProbe1.query(
                            "Parameter == 'Temperature'"
                        )
                        self.dataProbe1DFT = self.dataProbe1T.filter(
                            ["Measurement"], axis=1
                        )
                        newColumnNames = {
                            "Date and Time": "Date and Time Probe 0",
                            "Measurement": "Temperature Probe 0",
                        }
                        self.dataProbe1DFT.rename(columns=newColumnNames, inplace=True)
                        self.dataProbe1DFT = self.dataProbe1DFT.reset_index(drop=True)
                        self.dataProbe1DFT = pd.DataFrame(self.dataProbe1DFT)

                        # Probe 1 Humidity
                        self.dataProbe1H = self.dataProbe1.query(
                            "Parameter == 'Relative humidity'"
                        )
                        self.dataProbe1DFH = self.dataProbe1H.filter(
                            ["Date and Time", "Measurement"], axis=1
                        )
                        newColumnNames = {
                            "Date and Time": "Date and Time Probe 0",
                            "Measurement": "Humidity Probe 0",
                        }
                        self.dataProbe1DFH.rename(columns=newColumnNames, inplace=True)
                        self.dataProbe1DFH = self.dataProbe1DFH.reset_index(drop=True)
                        self.dataProbe1DFH = pd.DataFrame(self.dataProbe1DFH)

                        # Probe 2
                        self.dataProbe2 = self.data.query("Probe == 'Probe 2'")

                        # Probe 2 Temperature
                        self.dataProbe2T = self.dataProbe2.query(
                            "Parameter == 'Temperature'"
                        )
                        self.dataProbe2DFT = self.dataProbe2T.filter(
                            ["Measurement"], axis=1
                        )
                        newColumnNames = {"Measurement": "Temperature Probe 1"}
                        self.dataProbe2DFT.rename(columns=newColumnNames, inplace=True)
                        self.dataProbe2DFT = self.dataProbe2DFT.reset_index(drop=True)
                        self.dataProbe2DFT = pd.DataFrame(self.dataProbe2DFT)

                        # Probe 2 Humidity
                        self.dataProbe2H = self.dataProbe2.query(
                            "Parameter == 'Relative humidity'"
                        )
                        self.dataProbe2DFH = self.dataProbe2H.filter(
                            ["Date and Time", "Measurement"], axis=1
                        )
                        newColumnNames = {
                            "Date and Time": "Date and Time Probe 1",
                            "Measurement": "Humidity Probe 1",
                        }
                        self.dataProbe2DFH.rename(columns=newColumnNames, inplace=True)
                        self.dataProbe2DFH = self.dataProbe2DFH.reset_index(drop=True)
                        self.dataProbe2DFH = pd.DataFrame(self.dataProbe2DFH)
                        data2 = pd.concat(
                            [
                                self.dataProbe1DFH,
                                self.dataProbe1DFT,
                                self.dataProbe2DFH,
                                self.dataProbe2DFT,
                            ],
                            axis=1,
                        )
                    else:
                        self.data = pd.read_csv(
                            filepath,
                            # sep="\t",
                            skiprows=11,
                            low_memory=False,
                            encoding_errors="ignore",
                        )
                        name = "test"
                        self.savepath = filedialog.asksaveasfilename(
                            parent=self.root,
                            initialfile="{}_Log".format(name),
                            defaultextension=".xlsx",
                            filetypes=(
                                ("Excel Files", "*.xlsx"),
                                ("all files", ".*"),
                            ),
                            initialdir=os.getcwd(),
                        )
                        writer = pd.ExcelWriter(self.savepath, engine="xlsxwriter")
                        self.data.to_excel(
                            writer,
                            "Data",
                            index=False,
                            startrow=0,
                            startcol=0,
                        )

                    listData.append(data2)
                except:
                    return Message(
                        "error",
                        "Wrong file format."
                        "\nPlease contact with: piotr.klys92@gmail.com",
                        0,
                    )
            try:
                self.data = pd.concat(listData, axis=0, ignore_index=False)
                try:
                    self.data = self.data.sort_values(
                        by="Date and Time Probe 0", ascending=True
                    )
                except:
                    self.data = self.data.sort_values(
                        by="Date and Time Probe 1", ascending=True
                    )
                optionTest = {
                    "dewing": 0,
                    "secasi": False,
                    "angel": False,
                    "samples": False,
                    "timeabsOne": False,
                    "timeabsAll": False,
                    "indigo": True,
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
                    "hum": True,
                }
                return DataFormating(self.root, self.data, optionTest, "VaisalaINDIGO")
            except:
                return Message(
                    "error",
                    "Wrong file format."
                    "\nPlease contact with: piotr.klys92@gmail.com",
                    0,
                )

        def InsightSW(self):
            # Insight Software data processing:
            try:
                filepath = filedialog.askopenfilename(
                    parent=self.root,
                    filetypes=(
                        ("CSV Files", "*.csv"),
                        ("all files", ".*"),
                    ),
                    initialdir=os.getcwd(),
                )
                self.data = pd.read_csv(
                    filepath,
                    sep=";",
                    skiprows=8,
                    header=None,
                    low_memory=False,
                    encoding_errors="ignore",
                )
            except:
                self.WarningInfo("File was not selected")
                return
            try:
                newColumnNames = {0: "Date and Time", 1: "Temperature", 2: "Humidity"}
                self.data.rename(columns=newColumnNames, inplace=True)
                if self.data["Date and Time"].str.contains("Z").any() == False:
                    self.data["Date and Time"] = pd.to_datetime(
                        self.data["Date and Time"], errors="coerce"
                    )
                else:
                    return Message(
                        "error",
                        "Wrong file format."
                        "\nPlease contact with: piotr.klys92@gmail.com",
                        0,
                    )
                self.data["Temperature"] = self.data["Temperature"].str.replace(
                    ",", ".", regex=True
                )
                self.data["Temperature"] = pd.to_numeric(
                    self.data["Temperature"], errors="coerce"
                )
            except:
                return Message(
                    "error",
                    "Wrong file format."
                    "\nPlease contact with: piotr.klys92@gmail.com",
                    0,
                )

            try:
                self.data["Humidity"] = self.data["Humidity"].str.replace(
                    ",", ".", regex=True
                )
                self.data["Humidity"] = pd.to_numeric(
                    self.data["Humidity"], errors="coerce"
                )
                data2 = self.data.filter(
                    ["Date and Time", "Humidity", "Temperature"], axis=1
                )
                optionTest = {
                    "dewing": 0,
                    "secasi": False,
                    "angel": False,
                    "samples": False,
                    "timeabsOne": False,
                    "timeabsAll": False,
                    "indigo": False,
                    "insight": True,
                    "grafana": False,
                    "agilent": False,
                    "rotronic": False,
                    "keithleymanager": False,
                    "gasChamber": False,
                    "humBath": False,
                    "indexZUT": False,
                    "coolingSys": False,
                    "saltHum": False,
                    "hum": True,
                }
            except:
                data2 = self.data.filter(["Date and Time", "Temperature"], axis=1)
                optionTest = {
                    "dewing": 0,
                    "secasi": False,
                    "angel": False,
                    "samples": False,
                    "timeabsOne": False,
                    "timeabsAll": False,
                    "indigo": False,
                    "insight": True,
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
            DataFormating(self.root, data2, optionTest, "VasialaINSIGHT")

    def ReadKeithley(self):
        # Keithley data processing:
        # try:
        #     popupWin1 = tk.Toplevel(self.root)
        #     width = 223
        #     height = 147
        #     screenwidth = popupWin1.winfo_screenwidth()
        #     screenheight = popupWin1.winfo_screenheight()
        #     alignstr = "%dx%d+%d+%d" % (
        #         width,
        #         height,
        #         (screenwidth - width) / 2,
        #         (screenheight - height) / 2,
        #     )
        #     popupWin1.geometry(alignstr)
        #     popupWin1.resizable(width=False, height=False)
        #     popupWin1.attributes("-topmost", True)
        #     popupWin1.minsize(223, 147)
        #     popupWin1.maxsize(223, 147)

        #     buttonTemperature = btk.Button(
        #         popupWin1,
        #         text="Xlinx",
        #         command=lambda: [popupWin1.destroy(), Xlinx(self)],
        #         bootstyle=DARK,
        #     )
        #     buttonTemperature.place(x=20, y=50, width=82, height=30)

        #     buttonHumidity = btk.Button(
        #         popupWin1,
        #         text="Manager",
        #         command=lambda: [popupWin1.destroy(), KeithleyManager(self)],
        #         bootstyle=DARK,
        #     )
        #     buttonHumidity.place(x=120, y=50, width=82, height=30)

        #     labelData = tk.Label(popupWin1)
        #     labelData["anchor"] = "n"
        #     ft = tkFont.Font(family="Helvetica", size=15)
        #     labelData["font"] = ft
        #     labelData["fg"] = "#333333"
        #     labelData["justify"] = "center"
        #     labelData["text"] = "Logger type:"
        #     labelData.place(x=20, y=10, width=182, height=37)

        #     buttonQuit = btk.Button(
        #         popupWin1,
        #         text="Close",
        #         command=lambda: [
        #             popupWin1.destroy(),
        #         ],
        #         bootstyle=DANGER,
        #     )
        #     buttonQuit.place(x=70, y=100, width=82, height=30)
        # except:
        #     popupWin1.destroy()

        try:
            filepath = filedialog.askopenfilename(
                parent=self.root,
                filetypes=(
                    ("Text Files", "*.log"),
                    ("all files", ".*"),
                ),
                initialdir=os.getcwd(),
            )
            self.data = pd.read_csv(
                filepath,
                header=None,
                sep=",",
                skiprows=[1, 2, 3],
                low_memory=False,
                encoding_errors="ignore",
            )
        except:
            self.WarningInfo("File was not selected")
            return
        try:
            count1 = 1
            count2 = 1
            for i in self.data.keys()[1:]:
                if i % 2 == 0:
                    self.data = self.data.rename(columns={i: "time(abs)" + str(count1)})
                    count1 += 1
                else:
                    self.data = self.data.rename(columns={i: "Channel" + str(count2)})
                    count2 += 1
            listData = []
            count1 = 1
            for col in self.data.keys():
                try:
                    self.data["time(abs)" + str(count1)] = self.data[
                        "time(abs)" + str(count1)
                    ].str.replace(r"\.\d+", "", regex=True)
                    self.data["time(abs)" + str(count1)] = pd.to_datetime(
                        self.data["time(abs)" + str(count1)], format="%H:%M:%S"
                    ).dt.time()
                    count1 += 1
                except:
                    pass
                    listData.append(self.data)
            newColumnNames = {0: "Samples", "time(abs)1": "T"}
            self.data.rename(columns=newColumnNames, inplace=True)
            self.data = pd.DataFrame(self.data)
        except:
            return Message(
                "error",
                "Wrong file format." "\nPlease contact with: piotr.klys92@gmail.com",
                0,
            )

        # Data processing variants - popup:
        try:
            popupWin1 = tk.Toplevel(self.root)
            width = 243
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
            popupWin1.minsize(243, 147)
            popupWin1.maxsize(243, 147)

            buttonTemperature = btk.Button(
                popupWin1,
                text="All time(abs)",
                command=lambda: [popupWin1.destroy(), AllTimeabs(self)],
                bootstyle=DARK,
            )
            buttonTemperature.place(x=10, y=50, width=102, height=30)

            buttonHumidity = btk.Button(
                popupWin1,
                text="One time(abs)",
                command=lambda: [popupWin1.destroy(), OneTimeabs(self)],
                bootstyle=DARK,
            )
            buttonHumidity.place(x=130, y=50, width=102, height=30)

            labelData = tk.Label(popupWin1)
            labelData["anchor"] = "n"
            ft = tkFont.Font(family="Helvetica", size=15)
            labelData["font"] = ft
            labelData["fg"] = "#333333"
            labelData["justify"] = "center"
            labelData["text"] = "Data type"
            labelData.place(x=30, y=10, width=182, height=37)

            buttonQuit = btk.Button(
                popupWin1,
                text="Close",
                command=lambda: [
                    popupWin1.destroy(),
                ],
                bootstyle=DANGER,
            )
            buttonQuit.place(x=75, y=100, width=92, height=30)
        except:
            popupWin1.destroy()

        # def KeithleyManager(self):
        #     try:
        #         filepath = filedialog.askopenfilename(
        #             parent=self.root,
        #             filetypes=(
        #                 ("Text files", "*.txt"),
        #                 ("all files", ".*"),
        #             ),
        #         )
        #         self.data = pd.read_csv(
        #             filepath, sep=";", low_memory=False, encoding_errors="ignore"
        #         )
        #     except:
        #         return Message("warning","File was not selected", 0)
        #
        #     for i in self.data.keys():
        #         if i.startswith("Date and Time"):
        #             self.data[i] = pd.to_datetime(
        #                 self.data[i], yearfirst=True, errors="coerce"
        #             )
        #         elif i.startswith("Measurement"):
        #             self.data[i] = pd.to_numeric(self.data[i], errors="coerce")
        #     try:
        #         data2 = pd.DataFrame(self.data)
        #         optionTest = {
        #             "dewing": 0,
        #             "secasi": False,
        #             "angel": False,
        #             "samples": False,
        #             "timeabsOne": False,
        #             "timeabsAll": False,
        #             "indigo": False,
        #             "insight": False,
        #             "grafana": False,
        #             "agilent": False,
        #             "rotronic": False,
        #             "keithleymanager": True,
        #             "gasChamber": False,
        #             "humBath": False,
        #             "indexZUT": False,
        #             "coolingSys": False,
        #             "saltHum": False,
        #             "hum": False,
        #         }
        #         DataFormating(self.root, data2, optionTest)
        #     except:
        #         return Message(
        #             "error",
        #             "Wrong file format." "\nPlease contact with: piotr.klys92@gmail.com",
        #             0,
        #         )

        # def Xlinx(self):
        #     try:
        #         filepath = filedialog.askopenfilename(
        #             parent=self.root,
        #             filetypes=(
        #                 ("Text Files", "*.log"),
        #                 ("all files", ".*"),
        #             ),
        #         )
        #         self.data = pd.read_csv(
        #             filepath,
        #             header=None,
        #             sep=",",
        #             skiprows=[1, 2, 3],
        #             low_memory=False,
        #             encoding_errors="ignore",
        #         )
        #     except:
        #         self.WarningInfo("File was not selected")
        #         return
        #     try:
        #         count1 = 1
        #         count2 = 1
        #         for i in self.data.keys()[1:]:
        #             if i % 2 == 0:
        #                 self.data = self.data.rename(
        #                     columns={i: "time(abs)" + str(count1)}
        #                 )
        #                 count1 += 1
        #             else:
        #                 self.data = self.data.rename(
        #                     columns={i: "Channel" + str(count2)}
        #                 )
        #                 count2 += 1
        #         listData = []
        #         count1 = 1
        #         for col in self.data.keys():
        #             try:
        #                 self.data["time(abs)" + str(count1)] = self.data[
        #                     "time(abs)" + str(count1)
        #                 ].str.replace("\.\d+", "", regex=True)
        #                 self.data["time(abs)" + str(count1)] = pd.to_datetime(
        #                     self.data["time(abs)" + str(count1)], format="%H:%M:%S"
        #                 ).dt.time()
        #                 count1 += 1
        #             except:
        #                 pass
        #                 listData.append(self.data)
        #         newColumnNames = {0: "Samples", "time(abs)1": "T"}
        #         self.data.rename(columns=newColumnNames, inplace=True)
        #         self.data = pd.DataFrame(self.data)
        #     except:
        #         return Message(
        #             "error",
        #             "Wrong file format." "\nPlease contact with: piotr.klys92@gmail.com",
        #             0,
        #         )
        #         return

        #     # Data processing variants - popup:
        #     try:
        #         popupWin1 = tk.Toplevel(self.root)
        #         width = 243
        #         height = 147
        #         screenwidth = popupWin1.winfo_screenwidth()
        #         screenheight = popupWin1.winfo_screenheight()
        #         alignstr = "%dx%d+%d+%d" % (
        #             width,
        #             height,
        #             (screenwidth - width) / 2,
        #             (screenheight - height) / 2,
        #         )
        #         popupWin1.geometry(alignstr)
        #         popupWin1.resizable(width=False, height=False)
        #         popupWin1.attributes("-topmost", True)
        #         popupWin1.minsize(243, 147)
        #         popupWin1.maxsize(243, 147)

        #         buttonTemperature = btk.Button(
        #             popupWin1,
        #             text="All time(abs)",
        #             command=lambda: [popupWin1.destroy(), AllTimeabs(self)],
        #             bootstyle=DARK,
        #         )
        #         buttonTemperature.place(x=10, y=50, width=102, height=30)

        #         buttonHumidity = btk.Button(
        #             popupWin1,
        #             text="One time(abs)",
        #             command=lambda: [popupWin1.destroy(), OneTimeabs(self)],
        #             bootstyle=DARK,
        #         )
        #         buttonHumidity.place(x=130, y=50, width=102, height=30)

        #         labelData = tk.Label(popupWin1)
        #         labelData["anchor"] = "n"
        #         ft = tkFont.Font(family="Helvetica", size=15)
        #         labelData["font"] = ft
        #         labelData["fg"] = "#333333"
        #         labelData["justify"] = "center"
        #         labelData["text"] = "Data type"
        #         labelData.place(x=30, y=10, width=182, height=37)

        #         buttonQuit = btk.Button(
        #             popupWin1,
        #             text="Close",
        #             command=lambda: [
        #                 popupWin1.destroy(),
        #             ],
        #             bootstyle=DANGER,
        #         )
        #         buttonQuit.place(x=75, y=100, width=92, height=30)
        #     except:
        #         popupWin1.destroy()

        def AllTimeabs(self):
            newColumnNames = {"T": "time(abs)1"}
            self.data.rename(columns=newColumnNames, inplace=True)
            data2 = self.data
            optionTest = {
                "dewing": 0,
                "secasi": False,
                "angel": False,
                "samples": False,
                "timeabsOne": False,
                "timeabsAll": True,
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
            DataFormating(self.root, data2, optionTest, "Keithley")

        def OneTimeabs(self):
            try:
                self.data2 = self.data
                self.data2 = self.data2.loc[
                    :, ~self.data2.columns.str.startswith("time")
                ]
                newColumnNames = {"T": "time(abs)"}
                self.data2.rename(columns=newColumnNames, inplace=True)
                # Changing positions:
                temp_cols = self.data2.columns.tolist()
                index = self.data2.columns.get_loc("time(abs)")
                new_cols = (
                    temp_cols[index : index + 1]
                    + temp_cols[0:index]
                    + temp_cols[index + 1 :]
                )
                self.data2 = self.data2[new_cols]

                temp_cols = self.data2.columns.tolist()
                index = self.data2.columns.get_loc("Samples")
                new_cols = (
                    temp_cols[index : index + 1]
                    + temp_cols[0:index]
                    + temp_cols[index + 1 :]
                )
                data2 = self.data2[new_cols]

            except:
                self.popupWin2.destroy()
                return

            # Axis X variants - popup:
            try:
                popupWin4 = tk.Toplevel(self.root)
                width = 223
                height = 147
                screenwidth = popupWin4.winfo_screenwidth()
                screenheight = popupWin4.winfo_screenheight()
                alignstr = "%dx%d+%d+%d" % (
                    width,
                    height,
                    (screenwidth - width) / 2,
                    (screenheight - height) / 2,
                )
                popupWin4.geometry(alignstr)
                popupWin4.resizable(width=False, height=False)
                popupWin4.attributes("-topmost", True)
                popupWin4.minsize(223, 147)
                popupWin4.maxsize(223, 147)

                buttonTemperature = btk.Button(
                    popupWin4,
                    text="Samples",
                    command=lambda: [popupWin4.destroy(), Samples(self, data2)],
                    bootstyle=DARK,
                )
                buttonTemperature.place(x=20, y=50, width=82, height=30)

                buttonHumidity = btk.Button(
                    popupWin4,
                    text="time(abs)",
                    command=lambda: [popupWin4.destroy(), Time_abs(self, data2)],
                    bootstyle=DARK,
                )
                buttonHumidity.place(x=120, y=50, width=82, height=30)

                labelData = tk.Label(popupWin4)
                labelData["anchor"] = "n"
                ft = tkFont.Font(family="Helvetica", size=15)
                labelData["font"] = ft
                labelData["fg"] = "#333333"
                labelData["justify"] = "center"
                labelData["text"] = "Axis X data"
                labelData.place(x=20, y=10, width=182, height=37)

                buttonQuit = btk.Button(
                    popupWin4,
                    text="Close",
                    command=lambda: [
                        popupWin4.destroy(),
                    ],
                    bootstyle=DANGER,
                )
                buttonQuit.place(x=70, y=100, width=82, height=30)
            except:
                popupWin4.destroy()

            def Samples(self, data2):
                optionTest = {
                    "dewing": 0,
                    "secasi": False,
                    "angel": False,
                    "samples": True,
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
                DataFormating(self.root, data2, optionTest, "Keithley")

            def Time_abs(self, data2):
                optionTest = {
                    "dewing": 0,
                    "secasi": False,
                    "angel": False,
                    "samples": False,
                    "timeabsOne": True,
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
                DataFormating(self.root, data2, optionTest, "Keithley")

    def ReadGrafana(self):
        # Grafana and logger manager data processing:
        filepaths = filedialog.askopenfilenames(
            parent=self.root,
            filetypes=(
                ("CSV Files", "*.csv"),
                ("all files", ".*"),
            ),
            initialdir=os.getcwd(),
        )
        filepaths2 = list(filepaths)
        listData = []
        for filepath in filepaths2:
            try:
                try:
                    try:
                        try:
                            self.data = pd.read_csv(
                                filepath,
                                sep=";",
                                low_memory=False,
                                encoding_errors="ignore",
                            )
                        except ValueError:
                            return Message(
                                "error",
                                "Wrong file format."
                                "\nPlease contact with: piotr.klys92@gmail.com",
                                0,
                            )
                    except:
                        return Message("warning", "File was not selected", 0)
                    newColumnNames = {
                        "Data/czas": "Date and Time",
                        "Temperatura[C]": "Temperature",
                        "Wilgotnosc[rH]": "Humidity",
                    }
                    self.data.rename(columns=newColumnNames, inplace=True)
                    self.data["Date and Time"] = pd.to_datetime(
                        self.data["Date and Time"], errors="coerce"
                    )
                except:
                    try:
                        try:
                            try:
                                self.data = pd.read_csv(
                                    filepath,
                                    sep=",",
                                    low_memory=False,
                                    encoding_errors="ignore",
                                )
                            except ValueError:
                                return Message(
                                    "error",
                                    "Wrong file format."
                                    "\nPlease contact with: piotr.klys92@gmail.com",
                                    0,
                                )

                        except:
                            return Message("warning", "File was not selected", 0)
                        newColumnNames = {
                            "Time": "Date and Time",
                            "time": "Date and Time",
                            "temp": "Temperature",
                            "hum": "Humidity",
                            "temp1": "Temperature Probe 0",
                            "hum1": "Humidity Probe 0",
                            "temp2": "Temperature Probe 1",
                            "hum2": "Humidity Probe 1",
                            "Temperatura": "Temperature",
                            "Wilgotność": "Humidity",
                        }
                        self.data.rename(columns=newColumnNames, inplace=True)
                        self.data["Date and Time"] = pd.to_datetime(
                            self.data["Date and Time"], errors="coerce"
                        )

                    except:
                        try:
                            try:
                                self.data = pd.read_csv(
                                    filepath,
                                    skiprows=1,
                                    low_memory=False,
                                    encoding_errors="ignore",
                                )
                            except KeyError:
                                return Message(
                                    "error",
                                    "Wrong file format."
                                    "\nPlease contact with: piotr.klys92@gmail.com",
                                    0,
                                )

                        except:
                            return Message("warning", "File was not selected", 0)
                        try:
                            newColumnNames = {
                                "Time": "Date and Time",
                                "time": "Date and Time",
                                "temp": "Temperature",
                                "hum": "Humidity",
                                "temp1": "Temperature Probe 0",
                                "hum1": "Humidity Probe 0",
                                "temp2": "Temperature Probe 1",
                                "hum2": "Humidity Probe 1",
                                "Temperatura": "Temperature",
                                "Wilgotność": "Humidity",
                            }
                            self.data.rename(columns=newColumnNames, inplace=True)
                            self.data["Date and Time"] = pd.to_datetime(
                                self.data["Date and Time"], errors="coerce"
                            )
                        except KeyError:
                            return Message(
                                "error",
                                "Wrong file format."
                                "\nPlease contact with: piotr.klys92@gmail.com",
                                0,
                            )
                try:
                    self.data["Temperature"] = pd.to_numeric(
                        self.data["Temperature"], errors="coerce"
                    )
                    self.data["Humidity"] = pd.to_numeric(
                        self.data["Humidity"], errors="coerce"
                    )
                    data2 = self.data.filter(
                        ["Date and Time", "Humidity", "Temperature"], axis=1
                    )

                except:
                    try:
                        maxCol = len(self.data.columns)
                        index = 0
                        for col in range(0, maxCol):
                            try:
                                self.data["Temperature Probe " + str(index)] = (
                                    pd.to_numeric(
                                        self.data["Temperature Probe " + str(index)],
                                        errors="coerce",
                                    )
                                )
                                self.data["Humidity Probe " + str(index)] = (
                                    pd.to_numeric(
                                        self.data["Humidity Probe " + str(index)],
                                        errors="coerce",
                                    )
                                )
                                index += 1
                            except:
                                pass
                        data2 = self.data.filter(
                            [
                                "Date and Time",
                                "Humidity Probe 0",
                                "Temperature Probe 0",
                                "Humidity Probe 1",
                                "Temperature Probe 1",
                            ],
                            axis=1,
                        )

                    except:
                        maxCol = len(self.data.columns)
                        index = 0
                        for col in range(0, maxCol):
                            try:
                                self.data["Temperature Probe " + str(index)] = (
                                    pd.to_numeric(
                                        self.data["Temperature Probe " + str(index)],
                                        errors="coerce",
                                    )
                                )
                                self.data["Humidity Probe " + str(index)] = (
                                    pd.to_numeric(
                                        self.data["Humidity Probe " + str(index)],
                                        errors="coerce",
                                    )
                                )
                                index += 1
                            except:
                                pass
                        data2 = self.data.filter(
                            [
                                "Date and Time",
                                "Humidity Probe 0",
                                "Temperature Probe 0",
                                "Humidity Probe 1",
                                "Temperature Probe 1",
                            ],
                            axis=1,
                        )
                listData.append(data2)
            except:
                return Message(
                    "error",
                    "Wrong file format."
                    "\nPlease contact with: piotr.klys92@gmail.com",
                    0,
                )

        try:
            data2 = pd.concat(listData, axis=0, ignore_index=False)
            data2 = data2.sort_values(by="Date and Time", ascending=True)
            optionTest = {
                "dewing": 0,
                "secasi": False,
                "angel": False,
                "samples": False,
                "timeabsOne": False,
                "timeabsAll": False,
                "indigo": False,
                "insight": False,
                "grafana": True,
                "agilent": False,
                "rotronic": False,
                "keithleymanager": False,
                "gasChamber": False,
                "humBath": False,
                "indexZUT": False,
                "coolingSys": False,
                "saltHum": False,
                "hum": True,
            }
            return DataFormating(self.root, data2, optionTest, "Grafana")
        except:
            return Message(
                "error",
                "Wrong file format." "\nPlease contact with: piotr.klys92@gmail.com",
                0,
            )

    def ReadAgilent(self):
        # Agilent data processing:
        filepaths = filedialog.askopenfilenames(
            parent=self.root,
            filetypes=(
                ("Text File", "*.txt"),
                ("all files", ".*"),
            ),
            initialdir=os.getcwd(),
        )
        filepaths2 = list(filepaths)
        listData = []
        for filepath in filepaths2:
            try:
                self.data = pd.read_csv(
                    filepath, header=None, sep="; ", engine="python"
                )
                self.data.drop(self.data.index[[0, 1]], axis=0, inplace=True)
                empty_cols = [
                    col for col in self.data.columns if self.data[col].isnull().all()
                ]
                self.data.drop(empty_cols, axis=1, inplace=True)
                count = 1
                for i in self.data.keys()[1:]:
                    self.data = self.data.rename(columns={i: "Channel" + str(count)})
                    self.data["Channel" + str(count)] = self.data[
                        "Channel" + str(count)
                    ].str.replace(",", ".", regex=True)
                    self.data["Channel" + str(count)] = pd.to_numeric(
                        self.data["Channel" + str(count)], errors="coerce"
                    )
                    count += 1

                newColumnNames = {0: "Date and Time"}
                self.data.rename(columns=newColumnNames, inplace=True)
                self.data.drop(self.data.index[[0, 1]], axis=0, inplace=True)
                self.data["Date and Time"] = pd.to_datetime(
                    self.data["Date and Time"], dayfirst=True, errors="coerce"
                )
                data2 = pd.DataFrame(self.data)
                listData.append(data2)
            except UnicodeDecodeError:
                return Message(
                    "error",
                    "Wrong file format."
                    "\nPlease contact with: piotr.klys92@gmail.com",
                    0,
                )
        try:
            self.data = pd.concat(listData, axis=0, ignore_index=False)
            self.data = self.data.sort_values(by="Date and Time", ascending=True)
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
                "agilent": True,
                "rotronic": False,
                "keithleymanager": False,
                "gasChamber": False,
                "humBath": False,
                "indexZUT": False,
                "coolingSys": False,
                "saltHum": False,
                "hum": False,
            }
            DataFormating(self.root, self.data, optionTest, "Agilent")
        except:
            return Message(
                "error",
                "Wrong file format." "\nPlease contact with: piotr.klys92@gmail.com",
                0,
            )
