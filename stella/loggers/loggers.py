import pandas as pd
import tkinter as tk
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
            self.read_rotronic()
        elif option == 1:
            self.read_vaisala()
        elif option == 2:
            self.read_keithley()
        elif option == 3:
            self.read_grafana()
        elif option == 4:
            self.read_agilent()

    def read_rotronic(self):
        # Rotronic data processing:
        filepaths = filedialog.askopenfilenames(
            parent=self.root,
            filetypes=(
                ("Excel Files", "*.xls"),
                ("all files", ".*"),
            ),
        )
        filepaths = list(filepaths)
        # try:
        # Multiple data function:
        list_data = []
        for filepath in filepaths:
            indexNumber = str(filepaths.index(filepath))
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
                new_column_names = {
                    "Date    _Time    ": DT + indexNumber,
                    "Temperature": T + indexNumber,
                    "Humidity": H + indexNumber,
                }
                self.data.rename(columns=new_column_names, inplace=True)
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
                    new_column_names = {
                        "Date    _Time    ": DT + indexNumber,
                        "Temperature": T + indexNumber,
                        "Humidity": H + indexNumber,
                    }
                    self.data.rename(columns=new_column_names, inplace=True)
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
                    new_column_names = {
                        "Date    _Time    ": DT + indexNumber,
                        "Temperature": T + indexNumber,
                        "Humidity": H + indexNumber,
                    }
                    self.data.rename(columns=new_column_names, inplace=True)
            self.data[T + indexNumber] = pd.to_numeric(
                self.data[T + indexNumber], errors="coerce"
            )
            self.data[H + indexNumber] = pd.to_numeric(
                self.data[H + indexNumber], errors="coerce"
            )
            self.data = self.data.dropna(subset=[H + indexNumber, T + indexNumber])
            list_data.append(self.data)
        try:
            self.data = pd.concat(list_data, axis=1, ignore_index=False)
        except:
            return Message("warning", "File was not selected", 0)
        try:
            # Data processing variants - popup:
            popup_win_1 = tk.Toplevel(self.root)
            width = 223
            height = 147
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
            popup_win_1.minsize(223, 147)
            popup_win_1.maxsize(223, 147)

            button_date_format = btk.Button(
                popup_win_1,
                text="dd.mm.yyyy",
                command=lambda: [popup_win_1.destroy(), day_first(self)],
                bootstyle=DARK,
            )
            button_date_format.place(x=15, y=50, width=92, height=30)

            button_date_format_2 = btk.Button(
                popup_win_1,
                text="yyyy.mm.dd",
                command=lambda: [popup_win_1.destroy(), year_first(self)],
                bootstyle=DARK,
            )
            button_date_format_2.place(x=115, y=50, width=92, height=30)

            label_data = tk.Label(popup_win_1)
            label_data["anchor"] = "n"
            ft = tkFont.Font(family="Helvetica", size=15)
            label_data["font"] = ft
            label_data["fg"] = "#333333"
            label_data["justify"] = "center"
            label_data["text"] = "Select date format"
            label_data.place(x=20, y=10, width=182, height=37)

            button_close = btk.Button(
                popup_win_1,
                text="Close",
                command=lambda: [
                    popup_win_1.destroy(),
                ],
                bootstyle=DANGER,
            )
            button_close.place(x=65, y=100, width=92, height=30)
        except:
            popup_win_1.destroy()

        def day_first(self):
            max_col = len(self.data.columns)
            index = 0
            for col in range(0, max_col):
                try:
                    self.data["Date and Time Probe " + str(index)] = pd.to_datetime(
                        self.data["Date and Time Probe " + str(index)],
                        format="%d.%m.%y %H:%M:%S",
                        errors="coerce",
                    )
                    index += 1
                except:
                    pass
            rotronic_save(self)

        def year_first(self):
            max_col = len(self.data.columns)
            index = 0
            for col in range(0, max_col):
                try:
                    self.data["Date and Time Probe " + str(index)] = pd.to_datetime(
                        self.data["Date and Time Probe " + str(index)],
                        format="%y.%m.%d %H:%M:%S",
                        errors="coerce",
                    )
                    index += 1
                except:
                    pass
            rotronic_save(self)

        def rotronic_save(self):
            # File save:
            data2 = pd.DataFrame(self.data)
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
                "rotronic": True,
                "keithleymanager": False,
                "gasChamber": False,
                "humBath": False,
                "indexZUT": False,
                "coolingSys": False,
                "saltHum": False,
                "hum": True,
            }
            DataFormating(self.root, data2, option_test, "Rotronic")

    def read_vaisala(self):
        # Vaisala data processing:
        try:
            # Data processing variants - popup:
            popup_win_1 = tk.Toplevel(self.root)
            width = 223
            height = 147
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
            popup_win_1.minsize(223, 147)
            popup_win_1.maxsize(223, 147)

            button_indigo = btk.Button(
                popup_win_1,
                text="Indigo520",
                command=lambda: [popup_win_1.destroy(), indigo_logger(self)],
                bootstyle=DARK,
            )
            button_indigo.place(x=20, y=50, width=82, height=30)

            button_insight = btk.Button(
                popup_win_1,
                text="Insight",
                command=lambda: [popup_win_1.destroy(), insight_sw(self)],
                bootstyle=DARK,
            )
            button_insight.place(x=120, y=50, width=82, height=30)

            label_data = tk.Label(popup_win_1)
            label_data["anchor"] = "n"
            ft = tkFont.Font(family="Helvetica", size=15)
            label_data["font"] = ft
            label_data["fg"] = "#333333"
            label_data["justify"] = "center"
            label_data["text"] = "Used logger"
            label_data.place(x=20, y=10, width=182, height=37)

            button_close = btk.Button(
                popup_win_1,
                text="Close",
                command=lambda: [
                    popup_win_1.destroy(),
                ],
                bootstyle=DANGER,
            )
            button_close.place(x=70, y=100, width=82, height=30)
        except:
            popup_win_1.destroy()

        def indigo_logger(self):
            # INDIGO520 logger data processing:
            filepaths = filedialog.askopenfilenames(
                parent=self.root,
                filetypes=(
                    ("CSV Files", "*.csv"),
                    ("all files", ".*"),
                ),
            )
            filepaths = list(filepaths)
            list_data = []
            for filepath in filepaths:
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
                        self.data_probe_1 = self.data.query("Probe == 'Probe 1'")

                        # Probe 1 Temperature
                        self.data_probe_1_t = self.data_probe_1.query(
                            "Parameter == 'Temperature'"
                        )
                        self.data_probe_dft = self.data_probe_1_t.filter(
                            ["Measurement"], axis=1
                        )
                        new_column_names = {
                            "Date and Time": "Date and Time Probe 0",
                            "Measurement": "Temperature Probe 0",
                        }
                        self.data_probe_dft.rename(
                            columns=new_column_names, inplace=True
                        )
                        self.data_probe_dft = self.data_probe_dft.reset_index(drop=True)
                        self.data_probe_dft = pd.DataFrame(self.data_probe_dft)

                        # Probe 1 Humidity
                        self.data_probe_1_h = self.data_probe_1.query(
                            "Parameter == 'Relative humidity'"
                        )
                        self.data_probe_dfh = self.data_probe_1_h.filter(
                            ["Date and Time", "Measurement"], axis=1
                        )
                        new_column_names = {
                            "Date and Time": "Date and Time Probe 0",
                            "Measurement": "Humidity Probe 0",
                        }
                        self.data_probe_dfh.rename(
                            columns=new_column_names, inplace=True
                        )
                        self.data_probe_dfh = self.data_probe_dfh.reset_index(drop=True)
                        self.data_probe_dfh = pd.DataFrame(self.data_probe_dfh)

                        # Probe 2
                        self.data_probe_2 = self.data.query("Probe == 'Probe 2'")

                        # Probe 2 Temperature
                        self.data_probe_2_t = self.data_probe_2.query(
                            "Parameter == 'Temperature'"
                        )
                        self.data_probe_2_dft = self.data_probe_2_t.filter(
                            ["Measurement"], axis=1
                        )
                        new_column_names = {"Measurement": "Temperature Probe 1"}
                        self.data_probe_2_dft.rename(
                            columns=new_column_names, inplace=True
                        )
                        self.data_probe_2_dft = self.data_probe_2_dft.reset_index(
                            drop=True
                        )
                        self.data_probe_2_dft = pd.DataFrame(self.data_probe_2_dft)

                        # Probe 2 Humidity
                        self.data_probe_2_h = self.data_probe_2.query(
                            "Parameter == 'Relative humidity'"
                        )
                        self.data_probe_2_dfh = self.data_probe_2_h.filter(
                            ["Date and Time", "Measurement"], axis=1
                        )
                        new_column_names = {
                            "Date and Time": "Date and Time Probe 1",
                            "Measurement": "Humidity Probe 1",
                        }
                        self.data_probe_2_dfh.rename(
                            columns=new_column_names, inplace=True
                        )
                        self.data_probe_2_dfh = self.data_probe_2_dfh.reset_index(
                            drop=True
                        )
                        self.data_probe_2_dfh = pd.DataFrame(self.data_probe_2_dfh)
                        data2 = pd.concat(
                            [
                                self.data_probe_dfh,
                                self.data_probe_dft,
                                self.data_probe_2_dfh,
                                self.data_probe_2_dft,
                            ],
                            axis=1,
                        )
                    else:
                        self.data = pd.read_csv(
                            filepath,
                            skiprows=11,
                            low_memory=False,
                            encoding_errors="ignore",
                        )
                        name = "test"
                        self.savepath = filedialog.asksaveasfilename(
                            parent=self.root,
                            initialfile=f"{name}_Log",
                            defaultextension=".xlsx",
                            filetypes=(
                                ("Excel Files", "*.xlsx"),
                                ("all files", ".*"),
                            ),
                        )
                        writer = pd.ExcelWriter(self.savepath, engine="xlsxwriter")
                        self.data.to_excel(
                            writer,
                            "Data",
                            index=False,
                            startrow=0,
                            startcol=0,
                        )

                    list_data.append(data2)
                except:
                    return Message(
                        "error",
                        "Wrong file format."
                        "\nPlease contact with: piotr.klys92@gmail.com",
                        0,
                    )
            try:
                self.data = pd.concat(list_data, axis=0, ignore_index=False)
                try:
                    self.data = self.data.sort_values(
                        by="Date and Time Probe 0", ascending=True
                    )
                except:
                    self.data = self.data.sort_values(
                        by="Date and Time Probe 1", ascending=True
                    )
                option_test = {
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
                return DataFormating(self.root, self.data, option_test, "VaisalaINDIGO")
            except:
                return Message(
                    "error",
                    "Wrong file format." "\nPlease contact with: piotr.klys92@gmail.com",
                    0,
                )

        def insight_sw(self):
            # Insight Software data processing:
            try:
                filepath = filedialog.askopenfilename(
                    parent=self.root,
                    filetypes=(
                        ("CSV Files", "*.csv"),
                        ("all files", ".*"),
                    ),
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
                return Message("warning", "File was not selected", 0)
            try:
                new_column_names = {0: "Date and Time", 1: "Temperature", 2: "Humidity"}
                self.data.rename(columns=new_column_names, inplace=True)
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
                    "Wrong file format." "\nPlease contact with: piotr.klys92@gmail.com",
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
                option_test = {
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
                option_test = {
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
            DataFormating(self.root, data2, option_test, "VasialaINSIGHT")

    def read_keithley(self):
        # Keithley data processing:
        # try:
        #     popup_win_1 = tk.Toplevel(self.root)
        #     width = 223
        #     height = 147
        #     screenwidth = popup_win_1.winfo_screenwidth()
        #     screenheight = popup_win_1.winfo_screenheight()
        #     alignstr = "%dx%d+%d+%d" % (
        #         width,
        #         height,
        #         (screenwidth - width) / 2,
        #         (screenheight - height) / 2,
        #     )
        #     popup_win_1.geometry(alignstr)
        #     popup_win_1.resizable(width=False, height=False)
        #     popup_win_1.attributes("-topmost", True)
        #     popup_win_1.minsize(223, 147)
        #     popup_win_1.maxsize(223, 147)

        #     buttonTemperature = btk.Button(
        #         popup_win_1,
        #         text="Xlinx",
        #         command=lambda: [popup_win_1.destroy(), Xlinx(self)],
        #         bootstyle=DARK,
        #     )
        #     buttonTemperature.place(x=20, y=50, width=82, height=30)

        #     buttonHumidity = btk.Button(
        #         popup_win_1,
        #         text="Manager",
        #         command=lambda: [popup_win_1.destroy(), KeithleyManager(self)],
        #         bootstyle=DARK,
        #     )
        #     buttonHumidity.place(x=120, y=50, width=82, height=30)

        #     label_data = tk.Label(popup_win_1)
        #     label_data["anchor"] = "n"
        #     ft = tkFont.Font(family="Helvetica", size=15)
        #     label_data["font"] = ft
        #     label_data["fg"] = "#333333"
        #     label_data["justify"] = "center"
        #     label_data["text"] = "Logger type:"
        #     label_data.place(x=20, y=10, width=182, height=37)

        #     button_close = btk.Button(
        #         popup_win_1,
        #         text="Close",
        #         command=lambda: [
        #             popup_win_1.destroy(),
        #         ],
        #         bootstyle=DANGER,
        #     )
        #     button_close.place(x=70, y=100, width=82, height=30)
        # except:
        #     popup_win_1.destroy()

        try:
            filepath = filedialog.askopenfilename(
                parent=self.root,
                filetypes=(
                    ("Text Files", "*.log"),
                    ("all files", ".*"),
                ),
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
            return Message("warning", "File was not selected", 0)
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
            list_data = []
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
                    list_data.append(self.data)
            new_column_names = {0: "Samples", "time(abs)1": "T"}
            self.data.rename(columns=new_column_names, inplace=True)
            self.data = pd.DataFrame(self.data)
        except:
            return Message(
                "error",
                "Wrong file format." "\nPlease contact with: piotr.klys92@gmail.com",
                0,
            )

        # Data processing variants - popup:
        try:
            popup_win_1 = tk.Toplevel(self.root)
            width = 243
            height = 147
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
            popup_win_1.minsize(243, 147)
            popup_win_1.maxsize(243, 147)

            button_all_time = btk.Button(
                popup_win_1,
                text="All time(abs)",
                command=lambda: [popup_win_1.destroy(), all_timeabs(self)],
                bootstyle=DARK,
            )
            button_all_time.place(x=10, y=50, width=102, height=30)

            button_one_time = btk.Button(
                popup_win_1,
                text="One time(abs)",
                command=lambda: [popup_win_1.destroy(), one_timeabs(self)],
                bootstyle=DARK,
            )
            button_one_time.place(x=130, y=50, width=102, height=30)

            label_data = tk.Label(popup_win_1)
            label_data["anchor"] = "n"
            ft = tkFont.Font(family="Helvetica", size=15)
            label_data["font"] = ft
            label_data["fg"] = "#333333"
            label_data["justify"] = "center"
            label_data["text"] = "Data type"
            label_data.place(x=30, y=10, width=182, height=37)

            button_close = btk.Button(
                popup_win_1,
                text="Close",
                command=lambda: [
                    popup_win_1.destroy(),
                ],
                bootstyle=DANGER,
            )
            button_close.place(x=75, y=100, width=92, height=30)
        except:
            popup_win_1.destroy()

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
        #         option_test = {
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
        #         DataFormating(self.root, data2, option_test)
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
        #         list_data = []
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
        #                 list_data.append(self.data)
        #         new_column_names = {0: "Samples", "time(abs)1": "T"}
        #         self.data.rename(columns=new_column_names, inplace=True)
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
        #         popup_win_1 = tk.Toplevel(self.root)
        #         width = 243
        #         height = 147
        #         screenwidth = popup_win_1.winfo_screenwidth()
        #         screenheight = popup_win_1.winfo_screenheight()
        #         alignstr = "%dx%d+%d+%d" % (
        #             width,
        #             height,
        #             (screenwidth - width) / 2,
        #             (screenheight - height) / 2,
        #         )
        #         popup_win_1.geometry(alignstr)
        #         popup_win_1.resizable(width=False, height=False)
        #         popup_win_1.attributes("-topmost", True)
        #         popup_win_1.minsize(243, 147)
        #         popup_win_1.maxsize(243, 147)

        #         buttonTemperature = btk.Button(
        #             popup_win_1,
        #             text="All time(abs)",
        #             command=lambda: [popup_win_1.destroy(), AllTimeabs(self)],
        #             bootstyle=DARK,
        #         )
        #         buttonTemperature.place(x=10, y=50, width=102, height=30)

        #         buttonHumidity = btk.Button(
        #             popup_win_1,
        #             text="One time(abs)",
        #             command=lambda: [popup_win_1.destroy(), OneTimeabs(self)],
        #             bootstyle=DARK,
        #         )
        #         buttonHumidity.place(x=130, y=50, width=102, height=30)

        #         label_data = tk.Label(popup_win_1)
        #         label_data["anchor"] = "n"
        #         ft = tkFont.Font(family="Helvetica", size=15)
        #         label_data["font"] = ft
        #         label_data["fg"] = "#333333"
        #         label_data["justify"] = "center"
        #         label_data["text"] = "Data type"
        #         label_data.place(x=30, y=10, width=182, height=37)

        #         button_close = btk.Button(
        #             popup_win_1,
        #             text="Close",
        #             command=lambda: [
        #                 popup_win_1.destroy(),
        #             ],
        #             bootstyle=DANGER,
        #         )
        #         button_close.place(x=75, y=100, width=92, height=30)
        #     except:
        #         popup_win_1.destroy()

        def all_timeabs(self):
            new_column_names = {"T": "time(abs)1"}
            self.data.rename(columns=new_column_names, inplace=True)
            data2 = self.data
            option_test = {
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
            DataFormating(self.root, data2, option_test, "Keithley")

        def one_timeabs(self):
            try:
                self.data2 = self.data
                self.data2 = self.data2.loc[
                    :, ~self.data2.columns.str.startswith("time")
                ]
                new_column_names = {"T": "time(abs)"}
                self.data2.rename(columns=new_column_names, inplace=True)
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
                popup_win_4 = tk.Toplevel(self.root)
                width = 223
                height = 147
                screenwidth = popup_win_4.winfo_screenwidth()
                screenheight = popup_win_4.winfo_screenheight()
                alignstr = "%dx%d+%d+%d" % (
                    width,
                    height,
                    (screenwidth - width) / 2,
                    (screenheight - height) / 2,
                )
                popup_win_4.geometry(alignstr)
                popup_win_4.resizable(width=False, height=False)
                popup_win_4.attributes("-topmost", True)
                popup_win_4.minsize(223, 147)
                popup_win_4.maxsize(223, 147)

                buttonTemperature = btk.Button(
                    popup_win_4,
                    text="Samples",
                    command=lambda: [popup_win_4.destroy(), samples(self, data2)],
                    bootstyle=DARK,
                )
                buttonTemperature.place(x=20, y=50, width=82, height=30)

                buttonHumidity = btk.Button(
                    popup_win_4,
                    text="time(abs)",
                    command=lambda: [popup_win_4.destroy(), time_abs(self, data2)],
                    bootstyle=DARK,
                )
                buttonHumidity.place(x=120, y=50, width=82, height=30)

                label_data = tk.Label(popup_win_4)
                label_data["anchor"] = "n"
                ft = tkFont.Font(family="Helvetica", size=15)
                label_data["font"] = ft
                label_data["fg"] = "#333333"
                label_data["justify"] = "center"
                label_data["text"] = "Axis X data"
                label_data.place(x=20, y=10, width=182, height=37)

                button_close = btk.Button(
                    popup_win_4,
                    text="Close",
                    command=lambda: [
                        popup_win_4.destroy(),
                    ],
                    bootstyle=DANGER,
                )
                button_close.place(x=70, y=100, width=82, height=30)
            except:
                popup_win_4.destroy()

            def samples(self, data2):
                option_test = {
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
                DataFormating(self.root, data2, option_test, "Keithley")

            def time_abs(self, data2):
                option_test = {
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
                DataFormating(self.root, data2, option_test, "Keithley")

    def read_grafana(self):
        # Grafana and logger manager data processing:
        filepaths = filedialog.askopenfilenames(
            parent=self.root,
            filetypes=(
                ("CSV Files", "*.csv"),
                ("all files", ".*"),
            ),
        )
        filepaths = list(filepaths)
        list_data = []
        for filepath in filepaths:
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
                    new_column_names = {
                        "Data/czas": "Date and Time",
                        "Temperatura[C]": "Temperature",
                        "Wilgotnosc[rH]": "Humidity",
                    }
                    self.data.rename(columns=new_column_names, inplace=True)
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
                        new_column_names = {
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
                        self.data.rename(columns=new_column_names, inplace=True)
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
                            new_column_names = {
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
                            self.data.rename(columns=new_column_names, inplace=True)
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
                        max_col = len(self.data.columns)
                        index = 0
                        for col in range(0, max_col):
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
                        max_col = len(self.data.columns)
                        index = 0
                        for col in range(0, max_col):
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
                list_data.append(data2)
            except:
                return Message(
                    "error",
                    "Wrong file format." "\nPlease contact with: piotr.klys92@gmail.com",
                    0,
                )

        try:
            data2 = pd.concat(list_data, axis=0, ignore_index=False)
            data2 = data2.sort_values(by="Date and Time", ascending=True)
            option_test = {
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
            return DataFormating(self.root, data2, option_test, "Grafana")
        except:
            return Message(
                "error",
                "Wrong file format." "\nPlease contact with: piotr.klys92@gmail.com",
                0,
            )

    def read_agilent(self):
        # Agilent data processing:
        filepaths = filedialog.askopenfilenames(
            parent=self.root,
            filetypes=(
                ("Text File", "*.txt"),
                ("all files", ".*"),
            ),
        )
        filepaths = list(filepaths)
        list_data = []
        for filepath in filepaths:
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

                new_column_names = {0: "Date and Time"}
                self.data.rename(columns=new_column_names, inplace=True)
                self.data.drop(self.data.index[[0, 1]], axis=0, inplace=True)
                self.data["Date and Time"] = pd.to_datetime(
                    self.data["Date and Time"], dayfirst=True, errors="coerce"
                )
                data2 = pd.DataFrame(self.data)
                list_data.append(data2)
            except UnicodeDecodeError:
                return Message(
                    "error",
                    "Wrong file format." "\nPlease contact with: piotr.klys92@gmail.com",
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
            DataFormating(self.root, self.data, option_test, "Agilent")
        except:
            return Message(
                "error",
                "Wrong file format." "\nPlease contact with: piotr.klys92@gmail.com",
                0,
            )
