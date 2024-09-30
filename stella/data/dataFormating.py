import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import tkinter.font as tkFont
import ttkbootstrap as btk
from ttkbootstrap.constants import *
import openpyxl
import datetime
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import (
    Paragraph,
    ParagraphProperties,
    CharacterProperties,
    Font,
)
import os
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import matplotlib.dates as mdates
from ..utility.message import Message


class DataFormating:
    def __init__(self, root, data, option, name):
        # File save
        self.data = data
        self.option = option
        try:
            # Message during data processing:
            self.popupWin2 = tk.Toplevel(root)
            width = 205
            height = 64
            screenwidth = self.popupWin2.winfo_screenwidth()
            screenheight = self.popupWin2.winfo_screenheight()
            alignstr = "%dx%d+%d+%d" % (
                width,
                height,
                (screenwidth - width) / 2,
                (screenheight - height) / 6,
            )
            self.popupWin2.geometry(alignstr)
            self.popupWin2.resizable(width=False, height=False)
            self.popupWin2.attributes("-topmost", True)
            self.popupWin2.minsize(205, 64)
            self.popupWin2.maxsize(205, 64)

            popupLabel = tk.Label(self.popupWin2)
            popupLabel["anchor"] = "n"
            ft = tkFont.Font(family="Helvetica", size=15)
            popupLabel["font"] = ft
            popupLabel["fg"] = "#333333"
            popupLabel["justify"] = "center"
            popupLabel["text"] = "Save file and wait"
            popupLabel.place(x=10, y=10, width=182, height=37)

            self.savepath = filedialog.asksaveasfilename(
                parent=root,
                initialfile="{}_Log".format(name),
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
            pd.io.formats.excel.header_style = None
            writer.close()
            self.wb = openpyxl.load_workbook(self.savepath)
            self.sheet = self.wb["Data"]
            for col in self.sheet.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2) * 1.2
                self.sheet.column_dimensions[column].width = adjusted_width
            self.Combobox(root)
        except:
            self.popupWin2.destroy()
            Message("error", "File could not be saved", 0)
            return

    def Combobox(self, root):
        # Axis X range selection:
        self.popupWin5 = tk.Toplevel(root)
        width = 325
        height = 280
        screenwidth = self.popupWin5.winfo_screenwidth()
        screenheight = self.popupWin5.winfo_screenheight()
        alignstr = "%dx%d+%d+%d" % (
            width,
            height,
            (screenwidth - width) / 2,
            (screenheight - height) / 2,
        )
        self.popupWin5.geometry(alignstr)
        self.popupWin5.resizable(width=False, height=False)
        self.popupWin5.attributes("-topmost", True)
        self.popupWin5.minsize(325, 280)
        self.popupWin5.maxsize(325, 280)

        def ChartViewer(self):
            # Chart preview with data selection:
            try:
                if self.chartViewVar.get() == 1:
                    chartList = pd.read_excel(self.savepath)
                    chartList = pd.DataFrame(chartList)
                    self.popupWin6 = tk.Toplevel(self.popupWin5)
                    width = 700
                    height = 770
                    screenwidth = self.popupWin6.winfo_screenwidth()
                    screenheight = self.popupWin6.winfo_screenheight()
                    alignstr = "%dx%d-%d+%d" % (
                        width,
                        height,
                        (screenwidth - width) / 16,
                        (screenheight - height) / 2,
                    )
                    self.popupWin6.geometry(alignstr)
                    self.popupWin6.resizable(width=False, height=False)
                    self.popupWin6.attributes("-topmost", True)
                    self.popupWin6.minsize(700, 770)
                    self.popupWin6.maxsize(700, 770)

                    chartFrame = tk.Frame(self.popupWin6)
                    chartFrame.place(x=0, y=0, width=700, height=720)

                    buttonQuit = btk.Button(
                        self.popupWin6,
                        text="Close",
                        command=lambda: [Refresh(self), self.popupWin6.destroy()],
                        bootstyle=DANGER,
                    )
                    buttonQuit.place(x=240, y=685, width=82, height=30)

                    buttonRefresh = btk.Button(
                        self.popupWin6,
                        text="Refresh",
                        command=lambda: Refresh(self),
                        bootstyle=SUCCESS,
                    )
                    buttonRefresh.place(x=340, y=685, width=82, height=30)

                    labelSelect = tk.Label(self.popupWin6)
                    labelSelect["anchor"] = "n"
                    ft = tkFont.Font(family="Helvetica", size=11)
                    labelSelect["font"] = ft
                    labelSelect["fg"] = "#333333"
                    labelSelect["justify"] = "left"
                    labelSelect["text"] = (
                        "Press ',' button to select min. range (axis X and Y).\nPress '.' button to select max. range (axis X and Y)."
                    )
                    labelSelect.place(x=0, y=720, width=400, height=60)

                    def Refresh(self):
                        # Combobox refresh
                        self.cb1.delete(0, tk.END)
                        self.cb1.insert(tk.END, self.comboStart)
                        self.cb2.delete(0, tk.END)
                        self.cb2.insert(tk.END, self.comboEnd)
                        self.valueMin.delete(0, tk.END)
                        self.valueMin.insert(tk.END, self.comboMin)
                        self.valueMax.delete(0, tk.END)
                        self.valueMax.insert(tk.END, self.comboMax)

                    figure = Figure(figsize=(6, 4), dpi=100)
                    figure_canvas = FigureCanvasTkAgg(figure, chartFrame)
                    NavigationToolbar2Tk(figure_canvas, chartFrame)

                    colorsChambers = [
                        "#752524",
                        "#F80E0A",
                        "#173B96",
                        "#1F95F8",
                        "#2AA31E",
                        "#000000",
                    ]  # Climatic and TS chambers.
                    colorsHumBath = [
                        "#752524",
                        "#F80E0A",
                        "#0A6E44",
                        "#2AA31E",
                        "#000000",
                    ]  # Salt and splash.
                    colorsSaltHum = [
                        "#752524",
                        "#F80E0A",
                        "#0A6E44",
                        "#2AA31E",
                        "#173B96",
                        "#1F95F8",
                    ]  # Salt humidity
                    colorsGasChamber = [
                        "#752524",
                        "#F80E0A",
                        "#F49D1E",
                        "#173B96",
                        "#1F95F8",
                        "#0AB1C8",
                    ]
                    ax = figure.subplots()
                    if self.option["hum"] == True:
                        ax.set_ylabel("Temperature[°C] / Humidity[%rH]")
                    else:
                        ax.set_ylabel("Temperature [°C]")
                    if (
                        self.option["samples"] == True
                        or self.option["timeabsOne"] == True
                    ):
                        for i in range(2, int(len(chartList.columns))):
                            chartList.plot(
                                kind="line", x=0, y=int(i), ax=ax, x_compat=True
                            )
                    elif self.option["timeabsAll"] == True:
                        try:
                            data = 1
                            for col in chartList:
                                chartList.plot(
                                    kind="line", x=0, y=int(data), ax=ax, x_compat=True
                                )
                                data += 2
                        except:
                            pass
                    elif self.option["indigo"] == True:
                        try:
                            for i in range(1, 3):
                                chartList.plot(
                                    kind="line", x=0, y=i, ax=ax, x_compat=True
                                )
                        except:
                            pass
                        try:
                            for i in range(4, 6):
                                chartList.plot(
                                    kind="line", x=3, y=i, ax=ax, x_compat=True
                                )
                        except:
                            pass
                    elif self.option["grafana"] == True:
                        try:
                            for i in range(1, 3):
                                chartList.plot(
                                    kind="line", x=0, y=i, ax=ax, x_compat=True
                                )
                        except:
                            pass
                        try:
                            for i in range(4, 6):
                                chartList.plot(
                                    kind="line", x=0, y=i, ax=ax, x_compat=True
                                )
                        except:
                            pass
                    elif self.option["rotronic"] == True:
                        time = 0
                        hums = 1
                        temp = 2
                        try:
                            for col in chartList:
                                chartList.plot(
                                    kind="line", x=time, y=hums, ax=ax, x_compat=True
                                )
                                chartList.plot(
                                    kind="line", x=time, y=temp, ax=ax, x_compat=True
                                )
                                time += 3
                                hums += 3
                                temp += 3
                        except:
                            pass
                    elif self.option["keithleymanager"] == True:
                        time = 0
                        temp = 1
                        try:
                            for col in chartList:
                                chartList.plot(
                                    kind="line", x=time, y=temp, ax=ax, x_compat=True
                                )
                                time += 2
                                temp += 2
                        except:
                            pass
                    elif self.option["gasChamber"] == True:
                        x = 0
                        for i in range(1, 7):
                            chartList.plot(
                                kind="line",
                                x=0,
                                y=i,
                                color=colorsGasChamber[x],
                                ax=ax,
                                x_compat=True,
                            )
                            x += 1
                    elif (
                        self.option["humBath"] == True
                        and self.option["indexZUT"] == True
                    ):
                        x = 0
                        for i in range(1, 5):
                            chartList.plot(
                                kind="line",
                                x=0,
                                y=i,
                                color=colorsHumBath[x],
                                ax=ax,
                                x_compat=True,
                            )
                            x += 1
                    elif (
                        self.option["humBath"] == True
                        and self.option["indexZUT"] == False
                    ):
                        x = 0
                        for i in range(1, int(len(chartList.columns))):
                            chartList.plot(
                                kind="line",
                                x=0,
                                y=i,
                                color=colorsHumBath[x],
                                ax=ax,
                                x_compat=True,
                            )
                            x += 1
                    elif (
                        self.option["humBath"] == False
                        and self.option["indexZUT"] == True
                    ):
                        x = 0
                        for i in range(1, 5):
                            chartList.plot(
                                kind="line",
                                x=0,
                                y=i,
                                color=colorsChambers[x],
                                ax=ax,
                                x_compat=True,
                            )
                            x += 1
                    elif self.option["coolingSys"] == True:
                        x = 0
                        for i in range(1, 4):
                            chartList.plot(
                                kind="line",
                                x=0,
                                y=i,
                                color=colorsSaltHum[x],
                                ax=ax,
                                x_compat=True,
                            )
                            x += 1
                    elif self.option["saltHum"] == True:
                        x = 0
                        for i in range(1, int(len(chartList.columns))):
                            chartList.plot(
                                kind="line",
                                x=0,
                                y=i,
                                color=colorsSaltHum[x],
                                ax=ax,
                                x_compat=True,
                            )
                            x += 1
                    elif self.option["agilent"] == True:
                        x = 0
                        for i in range(1, int(len(chartList.columns))):
                            chartList.plot(kind="line", x=0, y=i, ax=ax, x_compat=True)
                            x += 1
                    elif self.option["insight"] == True:
                        if int(len(chartList.columns)) == 2:
                            chartList.plot(
                                kind="line",
                                x=0,
                                y=1,
                                color="#FFA500",
                                ax=ax,
                                x_compat=True,
                            )
                        else:
                            x = 0
                            for i in range(1, int(len(chartList.columns))):
                                chartList.plot(
                                    kind="line", x=0, y=i, ax=ax, x_compat=True
                                )
                                x += 1
                    elif self.option["secasi"] == True:
                        x = 0
                        for i in range(1, int(len(chartList.columns))):
                            chartList.plot(
                                kind="line",
                                x=0,
                                y=int(i),
                                color=colorsChambers[x],
                                ax=ax,
                                x_compat=True,
                            )
                            x += 1
                    else:
                        try:
                            x = 0
                            for i in range(1, int(len(chartList.columns))):
                                chartList.plot(
                                    kind="line",
                                    x=0,
                                    y=i,
                                    color=colorsChambers[x],
                                    ax=ax,
                                    x_compat=True,
                                )
                                x += 1
                        except:
                            pass
                    pos = ax.get_position()
                    ax.set_position([pos.x0, pos.y0, pos.width, pos.height * 0.98])
                    ax.legend(
                        loc="upper center",
                        fancybox=True,
                        fontsize="small",
                        bbox_to_anchor=(0.5, 1.15),
                        ncol=2,
                    )
                    if (
                        self.option["samples"] == True
                        or self.option["timeabsOne"] == True
                        or self.option["secasi"] == True
                    ):
                        pass
                    else:
                        ax.xaxis.set_major_formatter(
                            mdates.DateFormatter("%y-%m-%d %H:%M")
                        )
                    ax.grid(True)
                    figure_canvas.draw()
                    figure_canvas.get_tk_widget().pack(
                        side=tk.TOP, fill=tk.BOTH, expand=1
                    )

                    def onPress(event):
                        if event.key == ",":
                            ix = event.xdata
                            iy = event.ydata
                            if (
                                self.option["samples"] == True
                                or self.option["secasi"] == True
                                or self.option["timeabsOne"] == True
                                or self.option["timeabsAll"] == True
                            ):
                                self.dateStart = int(ix)
                            else:
                                try:
                                    self.dateStart = str(mdates.num2date(ix))
                                    self.dateStart = self.dateStart.split(".")
                                    self.dateStart = self.dateStart[0]
                                except:
                                    self.dateStart = datetime.datetime.fromtimestamp(
                                        ix
                                    ).strftime("%Y-%m-%d %H:%M:%S")
                            self.comboMin = int(iy)
                            self.comboStart = self.dateStart

                        elif event.key == ".":
                            ix = event.xdata
                            iy = event.ydata
                            if (
                                self.option["samples"] == True
                                or self.option["secasi"] == True
                                or self.option["timeabsOne"] == True
                                or self.option["timeabsAll"] == True
                            ):
                                self.dateEnd = int(ix)
                            else:
                                try:
                                    self.dateEnd = str(mdates.num2date(ix))
                                    self.dateEnd = self.dateEnd.split(".")
                                    self.dateEnd = self.dateEnd[0]
                                except:
                                    self.dateEnd = datetime.datetime.fromtimestamp(
                                        ix
                                    ).strftime("%Y-%m-%d %H:%M:%S")
                            self.comboMax = int(iy)
                            self.comboEnd = self.dateEnd

                    figure.canvas.mpl_connect("key_press_event", onPress)

                    # Cross-hair
                    value = chartList.loc[1, :].values[0]
                    if (
                        self.option["samples"] == True
                        or self.option["secasi"] == True
                        or self.option["timeabsOne"] == True
                        or self.option["timeabsAll"] == True
                    ):
                        pass
                    else:
                        value = mdates.date2num(value)
                    line1 = ax.axvline(x=value, color="r", lw=0.8, ls="--")
                    line2 = ax.axhline(y=0, color="r", lw=0.8, ls="--")
                    figure_canvas.draw()

                    def onMouseMove(event):
                        ix = event.xdata
                        iy = event.ydata
                        if iy == None or ix == None:
                            pass
                        else:
                            line1.set_xdata([ix])
                            line2.set_ydata([iy])
                            figure_canvas.draw()

                    figure.canvas.mpl_connect("motion_notify_event", onMouseMove)

                else:
                    self.popupWin6.destroy()
            except:
                try:
                    self.popupWin2.destroy()
                    self.popupWin5.destroy()
                    self.popupWin6.destroy()
                except:
                    pass
                Message(
                    "error",
                    "Something went wrong"
                    "\nPlease contact with: piotr.klys92@gmail.com",
                    0,
                )

        if self.option["secasi"] or self.option["samples"] == True:
            dataList = []
            for data in self.sheet.iter_rows(
                min_row=2, max_col=1, max_row=self.sheet.max_row, values_only=True
            ):
                dataList.append(data)
            dataList = pd.DataFrame(dataList)
            dataList = dataList[0].tolist()
        else:
            if self.option["timeabsAll"] == True or self.option["timeabsOne"] == True:
                dataList = []
                for data in self.sheet.iter_rows(
                    min_row=2, max_col=1, max_row=self.sheet.max_row, values_only=True
                ):
                    dataList.append(data)
                dataList = pd.DataFrame(dataList)
                dataList[0] = dataList[0].astype(str)
                dataList = dataList[0].tolist()
            else:
                dataList = []
                for data in self.sheet.iter_rows(
                    min_row=2, max_col=1, max_row=self.sheet.max_row, values_only=True
                ):
                    dataList.append(data)
                dataList = pd.DataFrame(dataList)
                dataList[0] = pd.to_datetime(
                    dataList[0], format="%y-%m-%d %H:%M:%S", errors="coerce"
                )
                dataList[0] = dataList[0].astype(str)
                dataList = dataList[0].tolist()

        if (
            self.option["samples"] == True
            or self.option["secasi"] == True
            or self.option["timeabsOne"] == True
            or self.option["timeabsAll"] == True
        ):

            def cb1Value(event):
                self.minValue = self.cb1_var.get()
                dataList2 = [
                    int(i) for i in dataList if int(i) > int(self.cb1_var.get())
                ]
                dataList2.sort()
                self.cb2.config(values=dataList2)

        else:

            def cb1Value(event):
                self.minValue = self.cb1_var.get()
                dataList2 = [i for i in dataList if i > self.cb1_var.get()]
                dataList2.sort()
                self.cb2.config(values=dataList2)

        self.cb1_var = tk.StringVar()
        self.cb1 = ttk.Combobox(
            self.popupWin5, values=dataList, textvariable=self.cb1_var, width=20
        )
        self.cb1.place(x=80, y=100, width=189, height=30)
        self.cb1.bind("<<ComboboxSelected>>", cb1Value)

        self.cb2_var = tk.StringVar()
        self.cb2 = ttk.Combobox(self.popupWin5, textvariable=self.cb2_var, width=20)
        self.cb2.place(x=80, y=135, width=189, height=30)
        try:
            buttonNewValue = btk.Button(
                self.popupWin5,
                text="Selected",
                command=lambda: [
                    self.popupWin5.destroy(),
                    PrintValues(
                        self,
                        self.cb1_var.get(),
                        self.cb2_var.get(),
                        self.minVariable.get(),
                        self.maxVariable.get(),
                        self.popupWin6.destroy(),
                    ),
                ],
                bootstyle=DARK,
            )
            buttonNewValue.place(x=20, y=50, width=82, height=30)

            buttonOldValue = btk.Button(
                self.popupWin5,
                text="Total",
                command=lambda: [
                    self.popupWin5.destroy(),
                    OldValues(self),
                    self.popupWin6.destroy(),
                ],
                bootstyle=DARK,
            )
            buttonOldValue.place(x=120, y=50, width=82, height=30)

            buttonAllValues = btk.Button(
                self.popupWin5,
                text="Both",
                command=lambda: [
                    self.popupWin5.destroy(),
                    PrintValues(
                        self,
                        self.cb1_var.get(),
                        self.cb2_var.get(),
                        self.minVariable.get(),
                        self.maxVariable.get(),
                        both=True,
                    ),
                    OldValues(self),
                    self.popupWin6.destroy(),
                ],
                bootstyle=DARK,
            )
            buttonAllValues.place(x=220, y=50, width=82, height=30)
        except:
            pass

        labelStart = tk.Label(self.popupWin5)
        labelStart["anchor"] = "n"
        ft = tkFont.Font(family="Helvetica", size=11)
        labelStart["font"] = ft
        labelStart["fg"] = "#333333"
        labelStart["justify"] = "center"
        labelStart["text"] = "From:"
        labelStart.place(x=10, y=105, width=50, height=30)

        labelEnd = tk.Label(self.popupWin5)
        labelEnd["anchor"] = "n"
        ft = tkFont.Font(family="Helvetica", size=11)
        labelEnd["font"] = ft
        labelEnd["fg"] = "#333333"
        labelEnd["justify"] = "center"
        labelEnd["text"] = "To:"
        labelEnd.place(x=10, y=140, width=50, height=30)

        labelData = tk.Label(self.popupWin5)
        labelData["anchor"] = "n"
        ft = tkFont.Font(family="Helvetica", size=15)
        labelData["font"] = ft
        labelData["fg"] = "#333333"
        labelData["justify"] = "center"
        labelData["text"] = "Axis X - scale"
        labelData.place(x=25, y=10, width=280, height=36)

        buttonQuit = btk.Button(
            self.popupWin5,
            text="Close",
            command=lambda: [
                self.popupWin5.destroy(),
                self.popupWin6.destroy(),
            ],
            bootstyle=DANGER,
        )
        buttonQuit.place(x=120, y=240, width=82, height=30)

        self.chartViewVar = tk.IntVar()
        self.chartValue = self.chartViewVar.get()
        chartView = tk.Checkbutton(
            self.popupWin5,
            text="Chart preview",
            variable=self.chartViewVar,
            command=lambda: ChartViewer(self),
        )
        ft = tkFont.Font(family="Helvetica", size=9)
        chartView["font"] = ft
        chartView["fg"] = "#333333"
        chartView["justify"] = "center"
        chartView["offvalue"] = 0
        chartView["onvalue"] = 1
        chartView.place(x=10, y=240, width=95, height=25)

        labelMin = tk.Label(self.popupWin5)
        labelMin["anchor"] = "n"
        ft = tkFont.Font(family="Helvetica", size=11)
        labelMin["font"] = ft
        labelMin["fg"] = "#333333"
        labelMin["justify"] = "center"
        labelMin["text"] = "Min:"
        labelMin.place(x=20, y=180, width=50, height=30)

        labelScale = tk.Label(self.popupWin5)
        labelScale["anchor"] = "n"
        ft = tkFont.Font(family="Helvetica", size=9)
        labelScale["font"] = ft
        labelScale["fg"] = "#333333"
        labelScale["justify"] = "center"
        labelScale["text"] = "Axis Y Scale:"
        labelScale.place(x=120, y=180, width=80, height=30)

        labelMax = tk.Label(self.popupWin5)
        labelMax["anchor"] = "n"
        ft = tkFont.Font(family="Helvetica", size=11)
        labelMax["font"] = ft
        labelMax["fg"] = "#333333"
        labelMax["justify"] = "center"
        labelMax["text"] = "Max:"
        labelMax.place(x=220, y=180, width=50, height=30)

        self.minVariable = tk.StringVar()
        self.valueMin = tk.Entry(
            self.popupWin5, justify="center", textvariable=self.minVariable
        )
        self.valueMin["bg"] = "#ffffff"
        ft = tkFont.Font(family="Helvetica", size=11)
        self.valueMin["font"] = ft
        self.valueMin["fg"] = "#333333"
        self.valueMin.place(x=20, y=200, width=82, height=30)

        self.axScale = tk.StringVar()
        self.scaleAX = tk.Entry(
            self.popupWin5, justify="center", textvariable=self.axScale
        )
        self.scaleAX["bg"] = "#ffffff"
        ft = tkFont.Font(family="Helvetica", size=11)
        self.scaleAX["font"] = ft
        self.scaleAX["fg"] = "#333333"
        self.scaleAX.insert(0, "5")
        self.scaleAX.place(x=120, y=200, width=82, height=30)

        self.maxVariable = tk.StringVar()
        self.valueMax = tk.Entry(
            self.popupWin5, justify="center", textvariable=self.maxVariable
        )
        self.valueMax["bg"] = "#ffffff"
        ft = tkFont.Font(family="Helvetica", size=11)
        self.valueMax["font"] = ft
        self.valueMax["fg"] = "#333333"
        self.valueMax.place(x=220, y=200, width=82, height=30)

        def PrintValues(
            self, comboStartVal, comboEndVal, comboMinVal, comboMaxVal, both
        ):
            try:
                if self.option["hum"] == True:
                    self.sheetChart = self.wb.create_chartsheet("Chart")
                    self.chart = ScatterChart()
                    self.chart.x_axis.number_format = "yyyy-mm-dd HH:MM:SS"
                    self.chart.x_axis.majorTimeUnit = "months"
                    self.chart.x_axis.tickLblPos = "low"
                    self.chart.legend.position = "b"
                    self.chart.x_axis.title = "Date and Time [yyyy-mm-dd hh:mm:ss]"
                    self.chart.y_axis.title = "Temperature[°C] / Humidity[%rH]"

                    # Text format - Axis
                    normal_font = Font(typeface="Calibri")
                    cp_text = CharacterProperties(latin=normal_font, sz=900, b=False)
                    self.chart.x_axis.txPr = RichText(
                        p=[
                            Paragraph(
                                pPr=ParagraphProperties(defRPr=cp_text),
                                endParaRPr=cp_text,
                            )
                        ]
                    )
                    self.chart.y_axis.txPr = RichText(
                        p=[
                            Paragraph(
                                pPr=ParagraphProperties(defRPr=cp_text),
                                endParaRPr=cp_text,
                            )
                        ]
                    )

                    # Text format - Title
                    cp_text1 = CharacterProperties(latin=normal_font, sz=1000, b=False)
                    pp1 = ParagraphProperties(defRPr=cp_text1)
                    self.chart.x_axis.title.tx.rich.p[0].pPr = pp1
                    self.chart.y_axis.title.tx.rich.p[0].pPr = pp1
                    self.chart.x_axis.txPr.properties.rot = "-2700000"
                    self.chart.y_axis.majorUnit = self.axScale.get()
                    self.minValue = comboStartVal
                    self.maxValue = comboEndVal
                    try:
                        self.chart.y_axis.scaling.min = comboMinVal
                        self.chart.y_axis.scaling.max = comboMaxVal
                    except:
                        pass
                    minScale = pd.Timestamp(self.minValue)
                    maxScale = pd.Timestamp(self.maxValue)
                    minConverted = (
                        minScale - pd.Timestamp("1899-12-30")
                    ).total_seconds() / 86400
                    maxConverted = (
                        maxScale - pd.Timestamp("1899-12-30")
                    ).total_seconds() / 86400
                    self.chart.x_axis.scaling.min = minConverted
                    self.chart.x_axis.scaling.max = maxConverted

                    if self.option["rotronic"] == True:
                        colTime = self.sheet.min_column
                        colMeasRangeMin = 2
                        colMeasRangeMax = 4
                        while colTime <= self.sheet.max_column:
                            data = Reference(
                                self.sheet,
                                min_col=colTime,
                                min_row=2,
                                max_row=self.sheet.max_row,
                            )
                            colTime += 3
                            for i in range(colMeasRangeMin, colMeasRangeMax):
                                values = Reference(
                                    self.sheet,
                                    min_col=i,
                                    min_row=1,
                                    max_row=self.sheet.max_row,
                                )
                                series = Series(values, data, title_from_data=True)
                                self.chart.series.append(series)
                            colMeasRangeMin += 3
                            colMeasRangeMax += 3
                    elif self.option["indigo"] == True:
                        data = Reference(
                            self.sheet, min_col=1, min_row=2, max_row=self.sheet.max_row
                        )
                        for i in range(2, 4):
                            values = Reference(
                                self.sheet,
                                min_col=i,
                                min_row=1,
                                max_row=self.sheet.max_row,
                            )
                            series = Series(values, data, title_from_data=True)
                            self.chart.series.append(series)
                        try:
                            data2 = Reference(
                                self.sheet,
                                min_col=4,
                                min_row=2,
                                max_row=self.sheet.max_row,
                            )
                            for i in range(5, 7):
                                values2 = Reference(
                                    self.sheet,
                                    min_col=i,
                                    min_row=1,
                                    max_row=self.sheet.max_row,
                                )
                                series2 = Series(values2, data2, title_from_data=True)
                                self.chart.series.append(series2)
                        except:
                            pass
                    else:
                        if self.option["indexZUT"] == True:
                            index = int(6)
                        elif self.option["coolingSys"] == True:
                            index = int(5)
                        else:
                            index = int(self.sheet.max_column + 1)
                        data = Reference(
                            self.sheet, min_col=1, min_row=2, max_row=self.sheet.max_row
                        )
                        for i in range(2, index):
                            values = Reference(
                                self.sheet,
                                min_col=i,
                                min_row=1,
                                max_row=self.sheet.max_row,
                            )
                            series = Series(values, data, title_from_data=True)
                            self.chart.series.append(series)
                    self.sheetChart.add_chart(self.chart)
                    for serie in self.chart.series:
                        serie.graphicalProperties.line.width = 15000
                    if (
                        self.option["grafana"] == True
                        or self.option["rotronic"] == True
                        or self.option["indigo"] == True
                    ):
                        pass
                    else:
                        if self.option["insight"] == True:
                            # Temperature
                            s1 = self.chart.series[0]
                            s1.graphicalProperties.line.solidFill = "F80E0A"

                            # Humidity
                            s2 = self.chart.series[0]
                            s2.graphicalProperties.line.solidFill = "1F95F8"
                        else:
                            # Temperature Set
                            s1 = self.chart.series[0]
                            s1.graphicalProperties.line.solidFill = "752524"

                            # Temperature
                            s2 = self.chart.series[1]
                            s2.graphicalProperties.line.solidFill = "F80E0A"

                            if self.option["gasChamber"] == True:
                                # Temperature Set
                                s1 = self.chart.series[0]
                                s1.graphicalProperties.line.solidFill = "752524"

                                # Temperature IN
                                s2 = self.chart.series[1]
                                s2.graphicalProperties.line.solidFill = "F80E0A"

                                # Temperature OUT
                                s3 = self.chart.series[2]
                                s3.graphicalProperties.line.solidFill = "F49D1E"

                                # Humidity Set
                                s4 = self.chart.series[3]
                                s4.graphicalProperties.line.solidFill = "173B96"

                                # Humidity IN
                                s5 = self.chart.series[4]
                                s5.graphicalProperties.line.solidFill = "1F95F8"

                                # Humidity OUT
                                s6 = self.chart.series[5]
                                s6.graphicalProperties.line.solidFill = "0AB1C8"
                            else:
                                if self.option["saltHum"] == True:
                                    # Humidifier Set
                                    s3 = self.chart.series[2]
                                    s3.graphicalProperties.line.solidFill = "0A6E44"

                                    # Humidifier
                                    s4 = self.chart.series[3]
                                    s4.graphicalProperties.line.solidFill = "2AA31E"

                                    # Humidity Set
                                    s5 = self.chart.series[4]
                                    s5.graphicalProperties.line.solidFill = "173B96"

                                    # Humidity
                                    s6 = self.chart.series[5]
                                    s6.graphicalProperties.line.solidFill = "1F95F8"
                                else:
                                    try:
                                        # Humidity Set
                                        s3 = self.chart.series[2]
                                        s3.graphicalProperties.line.solidFill = "173B96"
                                    except:
                                        pass

                                    try:
                                        # Humidity
                                        s4 = self.chart.series[3]
                                        s4.graphicalProperties.line.solidFill = "1F95F8"
                                    except:
                                        pass

                                    try:
                                        # Temperature Bath
                                        s5 = self.chart.series[4]
                                        s5.graphicalProperties.line.solidFill = "2AA31E"
                                    except:
                                        pass

                    self.wb.save(self.savepath)
                    self.popupWin2.destroy()
                    if both == True:
                        pass
                    else:
                        try:
                            os.startfile(self.savepath)
                        except:
                            Message(
                                "warning",
                                "File can not be opened by Stella"
                                "\nPlease open it manually",
                                0,
                            )
                else:
                    self.sheetChart = self.wb.create_chartsheet("Chart")
                    self.chart = ScatterChart()
                    if self.option["timeabsOne"] == True:
                        self.chart.x_axis.tickLblPos = "low"
                        self.chart.legend.position = "b"

                        self.chart.x_axis.title = "Samples"
                        self.chart.y_axis.title = "Temperature[°C]"

                        # Text format - Axis
                        normal_font = Font(typeface="Calibri")
                        cp_text = CharacterProperties(
                            latin=normal_font, sz=900, b=False
                        )
                        self.chart.x_axis.txPr = RichText(
                            p=[
                                Paragraph(
                                    pPr=ParagraphProperties(defRPr=cp_text),
                                    endParaRPr=cp_text,
                                )
                            ]
                        )
                        self.chart.y_axis.txPr = RichText(
                            p=[
                                Paragraph(
                                    pPr=ParagraphProperties(defRPr=cp_text),
                                    endParaRPr=cp_text,
                                )
                            ]
                        )

                        # Text format - Title
                        cp_text1 = CharacterProperties(
                            latin=normal_font, sz=1000, b=False
                        )
                        pp1 = ParagraphProperties(defRPr=cp_text1)
                        self.chart.x_axis.title.tx.rich.p[0].pPr = pp1
                        self.chart.y_axis.title.tx.rich.p[0].pPr = pp1
                        self.chart.x_axis.txPr.properties.rot = "-2700000"
                        self.chart.y_axis.majorUnit = self.axScale.get()
                        self.minValue = comboStartVal
                        self.maxValue = comboEndVal
                        try:
                            self.chart.y_axis.scaling.min = comboMinVal
                            self.chart.y_axis.scaling.max = comboMaxVal
                        except:
                            pass
                        self.chart.x_axis.scaling.min = self.minValue
                        self.chart.x_axis.scaling.max = self.maxValue

                        index = int(self.sheet.max_column + 1)

                        data = Reference(
                            self.sheet, min_col=2, min_row=2, max_row=self.sheet.max_row
                        )
                        for i in range(3, index):
                            values = Reference(
                                self.sheet,
                                min_col=i,
                                min_row=1,
                                max_row=self.sheet.max_row,
                            )
                            series = Series(values, data, title_from_data=True)
                            self.chart.series.append(series)
                        self.sheetChart.add_chart(self.chart)
                        for serie in self.chart.series:
                            serie.graphicalProperties.line.width = 15000
                    else:
                        if self.option["samples"] == True:
                            self.chart.x_axis.tickLblPos = "low"
                            self.chart.legend.position = "b"

                            self.chart.x_axis.title = "Samples"
                            self.chart.y_axis.title = "Temperature[°C]"

                            # Text format - Axis
                            normal_font = Font(typeface="Calibri")
                            cp_text = CharacterProperties(
                                latin=normal_font, sz=900, b=False
                            )
                            self.chart.x_axis.txPr = RichText(
                                p=[
                                    Paragraph(
                                        pPr=ParagraphProperties(defRPr=cp_text),
                                        endParaRPr=cp_text,
                                    )
                                ]
                            )
                            self.chart.y_axis.txPr = RichText(
                                p=[
                                    Paragraph(
                                        pPr=ParagraphProperties(defRPr=cp_text),
                                        endParaRPr=cp_text,
                                    )
                                ]
                            )

                            # Text format - Title
                            cp_text1 = CharacterProperties(
                                latin=normal_font, sz=1000, b=False
                            )
                            pp1 = ParagraphProperties(defRPr=cp_text1)
                            self.chart.x_axis.title.tx.rich.p[0].pPr = pp1
                            self.chart.y_axis.title.tx.rich.p[0].pPr = pp1
                            self.chart.x_axis.txPr.properties.rot = "-2700000"
                            self.chart.y_axis.majorUnit = self.axScale.get()
                            self.minValue = comboStartVal
                            self.maxValue = comboEndVal
                            try:
                                self.chart.y_axis.scaling.min = comboMinVal
                                self.chart.y_axis.scaling.max = comboMaxVal
                            except:
                                pass
                            self.chart.x_axis.scaling.min = self.minValue
                            self.chart.x_axis.scaling.max = self.maxValue

                            index = int(self.sheet.max_column + 1)

                            data = Reference(
                                self.sheet,
                                min_col=1,
                                min_row=2,
                                max_row=self.sheet.max_row,
                            )
                            for i in range(3, index):
                                values = Reference(
                                    self.sheet,
                                    min_col=i,
                                    min_row=1,
                                    max_row=self.sheet.max_row,
                                )
                                series = Series(values, data, title_from_data=True)
                                self.chart.series.append(series)
                            self.sheetChart.add_chart(self.chart)
                            for serie in self.chart.series:
                                serie.graphicalProperties.line.width = 15000
                        else:
                            if self.option["timeabsAll"] == True:
                                self.chart.x_axis.tickLblPos = "low"
                                self.chart.legend.position = "b"

                                self.chart.x_axis.title = "Samples"
                                self.chart.y_axis.title = "Temperature[°C]"

                                # Text format - Axis
                                normal_font = Font(typeface="Calibri")
                                cp_text = CharacterProperties(
                                    latin=normal_font, sz=900, b=False
                                )
                                self.chart.x_axis.txPr = RichText(
                                    p=[
                                        Paragraph(
                                            pPr=ParagraphProperties(defRPr=cp_text),
                                            endParaRPr=cp_text,
                                        )
                                    ]
                                )
                                self.chart.y_axis.txPr = RichText(
                                    p=[
                                        Paragraph(
                                            pPr=ParagraphProperties(defRPr=cp_text),
                                            endParaRPr=cp_text,
                                        )
                                    ]
                                )

                                # Text format - Title
                                cp_text1 = CharacterProperties(
                                    latin=normal_font, sz=1000, b=False
                                )
                                pp1 = ParagraphProperties(defRPr=cp_text1)
                                self.chart.x_axis.title.tx.rich.p[0].pPr = pp1
                                self.chart.y_axis.title.tx.rich.p[0].pPr = pp1
                                self.chart.x_axis.txPr.properties.rot = "-2700000"
                                self.chart.y_axis.majorUnit = self.axScale.get()
                                self.minValue = comboStartVal
                                self.maxValue = comboEndVal
                                try:
                                    self.chart.y_axis.scaling.min = comboMinVal
                                    self.chart.y_axis.scaling.max = comboMaxVal
                                except:
                                    pass
                                self.chart.x_axis.scaling.min = self.minValue
                                self.chart.x_axis.scaling.max = self.maxValue

                                colTime = 3
                                colMeas = 2
                                while colTime <= self.sheet.max_column:
                                    data = Reference(
                                        self.sheet,
                                        min_col=colTime,
                                        min_row=2,
                                        max_row=self.sheet.max_row,
                                    )
                                    colTime += 2
                                    if colMeas <= self.sheet.max_column:
                                        values = Reference(
                                            self.sheet,
                                            min_col=colMeas,
                                            min_row=1,
                                            max_row=self.sheet.max_row,
                                        )
                                        series = Series(
                                            values, data, title_from_data=True
                                        )
                                        self.chart.series.append(series)
                                    colMeas += 2
                                self.sheetChart.add_chart(self.chart)
                                for serie in self.chart.series:
                                    serie.graphicalProperties.line.width = 15000
                            else:
                                if self.option["secasi"] == True:
                                    self.chart.x_axis.tickLblPos = "low"
                                    self.chart.legend.position = "b"

                                    self.chart.x_axis.title = "Time [s]"
                                    self.chart.y_axis.title = "Temperature[°C]"

                                    # Text format - Axis
                                    normal_font = Font(typeface="Calibri")
                                    cp_text = CharacterProperties(
                                        latin=normal_font, sz=900, b=False
                                    )
                                    self.chart.x_axis.txPr = RichText(
                                        p=[
                                            Paragraph(
                                                pPr=ParagraphProperties(defRPr=cp_text),
                                                endParaRPr=cp_text,
                                            )
                                        ]
                                    )
                                    self.chart.y_axis.txPr = RichText(
                                        p=[
                                            Paragraph(
                                                pPr=ParagraphProperties(defRPr=cp_text),
                                                endParaRPr=cp_text,
                                            )
                                        ]
                                    )

                                    # Text format - Title
                                    cp_text1 = CharacterProperties(
                                        latin=normal_font, sz=1000, b=False
                                    )
                                    pp1 = ParagraphProperties(defRPr=cp_text1)
                                    self.chart.x_axis.title.tx.rich.p[0].pPr = pp1
                                    self.chart.y_axis.title.tx.rich.p[0].pPr = pp1
                                    self.chart.x_axis.txPr.properties.rot = "-2700000"
                                    self.chart.y_axis.majorUnit = self.axScale.get()
                                    self.minValue = comboStartVal
                                    self.maxValue = comboEndVal
                                    try:
                                        self.chart.y_axis.scaling.min = comboMinVal
                                        self.chart.y_axis.scaling.max = comboMaxVal
                                    except:
                                        pass
                                    self.chart.x_axis.scaling.min = self.minValue
                                    self.chart.x_axis.scaling.max = self.maxValue
                                else:
                                    self.chart.x_axis.number_format = (
                                        "yyyy-mm-dd HH:MM:SS"
                                    )
                                    self.chart.x_axis.majorTimeUnit = "months"
                                    self.chart.x_axis.tickLblPos = "low"
                                    self.chart.legend.position = "b"
                                    self.chart.x_axis.title = (
                                        "Date and Time [yyyy-mm-dd hh:mm:ss]"
                                    )
                                    self.chart.y_axis.title = "Temperature[°C]"

                                    # Text format - Axis
                                    normal_font = Font(typeface="Calibri")
                                    cp_text = CharacterProperties(
                                        latin=normal_font, sz=900, b=False
                                    )
                                    self.chart.x_axis.txPr = RichText(
                                        p=[
                                            Paragraph(
                                                pPr=ParagraphProperties(defRPr=cp_text),
                                                endParaRPr=cp_text,
                                            )
                                        ]
                                    )
                                    self.chart.y_axis.txPr = RichText(
                                        p=[
                                            Paragraph(
                                                pPr=ParagraphProperties(defRPr=cp_text),
                                                endParaRPr=cp_text,
                                            )
                                        ]
                                    )

                                    # Text format - Title
                                    cp_text1 = CharacterProperties(
                                        latin=normal_font, sz=1000, b=False
                                    )
                                    pp1 = ParagraphProperties(defRPr=cp_text1)
                                    self.chart.x_axis.title.tx.rich.p[0].pPr = pp1
                                    self.chart.y_axis.title.tx.rich.p[0].pPr = pp1
                                    self.chart.x_axis.txPr.properties.rot = "-2700000"
                                    self.chart.y_axis.majorUnit = self.axScale.get()
                                    self.minValue = comboStartVal
                                    self.maxValue = comboEndVal
                                    try:
                                        self.chart.y_axis.scaling.min = comboMinVal
                                        self.chart.y_axis.scaling.max = comboMaxVal
                                    except:
                                        pass
                                    minScale = pd.Timestamp(self.minValue)
                                    maxScale = pd.Timestamp(self.maxValue)
                                    minConverted = (
                                        minScale - pd.Timestamp("1899-12-30")
                                    ).total_seconds() / 86400
                                    maxConverted = (
                                        maxScale - pd.Timestamp("1899-12-30")
                                    ).total_seconds() / 86400
                                    self.chart.x_axis.scaling.min = minConverted
                                    self.chart.x_axis.scaling.max = maxConverted
                                if self.option["keithleymanager"] == True:
                                    colTime = 1
                                    colMeas = 2
                                    while colTime <= self.sheet.max_column:
                                        data = Reference(
                                            self.sheet,
                                            min_col=colTime,
                                            min_row=2,
                                            max_row=self.sheet.max_row,
                                        )
                                        colTime += 2
                                        if colMeas <= self.sheet.max_column:
                                            values = Reference(
                                                self.sheet,
                                                min_col=colMeas,
                                                min_row=1,
                                                max_row=self.sheet.max_row,
                                            )
                                            series = Series(
                                                values, data, title_from_data=True
                                            )
                                            self.chart.series.append(series)
                                        colMeas += 2
                                    self.sheetChart.add_chart(self.chart)
                                    for serie in self.chart.series:
                                        serie.graphicalProperties.line.width = 15000
                                else:
                                    if self.option["indexZUT"] == True:
                                        index = int(6)
                                    elif self.option["coolingSys"] == True:
                                        index = int(5)
                                    else:
                                        index = int(self.sheet.max_column + 1)
                                    data = Reference(
                                        self.sheet,
                                        min_col=1,
                                        min_row=2,
                                        max_row=self.sheet.max_row,
                                    )
                                    for i in range(2, index):
                                        values = Reference(
                                            self.sheet,
                                            min_col=i,
                                            min_row=1,
                                            max_row=self.sheet.max_row,
                                        )
                                        series = Series(
                                            values, data, title_from_data=True
                                        )
                                        self.chart.series.append(series)
                                    self.sheetChart.add_chart(self.chart)
                                    for serie in self.chart.series:
                                        serie.graphicalProperties.line.width = 15000

                                if (
                                    self.option["agilent"] == True
                                    or self.option["keithleymanager"] == True
                                ):
                                    pass
                                else:
                                    if self.option["insight"] == True:
                                        # Temperature
                                        s1 = self.chart.series[0]
                                        s1.graphicalProperties.line.solidFill = "F80E0A"
                                    else:
                                        # Temperature Hot Set
                                        s1 = self.chart.series[0]
                                        s1.graphicalProperties.line.solidFill = "752524"

                                        # Temperature Hot
                                        s2 = self.chart.series[1]
                                        s2.graphicalProperties.line.solidFill = "F80E0A"
                                        if self.option["humBath"] == True:
                                            # Humidifier Set
                                            s3 = self.chart.series[2]
                                            s3.graphicalProperties.line.solidFill = (
                                                "0A6E44"
                                            )

                                            # Humidifier
                                            s4 = self.chart.series[3]
                                            s4.graphicalProperties.line.solidFill = (
                                                "2AA31E"
                                            )
                                        else:
                                            try:
                                                # Temperature Cold Set
                                                s3 = self.chart.series[2]
                                                s3.graphicalProperties.line.solidFill = (
                                                    "173B96"
                                                )
                                            except:
                                                pass
                                            try:
                                                # Temperature Cold
                                                s4 = self.chart.series[3]
                                                s4.graphicalProperties.line.solidFill = (
                                                    "1F95F8"
                                                )
                                            except:
                                                pass
                                            try:
                                                # Temperature Basket
                                                s5 = self.chart.series[4]
                                                s5.graphicalProperties.line.solidFill = (
                                                    "2AA31E"
                                                )
                                            except:
                                                pass

                    self.wb.save(self.savepath)
                    self.popupWin2.destroy()
                    if both == True:
                        pass
                    else:
                        try:
                            os.startfile(self.savepath)
                        except:
                            Message(
                                "warning",
                                "File can not be opened by Stella"
                                "\nPlease open it manually",
                                0,
                            )
            except:
                try:
                    self.popupWin2.destroy()
                except:
                    pass
                Message(
                    "error",
                    "Something went wrong"
                    "\nPlease contact with: piotr.klys92@gmail.com",
                    0,
                )

        def OldValues(self):
            try:
                # Total duration chart plotting
                if self.option["hum"] == True:
                    self.sheetChart = self.wb.create_chartsheet("ChartAll")
                    self.chart = ScatterChart()
                    self.chart.x_axis.number_format = "yyyy-mm-dd HH:MM:SS"
                    self.chart.x_axis.majorTimeUnit = "months"
                    self.chart.x_axis.tickLblPos = "low"
                    self.chart.legend.position = "b"
                    self.chart.x_axis.title = "Date and Time [yyyy-mm-dd hh:mm:ss]"
                    self.chart.y_axis.title = "Temperature[°C] / Humidity[%rH]"

                    # Text format - Axis
                    normal_font = Font(typeface="Calibri")
                    cp_text = CharacterProperties(latin=normal_font, sz=900, b=False)
                    self.chart.x_axis.txPr = RichText(
                        p=[
                            Paragraph(
                                pPr=ParagraphProperties(defRPr=cp_text),
                                endParaRPr=cp_text,
                            )
                        ]
                    )
                    self.chart.y_axis.txPr = RichText(
                        p=[
                            Paragraph(
                                pPr=ParagraphProperties(defRPr=cp_text),
                                endParaRPr=cp_text,
                            )
                        ]
                    )

                    # Text format - Title
                    cp_text1 = CharacterProperties(latin=normal_font, sz=1000, b=False)
                    pp1 = ParagraphProperties(defRPr=cp_text1)
                    self.chart.x_axis.title.tx.rich.p[0].pPr = pp1
                    self.chart.y_axis.title.tx.rich.p[0].pPr = pp1
                    self.chart.x_axis.txPr.properties.rot = "-2700000"
                    self.chart.y_axis.majorUnit = self.axScale.get()
                    self.minValue = self.sheet["A" + str(self.sheet.min_row + 1)].value
                    self.maxValue = self.sheet["A" + str(self.sheet.max_row)].value
                    minScale = pd.Timestamp(self.minValue)
                    maxScale = pd.Timestamp(self.maxValue)
                    minConverted = (
                        minScale - pd.Timestamp("1899-12-30")
                    ).total_seconds() / 86400
                    maxConverted = (
                        maxScale - pd.Timestamp("1899-12-30")
                    ).total_seconds() / 86400
                    self.chart.x_axis.scaling.min = minConverted
                    self.chart.x_axis.scaling.max = maxConverted

                    if self.option["rotronic"] == True:
                        colTime = self.sheet.min_column
                        colMeasRangeMin = 2
                        colMeasRangeMax = 4
                        while colTime <= self.sheet.max_column:
                            data = Reference(
                                self.sheet,
                                min_col=colTime,
                                min_row=2,
                                max_row=self.sheet.max_row,
                            )
                            colTime += 3
                            for i in range(colMeasRangeMin, colMeasRangeMax):
                                values = Reference(
                                    self.sheet,
                                    min_col=i,
                                    min_row=1,
                                    max_row=self.sheet.max_row,
                                )
                                series = Series(values, data, title_from_data=True)
                                self.chart.series.append(series)
                            colMeasRangeMin += 3
                            colMeasRangeMax += 3
                        self.sheetChart.add_chart(self.chart)
                        for serie in self.chart.series:
                            serie.graphicalProperties.line.width = 15000
                    else:
                        if self.option["indigo"] == True:
                            data = Reference(
                                self.sheet,
                                min_col=1,
                                min_row=2,
                                max_row=self.sheet.max_row,
                            )
                            for i in range(2, 4):
                                values = Reference(
                                    self.sheet,
                                    min_col=i,
                                    min_row=1,
                                    max_row=self.sheet.max_row,
                                )
                                series = Series(values, data, title_from_data=True)
                                self.chart.series.append(series)
                            try:
                                data2 = Reference(
                                    self.sheet,
                                    min_col=4,
                                    min_row=2,
                                    max_row=self.sheet.max_row,
                                )
                                for i in range(5, 7):
                                    values2 = Reference(
                                        self.sheet,
                                        min_col=i,
                                        min_row=1,
                                        max_row=self.sheet.max_row,
                                    )
                                    series2 = Series(
                                        values2, data2, title_from_data=True
                                    )
                                    self.chart.series.append(series2)
                            except:
                                pass

                            self.sheetChart.add_chart(self.chart)
                            for serie in self.chart.series:
                                serie.graphicalProperties.line.width = 15000
                        else:
                            if self.option["indexZUT"] == True:
                                index = int(6)
                            elif self.option["coolingSys"] == True:
                                index = int(5)
                            else:
                                index = int(self.sheet.max_column + 1)
                            data = Reference(
                                self.sheet,
                                min_col=1,
                                min_row=2,
                                max_row=self.sheet.max_row,
                            )
                            for i in range(2, index):
                                values = Reference(
                                    self.sheet,
                                    min_col=i,
                                    min_row=1,
                                    max_row=self.sheet.max_row,
                                )
                                series = Series(values, data, title_from_data=True)
                                self.chart.series.append(series)
                            self.sheetChart.add_chart(self.chart)
                            for serie in self.chart.series:
                                serie.graphicalProperties.line.width = 15000

                            if self.option["grafana"] == True:
                                pass
                            else:
                                if self.option["insight"] == True:
                                    # Temperature
                                    s1 = self.chart.series[0]
                                    s1.graphicalProperties.line.solidFill = "F80E0A"

                                    # Humidity
                                    s2 = self.chart.series[0]
                                    s2.graphicalProperties.line.solidFill = "1F95F8"
                                else:
                                    # Temperature Set
                                    s1 = self.chart.series[0]
                                    s1.graphicalProperties.line.solidFill = "752524"

                                    # Temperature
                                    s2 = self.chart.series[1]
                                    s2.graphicalProperties.line.solidFill = "F80E0A"

                                    if self.option["gasChamber"] == True:
                                        # Temperature Set
                                        s1 = self.chart.series[0]
                                        s1.graphicalProperties.line.solidFill = "752524"

                                        # Temperature IN
                                        s2 = self.chart.series[1]
                                        s2.graphicalProperties.line.solidFill = "F80E0A"

                                        # Temperature OUT
                                        s3 = self.chart.series[2]
                                        s3.graphicalProperties.line.solidFill = "F49D1E"

                                        # Humidity Set
                                        s4 = self.chart.series[3]
                                        s4.graphicalProperties.line.solidFill = "173B96"

                                        # Humidity IN
                                        s5 = self.chart.series[4]
                                        s5.graphicalProperties.line.solidFill = "1F95F8"

                                        # Humidity OUT
                                        s6 = self.chart.series[5]
                                        s6.graphicalProperties.line.solidFill = "0AB1C8"
                                    else:
                                        if self.option["saltHum"] == True:
                                            # Humidifier Set
                                            s3 = self.chart.series[2]
                                            s3.graphicalProperties.line.solidFill = (
                                                "0A6E44"
                                            )

                                            # Humidifier
                                            s4 = self.chart.series[3]
                                            s4.graphicalProperties.line.solidFill = (
                                                "2AA31E"
                                            )

                                            # Humidity Set
                                            s5 = self.chart.series[4]
                                            s5.graphicalProperties.line.solidFill = (
                                                "173B96"
                                            )

                                            # Humidity
                                            s6 = self.chart.series[5]
                                            s6.graphicalProperties.line.solidFill = (
                                                "1F95F8"
                                            )
                                        else:
                                            try:
                                                # Humidity Set
                                                s3 = self.chart.series[2]
                                                s3.graphicalProperties.line.solidFill = (
                                                    "173B96"
                                                )
                                            except:
                                                pass

                                            try:
                                                # Humidity
                                                s4 = self.chart.series[3]
                                                s4.graphicalProperties.line.solidFill = (
                                                    "1F95F8"
                                                )
                                            except:
                                                pass

                                            try:
                                                # Temperature Bath
                                                s5 = self.chart.series[4]
                                                s5.graphicalProperties.line.solidFill = (
                                                    "2AA31E"
                                                )
                                            except:
                                                pass

                    self.wb.save(self.savepath)
                    self.popupWin2.destroy()
                    try:
                        os.startfile(self.savepath)
                    except:
                        Message(
                            "warning",
                            "File can not be opened by Stella"
                            "\nPlease open it manually",
                            0,
                        )

                else:
                    self.sheetChart = self.wb.create_chartsheet("ChartAll")
                    self.chart = ScatterChart()
                    if self.option["timeabsOne"] == True:
                        self.chart.x_axis.tickLblPos = "low"
                        self.chart.legend.position = "b"

                        self.chart.x_axis.title = "Samples"
                        self.chart.y_axis.title = "Temperature[°C]"

                        # Text format - Axis
                        normal_font = Font(typeface="Calibri")
                        cp_text = CharacterProperties(
                            latin=normal_font, sz=900, b=False
                        )
                        self.chart.x_axis.txPr = RichText(
                            p=[
                                Paragraph(
                                    pPr=ParagraphProperties(defRPr=cp_text),
                                    endParaRPr=cp_text,
                                )
                            ]
                        )
                        self.chart.y_axis.txPr = RichText(
                            p=[
                                Paragraph(
                                    pPr=ParagraphProperties(defRPr=cp_text),
                                    endParaRPr=cp_text,
                                )
                            ]
                        )

                        # Text format - Title
                        cp_text1 = CharacterProperties(
                            latin=normal_font, sz=1000, b=False
                        )
                        pp1 = ParagraphProperties(defRPr=cp_text1)
                        self.chart.x_axis.title.tx.rich.p[0].pPr = pp1
                        self.chart.y_axis.title.tx.rich.p[0].pPr = pp1
                        self.chart.x_axis.txPr.properties.rot = "-2700000"
                        self.chart.y_axis.majorUnit = self.axScale.get()
                        minValue = self.sheet["A" + str(self.sheet.min_row + 1)].value
                        maxValue = self.sheet["A" + str(self.sheet.max_row)].value
                        self.chart.x_axis.scaling.min = minValue
                        self.chart.x_axis.scaling.max = maxValue

                        index = int(self.sheet.max_column + 1)

                        data = Reference(
                            self.sheet, min_col=2, min_row=2, max_row=self.sheet.max_row
                        )
                        for i in range(3, index):
                            values = Reference(
                                self.sheet,
                                min_col=i,
                                min_row=1,
                                max_row=self.sheet.max_row,
                            )
                            series = Series(values, data, title_from_data=True)
                            self.chart.series.append(series)
                        self.sheetChart.add_chart(self.chart)
                        for serie in self.chart.series:
                            serie.graphicalProperties.line.width = 15000
                    else:
                        if self.option["samples"] == True:
                            self.chart.x_axis.tickLblPos = "low"
                            self.chart.legend.position = "b"

                            self.chart.x_axis.title = "Samples"
                            self.chart.y_axis.title = "Temperature[°C]"

                            # Text format - Axis
                            normal_font = Font(typeface="Calibri")
                            cp_text = CharacterProperties(
                                latin=normal_font, sz=900, b=False
                            )
                            self.chart.x_axis.txPr = RichText(
                                p=[
                                    Paragraph(
                                        pPr=ParagraphProperties(defRPr=cp_text),
                                        endParaRPr=cp_text,
                                    )
                                ]
                            )
                            self.chart.y_axis.txPr = RichText(
                                p=[
                                    Paragraph(
                                        pPr=ParagraphProperties(defRPr=cp_text),
                                        endParaRPr=cp_text,
                                    )
                                ]
                            )

                            # Text format - Title
                            cp_text1 = CharacterProperties(
                                latin=normal_font, sz=1000, b=False
                            )
                            pp1 = ParagraphProperties(defRPr=cp_text1)
                            self.chart.x_axis.title.tx.rich.p[0].pPr = pp1
                            self.chart.y_axis.title.tx.rich.p[0].pPr = pp1
                            self.chart.x_axis.txPr.properties.rot = "-2700000"
                            self.chart.y_axis.majorUnit = self.axScale.get()
                            minValue = self.sheet[
                                "A" + str(self.sheet.min_row + 1)
                            ].value
                            maxValue = self.sheet["A" + str(self.sheet.max_row)].value
                            self.chart.x_axis.scaling.min = minValue
                            self.chart.x_axis.scaling.max = maxValue

                            index = int(self.sheet.max_column + 1)

                            data = Reference(
                                self.sheet,
                                min_col=1,
                                min_row=2,
                                max_row=self.sheet.max_row,
                            )
                            for i in range(3, index):
                                values = Reference(
                                    self.sheet,
                                    min_col=i,
                                    min_row=1,
                                    max_row=self.sheet.max_row,
                                )
                                series = Series(values, data, title_from_data=True)
                                self.chart.series.append(series)
                            self.sheetChart.add_chart(self.chart)
                            for serie in self.chart.series:
                                serie.graphicalProperties.line.width = 15000
                        else:
                            if self.option["timeabsAll"] == True:
                                self.chart.x_axis.tickLblPos = "low"
                                self.chart.legend.position = "b"

                                self.chart.x_axis.title = "Samples"
                                self.chart.y_axis.title = "Temperature[°C]"

                                # Text format - Axis
                                normal_font = Font(typeface="Calibri")
                                cp_text = CharacterProperties(
                                    latin=normal_font, sz=900, b=False
                                )
                                self.chart.x_axis.txPr = RichText(
                                    p=[
                                        Paragraph(
                                            pPr=ParagraphProperties(defRPr=cp_text),
                                            endParaRPr=cp_text,
                                        )
                                    ]
                                )
                                self.chart.y_axis.txPr = RichText(
                                    p=[
                                        Paragraph(
                                            pPr=ParagraphProperties(defRPr=cp_text),
                                            endParaRPr=cp_text,
                                        )
                                    ]
                                )

                                # Text format - Title
                                cp_text1 = CharacterProperties(
                                    latin=normal_font, sz=1000, b=False
                                )
                                pp1 = ParagraphProperties(defRPr=cp_text1)
                                self.chart.x_axis.title.tx.rich.p[0].pPr = pp1
                                self.chart.y_axis.title.tx.rich.p[0].pPr = pp1
                                self.chart.x_axis.txPr.properties.rot = "-2700000"
                                self.chart.y_axis.majorUnit = self.axScale.get()
                                minValue = self.sheet[
                                    "A" + str(self.sheet.min_row + 1)
                                ].value
                                maxValue = self.sheet[
                                    "A" + str(self.sheet.max_row)
                                ].value
                                self.chart.x_axis.scaling.min = minValue
                                self.chart.x_axis.scaling.max = maxValue

                                colTime = 3
                                colMeas = 2
                                while colTime <= self.sheet.max_column:
                                    data = Reference(
                                        self.sheet,
                                        min_col=colTime,
                                        min_row=2,
                                        max_row=self.sheet.max_row,
                                    )
                                    colTime += 2
                                    if colMeas <= self.sheet.max_column:
                                        values = Reference(
                                            self.sheet,
                                            min_col=colMeas,
                                            min_row=1,
                                            max_row=self.sheet.max_row,
                                        )
                                        series = Series(
                                            values, data, title_from_data=True
                                        )
                                        self.chart.series.append(series)
                                    colMeas += 2
                                self.sheetChart.add_chart(self.chart)
                                for serie in self.chart.series:
                                    serie.graphicalProperties.line.width = 15000
                            else:
                                if self.option["secasi"] == True:
                                    self.chart.x_axis.tickLblPos = "low"
                                    self.chart.legend.position = "b"

                                    self.chart.x_axis.title = "Time [s]"
                                    self.chart.y_axis.title = "Temperature[°C]"

                                    # Text format - Axis
                                    normal_font = Font(typeface="Calibri")
                                    cp_text = CharacterProperties(
                                        latin=normal_font, sz=900, b=False
                                    )
                                    self.chart.x_axis.txPr = RichText(
                                        p=[
                                            Paragraph(
                                                pPr=ParagraphProperties(defRPr=cp_text),
                                                endParaRPr=cp_text,
                                            )
                                        ]
                                    )
                                    self.chart.y_axis.txPr = RichText(
                                        p=[
                                            Paragraph(
                                                pPr=ParagraphProperties(defRPr=cp_text),
                                                endParaRPr=cp_text,
                                            )
                                        ]
                                    )

                                    # Text format - Title
                                    cp_text1 = CharacterProperties(
                                        latin=normal_font, sz=1000, b=False
                                    )
                                    pp1 = ParagraphProperties(defRPr=cp_text1)
                                    self.chart.x_axis.title.tx.rich.p[0].pPr = pp1
                                    self.chart.y_axis.title.tx.rich.p[0].pPr = pp1
                                    self.chart.x_axis.txPr.properties.rot = "-2700000"
                                    self.chart.y_axis.majorUnit = self.axScale.get()
                                    minValue = self.sheet[
                                        "A" + str(self.sheet.min_row + 1)
                                    ].value
                                    maxValue = self.sheet[
                                        "A" + str(self.sheet.max_row)
                                    ].value
                                    self.chart.x_axis.scaling.min = minValue
                                    self.chart.x_axis.scaling.max = maxValue
                                else:
                                    self.chart.x_axis.number_format = (
                                        "yyyy-mm-dd HH:MM:SS"
                                    )
                                    self.chart.x_axis.majorTimeUnit = "months"
                                    self.chart.x_axis.tickLblPos = "low"
                                    self.chart.legend.position = "b"
                                    self.chart.x_axis.title = (
                                        "Date and Time [yyyy-mm-dd hh:mm:ss]"
                                    )
                                    self.chart.y_axis.title = "Temperature[°C]"

                                    # Text format - Axis
                                    normal_font = Font(typeface="Calibri")
                                    cp_text = CharacterProperties(
                                        latin=normal_font, sz=900, b=False
                                    )
                                    self.chart.x_axis.txPr = RichText(
                                        p=[
                                            Paragraph(
                                                pPr=ParagraphProperties(defRPr=cp_text),
                                                endParaRPr=cp_text,
                                            )
                                        ]
                                    )
                                    self.chart.y_axis.txPr = RichText(
                                        p=[
                                            Paragraph(
                                                pPr=ParagraphProperties(defRPr=cp_text),
                                                endParaRPr=cp_text,
                                            )
                                        ]
                                    )

                                    # Text format - Title
                                    cp_text1 = CharacterProperties(
                                        latin=normal_font, sz=1000, b=False
                                    )
                                    pp1 = ParagraphProperties(defRPr=cp_text1)
                                    self.chart.x_axis.title.tx.rich.p[0].pPr = pp1
                                    self.chart.y_axis.title.tx.rich.p[0].pPr = pp1
                                    self.chart.x_axis.txPr.properties.rot = "-2700000"
                                    self.chart.y_axis.majorUnit = self.axScale.get()
                                    minValue = self.sheet[
                                        "A" + str(self.sheet.min_row + 1)
                                    ].value
                                    maxValue = self.sheet[
                                        "A" + str(self.sheet.max_row)
                                    ].value
                                    minScale = pd.Timestamp(minValue)
                                    maxScale = pd.Timestamp(maxValue)
                                    minConverted = (
                                        minScale - pd.Timestamp("1899-12-30")
                                    ).total_seconds() / 86400
                                    maxConverted = (
                                        maxScale - pd.Timestamp("1899-12-30")
                                    ).total_seconds() / 86400
                                    self.chart.x_axis.scaling.min = minConverted
                                    self.chart.x_axis.scaling.max = maxConverted
                                if self.option["keithleymanager"] == True:
                                    colTime = 1
                                    colMeas = 2
                                    while colTime <= self.sheet.max_column:
                                        data = Reference(
                                            self.sheet,
                                            min_col=colTime,
                                            min_row=2,
                                            max_row=self.sheet.max_row,
                                        )
                                        colTime += 2
                                        if colMeas <= self.sheet.max_column:
                                            values = Reference(
                                                self.sheet,
                                                min_col=colMeas,
                                                min_row=1,
                                                max_row=self.sheet.max_row,
                                            )
                                            series = Series(
                                                values, data, title_from_data=True
                                            )
                                            self.chart.series.append(series)
                                        colMeas += 2
                                    self.sheetChart.add_chart(self.chart)
                                    for serie in self.chart.series:
                                        serie.graphicalProperties.line.width = 15000
                                else:
                                    if self.option["indexZUT"] == True:
                                        index = int(6)
                                    elif self.option["coolingSys"] == True:
                                        index = int(5)

                                    else:
                                        index = int(self.sheet.max_column + 1)
                                    data = Reference(
                                        self.sheet,
                                        min_col=1,
                                        min_row=2,
                                        max_row=self.sheet.max_row,
                                    )
                                    for i in range(2, index):
                                        values = Reference(
                                            self.sheet,
                                            min_col=i,
                                            min_row=1,
                                            max_row=self.sheet.max_row,
                                        )
                                        series = Series(
                                            values, data, title_from_data=True
                                        )
                                        self.chart.series.append(series)
                                    self.sheetChart.add_chart(self.chart)
                                    for serie in self.chart.series:
                                        serie.graphicalProperties.line.width = 15000
                                if (
                                    self.option["agilent"] == True
                                    or self.option["keithleymanager"] == True
                                ):
                                    pass
                                else:
                                    if self.option["insight"] == True:
                                        # Temperature
                                        s1 = self.chart.series[0]
                                        s1.graphicalProperties.line.solidFill = "F80E0A"
                                    else:
                                        # Temperature Hot Set
                                        s2 = self.chart.series[0]
                                        s2.graphicalProperties.line.solidFill = "752524"

                                        # Temperature Hot
                                        s3 = self.chart.series[1]
                                        s3.graphicalProperties.line.solidFill = "F80E0A"

                                        if self.option["humBath"] == True:
                                            # Humidifier Set
                                            s4 = self.chart.series[2]
                                            s4.graphicalProperties.line.solidFill = (
                                                "0A6E44"
                                            )

                                            # Humidifier
                                            s5 = self.chart.series[3]
                                            s5.graphicalProperties.line.solidFill = (
                                                "2AA31E"
                                            )
                                        else:
                                            try:
                                                # Temperature Cold Set
                                                s6 = self.chart.series[2]
                                                s6.graphicalProperties.line.solidFill = (
                                                    "173B96"
                                                )
                                            except:
                                                pass
                                            try:
                                                # Temperature Cold
                                                s7 = self.chart.series[3]
                                                s7.graphicalProperties.line.solidFill = (
                                                    "1F95F8"
                                                )
                                            except:
                                                pass
                                            try:
                                                # Temperature Basket
                                                s8 = self.chart.series[4]
                                                s8.graphicalProperties.line.solidFill = (
                                                    "2AA31E"
                                                )
                                            except:
                                                pass

                    self.wb.save(self.savepath)
                    self.popupWin2.destroy()
                    try:
                        os.startfile(self.savepath)
                    except:
                        Message(
                            "warning",
                            "File can not be opened by Stella"
                            "\nPlease open it manually",
                            0,
                        )
            except:
                try:
                    self.popupWin2.destroy()
                except:
                    pass
                Message(
                    "error",
                    "Something went wrong"
                    "\nPlease contact with: piotr.klys92@gmail.com",
                    0,
                )
