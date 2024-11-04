"""
Stella Tool for data processing of temperature and humidity measurements from different equipment.
Created on Mon Feb 06 2023 (full version)
Modified on Mon Sep 30 2024 (demo version)

@author: piotrklys
@mail: piotr.klys92@gmail.com
"""

import tkinter as tk
from stella.config.config import ConfigInformations
from stella.__main__ import App

if __name__ == "__main__":
    root = tk.Tk()
    img = tk.PhotoImage(data=ConfigInformations.icon_path)
    root.iconphoto(False, img)
    app = App(root)
    root.mainloop()
