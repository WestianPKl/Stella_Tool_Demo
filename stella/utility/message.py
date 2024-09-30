from tkinter import messagebox


class Message:
    def __init__(self, type, msg, parent):
        if parent != 0:
            if type == "info":
                messagebox.showinfo(title="Info", message=msg, parent=parent)
            elif type == "warning":
                messagebox.showwarning(title="Warning", message=msg, parent=parent)
            elif type == "error":
                messagebox.showerror(title="Error", message=msg, parent=parent)
        elif parent == 0:
            if type == "info":
                messagebox.showinfo(title="Info", message=msg)
            elif type == "warning":
                messagebox.showwarning(title="Warning", message=msg)
            elif type == "error":
                messagebox.showerror(title="Error", message=msg)
