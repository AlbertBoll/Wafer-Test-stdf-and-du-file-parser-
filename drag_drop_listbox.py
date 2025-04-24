import tkinter as tk
from tkinterdnd2 import Tk, DND_FILES, TkinterDnD
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from enum import Enum


class FileDialog(Enum):
    ASK_DIRECTORY = 1,
    ASK_OPEN_FILE = 2,
    ASK_SAVE_FILE = 3

class FileListbox(tk.Listbox):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self.on_drop)
        self.dnd_bind('<<DragEnter>>', self.on_drag_enter)
        self.dnd_bind('<<DragLeave>>', self.on_drag_leave)

    def on_drop(self, event):
        files = self.tk.splitlist(event.data)
        for file in files:
          if file not in self.get(0, tk.END):
            self.insert(tk.END, file)

    def on_drag_enter(self, event):
        self.config(relief="sunken")

    def on_drag_leave(self, event):
        self.config(relief="flat")


class FilesListBoxImporter(ttk.Frame):
    def __init__(self, parent, width, height, row, column, column_span, sticky, pad_x=0, pad_y=0):
        super().__init__(master=parent)
        self.grid(column=column, columnspan=column_span, row=row, sticky=sticky)
        self.list_box = FileListbox(self, selectmode='multiple', width=width, height=height)
        self.list_box.configure(background="sky blue", foreground="red")
        self.list_box.pack(expand=True, padx=pad_x, pady=pad_y, side=LEFT)


class RadioButton(ttk.Frame):
    def __init__(self, parent, row, column, column_span, sticky, radio_list, selector_var):
        super().__init__(master=parent)
        self.grid(column=column, columnspan=column_span, row=row, sticky=sticky)
        type_lbl = ttk.Label(self, text="Filter", width=15)
        type_lbl.pack(side=LEFT, padx=(15, 0), pady=(50, 30))

        text_0, value_0 = radio_list[0]
        self.all_runs_opt = ttk.Radiobutton(
            master=self,
            text=text_0,
            variable=selector_var,
            value=value_0
        )
        self.all_runs_opt.pack(side=LEFT, pady=(50, 30))

        text_1, value_1 = radio_list[1]
        self.passed_run_opt = ttk.Radiobutton(
            master=self,
            text=text_1,
            variable=selector_var,
            value=value_1
        )
        self.passed_run_opt.pack(side=LEFT, padx=40, pady=(50, 30))

        self.all_runs_opt.invoke()


class PathBrowser(ttk.Frame):
    def __init__(self, parent, row, column, column_span, on_browse, path_var, info, sticky, width=20,
                 file_dialog=FileDialog.ASK_DIRECTORY):
        super().__init__(master=parent)
        self.grid(column=column, columnspan=column_span, row=row, sticky=sticky)
        path_lbl = ttk.Label(self, text=info, width=12)
        path_lbl.pack(side=LEFT, padx=(1, 0))
        path_ent = ttk.Entry(self, textvariable=path_var, width=width)
        path_ent.pack(side=LEFT, fill=X, expand=YES, padx=1)
        browse_btn = ttk.Button(
            master=self,
            text="Load",
            command=lambda: on_browse(path_var, file_dialog),
            width=8
        )
        browse_btn.pack(side=LEFT, padx=5, pady=2)
