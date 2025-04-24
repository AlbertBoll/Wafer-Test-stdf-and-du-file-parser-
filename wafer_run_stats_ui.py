import multiprocessing
import os
from threading import Thread
from tkinter.filedialog import askopenfilename, asksaveasfile, askdirectory
import ttkbootstrap as ttk
import tkinter as tk

from PIL import ImageTk
from tkinterdnd2 import TkinterDnD
from ttkbootstrap.constants import *
import pathlib
from stdf import *
from abbott_wafer_run_du_parser import *
from drag_drop_listbox import FileListbox, FilesListBoxImporter, PathBrowser, FileDialog, RadioButton
import warnings
warnings.filterwarnings('ignore')

src_dict = {}





class GenerateStatisticReportUI(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=15)

        self.pack(fill=BOTH, expand=YES)

        self.wafer_run_parser = WaferRunCSVParser()
        # application variables
        _import_file_path = pathlib.Path().absolute().as_posix()
        _output_path = pathlib.Path().absolute().as_posix()

        self.import_file_path_var = ttk.StringVar(value=_import_file_path)
        self.output_path_var = ttk.StringVar(value=_output_path)

        self.report_selector_var = ttk.StringVar(value='passed runs')
        # self.all_runs_var = ttk.StringVar(value='all runs')
        # self.passed_runs_var = ttk.StringVar(value='passed runs')

        # header and labelframe option container
        option_text = "select du(csv) files to generate statistic report"
        # self.option_lf = ttk.Labelframe(self, text=option_text, padding=15, width=50, height=500)
        self.option_lf = ttk.Labelframe(self, text=option_text, padding=15)
        self.option_lf.pack(fill=BOTH, expand=YES, anchor=N)
        self.option_lf.rowconfigure(0, weight=1)
        self.option_lf.rowconfigure(1, weight=2)
        self.option_lf.rowconfigure(2, weight=1)
        self.option_lf.rowconfigure(3, weight=1)
        # self.option_lf.rowconfigure(2, weight=2)
        self.option_lf.columnconfigure(0, weight=15)
        self.option_lf.columnconfigure(1, weight=1)
        self.option_lf.columnconfigure(2, weight=15)
        self.list_dir = []
        PathBrowser(self.option_lf, 0, 0, 2, self.on_browse, self.import_file_path_var,
                    "import path", "wn", 25)
        self.file_list_left = FilesListBoxImporter(self.option_lf, 50, 18, 1, 0, 1, "wn", pad_y=(10, 1))
        self.file_list_right = FilesListBoxImporter(self.option_lf, 50, 18, 1, 2, 1, "en", pad_y=(10, 1))

        right_image = Image.open("right arrow.png").resize((30, 30)) # Replace with your image path
        self.right_photo = ImageTk.PhotoImage(right_image)

        left_image = Image.open("left arrow.png").resize((30, 30))  # Replace with your image path
        self.left_photo = ImageTk.PhotoImage(left_image)

        self.btn1 = ttk.Button(
            master=self.option_lf,
            image=self.right_photo,
            command=self.arrow_right_button_selected_items,
            bootstyle=SUCCESS,
            width=5,
            style='secondary.TButton'
        )
        # self.btn1.configure(background="grey")
        self.btn1.grid(row=1, column=1, padx=75, pady=135, ipadx=1, ipady=25, sticky="new")

        self.btn2 = ttk.Button(
            master=self.option_lf,
            image=self.left_photo,
            command=self.arrow_left_button_selected_items,
            bootstyle=SUCCESS,
            width=5,
            style='secondary.TButton'
        )
        self.btn2.grid(row=1, column=1, padx=75, pady=245, ipadx=1, ipady=25, sticky="new")
        # self.btn.grid(row=0, column=1, pady=200, ipady=20, sticky="n")
        # self.btn.
        # self.create_path_row(self.import_file_path_var, "import file")
        # # self.create_path_row(self.dut_summary_path_var, "dut summary")
        # self.create_path_row(self.output_path_var, "save as", False)
        # self.create_selectors_row()
        # self.create_combo_box()
        # self.slider_l_lim = self.create_slider_row("expand LLim")
        # self.slider_u_lim = self.create_slider_row("expand ULim")

        # self.list_box = self.create_drag_and_drop_list_box()
        RadioButton(self.option_lf, 2, 0, 1, "wn", [("None", "all runs"),
                                                    ("Outliers", "passed runs")], self.report_selector_var)


        # configure generate button
        # noinspection PyArgumentList
        self.generate_btn = ttk.Button(
            master=self.option_lf,
            text="Generate",
            command=self.on_generate,
            bootstyle=SUCCESS,
            width=8
        )
        self.generate_btn.grid(row=3, column=0, sticky='wn')
        # self.generate_btn.pack(side=LEFT, padx=5, pady=10)

        # # configure progress bar
        # # noinspection PyArgumentList
        self.progressbar = ttk.Progressbar(
            master=self.option_lf,
            mode=DETERMINATE,
            bootstyle=(STRIPED, SUCCESS),
            maximum=100.00,
            length=1180,
            value=0
        )
        self.progressbar.grid(row=3, column=0, columnspan=3, ipady=10, sticky='ne')
        self.percentage_of_completion = ttk.Label(self.option_lf, width=8, text=str(0) + "%")
        self.percentage_of_completion.grid(row=3, column=1, columnspan=1, sticky='ew', pady=40, padx=100)
        # # Readme File button
        # # noinspection PyArgumentList
        readme_btn = ttk.Button(
            master=self,
            text="Readme File",
            command=self.show_readme,
            bootstyle=INFO,
            width=11
        )
        readme_btn.pack(side=LEFT, padx=5, pady=10)
        #
        # # self.progressbar.pack(side=LEFT)
        # self.progressbar.pack(side=LEFT, fill=X, expand=YES)
        #
        # # start with 0 percent progress
        # self.percentage_of_completion = ttk.Label(self.option_lf, width=5, text=str(0) + "%")
        # self.percentage_of_completion.pack(fill=BOTH, expand=TRUE,  pady=0)

    def check_files_du_extension(self, filename_list):
        if len(filename_list) == 0:
            messagebox.showinfo("warn", "please select du or csv file")
            return False
        for file_name in filename_list:
            is_du_extension = (os.path.splitext(os.path.basename(file_name))[1] == ".du" or
                               os.path.splitext(os.path.basename(file_name))[1] == ".csv")
            if not is_du_extension:
                messagebox.showinfo("warn", "please select du or csv file")
                return False

        return True

    def on_generate(self):
        generate_list = [element for element in self.file_list_right.list_box.get(0, END)]

        if self.check_files_du_extension(generate_list):
            self.generate_btn.state(["disabled"])  # first inactive the generate button, prevent user to multiple click
            # button when the
            export_path = self.output_path_var.get()
            print(export_path)

            select = self.report_selector_var.get()

            # using enum instead
            select_filter = FilterType.ALL_RUN
            if select == "all runs":
                select_filter = FilterType.ALL_RUN
            elif select == "passed runs":
                select_filter = FilterType.PASSED_RUN

            Thread(
                target=self.wafer_run_parser.convert_to_stats_report_pdfs,
                args=(generate_list, select_filter, self.progressbar, self.generate_btn,
                      self.percentage_of_completion, src_dict),
                daemon=True
            ).start()
        # else:
        #     Thread(
        #         target=generate_stats_report,
        #         args=(file_name, select_filter, export_path, self.progressbar, self.generate_btn,
        #               self.percentage_of_completion, src_dict),
        #         daemon=True
        #     ).start()

            # select_bool = False if select == "passed runs" else True
            # Thread(
            #     target=generate_stats_report_,
            #     args=(file_name, select_filter, export_path, self.progressbar, self.percentage_of_completion, src_dict),
            #     daemon=True
            # ).start()

    def create_drag_and_drop_list_box(self):
        type_row = ttk.Frame(self.option_lf)
        type_row.pack(fill=X, expand=YES)
        type_lbl = ttk.Label(type_row, text="drag and drop .du/.csv files", width=25)
        type_lbl.pack(side=LEFT, padx=(15, 0))

        listbox = FileListbox(type_row, selectmode='multiple')
        listbox.pack(fill=tk.BOTH, expand=True, pady=8)

        self.clear_button = ttk.Button(type_row, text="Clear All", command=self.clear_listbox)
        self.clear_button.pack(side=tk.RIGHT, padx=4)

        self.delete_button = ttk.Button(type_row, text="Delete Selected", command=self.delete_selected_items)
        self.delete_button.pack(side=tk.RIGHT)
        return listbox

    def arrow_right_button_selected_items(self):
        selected_indices = self.file_list_left.list_box.curselection()
        for index in selected_indices:
            left_item = self.file_list_left.list_box.get(index)
            self.file_list_right.list_box.insert(END, left_item)
        for index in reversed(selected_indices):
            self.file_list_left.list_box.delete(index)

    def arrow_left_button_selected_items(self):
        selected_indices = self.file_list_right.list_box.curselection()
        for index in reversed(selected_indices):
            right_item = self.file_list_right.list_box.get(index)
            self.file_list_left.list_box.insert(END, right_item)
        for index in reversed(selected_indices):
            self.file_list_right.list_box.delete(index)

    def delete_selected_items(self):
        selected_indices = self.list_box.curselection()
        for index in reversed(selected_indices):
            self.list_box.delete(index)

    def clear_listbox(self):
        self.list_box.delete(0, END)

    def create_selectors_row(self):
        """Add selector row to labelframe"""
        type_row = ttk.Frame(self.option_lf)
        type_row.pack(fill=X, expand=YES)
        type_lbl = ttk.Label(type_row, text="filter", width=15)
        type_lbl.pack(side=LEFT, padx=(15, 0))

        all_runs_opt = ttk.Radiobutton(
            master=type_row,
            text="None",
            variable=self.report_selector_var,
            value="all runs",
            command=self.invoke_slider_combo
        )
        all_runs_opt.pack(side=LEFT)

        passed_run_opt = ttk.Radiobutton(
            master=type_row,
            text="Outliers",
            variable=self.report_selector_var,
            value="passed runs",
            command=self.invoke_slider_combo
        )
        passed_run_opt.pack(side=LEFT, padx=10)

        all_runs_opt.invoke()
        # self.invoke_slider()

    def on_browse(self, path_var_, file_dialog):
        """Callback for directory browse"""
        file_path = None
        try:
            if file_dialog == FileDialog.ASK_OPEN_FILE:
                file_path = askopenfilename(filetypes=(
                                     ("du files", "*.du"),
                                     ("xlsx files", "*.xlsx"),
                                     ("csv files", "*.csv"),), initialdir=os.getcwd())
                #file_path = asksaveasfile(mode="w")
            elif file_dialog == FileDialog.ASK_SAVE_FILE:
                file_path = asksaveasfile(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
                                          initialdir=os.getcwd()).name
            else:
                file_path = askdirectory(title="Browse directory")
                if file_path:
                    files = os.listdir(file_path)
                    self.file_list_left.list_box.delete(0, END)
                    for file in files:
                        if (os.path.isfile(file) and
                                (os.path.splitext(file)[1] == ".csv" or os.path.splitext(file)[1] == ".du")):
                            self.file_list_left.list_box.insert(END, file)

        except:
            pass
        if file_path:
            path_var_.set(file_path)

    def create_path_row(self, path_var_, info, is_file_path=True):
        """Add path row to labelframe"""
        path_row = ttk.Frame(self.option_lf)
        path_row.pack(fill=X, expand=YES)
        path_lbl = ttk.Label(path_row, text=info, width=14)
        path_lbl.pack(side=LEFT, padx=(15, 0))
        path_ent = ttk.Entry(path_row, textvariable=path_var_)
        path_ent.pack(side=LEFT, fill=X, expand=YES, padx=5)
        browse_btn = ttk.Button(
            master=path_row,
            text="Browse",
            command=lambda: self.on_browse(path_var_, is_file_path),
            width=8
        )
        browse_btn.pack(side=LEFT, padx=5, pady=2)

    def create_path_row_(self, row, column, column_span, path_var_, info, is_file_path=True):
        """Add path row to labelframe"""
        path_row = ttk.Frame(self.option_lf)
        # path_row.grid(column=column, columnspan=column_span, row=row, sticky="nsew")
        path_row.pack(fill=X, expand=YES)
        path_lbl = ttk.Label(path_row, text=info, width=14)
        path_lbl.pack(side=LEFT, padx=(15, 0))
        path_ent = ttk.Entry(path_row, textvariable=path_var_)
        path_ent.pack(side=LEFT, fill=X, expand=YES, padx=5)
        browse_btn = ttk.Button(
            master=path_row,
            text="Browse",
            command=lambda: self.on_browse(path_var_, is_file_path),
            width=8
        )
        browse_btn.pack(side=LEFT, padx=5, pady=2)

    def create_combo_box(self):
        type_row = ttk.Frame(self.option_lf)
        type_row.pack(fill=X, expand=YES)
        type_lbl = ttk.Label(type_row, text="select test", width=15)
        type_lbl.pack(side=LEFT, padx=(15, 0), pady=10)
        self.combo = ttk.Combobox(type_row, state="readonly", values=[""], width=40)
        self.combo.pack(side=LEFT)

        self.upload_btn = ttk.Button(
            master=type_row,
            text="load",
            width=8,
            #command=lambda: self.on_browse(path_var_, is_file_path),
            command=self.update_combo_list
        )
        self.upload_btn.pack(side=LEFT, padx=5, pady=2)

    # def on_upload_button_click(self):

    def update_combo_list(self):
        if not self.wafer_run_parser.b_loaded:
            file_name = self.import_file_path_var.get()
            is_du_extension = (os.path.splitext(os.path.basename(file_name))[1] == ".du" or
                               os.path.splitext(os.path.basename(file_name))[1] == ".csv")
            if not os.path.isfile(file_name):
                messagebox.showinfo("info", "Please select a import file")

            elif (os.path.splitext(os.path.basename(file_name))[1] != ".du" and
                  os.path.splitext(os.path.basename(file_name))[1] != ".csv" and
                  os.path.splitext(os.path.basename(file_name))[1] != ".xlsx"):
                messagebox.showinfo("warn", "please select du, csv or xlsx file")
            else:
                if is_du_extension:
                    self.wafer_run_parser.parse(file_name).transform_data()
                lis = [f'{test_id}: {test_description}' for test_id, test_description in
                       zip(self.wafer_run_parser.test_id_list, self.wafer_run_parser.test_description_list)]
                self.combo['values'] = lis
            self.wafer_run_parser.b_loaded = True
        # self.wafer_run_parser.parse(file_name)
        # if is_du_extension:
        #     Thread(
        #         target=self.invoke_update,
        #         daemon=True
        #     ).start()


    def invoke_update(self):
        # file_name = self.import_file_path_var.get()
        self.wafer_run_parser.transform_data()
        lis = [f'{test_id}: {test_description}' for test_id, test_description in
               zip(self.wafer_run_parser.test_id_list, self.wafer_run_parser.test_description_list)]
        self.combo['values'] = lis

    def create_slider_row(self, txt):
        type_row = ttk.Frame(self.option_lf)
        type_row.pack(fill=X, expand=YES)
        type_lbl = ttk.Label(type_row, text=txt, width=15)
        type_lbl.pack(side=LEFT, padx=(15, 0), pady=8)

        # Create a slider (Scale widget in Tkinter)
        slider = ttk.Scale(type_row, from_=-6.0, to=6.0, orient=ttk.HORIZONTAL, length=400,
                           command=lambda val: label.config(text=f'{slider.get():.1f}\u03C3'))

        slider.pack(side=LEFT)
        label = ttk.Label(type_row, text="0\u03C3", font=("Helvetica", 10))
        label.pack()
        slider.set(0)
        return slider
        # Disable the slider

    def invoke_slider_combo(self):
        select = self.report_selector_var.get()
        if hasattr(self, "slider_l_lim") and hasattr(self, "slider_u_lim") and hasattr(self, "combo"):
            if select != "all runs":
                self.slider_u_lim.config(state="disabled")
                self.slider_l_lim.config(state="disabled")
                self.upload_btn.state(["disabled"])
                self.combo.config(state="disabled")
            else:
                self.slider_l_lim.config(state="enabled")
                self.slider_u_lim.config(state="enabled")
                self.upload_btn.state(["!disabled"])
                self.combo.config(state="readonly")

    def show_readme(self):
        try:
            with open("../wafer run stats report tool/Readme - wafer run stats report Guide.txt", "r") as file:
                readme_content = file.read()
        except FileNotFoundError:
            readme_content = "README.md not found."

        popup = ttk.Toplevel()
        popup.title("README")

        text_area = ttk.Text(popup, wrap="word")
        text_area.insert("1.0", readme_content)
        text_area.config(state="disabled")  # Make it read-only
        text_area.pack(expand=True, fill="both")

        close_button = ttk.Button(popup, text="Close", command=popup.destroy)
        close_button.pack(pady=5)


class Window(ttk.Window, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)


if __name__ == '__main__':
    multiprocessing.freeze_support()
    app = Window("Generate DXV Statistic Report", "journal")
    app.resizable(False, False)
    # app = TkinterDnD.Tk()
    app.title("Generate DXV Statistic Report")
    app.iconbitmap("Abbott.ico")
    app.minsize(1100, 720)
    GenerateStatisticReportUI(app)
    app.mainloop()

    # remove all src dir list
    if "src_path" in src_dict:
        remove_list = src_dict["src_path"]
        for file_path in src_dict["src_path"]:
            os.remove(file_path)
