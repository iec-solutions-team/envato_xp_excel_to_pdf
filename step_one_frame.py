import customtkinter as ctk
from tkinter import ttk
from tkinter import *
from tkinter import filedialog
from CTkMessagebox import CTkMessagebox
import pandas as pd
import os
import xlrd
import pyperclip
from config import (
    FRAME_WIDTH,
    FRAME_HEIGHT,
    MAIN_TITLE_FONT,
    PADDING_X_FROM_SIDEBAR,
    EVEN_TREEVIEW_BG,
    ODD_TREEVIEW_BG,
    EMPTY_VALUES_LIST,
)


class StepOneFrame(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, width=FRAME_WIDTH, height=FRAME_HEIGHT, corner_radius=0, **kwargs)
        self.place(x=215, y=0)

        self.controller = master
        self.dialog_find_excel_file = None
        self.excel_path = None

        # print(f"\nself.controller.controller_excel_dataframe : {self.controller.controller_excel_dataframe}\n")
        # print(f"\nself.controller.controller_select_columns_holder : {self.controller.controller_select_columns_holder}\n")

        self.column_entry_var = ctk.StringVar(None)
        self.alternate_column_entry_var = ctk.StringVar(None)
        self.prefix_entry_var = ctk.StringVar(None)
        self.suffix_entry_var = ctk.StringVar(None)

        # Execute 'create_widgets' to load Gui elements
        self.create_widgets()

        self.populate_excel_columns_table_with_data()
        self.initial_populate_select_columns_table_with_data()


    def dialog_window(self):
        root = ctk.CTkToplevel()
        root.geometry("600x300")

        excel_path_entry_var = ctk.StringVar(None)
        excel_path = None
        excel_columns_list = []
        first_sheet_checkbox_var = ctk.StringVar(None)
        sheet_name_entry_var =  ctk.StringVar(None)


        def on_click_first_sheet_checkbox():
            if int(first_sheet_checkbox_var.get()) == 0:
                sheet_name_entry.configure(state="normal")
            else:
                sheet_name_entry.configure(state="disabled")
                sheet_name_entry_var.set("")


        def on_click_choose_excel_file():
            """
            Opens a file dialog to select multiple Excel files and updates the Excel path entry with the chosen Excel paths.

            :return: None
            """

            selected_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

            # Check if any files were selected
            if selected_file:
                self.excel_path = selected_file

                print(f"\n selected_file : {selected_file}\n")

                # Extract just the filenames
                excel_filename = os.path.basename(selected_file)

                print(f"\n excel_filename : {excel_filename}\n")

                # Set the variable to a comma-separated list of filenames
                excel_path_entry_var.set(excel_filename)

        def on_click_read_excel_data():
            print(f"\non_click_read_excel_data\n")
            df = pd.DataFrame()
            is_success = False
            error_msg = ""
            file_path = self.excel_path
            sheet_name = sheet_name_entry_var.get()
            is_first_sheet = True if int(first_sheet_checkbox_var.get()) == 1 else False

            print(f"sheet_name : {sheet_name}")
            print(f"is_first_sheet : {is_first_sheet}")

            if not is_first_sheet and sheet_name in EMPTY_VALUES_LIST:
                CTkMessagebox(
                    title="Error",
                    message="Please provide sheet name to read.",
                    icon="warning",
                    option_1="Ok",
                )
                return

            if is_first_sheet:
                # Create an ExcelFile object
                xls = pd.ExcelFile(file_path)

                # Get the sheet names as a list
                sheet_names = xls.sheet_names
                sheet_name = sheet_names[0]
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                is_success = True
            else:
                try:
                # Try reading the specified sheet into a DataFrame
                    df = pd.read_excel(file_path, sheet_name=sheet_name)
                    is_success = True
                except FileNotFoundError:
                    error_msg = f"The Excel file '{file_path}' does not exist."
                except (xlrd.biffh.XLRDError, ValueError) as e:
                    error_msg = f"The sheet '{sheet_name}' does not exist in the Excel file."


            if not is_success:
                CTkMessagebox(
                    title="Error",
                    message=error_msg,
                    icon="warning",
                    option_1="Ok",
                )
                return

            self.controller.controller_excel_columns = df.columns.tolist()
            self.controller.controller_excel_dataframe = df
            self.populate_excel_columns_table_with_data()
            root.destroy()

        choose_excel_file_button = ctk.CTkButton(
            root,
            text="Choose Excel File",
            command=on_click_choose_excel_file,
            font=("Arial", 14),
            width=80,
        )
        choose_excel_file_button.place(x=40, y=50)


        excel_path_entry = ctk.CTkEntry(
            root,
            width=300,
            textvariable=excel_path_entry_var,
            font=("Arial", 14),
            state="readonly",
        )
        excel_path_entry.place(x=200, y=50)


        first_sheet_checkbox = ctk.CTkCheckBox(
            root,
            text="Use the first sheet found excel file",
            command=on_click_first_sheet_checkbox,
            variable=first_sheet_checkbox_var,
            onvalue=True,
            offvalue=False,
        )
        first_sheet_checkbox.place(x=40, y=120)
        first_sheet_checkbox.select()


        title_label = ctk.CTkLabel(
            root,
            text="Excel Sheet Name",
            font=("Arial", 14),
        )
        title_label.place(x=40, y=180)

        sheet_name_entry = ctk.CTkEntry(
            root,
            width=300,
            textvariable=sheet_name_entry_var,
            font=("Arial", 14),
        )
        sheet_name_entry.configure(state="disabled")
        sheet_name_entry.place(x=200, y=180)


        read_excel_data_button = ctk.CTkButton(
            root,
            text="Read Excel Data",
            command=on_click_read_excel_data,
            font=("Arial", 14),
            width=250
        )
        read_excel_data_button.place(x=180, y=250)



    def create_widgets(self):
        self.find_excel_button = ctk.CTkButton(
            self,
            text="Find Excel File",
            command=self.on_click_find_excel_button,
            font=("Arial", 14),
            width=220,
            height=40,
            )
        self.find_excel_button.place(x=PADDING_X_FROM_SIDEBAR, y=40)


        self.excel_columns_label = ctk.CTkLabel(
            self,
            text="Excel Columns",
            font=("Arial", 14),
        )
        self.excel_columns_label.place(x=PADDING_X_FROM_SIDEBAR, y=100)

        # Excel Column
        self.excel_columns_table = ttk.Treeview(self, selectmode="browse")
        self.excel_columns_table["columns"] = ("ExcelColumn")
        self.excel_columns_table.column("#0", width=0, stretch=NO, anchor=CENTER)
        self.excel_columns_table.column("ExcelColumn",  width=220, anchor=CENTER)
        # Configure columns header value of messages table
        self.excel_columns_table.heading("ExcelColumn", text="Excel Column", anchor=CENTER)
        # Configure demensions of the messages table
        self.excel_columns_table.place(x=PADDING_X_FROM_SIDEBAR, y=100, width=240, height=520)
        # Adds scrollbar for the messages table
        scrollbar = Scrollbar(self)
        scrollbar.place(x=260, y=100, width=10, height=520)
        self.excel_columns_table.configure(yscrollcommand=scrollbar.set)
        scrollbar.configure(command=self.excel_columns_table.yview)
        # Add alternating background colors to the items of messages table
        self.excel_columns_table.tag_configure("even", background=EVEN_TREEVIEW_BG)
        self.excel_columns_table.tag_configure("odd", background=ODD_TREEVIEW_BG)
        # Bind helper function to the messages table
        self.excel_columns_table.bind("<<TreeviewSelect>>", self.on_select_excel_columns_table_item)



        self.select_your_columns_label = ctk.CTkLabel(
            self,
            text="Select your Columns",
            font=("Arial", 18, "underline"),
        )
        self.select_your_columns_label.place(x=490, y=40)


        self.column_label = ctk.CTkLabel(
            self,
            text="Column",
            font=("Arial", 14),
        )
        self.column_label.place(x=410, y=85)
        self.column_entry = ctk.CTkEntry(
            self,
            width=250,
            textvariable=self.column_entry_var,
            font=("Arial", 14),
        )
        self.column_entry.place(x=310, y=110)


        self.alternate_column_label = ctk.CTkLabel(
            self,
            text="Alternate Column Name",
            font=("Arial", 14),
        )
        self.alternate_column_label.place(x=645, y=85)
        self.alternate_column_entry = ctk.CTkEntry(
            self,
            width=250,
            textvariable=self.alternate_column_entry_var,
            font=("Arial", 14),
        )
        self.alternate_column_entry.place(x=590, y=110)




        self.prefix_label = ctk.CTkLabel(
            self,
            text="Prefix Value",
            font=("Arial", 14),
        )
        self.prefix_label.place(x=390, y=145)
        self.prefix_entry = ctk.CTkEntry(
            self,
            width=250,
            textvariable=self.prefix_entry_var,
            font=("Arial", 14),
        )
        self.prefix_entry.place(x=310, y=170)


        self.suffix_label = ctk.CTkLabel(
            self,
            text="Suffix Value",
            font=("Arial", 14),
        )
        self.suffix_label.place(x=675, y=145)
        self.suffix_entry = ctk.CTkEntry(
            self,
            width=250,
            textvariable=self.suffix_entry_var,
            font=("Arial", 14),
        )
        self.suffix_entry.place(x=590, y=170)


        self.add_button = ctk.CTkButton(
            self,
            text="Add",
            command=self.on_click_add_button,
            font=("Arial", 14),
            width=250,
            )
        self.add_button.place(x=310, y=205)

        self.delete_button = ctk.CTkButton(
            self,
            text="Delete",
            command=self.on_click_delete_button,
            font=("Arial", 14),
            width=250,
            )
        self.delete_button.place(x=590, y=205)


        # Excel Column
        self.select_columns_table = ttk.Treeview(self, selectmode="browse")
        self.select_columns_table["columns"] = ("ExcelColumn", "Alternate", "Prefix", "Suffix")
        self.select_columns_table.column("#0", width=0, stretch=NO, anchor=CENTER)
        self.select_columns_table.column("ExcelColumn",  width=220, anchor=CENTER)
        self.select_columns_table.column("Alternate",  width=150, anchor=CENTER)
        self.select_columns_table.column("Prefix",  width=80, anchor=CENTER)
        self.select_columns_table.column("Suffix",  width=80, anchor=CENTER)
        # Configure columns header value of messages table
        self.select_columns_table.heading("ExcelColumn", text="Excel Column", anchor=CENTER)
        self.select_columns_table.heading("Alternate", text="Alternate", anchor=CENTER)
        self.select_columns_table.heading("Prefix", text="Prefix", anchor=CENTER)
        self.select_columns_table.heading("Suffix", text="Suffix", anchor=CENTER)
        # Configure demensions of the messages table
        self.select_columns_table.place(x=310, y=240, width=550, height=375)
        # Adds scrollbar for the messages table
        scrollbar = Scrollbar(self)
        scrollbar.place(x=860, y=240, width=10, height=375)
        self.select_columns_table.configure(yscrollcommand=scrollbar.set)
        scrollbar.configure(command=self.select_columns_table.yview)
        # Add alternating background colors to the items of messages table
        self.select_columns_table.tag_configure("even", background=EVEN_TREEVIEW_BG)
        self.select_columns_table.tag_configure("odd", background=ODD_TREEVIEW_BG)
        # Bind helper function to the messages table
        self.select_columns_table.bind("<<TreeviewSelect>>")


    def on_select_excel_columns_table_item(self, event):
        selection = self.excel_columns_table.selection()
        if selection:
            # Set the selected item
            selected_column = self.excel_columns_table.item(selection[0])["values"][0]
            pyperclip.copy(selected_column)


    def on_click_add_button(self):
        print(f"\non_click_add_button\n")
        column = self.column_entry_var.get()
        alternate_column = self.alternate_column_entry_var.get()
        prefix = self.prefix_entry_var.get()
        suffix = self.suffix_entry_var.get()

        if column in EMPTY_VALUES_LIST:
            CTkMessagebox(
                title="Error",
                message="Need provide column name on Column textbox",
                icon="warning",
                option_1="Ok",
            )
            return

        if column not in self.controller.controller_excel_columns:
            CTkMessagebox(
                title="Error",
                message="Invalid column provided on Column textbox",
                icon="warning",
                option_1="Ok",
            )
            return

        tags = "even" if len(self.select_columns_table.get_children()) % 2 == 0 else "odd"
        self.select_columns_table.insert("", "end", values=(column, alternate_column, prefix, suffix), tags=(tags,))

        self.store_controller_select_columns_holder()

        self.column_entry_var.set("")
        self.alternate_column_entry_var.set("")
        self.prefix_entry_var.set("")
        self.suffix_entry_var.set("")

    def on_click_delete_button(self):
        print(f"\non_click_delete_button\n")

        selected_item = self.select_columns_table.selection()
        if not selected_item:
            CTkMessagebox(
                title="Error",
                message="Please select column to delete",
                icon="warning",
                option_1="Ok",
            )
        else:
            selected_column = selected_item[0]
            self.select_columns_table.delete(selected_column)

        self.store_controller_select_columns_holder()


    def on_click_find_excel_button(self):
        print("\non_click_find_excel_button\n")
        self.dialog_window()


    def store_controller_select_columns_holder(self):
        self.controller.controller_select_columns_holder = []
        for item in self.select_columns_table.get_children():
            self.controller.controller_select_columns_holder.append(self.select_columns_table.item(item)['values'])


    def populate_excel_columns_table_with_data(self):
        self.excel_columns_table.delete(*self.excel_columns_table.get_children())  # Clear existing items in the Treeview

        # Loop through records fetched from database and insert into Treeview
        for index, column in enumerate(self.controller.controller_excel_columns):
            # Determine the tag for the current item based on even or odd index
            tag = "even" if index % 2 == 0 else "odd"

            # Insert the item into the Treeview with the appropriate tag
            self.excel_columns_table.insert("", "end", values=(column,), tags=(tag,))


    def initial_populate_select_columns_table_with_data(self):
        if self.controller.controller_select_columns_holder:
            for index, row_data in enumerate(self.controller.controller_select_columns_holder):
                tag = "even" if index % 2 == 0 else "odd"
                self.select_columns_table.insert('', 'end', values=row_data, tags=(tag,))
