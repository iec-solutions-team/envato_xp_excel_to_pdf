from tkinter import filedialog
import customtkinter as ctk
from CTkMessagebox import CTkMessagebox
import pandas as pd
import os
import xlrd
from config import (
    FRAME_WIDTH,
    FRAME_HEIGHT,
    MAIN_TITLE_FONT,
    PADDING_X_FROM_SIDEBAR,
    EMPTY_VALUES_LIST,
)


class DialogFindExcelFile(ctk.CTkToplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.geometry("600x300")

        self.excel_path_entry_var = ctk.StringVar(None)
        self.excel_path = None
        # self.excel_columns_list = []

        self.first_sheet_checkbox_var = ctk.StringVar(None)
        self.sheet_name_entry_var =  ctk.StringVar(None)

        self.choose_excel_file_button = ctk.CTkButton(
            self,
            text="Choose Excel File",
            command=self.on_click_choose_excel_file,
            font=("Arial", 14),
            width=80,
        )
        self.choose_excel_file_button.place(x=40, y=50)


        self.excel_path_entry = ctk.CTkEntry(
            self,
            width=300,
            textvariable=self.excel_path_entry_var,
            font=("Arial", 14),
            state="readonly",
        )
        self.excel_path_entry.place(x=200, y=50)


        self.first_sheet_checkbox = ctk.CTkCheckBox(
            self,
            text="Use the first sheet found excel file",
            command=self.on_click_first_sheet_checkbox,
            variable=self.first_sheet_checkbox_var,
            onvalue=True,
            offvalue=False,
        )
        self.first_sheet_checkbox.place(x=40, y=120)
        self.first_sheet_checkbox.select()


        self.title_label = ctk.CTkLabel(self,
            text="Excel Sheet Name", font=("Arial", 14))
        self.title_label.place(x=40, y=180)

        self.sheet_name_entry = ctk.CTkEntry(
            self,
            width=300,
            textvariable=self.sheet_name_entry_var,
            font=("Arial", 14),
        )
        self.sheet_name_entry.configure(state="disabled")
        self.sheet_name_entry.place(x=200, y=180)


        self.read_excel_data_button = ctk.CTkButton(
            self,
            text="Read Excel Data",
            command=self.on_click_read_excel_data,
            font=("Arial", 14),
            width=250
        )
        self.read_excel_data_button.place(x=180, y=250)


    def on_click_read_excel_data(self):
        print(f"\non_click_read_excel_data\n")
        df = pd.DataFrame()
        is_success = False
        error_msg = ""
        file_path = self.excel_path
        sheet_name = self.sheet_name_entry_var.get()
        is_first_sheet = True if int(self.first_sheet_checkbox_var.get()) == 1 else False

        print(f"sheet_name : {sheet_name}")
        print(f"is_first_sheet : {is_first_sheet}")

        if not is_first_sheet and sheet_name in EMPTY_VALUES_LIST:
            CTkMessagebox(
                title="Error",
                message="Please provide sheet name to read.",
                icon="warning",
                option_1="Ok",
                option_focus=1,
                sound=True,
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
                option_focus=1,
                sound=True,
            )
            return

        self.parent_excel_columns_list = df.columns.tolist()
        print(f"\n\nself.parent_excel_columns_list : {self.parent_excel_columns_list}\n\n")


        # print(f"\nlen df : {len(df)}\n")
        # print(f"\n\ndf : {df}\n\n")

        #self.destroy()


    def on_click_first_sheet_checkbox(self):
        if int(self.first_sheet_checkbox_var.get()) == 0:
            self.sheet_name_entry.configure(state="normal")
        else:
            self.sheet_name_entry.configure(state="disabled")
            self.sheet_name_entry_var.set("")


    def on_click_choose_excel_file(self):
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
            self.excel_path_entry_var.set(excel_filename)