import customtkinter as ctk
# from CTkTable import *
from tkinter import ttk
from tkinter import *
from config import (
    FRAME_WIDTH,
    FRAME_HEIGHT,
    MAIN_TITLE_FONT,
    PADDING_X_FROM_SIDEBAR,
    EVEN_TREEVIEW_BG,
    ODD_TREEVIEW_BG,
)
from dialog_find_excel import (
    DialogFindExcelFile,
)


class StepOneFrame(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, width=FRAME_WIDTH, height=FRAME_HEIGHT, corner_radius=0, **kwargs)
        self.place(x=215, y=0)

        self.dialog_find_excel_file = None

        # Execute 'create_widgets' to load Gui elements
        self.create_widgets()


    def create_widgets(self):
        # self.label = ctk.CTkLabel(self, text="Select your Excel File", font=MAIN_TITLE_FONT)
        # self.label.place(x=30, y=20)


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
        self.excel_columns_table.place(x=PADDING_X_FROM_SIDEBAR, y=100, width=240, height=470)
        # Adds scrollbar for the messages table
        scrollbar = Scrollbar(self)
        scrollbar.place(x=260, y=100, width=10, height=470)
        self.excel_columns_table.configure(yscrollcommand=scrollbar.set)
        scrollbar.configure(command=self.excel_columns_table.yview)
        # Add alternating background colors to the items of messages table
        self.excel_columns_table.tag_configure("even", background=EVEN_TREEVIEW_BG)
        self.excel_columns_table.tag_configure("odd", background=ODD_TREEVIEW_BG)
        # Bind helper function to the messages table
        self.excel_columns_table.bind("<<TreeviewSelect>>", self.select_item)



    def select_item(self):
        print("\nselect_item\n")


    def on_click_find_excel_button(self):
        print("\non_click_find_excel_button\n")
        if self.dialog_find_excel_file is None or not self.dialog_find_excel_file.winfo_exists():
            self.dialog_find_excel_file = DialogFindExcelFile(self)  # create window if its None or destroyed
        else:
            self.dialog_find_excel_file.focus()  # if window exists focus it

        print("\n\n\n99999999\n\n\n")