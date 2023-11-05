import customtkinter as ctk
from tkinter import filedialog
from tkinter import ttk
from tkinter import *
from PIL import Image
import pyperclip
from config import (
    FRAME_WIDTH,
    FRAME_HEIGHT,
    PADDING_X_FROM_SIDEBAR,
    TEMPLATES,
    NORMAL_FONT,
    SMALL_FONT,
    EVEN_TREEVIEW_BG,
    ODD_TREEVIEW_BG,
)

class StepTwoFrame(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, width=FRAME_WIDTH, height=FRAME_HEIGHT, corner_radius=0, **kwargs)
        self.place(x=215, y=0)

        self.controller = master

        # print(f"\nself.controller.controller_selected_template : {self.controller.controller_selected_template}\ntype : {type(self.controller.controller_selected_template)}\n")
        # print(f"\nself.controller.controller_template_main_title : {self.controller.controller_template_main_title}")
        # print(f"\nself.controller.controller_template_sub_title : {self.controller.controller_template_sub_title}")

        self.template_selection_combobox_var = ctk.StringVar(self, self.controller.controller_selected_template)
        self.template_main_title_entry_var = ctk.StringVar(self, self.controller.controller_template_main_title)
        self.template_sub_title_entry_var = ctk.StringVar(self, self.controller.controller_template_sub_title)
        self.output_folder_name_entry_var = ctk.StringVar(self, self.controller.controller_output_folder_name)
        self.create_zipfile_radio_var = ctk.IntVar(self, self.controller.controller_create_zipfile)
        self.output_path_entry_var = ctk.StringVar(self, self.controller.controller_output_path)
        self.output_filename_entry_var = ctk.StringVar(self, self.controller.controller_output_filename)
        self.output_filename_format_entry_var = ctk.StringVar(self, ".pdf")



        # Execute 'create_widgets' to load Gui elements
        self.create_widgets()


    def on_select_template_selection_combobox(self, choice):
        selected_option = choice
        image_preview_path = TEMPLATES[selected_option]["image"]
        print(f"\nselected_option : {selected_option}\n")

        self.template_image_preview.configure(
            light_image=Image.open(image_preview_path),
        )

        self.controller.controller_selected_template = selected_option
        self.controller.controller_selected_template_image_preview_path = image_preview_path


    def on_template_main_title_entry(self, event):
        main_title = self.template_main_title_entry_var.get()
        self.controller.controller_template_main_title = main_title


    def on_template_sub_title_entry(self, event):
        sub_title = self.template_sub_title_entry_var.get()
        self.controller.controller_template_sub_title = sub_title


    def on_output_folder_name_entry(self, event):
        output_folder_name = self.output_folder_name_entry_var.get()
        self.controller.controller_output_folder_name = output_folder_name


    def on_create_zipfile_radio(self):
        create_zipfile = self.create_zipfile_radio_var.get()
        self.controller.controller_create_zipfile = create_zipfile


    def choose_output_folder_path(self):
        output_path = filedialog.askdirectory()
        self.output_path_entry_var.set(output_path)
        self.controller.controller_output_path = output_path


    def on_output_filename_entry(self, event):
        output_filename = self.output_filename_entry_var.get()
        self.controller.controller_output_filename = output_filename


    def on_format_one_button(self):
        pyperclip.copy("{today}")


    def on_format_two_button(self):
        pyperclip.copy("{increment}")


    def on_format_three_button(self):
        pyperclip.copy("{timestamp}")


    def dialog_show_excel_columns(self):
        root = ctk.CTkToplevel()

        # Get the screen width and height
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # Calculate the center of the screen
        window_size = (300,600)
        center_x = (screen_width - window_size[0]) // 2
        center_y = (screen_height - window_size[1]) // 2

        # Window Configuration
        root.title("Excel Columns")
        root.geometry(f"{window_size[0]}x{window_size[1]}+{center_x}+{center_y}")
        root.resizable(False, False)

        root.excel_columns_note_label = ctk.CTkLabel(
            root,
            text="(You can click on the column you want and",
            font=SMALL_FONT,
        )
        root.excel_columns_note_label.place(x=30, y=10)
        root.second_excel_columns_note_label = ctk.CTkLabel(
            root,
            text="it will get copy to your clipboard)",
            font=SMALL_FONT,
        )
        root.second_excel_columns_note_label.place(x=30, y=30)


        def on_select_excel_show_columns_table_item(event):
            selection = root.excel_show_columns_table.selection()
            if selection:
                # Set the selected item
                selected_column = root.excel_show_columns_table.item(selection[0])["values"][0]
                selected_column = "{" + f"{selected_column}" + "}"
                pyperclip.copy(selected_column)


        def initial_populate_select_columns_table_with_data():
            print(f"\n\nself.controller.controller_excel_columns : {self.controller.controller_excel_columns}\n\n")
            if self.controller.controller_excel_columns:
                for index, row_data in enumerate(self.controller.controller_excel_columns):
                    print(f"<<<<<<<< index : {index} , row_data: {row_data} >>>>>>")
                    tag = "even" if index % 2 == 0 else "odd"
                    root.excel_show_columns_table.insert('', 'end', values=((row_data, )), tags=(tag,))

        # Excel Column
        root.excel_show_columns_table = ttk.Treeview(root, selectmode="browse")
        root.excel_show_columns_table["columns"] = ("ExcelColumn")
        root.excel_show_columns_table.column("#0", width=0, stretch=NO, anchor=CENTER)
        root.excel_show_columns_table.column("ExcelColumn",  width=220, anchor=CENTER)
        # Configure columns header value of messages table
        root.excel_show_columns_table.heading("ExcelColumn", text="Excel Column", anchor=CENTER)
        # Configure demensions of the messages table
        root.excel_show_columns_table.place(x=PADDING_X_FROM_SIDEBAR, y=70, width=240, height=520)
        # Adds scrollbar for the messages table
        scrollbar = Scrollbar(root)
        scrollbar.place(x=260, y=70, width=10, height=520)
        root.excel_show_columns_table.configure(yscrollcommand=scrollbar.set)
        scrollbar.configure(command=root.excel_show_columns_table.yview)
        # Add alternating background colors to the items of messages table
        root.excel_show_columns_table.tag_configure("even", background=EVEN_TREEVIEW_BG)
        root.excel_show_columns_table.tag_configure("odd", background=ODD_TREEVIEW_BG)
        # Bind helper function to the messages table
        root.excel_show_columns_table.bind("<<TreeviewSelect>>", on_select_excel_show_columns_table_item)


        initial_populate_select_columns_table_with_data()



    def on_format_four_button(self):
        print("\non_format_four_button\n")
        self.dialog_show_excel_columns()


    def create_widgets(self):
        # 1st section
        self.main_title_1 = ctk.CTkLabel(
            self,
            text="Select Template",
            font=("Arial", 18, "underline"),
        )
        self.main_title_1.place(x=140, y=40)


        self.template_selection_label = ctk.CTkLabel(
            self,
            text="Template*",
            font=NORMAL_FONT,
        )
        self.template_selection_label.place(x=PADDING_X_FROM_SIDEBAR, y=100)
        self.template_selection_combobox = ctk.CTkComboBox(
            self,
            values=list(TEMPLATES.keys()),
            variable=self.template_selection_combobox_var,
            width=250,
            font=NORMAL_FONT,
            state="readonly",
            command=self.on_select_template_selection_combobox,
        )
        self.template_selection_combobox.place(x=120, y=100)


        self.template_main_title_label = ctk.CTkLabel(
            self,
            text="Main Title",
            font=NORMAL_FONT,
        )
        self.template_main_title_label.place(x=PADDING_X_FROM_SIDEBAR, y=150)
        self.template_main_title_entry = ctk.CTkEntry(
            self,
            width=250,
            textvariable=self.template_main_title_entry_var,
            font=NORMAL_FONT,
        )
        self.template_main_title_entry.bind("<KeyRelease>", self.on_template_main_title_entry)
        self.template_main_title_entry.place(x=120, y=150)


        self.template_sub_title_label = ctk.CTkLabel(
            self,
            text="Sub Title",
            font=NORMAL_FONT,
        )
        self.template_sub_title_label.place(x=PADDING_X_FROM_SIDEBAR, y=200)
        self.template_sub_title_entry = ctk.CTkEntry(
            self,
            width=250,
            textvariable=self.template_sub_title_entry_var,
            font=NORMAL_FONT,
        )
        self.template_sub_title_entry.bind("<KeyRelease>", self.on_template_sub_title_entry)
        self.template_sub_title_entry.place(x=120, y=200)

        self.template_image_preview = ctk.CTkImage(
            light_image=Image.open(self.controller.controller_selected_template_image_preview_path),
            # dark_image=Image.open("./templates/simple_preview.png"),
            size=(320, 400),
        )
        self.template_image_preview_label = ctk.CTkLabel(
            self,
            image=self.template_image_preview,
            text="",
        )
        self.template_image_preview_label.place(x=50, y=240)




        # 2nd section
        self.main_title_2 = ctk.CTkLabel(
            self,
            text="Output Options",
            font=("Arial", 18, "underline"),
        )
        self.main_title_2.place(x=570, y=40)


        self.output_folder_name_label = ctk.CTkLabel(
            self,
            text="Output Folder Name*",
            font=NORMAL_FONT,
        )
        self.output_folder_name_label.place(x=430, y=100)
        self.output_folder_name_entry = ctk.CTkEntry(
            self,
            width=250,
            textvariable=self.output_folder_name_entry_var,
            font=NORMAL_FONT,
        )
        self.output_folder_name_entry.bind("<KeyRelease>", self.on_output_folder_name_entry)
        self.output_folder_name_entry.place(x=580, y=100)



        self.shall_create_zipfile_label = ctk.CTkLabel(
            self,
            text="Shall Create Zip file ?",
            font=NORMAL_FONT,
        )
        self.shall_create_zipfile_label.place(x=430, y=150)
        self.yes_create_zipfile_radio = ctk.CTkRadioButton(
            self,
            text="Yes",
            command=self.on_create_zipfile_radio,
            variable=self.create_zipfile_radio_var,
            value=1,
        )
        self.yes_create_zipfile_radio.place(x=430, y=190)
        self.no_create_zipfile_radio = ctk.CTkRadioButton(
            self,
            text="No",
            command=self.on_create_zipfile_radio,
            variable=self.create_zipfile_radio_var,
            value=0,
        )
        self.no_create_zipfile_radio.place(x=430, y=230)

        self.output_path_button = ctk.CTkButton(
            self,
            text="Output Path*",
            command=self.choose_output_folder_path,
            font=NORMAL_FONT,
        )
        self.output_path_button.place(x=430, y=280)
        self.output_path_entry = ctk.CTkEntry(
            self,
            width=250,
            textvariable=self.output_path_entry_var,
            font=NORMAL_FONT,
            state="readonly",
        )
        self.output_path_entry.place(x=580, y=280)


        self.output_filename_label = ctk.CTkLabel(
            self,
            text="Output File Name*",
            font=NORMAL_FONT,
        )
        self.output_filename_label.place(x=430, y=330)
        self.output_filename_entry = ctk.CTkEntry(
            self,
            width=360,
            textvariable=self.output_filename_entry_var,
            font=NORMAL_FONT,
        )
        self.output_filename_entry.bind("<KeyRelease>", self.on_output_filename_entry)
        self.output_filename_entry.place(x=430, y=360)
        self.output_filename_format_entry = ctk.CTkEntry(
            self,
            width=40,
            textvariable=self.output_filename_format_entry_var,
            font=NORMAL_FONT,
            state="readonly"
        )
        self.output_filename_format_entry.place(x=790, y=360)


        self.format_output_filename_label = ctk.CTkLabel(
            self,
            text="Info about available formats for output filename",
            font=("Arial", 14, "underline"),
        )
        self.format_output_filename_label.place(x=430, y=410)


        self.format_one_button = ctk.CTkButton(
            self,
            text="Copy",
            command=self.on_format_one_button,
            font=SMALL_FONT,
            width=50,
            height=20,
        )
        self.format_one_button.place(x=390, y=445)
        self.format_one_label = ctk.CTkLabel(
            self,
            text="- {today} : example as 2023-08-15",
            font=NORMAL_FONT,
        )
        self.format_one_label.place(x=450, y=440)


        self.format_two_button = ctk.CTkButton(
            self,
            text="Copy",
            command=self.on_format_two_button,
            font=SMALL_FONT,
            width=50,
            height=20,
        )
        self.format_two_button.place(x=390, y=475)
        self.format_two_label = ctk.CTkLabel(
            self,
            text="- {increment}: increment number starts from 1",
            font=NORMAL_FONT,
        )
        self.format_two_label.place(x=450, y=470)


        self.format_three_button = ctk.CTkButton(
            self,
            text="Copy",
            command=self.on_format_three_button,
            font=SMALL_FONT,
            width=50,
            height=20,
        )
        self.format_three_button.place(x=390, y=505)
        self.format_three_label = ctk.CTkLabel(
            self,
            text="- {timestamp}: full timezone string with date, hour, minute, second",
            font=NORMAL_FONT,
        )
        self.format_three_label.place(x=450, y=500)


        self.format_four_button = ctk.CTkButton(
            self,
            text="Show",
            command=self.on_format_four_button,
            font=SMALL_FONT,
            width=50,
            height=20,
        )
        self.format_four_button.place(x=390, y=535)
        self.format_four_label = ctk.CTkLabel(
            self,
            text="- {any excel column name to use its value}",
            font=NORMAL_FONT,
        )
        self.format_four_label.place(x=450, y=530)
