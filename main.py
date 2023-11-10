import os
import threading
import pandas as pd
import customtkinter as ctk
from CTkMessagebox import CTkMessagebox
from config import (
    PAGE_WIDTH,
    PAGE_HEIGHT,
    SIDEBAR_WIDTH,
    BLUE_COLOR,
    WHITE_COLOR,
    RED_COLOR,
    HOVER_RED_COLOR,
    GREEN_COLOR,
    HOVER_GREEN_COLOR,
    MAIN_TITLE_FONT,
    APP_TITLE,
    TEMPLATES,
)
from step_one_frame import (
    StepOneFrame,
)
from step_two_frame import (
    StepTwoFrame,
)
from initial_page import (
    InitialPage,
)
from services import (
    process_pdf_filename_format,
    handle_folder_creation,
    create_zip,
    is_validate_filename_input,
)


# FOR PDF generation
from jinja2 import Environment, FileSystemLoader
import pdfkit
env = Environment(loader=FileSystemLoader('.'))
# Set the options for PDF generation
pdf_options = {
    'quiet': '',
    'page-size': 'A4',
    'margin-top': '0mm',
    'margin-right': '0mm',
    'margin-bottom': '0mm',
    'margin-left': '0mm',
    'enable-local-file-access': None
}


class Dashboard(ctk.CTk):

    def __init__(self):
        super().__init__()

        # data controller
        # used by step 1 page
        self.controller_excel_columns = []
        self.controller_excel_dataframe = None
        self.controller_select_columns_holder = None
        # used by step 2 page
        self.controller_selected_template = None
        self.controller_selected_template_image_preview_path = "./templates/placeholder.png"
        self.controller_template_main_title = ""
        self.controller_template_sub_title = ""
        self.controller_output_folder_name = "XP_PDF" # default folder name
        self.controller_create_zipfile = 0 # default as false
        self.controller_output_path = None
        self.controller_output_filename = None
        self.path_to_wkhtmltopdf = './wkhtmltopdf'

        # Get the screen width and height
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # Calculate the center of the screen
        window_size = (PAGE_WIDTH,PAGE_HEIGHT)
        center_x = (screen_width - window_size[0]) // 2
        center_y = (screen_height - window_size[1]) // 2

        # Window Configuration
        self.title(APP_TITLE)
        self.geometry(f"{window_size[0]}x{window_size[1]}+{center_x}+{center_y}")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Styles
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.sidebar_frame = ctk.CTkFrame(master=self, width=SIDEBAR_WIDTH, height=PAGE_HEIGHT, corner_radius=0)
        self.sidebar_frame.place(x=0, y=0)

        self.initial_page = InitialPage(self)
        self.step_one_frame = None
        self.step_two_frame = None

        # Execute 'create_widgets' to load Gui elements
        self.create_widgets()


    def create_widgets(self):
        self.sidebar_main_title = ctk.CTkLabel(self.sidebar_frame, text=APP_TITLE, font=MAIN_TITLE_FONT)
        self.sidebar_main_title.place(x=30, y=20)


        self.sidebar_select_excel_button = ctk.CTkButton(
            self.sidebar_frame,
            text="Step (1) - Select Excel File",
            command=self.switch_to_step_one_frame,
            corner_radius=0,
            width=SIDEBAR_WIDTH,
            height=47,
        )
        self.sidebar_select_excel_button.place(
            x=0,
            y=200,
        )


        self.sidebar_configure_pdf_button = ctk.CTkButton(
            self.sidebar_frame,
            text="Step (2) - Configure PDF",
            command=self.switch_to_step_two_frame,
            corner_radius=0,
            width=SIDEBAR_WIDTH,
            height=47,
        )
        self.sidebar_configure_pdf_button.place(
            x=0,
            y=300,
        )


        self.sidebar_loading_label = ctk.CTkLabel(
            self.sidebar_frame,
            text="Generating PDF Files...",
            font=MAIN_TITLE_FONT,
            text_color=BLUE_COLOR,
        )


        self.sidebar_start_button = ctk.CTkButton(
            self.sidebar_frame,
            text="Generate the PDF",
            command=self.on_start,
            corner_radius=0,
            fg_color=GREEN_COLOR,
            hover_color=HOVER_GREEN_COLOR,
            width=SIDEBAR_WIDTH,
            height=50,
        )
        self.sidebar_start_button.place(
            x=0,
            y=(PAGE_HEIGHT - 100),
        )


        self.sidebar_exit_button = ctk.CTkButton(
            self.sidebar_frame,
            text="Exit",
            command=self.on_closing,
            corner_radius=0,
            fg_color=RED_COLOR,
            hover_color=HOVER_RED_COLOR,
            width=SIDEBAR_WIDTH,
            height=50,
        )
        self.sidebar_exit_button.place(
            x=0,
            y=(PAGE_HEIGHT - 50),
        )

    def update_loading_text(self):
        if self.sidebar_loading_label.cget("text") == "Generating PDF Files...":
            self.sidebar_loading_label.configure(text="Generating PDF Files")
        else:
            self.sidebar_loading_label.configure(text=self.sidebar_loading_label.cget("text") + ".")

    def start_loading_animation(self):
        self.sidebar_loading_label.place(x=20, y=(PAGE_HEIGHT - 140))
        self.update_loading_text()

        update_id = self.after(500, self.start_loading_animation)
        # Store the update_id so we can cancel it later
        self.sidebar_loading_label.update_id = update_id

    def stop_loading_animation(self):
        self.sidebar_loading_label.place_forget()
        if hasattr(self.sidebar_loading_label, "update_id"):
            self.after_cancel(self.sidebar_loading_label.update_id)

    # This function remove any active page on dashbaord, used to switch between pages
    def destory_all_pages(self):
        if self.step_one_frame is not None:
            self.step_one_frame.destroy()

        if self.step_two_frame is not None:
            self.step_two_frame.destroy()

        if self.initial_page is not None:
            self.initial_page.destroy()


    def switch_to_step_one_frame(self):
        self.destory_all_pages()
        self.step_one_frame = StepOneFrame(self)


    def switch_to_step_two_frame(self):
        self.destory_all_pages()
        self.step_two_frame = StepTwoFrame(self)


    def validate_fields(self):
        has_error = False
        error_msg = ""

        # Step 1 validations
        if not isinstance(self.controller_excel_dataframe, pd.DataFrame):
            has_error = True
            error_msg = "Please find your excel file on (Step 1)"
        elif len(self.controller_excel_dataframe) == 0:
            has_error = True
            error_msg = "Your excel file seems to be empty, please find another one on (Step 1)"
        elif len(self.controller_excel_columns) == 0:
            has_error = True
            error_msg = "Your excel file seems to be empty, please find another one on (Step 1)"
        elif not self.controller_select_columns_holder:
            has_error = True
            error_msg = "Please select your excel columns on (Step 1)"
        # Step 2 validations
        elif not self.controller_selected_template:
            has_error = True
            error_msg = "Please select your template on (Step 2)"
        elif self.controller_output_folder_name in ("", " ", None):
            has_error = True
            error_msg = "Please type the output folder name on (Step 2)"
        elif not self.controller_output_path:
            has_error = True
            error_msg = "Please choose your output path on (Step 2)"
        elif not self.controller_output_filename:
            has_error = True
            error_msg = "Please type the output filename (Step 2)"
        elif not is_validate_filename_input(self.controller_output_filename):
            has_error = True
            error_msg = "Your output filename is incorrect, please make sure it contains only (letters, numbers, spaces, underscores, and periods)"

        return has_error, error_msg


    def handle_on_start(self):
        self.start_loading_animation()
        print("\nclicked start (Generate the PDF) button\n")

        BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        BASE_DIR = BASE_DIR + "/xp_excel_to_pdf/templates"

        # Print the absolute path
        print("\n\nAbsolute path:", BASE_DIR)
        print("\n\n")

        has_error, error_msg = self.validate_fields()
        print(f"\nhas_error : {has_error}, error_msg : {error_msg}\n")
        if has_error:
            self.stop_loading_animation()
            CTkMessagebox(
                title="Error",
                message=error_msg,
                icon="warning",
                option_1="Ok",
                option_focus=1,
                sound=True,
            )
            return

        # TODO:
        # DONE: 1 - start a loop over the dataframe that holds excel records
        # DONE: 2 - have a dict:
        #               - it's keys is the selected columns (use alternate column if found)
        #               - it's value is the excel value + the prefix/suffix
        # DONE: 3 - use this dict to create html code with value as <tr> and pass that to {{table_rows}}
        # DONE: 4 - get the value of main_title and sub_title and pass them to {{main_title}} and to {{sub_title}}
        # DONE: 5 - Process the output pdf filename for special formats
        # DONE: 6 - handle pdf folder name and output path
        # DONE: 7 - handle zip file creation
        # DONE: 8 - handle selected template
        # DONE: 9 - Add validations + display error dialog box to user:
        #                       - validates on click on generate button, make sure all required fields are filled
        #                       - make UI identifier if textbox is required
        #                       - add validation on dialog of excel file selection if (file not exists, sheet is not exists or wrong)
        #                       - add validation of output path if exists or not
        #                       - add validation if output pdf file name is valid or not
        # TODO!IMPORTANT: 10 - Clean up code

        # ON GOING: 11 - Do much test cases with different kind of excel files

        # >>>>>>>
        # DONE!!!!!!!!: 12 - Add loading animation or identifer when some process is happening, as:
        #                       - during program is reading the excel file
        #                       - during program is generating pdf files
        # >>>>>>>

        # DONE: 13 - Add success/error dialog message box to user, when actions complete or fails, eg. (completed generate pdf files)

        # ON HOLD: 14 - Test the code in windows os
                                # - TODO TODO !!FIX THE INCORRECT ALIGNED GUI TABLES ON WINDOWS OS

        # TODO!IMPORTANT: 15 - Test creating exe and check it if runs

        # DONE: 16 - Add more templates

        # TODO!IMPORTANT: 17 - Read the TEMPLATES from json file instead of config.py, so that users can able add thier own templates and modify the json file

        # TODO!IMPORTANT: 18 - Make the gui looking betters (sidebar buttons, loading text..)

        # TODO!IMPORRTANT : 19 - remove 'first_10_records' and any testing code

        has_error, error_msg = handle_folder_creation(self.controller_output_path, self.controller_output_folder_name)
        if has_error:
            self.stop_loading_animation()
            CTkMessagebox(
                title="Error",
                message=error_msg,
                icon="warning",
                option_1="Ok",
                option_focus=1,
                sound=True,
            )
            return

        # ONLY FOR TESTING, FOR REAL USE - TAKE THE WHOLE RECORDS AND NOT ONLY THE FIRST 10 RECORDS
        # first_10_records = self.controller_excel_dataframe.head(10)
        first_10_records = self.controller_excel_dataframe
        # for testing with 10 records
        print(f"\n\nfirst_10_records : {first_10_records}\n\n")

        main_title = self.controller_template_main_title
        sub_title = self.controller_template_sub_title

        print(f"\nmain_title : {main_title}\n")
        print(f"\nsub_title : {sub_title}\n")
        print(f"\ncontroller_select_columns_holder : {self.controller_select_columns_holder}\n")

        selected_template = TEMPLATES[self.controller_selected_template]["html"]
        print(f"\nselected_template : {selected_template}\n")

        # raise Exception("-----STOP----")

        template = env.get_template(selected_template)

        # table_rows = ""


        for index, row in first_10_records.iterrows():
            # temp output demo filename
            # output_pdf_filename = f"testing{index}.pdf"
            table_rows = ""
            render_values_dict = {
                "main_title": main_title,
                "sub_title": sub_title,
                "abs_path": BASE_DIR,
            }

            output_pdf_filename = process_pdf_filename_format(
                pdf_folder_path=self.controller_output_path,
                pdf_folder_name=self.controller_output_folder_name,
                pdf_filename_format=self.controller_output_filename,
                all_excel_columns=self.controller_excel_columns,
                excel_row=row,
                excel_row_index=index,
            )

            for column_list in self.controller_select_columns_holder:
                # column_list[0] is column name
                # column_list[1] is alternate column name
                # column_list[2] is prefix
                # column_list[3] is suffix

                column_name = column_list[0]
                alternate = column_list[1]
                prefix = column_list[2] or ""
                suffix = column_list[3] or ""

                table_column = alternate or column_name
                table_cell = str(prefix) + str(row[column_name]) + str(suffix)

                if self.controller_selected_template == "Modren Light":
                    table_rows += f'''
                    <tr class="spacer"><td colspan="100"></td></tr>
                    <tr scope="row">
                    <th scope="row"></th>
                    <td>{table_column}</td>
                    <td>{table_cell}</td>
                    </tr>
                    '''
                else:
                    table_rows += f"<tr><td>{table_column}</td><td>{table_cell}</td></tr>"

            print(f"\noutput_pdf_filename : {output_pdf_filename}\n")

            print(f"\n\ntable_rows : {table_rows}\n\n")

            render_values_dict['table_rows'] = table_rows
            html_content = template.render(render_values_dict)

            # Generate the PDF from the HTML content
            pdfkit.from_string(html_content, output_pdf_filename, options=pdf_options, configuration=pdfkit.configuration(wkhtmltopdf=self.path_to_wkhtmltopdf))

        if int(self.controller_create_zipfile) == 1:
            full_path = os.path.join(self.controller_output_path, self.controller_output_folder_name)
            full_path_zip_file = f"{full_path}.zip"

            if os.path.exists(full_path_zip_file):
                self.stop_loading_animation()
                CTkMessagebox(
                    title="Warning",
                    message=f"PDF files are generated, zip file is skipped as '{full_path_zip_file}' already exists.",
                    icon="warning",
                    option_1="Ok",
                    option_focus=1,
                    sound=True,
                )
                return
            else:
                create_zip(full_path, full_path_zip_file)

        CTkMessagebox(
            title="Success",
            message=f"PDF files are generated successfully.",
            icon="check",
            option_1="Ok",
            option_focus=1,
            sound=True,
        )

        self.stop_loading_animation()

    # Create a function to run the long-running function in a separate thread
    def on_start(self):
        # Create a thread to run the long-running function
        thread = threading.Thread(target=self.handle_on_start)
        # Start the thread
        thread.start()

    def on_closing(self):
        msg = CTkMessagebox(
            title="Exit",
            message="Do you want to close the program?",
            icon="question",
            option_1="Yes",
            option_2="No",
            option_focus=1,
            sound=True,
        )
        response = msg.get()

        if response=="Yes":
            self.destroy()