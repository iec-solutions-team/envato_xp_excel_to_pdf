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

class Dashboard(ctk.CTk):

    def __init__(self):
        super().__init__()

        # data controller
        self.controller_excel_columns = []
        self.controller_excel_dataframe = None
        self.controller_select_columns_holder = None

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


    def on_start(self):
        print("\nclicked start (Generate the PDF) button\n")


    def on_closing(self):
        msg = CTkMessagebox(title="Exit", message="Do you want to close the program?",
                        icon="question", option_1="Yes", option_2="No")
        response = msg.get()

        if response=="Yes":
            self.destroy()