import customtkinter as ctk
from CTkMessagebox import CTkMessagebox
from config import (
    PAGE_WIDTH,
    PAGE_HEIGHT,
    BLUE_COLOR,
    WHITE_COLOR,
    RED_COLOR,
    HOVER_RED_COLOR,
    GREEN_COLOR,
    HOVER_GREEN_COLOR,
    MAIN_TITLE_FONT,
    APP_TITLE,
)

class Dashboard(ctk.CTk):

    def __init__(self):
        super().__init__()

        window_size = (PAGE_WIDTH,PAGE_HEIGHT)

        # Get the screen width and height
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # Calculate the center of the screen
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

        self.sidebar_frame = ctk.CTkFrame(master=self, width=215, height=PAGE_HEIGHT)
        self.sidebar_frame.place(x=0, y=0)

        # Execute 'create_widgets' to load Gui elements
        self.create_widgets()


    def create_widgets(self):
        self.sidebar_main_title = ctk.CTkLabel(self.sidebar_frame, text=APP_TITLE, font=MAIN_TITLE_FONT)
        self.sidebar_main_title.place(x=30, y=20)

        self.sidebar_select_excel_button = ctk.CTkButton(
            self.sidebar_frame,
            text="Step (1) - Select Excel File",
            command=self.on_click_select_excel_button,
            corner_radius=0,
            width=215,
            height=47,
        )
        self.sidebar_select_excel_button.place(
            x=0,
            y=200,
        )

        self.sidebar_configure_pdf_button = ctk.CTkButton(
            self.sidebar_frame,
            text="Step (2) - Configure PDF",
            command=self.on_click_configure_pdf_button,
            corner_radius=0,
            width=215,
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
            width=215,
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
            width=215,
            height=50,
        )
        self.sidebar_exit_button.place(
            x=0,
            y=(PAGE_HEIGHT - 50),
        )


    def on_click_select_excel_button(self):
        print("\nclicked select excel button\n")

    def on_click_configure_pdf_button(self):
        print("\nclicked select excel button\n")

    def on_start(self):
        print("\nclicked start (Generate the PDF) button\n")

    def on_closing(self):
        msg = CTkMessagebox(title="Exit", message="Do you want to close the program?",
                        icon="question", option_1="Yes", option_2="No")
        response = msg.get()

        if response=="Yes":
            self.destroy()