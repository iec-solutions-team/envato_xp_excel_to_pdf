import customtkinter as ctk
from config import (
    FRAME_WIDTH,
    FRAME_HEIGHT,
    MAIN_TITLE_FONT,
    APP_TITLE,
)


class InitialPage(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, width=FRAME_WIDTH, height=FRAME_HEIGHT, corner_radius=0, **kwargs)
        self.place(x=215, y=0)

        # Execute 'create_widgets' to load Gui elements
        self.create_widgets()


    def create_widgets(self):
        self.label = ctk.CTkLabel(self, text=f"Welcome to {APP_TITLE}", font=("Arial", 25))
        self.label.place(relx=0.5, rely=0.5, anchor="center")