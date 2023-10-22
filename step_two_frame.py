import customtkinter as ctk
from config import (
    FRAME_WIDTH,
    FRAME_HEIGHT,
)


class StepTwoFrame(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, width=FRAME_WIDTH, height=FRAME_HEIGHT, corner_radius=0, **kwargs)
        self.place(x=215, y=0)

        # add widgets onto the frame, for example:
        self.label = ctk.CTkLabel(self, text="Step Two")
        self.label.place(x=30, y=20)