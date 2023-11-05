

PAGE_WIDTH = 1100
PAGE_HEIGHT = 650
BLUE_COLOR = "#5E95FF"
WHITE_COLOR = "#FFFFFF"
RED_COLOR = "#D22B2B"
HOVER_RED_COLOR = "#EE4B2B"
GREEN_COLOR = "#4CBB17"
HOVER_GREEN_COLOR = "#0BDA51"
MAIN_TITLE_FONT = ("Arial", 18)
NORMAL_FONT = ("Arial", 14)
SMALL_FONT = ("Arial", 12)
APP_TITLE = "XP - Excel to PDF"

SIDEBAR_WIDTH = 215
FRAME_WIDTH = PAGE_WIDTH - SIDEBAR_WIDTH
FRAME_HEIGHT = PAGE_HEIGHT

PADDING_X_FROM_SIDEBAR = 30

EVEN_TREEVIEW_BG = "#FFFFFF"    # Background color for even rows in TreeView
ODD_TREEVIEW_BG = "#E7EBF5"     # Background color for odd rows in TreeView


EMPTY_VALUES_LIST = ["", " ", None]


TEMPLATES = {
    "Simple": {
        "html": "./templates/simple.html",
        "image": "./templates/simple_preview.png",
    },
    "Zebra Striping": {
        "html": "./templates/zebra_striping.html",
        "image": "./templates/zebra_striping_preview.png",
    },
    "Modren Light": {
        "html": "./templates/modren_table_light.html",
        "image": "./templates/modren_table_light_preview.png",
    },
    "Modren Dark": {
        "html": "./templates/modren_table_dark.html",
        "image": "./templates/modren_table_dark_preview.png",
    },
}