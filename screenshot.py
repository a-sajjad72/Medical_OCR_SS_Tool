import time
import tkinter as tk
import sys
import subprocess
import tempfile


import pyautogui
from PIL import Image, ImageDraw, ImageEnhance, ImageTk


class SnipTool(tk.Toplevel):
    def __init__(self, parent, screenshot):
        super().__init__(parent)
        self.parent = parent
        self.attributes("-topmost", True)
        self.attributes("-fullscreen", True)

        # Use the provided screenshot instead of capturing a new one
        self.screenshot = screenshot
        self.screen_width, self.screen_height = self.screenshot.size

        # Create a darkened version of the screenshot
        self.dark_screenshot = ImageEnhance.Brightness(self.screenshot).enhance(0.3)
        self.tk_dark = ImageTk.PhotoImage(self.dark_screenshot)

        # Setup canvas
        self.canvas = tk.Canvas(
            self,
            width=self.screen_width,
            height=self.screen_height,
            highlightthickness=0,
            cursor="crosshair",
        )
        self.canvas.pack()
        self.image_on_canvas = self.canvas.create_image(
            0, 0, anchor="nw", image=self.tk_dark
        )

        # Variables for selection
        self.start_x = None
        self.start_y = None
        self.rect = None
        self.selection = None

        # Bind mouse events
        self.canvas.bind("<ButtonPress-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_button_release)

    def on_button_press(self, event):
        self.start_x = event.x
        self.start_y = event.y
        self.rect = self.canvas.create_rectangle(
            self.start_x,
            self.start_y,
            self.start_x,
            self.start_y,
            outline="red",
            width=2,
        )

    def on_mouse_drag(self, event):
        self.canvas.coords(self.rect, self.start_x, self.start_y, event.x, event.y)
        x1 = min(self.start_x, event.x)
        y1 = min(self.start_y, event.y)
        x2 = max(self.start_x, event.x)
        y2 = max(self.start_y, event.y)
        mask = Image.new("L", (self.screen_width, self.screen_height), 0)
        draw = ImageDraw.Draw(mask)
        draw.rectangle((x1, y1, x2, y2), fill=255)
        composite = self.dark_screenshot.copy()
        composite.paste(self.screenshot, mask=mask)
        self.tk_composite = ImageTk.PhotoImage(composite)
        self.canvas.itemconfig(self.image_on_canvas, image=self.tk_composite)

    def on_button_release(self, event):
        x1 = min(self.start_x, event.x)
        y1 = min(self.start_y, event.y)
        x2 = max(self.start_x, event.x)
        y2 = max(self.start_y, event.y)
        self.selection = (x1, y1, x2, y2)
        self.destroy()


def select_region(parent, screenshot):
    snip = SnipTool(parent, screenshot)
    parent.wait_window(snip)
    return snip.selection

def capture_screenshot(root_win):
    # Capture the screenshot of the entire screen
    if sys.platform == "win32":
        screenshot = pyautogui.screenshot().convert("RGB")
        region = select_region(root_win, screenshot)
        if region:
            x1, y1, x2, y2 = region
            snip_img = screenshot.crop((x1, y1, x2, y2))
        return snip_img
    else:
        temp_dir = tempfile.gettempdir()
        screenshot_path = os.path.join(temp_dir, 'screenshot.png')
        subprocess.call(['screencapture', '-ix', screenshot_path])
        image = Image.open(screenshot_path)
        return image




# Example integration
if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("400x300")
    root.title("Main Application")

    def do_snip():
        root.withdraw()
        root.update()  # Force Tkinter to process the hide request
        time.sleep(0.1)  # Give the OS time to hide the window (adjust as needed)

        # Now take the screenshot
        snip_img = capture_screenshot(root)
        if snip_img:
            snip_img.save("snip_from_app.png")
            snip_img.show()
        root.deiconify()

    btn = tk.Button(root, text="Select Region", command=do_snip)
    btn.pack(pady=20)

    root.mainloop()
