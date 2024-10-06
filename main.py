import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageTk
import os
import logging
from paddleOCR import initialize_ocr_SLANet_LCNetV2, process_image, group_into_rows, save_as_xlsx, draw_bounding_boxes
import threading
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

# Set up logging
logging.basicConfig(level=logging.INFO, format='[%(asctime)s] [%(levelname)8s] %(message)s')
logger = logging.getLogger(__name__)

class OCRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("OCR to Excel Converter")
        self.root.geometry("1280x720")
        # Set application icon to 'icon.png'
        self.icon_image = tk.PhotoImage(file='icon.png')
        self.root.iconphoto(False, self.icon_image)
        self.style = ttk.Style('cosmo')  # Use a modern theme
        self.ocr_engine = tk.StringVar()
        self.ocr_engine.set("PaddleOCR")
        self.setup_ui()

    def setup_ui(self):
        # Main frame
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # Center frame for initial view
        self.center_frame = ttk.Frame(self.main_frame)
        self.center_frame.pack(expand=True)

        # Logo
        logo_img = Image.open("icon.png")
        logo_img = logo_img.resize((100, 100), Image.LANCZOS)
        self.logo_photo = ImageTk.PhotoImage(logo_img)
        self.logo_label = ttk.Label(self.center_frame, image=self.logo_photo)
        self.logo_label.pack(pady=(20, 10))

        # OCR Engine Selection
        ocr_label = ttk.Label(self.center_frame, text="Select OCR Engine:")
        ocr_label.pack(pady=(10, 5))
        ocr_dropdown = ttk.Combobox(self.center_frame, textvariable=self.ocr_engine, state="readonly", width=30)
        ocr_dropdown['values'] = ('PaddleOCR', 'Tesseract', 'EasyOCR', 'Google Vision AI')
        ocr_dropdown.pack(pady=(0, 20))

        # Upload Button
        self.upload_button = ttk.Button(self.center_frame, text="Upload Image", command=self.select_image, width=20)
        self.upload_button.pack(pady=(0, 20))

        # Status Label
        self.status_label = ttk.Label(self.center_frame, text="")
        self.status_label.pack(pady=(0, 10))

        # Progress bar (hidden initially)
        self.progress_bar = ttk.Progressbar(self.center_frame, mode='indeterminate')
        self.progress_bar.pack(pady=(0, 10))
        self.progress_bar.pack_forget()

        # Prepare frames for results
        self.top_frame = ttk.Frame(self.main_frame)
        self.left_frame = ttk.Frame(self.main_frame)
        self.right_frame = ttk.Frame(self.main_frame)

    def reset_ui(self):
        # Hide the top, left, and right frames
        self.top_frame.pack_forget()
        self.left_frame.pack_forget()
        self.right_frame.pack_forget()

        # Clear images from left and right frames
        for widget in self.left_frame.winfo_children():
            widget.destroy()
        for widget in self.right_frame.winfo_children():
            widget.destroy()

        # Show the center frame
        self.center_frame.pack(expand=True)

        # Reset status label
        self.status_label.config(text="")

        # Remove progress bar
        self.progress_bar.pack_forget()

    def select_image(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Image files", ("*.png", "*.jpg", "*.jpeg", "*.bmp", "*.tiff"))]
        )
        if file_path:
            logger.info(f"Selected image: {file_path}")
            self.reset_ui()  # Reset the UI before processing a new image
            self.process_image(file_path)

    def process_image(self, file_path):
        # Start a new thread for processing
        processing_thread = threading.Thread(target=self._process_image_thread, args=(file_path,))
        processing_thread.start()

    def _process_image_thread(self, file_path):
        try:
            ocr_engine = self.ocr_engine.get()

            # Show progress bar in the center frame
            self.progress_bar.pack(pady=(0, 10))
            self.progress_bar.start()

            if ocr_engine == "PaddleOCR":
                self.process_with_paddleocr(file_path)
            elif ocr_engine == "Tesseract":
                self.status_label.config(text="Tesseract OCR not implemented yet")
            elif ocr_engine == "EasyOCR":
                self.status_label.config(text="EasyOCR not implemented yet")
            elif ocr_engine == "Google Vision AI":
                self.status_label.config(text="Google Vision AI not implemented yet")
            else:
                raise ValueError("Please select an OCR engine.")
        except Exception as e:
            logger.error(f"Error processing image: {str(e)}")
            self.status_label.config(text=f"Error: {str(e)}")
        finally:
            # Hide progress bar
            self.progress_bar.stop()
            self.progress_bar.pack_forget()

    def process_with_paddleocr(self, file_path):
        ocr = initialize_ocr_SLANet_LCNetV2()
        data = process_image(file_path, ocr)
        rows = group_into_rows(data)

        if not rows:
            raise ValueError("No data extracted from image.")

        output_xlsx = os.path.splitext(file_path)[0] + "_output.xlsx"
        save_as_xlsx(rows, output_xlsx)

        output_image_path = os.path.splitext(file_path)[0] + "_output_image.jpg"
        draw_bounding_boxes(file_path, data, output_image_path)

        self.status_label.config(text=f"Excel file saved: {output_xlsx}")
        self.display_results(output_image_path, output_xlsx)

    def display_results(self, image_path, excel_path):
        self.reorganize_layout()

        # Clear previous images
        for widget in self.left_frame.winfo_children():
            widget.destroy()
        for widget in self.right_frame.winfo_children():
            widget.destroy()

        # Display image with bounding boxes
        self.display_image(image_path, self.left_frame)

        # Display Excel image
        excel_image_path = os.path.splitext(excel_path)[0] + "_excel_image.png"
        self.generate_excel_image(excel_path, excel_image_path)

        # Add padding to the right frame
        padding_frame = ttk.Frame(self.right_frame, padding=20)
        padding_frame.pack(fill=tk.BOTH, expand=True)
        self.display_image(excel_image_path, padding_frame)

    def reorganize_layout(self):
        # Hide the center frame
        self.center_frame.pack_forget()

        # Clear the top_frame
        for widget in self.top_frame.winfo_children():
            widget.destroy()

        # Set up top frame
        self.top_frame.pack(side=tk.TOP, fill=tk.X, pady=(10, 0))

        # Create a frame inside top_frame to center widgets
        top_inner_frame = ttk.Frame(self.top_frame)
        top_inner_frame.pack(anchor='center')

        # Add small logo
        small_logo_img = Image.open("icon.png")
        small_logo_img = small_logo_img.resize((50, 50), Image.LANCZOS)
        self.small_logo_photo = ImageTk.PhotoImage(small_logo_img)
        logo_label = ttk.Label(top_inner_frame, image=self.small_logo_photo)
        logo_label.pack(side=tk.LEFT, padx=(10, 5))

        ocr_label = ttk.Label(top_inner_frame, text="Select OCR Engine:")
        ocr_label.pack(side=tk.LEFT, padx=(10, 5))
        ocr_dropdown = ttk.Combobox(top_inner_frame, textvariable=self.ocr_engine, state="readonly", width=20)
        ocr_dropdown['values'] = ('PaddleOCR', 'Tesseract', 'EasyOCR', 'Google Vision AI')
        ocr_dropdown.pack(side=tk.LEFT, padx=(0, 10))
        upload_button = ttk.Button(top_inner_frame, text="Upload Image", command=self.select_image)
        upload_button.pack(side=tk.LEFT, padx=(0, 10))

        # Set up left and right frames
        self.left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

    def display_image(self, image_path, panel):
        image = Image.open(image_path)

        # Create a canvas to display the image
        canvas = tk.Canvas(panel, bg='white')
        canvas.pack(fill=tk.BOTH, expand=True)

        # Function to resize the image when the panel size changes
        def resize_image(event):
            # Get the size of the canvas
            canvas_width = event.width
            canvas_height = event.height

            # Calculate the scaling factor
            width_ratio = canvas_width / image.width
            height_ratio = canvas_height / image.height
            scale_factor = min(width_ratio, height_ratio, 1.0)  # Do not upscale images

            new_width = int(image.width * scale_factor)
            new_height = int(image.height * scale_factor)

            resized_image = image.resize((new_width, new_height), Image.LANCZOS)
            photo = ImageTk.PhotoImage(resized_image)

            canvas.delete("all")
            canvas.create_image(canvas_width/2, canvas_height/2, image=photo, anchor='center')
            canvas.image = photo  # Keep a reference

        # Bind the resize event to the function
        canvas.bind("<Configure>", resize_image)

    def generate_excel_image(self, excel_path, output_image_path):
        from openpyxl import load_workbook
        from PIL import Image, ImageDraw, ImageFont

        wb = load_workbook(excel_path)
        ws = wb.active

        # Increase cell size and font size
        cell_width = 90
        cell_height = 30
        font_size = 60
        border_size = 40  # Increased border size

        max_col = ws.max_column
        max_row = ws.max_row
        image_width = cell_width * max_col + 2 * border_size
        image_height = cell_height * max_row + 2 * border_size

        image = Image.new('RGB', (image_width, image_height), 'white')
        draw = ImageDraw.Draw(image)
        try:
            font = ImageFont.truetype("arial.ttf", font_size)
        except:
            font = ImageFont.load_default()

        for row in ws.iter_rows():
            for cell in row:
                col_idx = cell.column - 1
                row_idx = cell.row - 1
                x1 = col_idx * cell_width + border_size
                y1 = row_idx * cell_height + border_size
                x2 = x1 + cell_width
                y2 = y1 + cell_height

                # Draw cell background
                fill_color = 'FFFFFF'  # Default fill
                if cell.fill and cell.fill.fgColor:
                    if cell.fill.fgColor.type == 'rgb':
                        fill_color = cell.fill.fgColor.rgb[-6:]
                    elif cell.fill.fgColor.type == 'indexed':
                        fill_color = 'FFFFFF'  # Handle indexed colors as white

                draw.rectangle([x1, y1, x2, y2], fill=f'#{fill_color}', outline='black')

                # Draw cell text
                text = str(cell.value) if cell.value is not None else ''
                bbox = font.getbbox(text)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]
                text_x = x1 + (cell_width - text_width) / 2
                text_y = y1 + (cell_height - text_height) / 2
                draw.text((text_x, text_y), text, fill='black', font=font)

        image.save(output_image_path)

if __name__ == "__main__":
    root = ttk.Window(themename="cosmo")
    app = OCRApp(root)
    root.mainloop()
