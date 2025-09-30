import sqlite3
from kivy.app import App
from kivy.uix.image import Image
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.textinput import TextInput
from kivy.uix.filechooser import FileChooserIconView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.scrollview import ScrollView
from kivy.core.window import Window
from kivy.graphics import Rectangle, Color, Line
from kivy.uix.popup import Popup
from tkinter import filedialog, Tk
import os
import sys
import fitz  # PyMuPDF
import re
from difflib import SequenceMatcher

from kivy.uix.button import Button
from kivy.properties import NumericProperty
from kivy.animation import Animation
from kivy.core.window import Window

import pytesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Users\shashank.aswath\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"

from PIL import Image as PILImage

import io


REF_ID_REGEX = r'\b[A-Z]{2}-\d{9}\b'



def extract_location_text(concession_text):
    lines = concession_text.splitlines()
    location_data = []
    capture = False
    for i, line in enumerate(lines):
        if "LOCATION" in line.upper():
            capture = True
            if ':' in line:
                location_data.append(line.split(":", 1)[1].strip())
            continue
        if capture:
            if any(kw in line.upper() for kw in ["NONCONFORMITY", "CHARACTERISTIC", "DESCRIPTION", "PART NUMBER"]):
                break
            if line.strip():
                location_data.append(line.strip())
    return " ".join(location_data)


def normalize_item(item):
    item = item.upper().strip()
    item = item.replace("–", "-").replace("—", "-")
    item = re.sub(r"[^A-Z0-9\- RLH]+", "", item)
    item = re.sub(r"\s+", " ", item)
    item = item.replace(" L H", " LH").replace(" R H", " RH")
    item = re.sub(r"0+(\d+[A-Z]?)", r"\1", item)
    item = item.replace("FROM ", "")
    item = item.replace("TO ", "")
    item = re.sub(r"FRAME\s*(\d+)([A-Z]?)\s*-\s*(\d+)([A-Z]?)", r"FRAME \1\2-\3\4", item)
    item = re.sub(r"FRAME\s*(\d+)([A-Z]?)", r"FRAME \1\2", item)
    item = re.sub(r"STRINGER\s*(\d+)([RL]H)?\s*-\s*(\d+)([RL]H)?", r"STRINGER \1\2-\3\4", item)
    item = re.sub(r"STRINGER\s*(\d+)([RL]H)?", r"STRINGER \1\2", item)
    return item.strip()


def fuzzy_match(a, b):
    return SequenceMatcher(None, a, b).ratio() * 100


def find_location_matches(sketch_pages, location_items):
    normalized_items = [normalize_item(x) for x in location_items if x]
    for _, text in sketch_pages:
        lines = text.splitlines()
        for line in lines:
            norm_line = normalize_item(line)
            for item in normalized_items:
                if item and (item in norm_line or norm_line in item or fuzzy_match(item, norm_line) >= 85):
                    return True
    return False


class TableCell(BoxLayout):
    def __init__(self, text, bold=False, fixed_height=30, col_ratio=1, bg_color=(1, 1, 1, 0), **kwargs):
        super().__init__(**kwargs)
        self.size_hint_y = None
        self.size_hint_x = col_ratio
        self.height = fixed_height
        self.padding = 5
        # with self.canvas.before:
        #     Color(*bg_color)
        #     self.bg_rect = Rectangle(size=self.size, pos=self.pos)
        #     Color(0, 0, 0, 1)
        #     self.border_line = Line(rectangle=(self.x, self.y, self.width, self.height), width=1)
        with self.canvas.before:
            self.bg_color_instruction = Color(*bg_color)  # Store Color reference
            self.bg_rect = Rectangle(size=self.size, pos=self.pos)
            Color(0, 0, 0, 1)
            self.border_line = Line(rectangle=(self.x, self.y, self.width, self.height), width=1)

        self.bind(pos=self.update_graphics, size=self.update_graphics)

        self.label = Label(
            text=text,
            bold=bold,
            color=(0, 0, 0, 1),
            font_size=26,
            halign='center',
            valign='middle',
            size_hint=(1, 1),
            text_size=(0, None),
        )
        self.label.bind(size=self._update_text_size)
        self.add_widget(self.label)

    def _update_text_size(self, instance, value):
        self.label.text_size = (self.label.width, None)

    def update_graphics(self, *args):
        self.bg_rect.size = self.size
        self.bg_rect.pos = self.pos
        self.border_line.rectangle = (self.x, self.y, self.width, self.height)






class HoverButton(Button):
    scale = NumericProperty(1.0)

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        Window.bind(mouse_pos=self.on_mouse_pos)
        self.original_color = self.background_color
        self.hovered = False

    def on_mouse_pos(self, *args):
        if not self.get_root_window():
            return
        inside = self.collide_point(*self.to_widget(*args[1]))
        if inside and not self.hovered:
            self.hovered = True
            self.animate_hover(True)
        elif not inside and self.hovered:
            self.hovered = False
            self.animate_hover(False)

    def animate_hover(self, hover):
        if hover:
            self.background_color = (0.2, 0.6, 1, 1)
        else:
            self.background_color = self.original_color

    def on_press(self):
        Animation(scale=0.95, d=0.1).start(self)

    def on_release(self):
        Animation(scale=1.0, d=0.1).start(self)




#self.last_uploaded_pdf_path = None


class FlexibleKivyApp(App):
    def build(self):
        Window.size = (800, 600)
        Window.fullscreen = False
        self.current_dir = os.path.dirname(os.path.abspath(__file__))
        self.bg_path = os.path.join(self.current_dir, "Assets", "A350_1000.jpg")
        self.logo_path = os.path.join(self.current_dir, "Assets", "AXISCADES_Logo.png")
        self.settings_path = os.path.join(self.current_dir, "Assets", "Settings.png")

        if not all(os.path.exists(p) for p in [self.bg_path, self.logo_path, self.settings_path]):
            print("One or more asset files missing.", file=sys.stderr)
            return None

        self.main_root = FloatLayout()
        self.main_root = FloatLayout()
        self.last_uploaded_pdf_path = None  # ✅ Remember last uploaded file
        self.switch_to_main()

        self.switch_to_main()
        return self.main_root
    
    


    def build_main_screen(self):
        root = FloatLayout()

        bg = Image(source=self.bg_path, allow_stretch=True, keep_ratio=False,
                   size_hint=(1, 1), pos_hint={'x': 0, 'y': 0})
        root.add_widget(bg)

        header = self.create_header()
        root.add_widget(header)

        main_container = FloatLayout(size_hint=(0.9, 0.9),
                                     pos_hint={'center_x': 0.5, 'center_y': 0.45})
        with main_container.canvas.before:
            Color(1, 1, 1, 0.7)
            self.main_rect = Rectangle(size=(0, 0), pos=(0, 0))
        main_container.bind(size=self._update_main_rect, pos=self._update_main_rect)

        self.checkpoint_list = [
            "Concession number displayed on Header in all sketches",
            "Sketch numbers are in ascending order (Sketch1,2,3 etc)",
            "Location details( Section, Frame, Stringer, SSRH,SSLH,US,LS, Exter surface affected/not affected ) are correct on cover page and sketches",
            "Type of Nonconformity is correct on cover page and sketches with respect to non-conformity description",
            "Nonconformity details (Hole code, Nominal fasteners, collars, Actual drilled code) are correct on cover page and sketches as per drawing.",
            "Number of Nonconformities mentioned on cover page are matching with sketches",
            "Hole code and installed fasteners are matching",
            "Material compatibility of installed fasteners are as per Drawing requirements",
            "Is Fastener and collar details are matching for both nominal and actual case"
        ]

        ratios = [1 / 9, 6 / 9, 2 / 9]  # Adjusted for 3 columns


        table_wrapper = BoxLayout(orientation='vertical',
                                  size_hint=(0.9, 0.95),
                                  pos_hint={'center_x': 0.5, 'center_y': 0.5})


        upload_btn_row = BoxLayout(orientation='horizontal', size_hint=(1, None), height=40)
        upload_btn = HoverButton(text="Upload Concession Document", size_hint=(None, None), size=(300, 40),
                         pos_hint={'x': 0})
        upload_btn.bind(on_press=self.upload_concession_pdf)
        upload_btn_row.add_widget(upload_btn)
        clear_btn = HoverButton(text="Clear", size_hint=(None, None), size=(100, 40))
        clear_btn.bind(on_press=self.clear_results)
        upload_btn_row.add_widget(clear_btn)

        upload_btn_row.add_widget(Label())  # Filler to push content left
        table_wrapper.add_widget(upload_btn_row)


        header_row = GridLayout(cols=3, size_hint=(1, None), height=50)
        header_row.add_widget(TableCell("S.No", bold=True, fixed_height=50, col_ratio=ratios[0]))
        header_row.add_widget(TableCell("Check Points", bold=True, fixed_height=50, col_ratio=ratios[1]))
        header_row.add_widget(TableCell("Status", bold=True, fixed_height=50, col_ratio=ratios[2]))
        #header_row.add_widget(TableCell("Document Viewer", bold=True, fixed_height=50, col_ratio=ratios[3]))
        table_wrapper.add_widget(header_row)

        scroll_view = ScrollView(size_hint=(1, 1))
        self.table_body = GridLayout(cols=3, spacing=0, size_hint_y=None)
        self.table_body.bind(minimum_height=self.table_body.setter('height'))

        self.status_cells = []

        for i, checkpoint in enumerate(self.checkpoint_list, start=1):
            self.table_body.add_widget(TableCell(str(i), fixed_height=120, col_ratio=ratios[0]))
            self.table_body.add_widget(TableCell(checkpoint, fixed_height=120, col_ratio=ratios[1]))

            status_cell = TableCell("", fixed_height=120, col_ratio=ratios[2])
            self.status_cells.append(status_cell)
            self.table_body.add_widget(status_cell)

            #self.table_body.add_widget(TableCell("", fixed_height=120, col_ratio=ratios[3]))

        scroll_view.add_widget(self.table_body)
        table_wrapper.add_widget(scroll_view)

        main_container.add_widget(table_wrapper)

       # === Filename label box (top-right corner inside main_container) ===
        self.file_label_box = FloatLayout(
            size_hint=(None, None),
            size=(220, 50),
            pos_hint={'right': 0.950, 'top': 0.993}  # Move to top-right inside the white container
        )

        self.file_label = Label(
            text="",
            font_size=20,
            color=(0, 0, 0, 1),
            halign='center',
            valign='middle',
            size_hint=(1, 1),
            pos_hint={'right': 1.0, 'top': 1.0}  # Ensure it aligns within the parent FloatLayout
        )
        self.file_label.bind(size=self._update_label_textbox)
        self.file_label_box.add_widget(self.file_label)

        with self.file_label_box.canvas.after:
            Color(0, 0, 0, 1)
            self.file_label_border = Line(rectangle=(0, 0, 320, 50), width=1.3)

        self.file_label_box.bind(pos=self._update_label_box_border, size=self._update_label_box_border)
        main_container.add_widget(self.file_label_box)


        root.add_widget(main_container)

        return root
    
    def _update_label_textbox(self, instance, value):
        self.file_label.text_size = self.file_label.size

    def _update_label_box_border(self, instance, value):
        self.file_label_border.rectangle = (
            self.file_label_box.x,
            self.file_label_box.y,
            self.file_label_box.width,
            self.file_label_box.height
        )

    def upload_concession_pdf(self, instance):
        try:
            root = Tk()
            root.withdraw()
            file_path = filedialog.askopenfilename(title="Select Concession PDF", filetypes=[("PDF files", "*.pdf")])
            root.destroy()
            if file_path:
                self.process_pdf_for_checkpoint_1(file_path)
                self.process_pdf_for_checkpoint_2(file_path)
                self.process_pdf_for_checkpoint_3(file_path)
                self.process_pdf_for_checkpoint_4(file_path)
                self.process_pdf_for_checkpoint_5(file_path)
                self.process_pdf_for_checkpoint_6(file_path)
                self.process_pdf_for_checkpoint_7(file_path)
                self.process_pdf_for_checkpoint_8(file_path)
                self.process_pdf_for_checkpoint_9(file_path)

        except Exception as e:
            print(f"Error opening file dialog: {e}")

        self.file_label.text = self.extract_filename_display_text(file_path)
        if file_path:
            self.last_uploaded_pdf_path = file_path



    # Checkpoint 1 
    def process_pdf_for_checkpoint_1(self, pdf_path):
        try:
            doc = fitz.open(pdf_path)
        except Exception as e:
            print(f"❌ Failed to open PDF: {e}")
            return

        page_types = []
        ref_id = None
        mismatches = []

        for i, page in enumerate(doc):
            rect = page.rect
            top_quarter = fitz.Rect(rect.x0, rect.y0, rect.x1, rect.y1 * 0.25)
            top_text = page.get_text(clip=top_quarter).upper()

            if "CONCESSION" in top_text:
                page_types.append((i, 'CONCESSION'))
            elif "SKETCH" in top_text:
                page_types.append((i, 'SKETCH'))

        first_page_text = doc[0].get_text()
        match = re.search(REF_ID_REGEX, first_page_text)
        if match:
            ref_id = match.group()
        else:
            self.update_status_cell(0, "Not Matching", (1, 0, 0, 0.7))
            return

        for i, page_type in page_types:
            if i == 0:
                continue
            text = doc[i].get_text()
            found_ids = re.findall(REF_ID_REGEX, text)
            for fid in found_ids:
                if fid != ref_id:
                    mismatches.append((i + 1, page_type, fid))

        if mismatches:
            self.update_status_cell(0, "Not Matching", (1, 0, 0, 0.7))
        else:
            self.update_status_cell(0, "Matching", (0, 1, 0, 0.7))

    # def update_status_cell(self, index, text, bg_color):
    #     cell = self.status_cells[index]
    #     cell.label.text = text
    #     with cell.canvas.before:
    #         Color(*bg_color)
    #         cell.bg_rect = Rectangle(size=cell.size, pos=cell.pos)
    #     cell.bind(pos=cell.update_graphics, size=cell.update_graphics)

    def update_status_cell(self, index, text, bg_color):
        cell = self.status_cells[index]
        cell.label.text = text

        # ✅ Update the original background color instruction instead of replacing canvas
        if hasattr(cell, 'bg_color_instruction'):
            cell.bg_color_instruction.rgba = bg_color

        # ✅ Force redraw
        if hasattr(cell, 'bg_rect'):
            cell.bg_rect.size = cell.size
            cell.bg_rect.pos = cell.pos

        cell.bind(pos=cell.update_graphics, size=cell.update_graphics)


    def _update_main_rect(self, instance, value):
        self.main_rect.size = instance.size
        self.main_rect.pos = instance.pos

    def create_header(self):
        header_height = 50
        header = BoxLayout(orientation='horizontal',
                           size_hint=(1, None), height=header_height,
                           pos_hint={'top': 1}, padding=10, spacing=10)
        with header.canvas.before:
            Color(30 / 255, 30 / 255, 30 / 255, 1)
            self.rect = Rectangle(size=header.size, pos=header.pos)
        header.bind(size=self._update_rect, pos=self._update_rect)

        logo = Image(source=self.logo_path, size_hint=(None, None), size=(150, 60),
                     allow_stretch=True, pos_hint={'center_y': 0.5})
        header.add_widget(logo)

        title = Label(text="T100-Automated Checker", font_size=25, bold=True,
                      color=(1, 1, 1, 1), halign='center', valign='middle')
        header.add_widget(title)

        settings_btn = Button(size_hint=(None, None), size=(35, 35), pos_hint={'center_y': 0.5},
                              background_normal=self.settings_path, background_color=(1, 1, 1, 1),
                              border=(0, 0, 0, 0))
        settings_btn.bind(on_press=self.on_settings_click)
        header.add_widget(settings_btn)

        return header

    def _update_rect(self, instance, value):
        self.rect.size = instance.size
        self.rect.pos = instance.pos

    def on_settings_click(self, instance):
        overlay = FloatLayout(size_hint=(1, 1))
        with overlay.canvas.before:
            Color(1, 1, 1, 0.07)
            Rectangle(size=Window.size, pos=(0, 0))

        bg = Image(source=self.bg_path, allow_stretch=True, keep_ratio=False,
                   size_hint=(1, 1), pos_hint={'x': 0, 'y': 0})
        overlay.add_widget(bg)

        header = self.create_header()
        overlay.add_widget(header)

        settings_window = FloatLayout(size_hint=(0.8, 0.8),
                                      pos_hint={'center_x': 0.5, 'center_y': 0.45})
        with settings_window.canvas.before:
            Color(1, 1, 1, 0.7)
            self.settings_rect = Rectangle(size=(0, 0), pos=(0, 0))
        settings_window.bind(size=self._update_settings_rect, pos=self._update_settings_rect)

        title_label = Label(text="Settings Page", font_size=22, bold=True, color=(0, 0, 0, 1),
                            size_hint=(None, None), size=(200, 40), pos_hint={'center_x': 0.5, 'top': 0.98})
        settings_window.add_widget(title_label)

        db_section = BoxLayout(orientation='horizontal', size_hint=(0.9, None), height=40,
                               pos_hint={'center_x': 0.5, 'top': 0.88}, spacing=10)
        label = Label(text="Database location:", bold=True, color=(0, 0, 0, 1),
                      size_hint=(None, 1), width=150)
        #self.db_input = TextInput(multiline=False, size_hint=(0.7, 1))

        # ✅ Updated line to prefill relative DB path
        self.db_input = TextInput(
            multiline=False,
            size_hint=(0.7, 1),
            text=r"C:\Users\shashank.aswath\Documents\airbus_t100\Assets\Hole codes and fasteners.db"
        )
        upload_btn = Button(text="Upload", size_hint=(None, 1), width=100)
        upload_btn.bind(on_press=self.open_file_chooser)

        db_section.add_widget(label)
        db_section.add_widget(self.db_input)
        db_section.add_widget(upload_btn)
        settings_window.add_widget(db_section)

        go_home_btn = Button(text="Go Home", size_hint=(None, None), size=(100, 40),
                             pos_hint={'right': 0.98, 'top': 0.98})
        go_home_btn.bind(on_press=lambda x: self.switch_to_main())
        settings_window.add_widget(go_home_btn)

        overlay.add_widget(settings_window)
        self.main_root.clear_widgets()
        self.main_root.add_widget(overlay)

    def _update_settings_rect(self, instance, value):
        self.settings_rect.size = instance.size
        self.settings_rect.pos = instance.pos

    def open_file_chooser(self, instance):
        root = Tk()
        root.withdraw()
        file_path = filedialog.askopenfilename(title="Select Database File", filetypes=[("SQLite DB", "*.db")])
        root.destroy()
        if file_path:
            self.db_input.text = file_path

    # Checkpoint 2
    def process_pdf_for_checkpoint_2(self, pdf_path):
        sketch_data = self.extract_sketch_numbers(pdf_path)
        if not sketch_data:
            self.update_status_cell(1, "Not In Order", (1, 0, 0, 0.7))
            return
        in_order = self.check_sketch_order(sketch_data)
        if in_order:
            self.update_status_cell(1, "In Order", (0, 1, 0, 0.5))
        else:
            self.update_status_cell(1, "Not In Order", (1, 0, 0, 0.7))
    
    def extract_sketch_numbers(self, pdf_path):
        try:
            doc = fitz.open(pdf_path)
        except Exception as e:
            print(f"❌ Error opening PDF: {e}")
            return []

        sketch_numbers = []
        for i in range(len(doc)):
            page = doc[i]
            rect = page.rect
            top_quarter = fitz.Rect(rect.x0, rect.y0, rect.x1, rect.y1 * 0.25)
            top_text = page.get_text(clip=top_quarter).upper()

            match = re.search(r"SKETCH\s+(\d+)", top_text)
            if match:
                sketch_num = int(match.group(1))
                sketch_numbers.append((i + 1, sketch_num))

        return sketch_numbers
    
    def check_sketch_order(self, sketch_numbers):
        expected = 1
        success = True
        for page_num, actual_num in sketch_numbers:
            if actual_num != expected:
                success = False
            expected += 1
        return success
    
    # Checkpoint 3
    def process_pdf_for_checkpoint_3(self, pdf_path):
        try:
            doc = fitz.open(pdf_path)
        except Exception as e:
            print(f"❌ Failed to open PDF: {e}")
            return

        location_texts = []
        sketch_pages = []

        for page_num, page in enumerate(doc):
            text = page.get_text()
            top_text = page.get_textbox((0, 0, page.rect.width, page.rect.height * 0.25))

            text = text.decode('utf-8', errors='ignore') if isinstance(text, bytes) else text
            top_text = top_text.decode('utf-8', errors='ignore') if isinstance(top_text, bytes) else top_text

            if "CONCESSION" in top_text.upper():
                extracted = extract_location_text(text)
                if extracted:
                    location_texts.append(extracted)
            elif "SKETCH" in top_text.upper():
                sketch_pages.append((page_num + 1, text))

        if not location_texts:
            self.update_status_cell(2, "Not Correct", (1, 0, 0, 0.7))
            return

        combined_text = " ".join(location_texts)
        location_items = [item.strip() for item in re.split(r",| {2,}|\n|;", combined_text.upper()) if item.strip()]

        matched = find_location_matches(sketch_pages, location_items)

        if matched:
            self.update_status_cell(2, "Correct", (0, 1, 0, 0.6))
        else:
            self.update_status_cell(2, "Not Correct", (1, 0, 0, 0.7))



    # Checkpoint 4
    def process_pdf_for_checkpoint_4(self, pdf_path):
        try:
            doc = fitz.open(pdf_path)
        except Exception as e:
            print(f"❌ Failed to open PDF: {e}")
            self.update_status_cell(3, "Not Correct", (1, 0, 0, 0.7))
            return

        def extract_nc_lines(text):
            """Extracts all NONCONFORMITY values in the given page text."""
            pattern = r"NON\s*CONFORMITY(?:\s*\d+)?\s*:\s*(.+?)(?=\n|\s{2,}|$)"
            return [m.strip().upper() for m in re.findall(pattern, text, re.IGNORECASE)]

        def normalize_nc(text):
            """Normalize NC values for comparison."""
            text = text.upper()
            text = re.sub(r"\b1X\b", "", text)  # remove quantity prefix
            text = re.sub(r"[^A-Z0-9 ]+", "", text)  # remove non-alphanum
            return text.strip()

        # Extract NCs from all Concession pages
        cover_nc_values = []
        for page in doc:
            text = page.get_text()
            if "CONCESSION" in text.upper():
                cover_nc_values += extract_nc_lines(text)

        if not cover_nc_values:
            print("❌ No NONCONFORMITY entries found in Concession pages.")
            self.update_status_cell(3, "Not Correct", (1, 0, 0, 0.7))
            return

        normalized_cover_nc = [normalize_nc(nc) for nc in cover_nc_values]
        matched_flags = [False] * len(normalized_cover_nc)

        # Scan SKETCH pages and look for matches
        for page in doc:
            text = page.get_text()
            if "SKETCH" in text.upper():
                sketch_lines = text.splitlines()
                for line in sketch_lines:
                    norm_line = normalize_nc(line)
                    for idx, nc in enumerate(normalized_cover_nc):
                        if nc in norm_line:
                            matched_flags[idx] = True

        if all(matched_flags):
            self.update_status_cell(3, "Correct", (0, 1, 0, 0.6))
        else:
            print("❌ Some NONCONFORMITY entries were not found in Sketch pages.")
            self.update_status_cell(3, "Not Correct", (1, 0, 0, 0.7))


    # Checkpoint 5
    def process_pdf_for_checkpoint_5(self, pdf_path):
        

        try:
            doc = fitz.open(pdf_path)
        except Exception as e:
            print(f"❌ Failed to open PDF: {e}")
            self.update_status_cell(4, "Not Correct", (1, 0, 0, 0.7))
            return

        def extract_specifications(concession_text: str):
            lines = concession_text.splitlines()
            specs = {
                "Nominal hole specification": None,
                "Nominal Fastener Sys-fastener": None,
                "Drilled Hole Specification": None,
            }

            for line in lines:
                for key in specs.keys():
                    if key.upper() in line.upper() and ":" in line:
                        value = line.split(":", 1)[1].strip()
                        if value:
                            specs[key] = value
            return specs

        def normalize(text):
            return text.lower().replace(",", ".").replace("⌀", "").replace("ø", "").strip() if text else ""

        def extract_text_with_ocr(page):
            try:
                pix = page.get_pixmap(dpi=300)
                img_data = pix.tobytes("png")
                image = PILImage.open(io.BytesIO(img_data))  # ← FIXED!
                return pytesseract.image_to_string(image)
            except Exception as e:
                print(f"OCR error: {e}")
                return ""



        def search_sketch_fields(sketch_pages, specs):
            norm_specs = {k: normalize(v) for k, v in specs.items() if v}

            for sheet_num, text in sketch_pages:
                text_lines = [line.strip() for line in text.splitlines() if line.strip()]
                full_text = " ".join(text_lines)
                full_text += "\n" + sketch_ocr.get(sheet_num, "")
                full_text = full_text.lower()

                matches = []
                for label, value in norm_specs.items():
                    if value and value in full_text:
                        matches.append(label)

                if len(matches) == len(norm_specs):
                    return True

            return False

        concession_text = ""
        sketch_pages = []
        sketch_ocr = {}
        sketch_sheet_counter = 0
        found_concession = False

        for page_num, page in enumerate(doc):
            text = page.get_text()
            top_text = page.get_textbox((0, 0, page.rect.width, page.rect.height * 0.25))

            if not found_concession and "CONCESSION" in top_text.upper():
                concession_text = text
                found_concession = True
            elif "SKETCH" in top_text.upper():
                sketch_sheet_counter += 1
                #ocr_text = extract_text_with_ocr(page)
                try:
                    ocr_text = extract_text_with_ocr(page)
                except Exception as e:
                    print(f"OCR failed on page {page_num + 1}: {e}")
                    ocr_text = ""

                sketch_pages.append((sketch_sheet_counter, text))
                sketch_ocr[sketch_sheet_counter] = ocr_text

        if not concession_text:
            print("❌ No CONCESSION sheet found.")
            self.update_status_cell(4, "Not Correct", (1, 0, 0, 0.7))
            return

        specs = extract_specifications(concession_text)
        match_found = search_sketch_fields(sketch_pages, specs)

        if match_found:
            self.update_status_cell(4, "Correct", (0, 1, 0, 0.6))
        else:
            self.update_status_cell(4, "Not Correct", (1, 0, 0, 0.7))


    # Checkpoint 6
    def process_pdf_for_checkpoint_6(self, pdf_path):
        def normalize_nc(text):
            text = re.sub(r"[^A-Z0-9]", "", text.upper())
            match = re.match(r"(NONCONFORMITY)(0*)(\d+)$", text)
            if match:
                return f"{match.group(1)}{int(match.group(3))}"  # strip leading zeros
            return text

        def extract_nonconformities(text):
            nc_set = set()

            # Numbered NCs
            numbered_matches = re.findall(r'NON\s*CONFORMITY\s*0*(\d+)', text, re.IGNORECASE)
            for match in numbered_matches:
                nc_set.add(f"NONCONFORMITY{int(match)}")

            # Ranges like NONCONFORMITIES 1-3
            range_matches = re.findall(
                r'NONCONFORMIT(?:Y|IES)[\s:]*0*([0-9]+)[\s\-–to]+0*([0-9]+)',
                text,
                re.IGNORECASE,
            )
            for start, end in range_matches:
                for i in range(int(start), int(end) + 1):
                    nc_set.add(f"NONCONFORMITY{i}")

            # Generic NC (no number)
            generic_matches = re.findall(r'\bNON\s*CONFORMITY\b(?!\s*\d)', text, re.IGNORECASE)
            if generic_matches:
                nc_set.add("NONCONFORMITY")

            return {normalize_nc(nc) for nc in nc_set}

        try:
            doc = fitz.open(pdf_path)
        except Exception as e:
            print(f"❌ Failed to open PDF: {e}")
            self.update_status_cell(5, "Not Matching", (1, 0, 0, 0.7))
            return

        concession_pages = {}
        sketch_pages = {}

        for i, page in enumerate(doc):
            text = page.get_text()
            lines = text.splitlines()
            first_quarter = lines[:max(1, len(lines) // 4)]
            is_sketch = any("SKETCH" in line.upper() for line in first_quarter)
            if "CONCESSION" in text.upper():
                concession_pages[i + 1] = text
            if is_sketch:
                sketch_pages[i + 1] = text

        if not concession_pages or not sketch_pages:
            self.update_status_cell(5, "Not Matching", (1, 0, 0, 0.7))
            return

        # Extract NCs
        all_concession_ncs = set()
        for text in concession_pages.values():
            all_concession_ncs.update(extract_nonconformities(text))

        all_sketch_ncs = set()
        for text in sketch_pages.values():
            all_sketch_ncs.update(extract_nonconformities(text))

        # --- Matching logic ---
        cover_count = len([nc for nc in all_concession_ncs if nc.startswith("NONCONFORMITY") and nc != "NONCONFORMITY"])
        sketch_count = len([nc for nc in all_sketch_ncs if nc.startswith("NONCONFORMITY") and nc != "NONCONFORMITY"])

        # Case 1: Cover has only 1 NC, sketches can have generic
        if cover_count == 1:
            if sketch_count == 1 or "NONCONFORMITY" in all_sketch_ncs:
                print("✅ Matching (Single NC case).")
                self.update_status_cell(5, "Matching", (0, 1, 0, 0.6))
                return

        # Case 2: Cover has multiple NCs, sketches must match count
        if cover_count > 1:
            if cover_count == sketch_count:
                print("✅ Matching (Multiple NC case).")
                self.update_status_cell(5, "Matching", (0, 1, 0, 0.6))
                return

        # ❌ Not matching
        print(f"❌ Not Matching: Cover NCs = {cover_count}, Sketch NCs = {sketch_count}")
        self.update_status_cell(5, "Not Matching", (1, 0, 0, 0.7))

    # Before checkpoint 7 processing
        self.db_path = self.db_input.text  # get current DB path
        print("Using DB path:", self.db_path)
            

    # Checkpoint 7
    # def process_pdf_for_checkpoint_7(self, pdf_path):
    #     try:
    #         doc = fitz.open(pdf_path)
    #     except Exception as e:
    #         print(f"❌ Failed to open PDF: {e}")
    #         self.update_status_cell(6, "Not Extracted", (1, 0, 0, 0.7))
    #         return

    #     extracted_data = []

    #     for page_num, page in enumerate(doc):
    #         text = page.get_text()

    #         if "NON CONFORMITY: 1X" not in text.upper():
    #             continue

    #         lines = text.splitlines()
    #         current_section = None
    #         nominal_code = ""
    #         actual_fasteners = ""

    #         for line in lines:
    #             line_upper = line.strip().upper()

    #             if line_upper.startswith("NOMINAL"):
    #                 current_section = "NOMINAL"
    #             elif line_upper.startswith("ACTUAL"):
    #                 current_section = "ACTUAL"
    #             elif "HOLE CODE" in line_upper and current_section == "NOMINAL":
    #                 nominal_code = line.split(":", 1)[-1].strip()
    #             elif "FASTENERS" in line_upper and current_section == "ACTUAL":
    #                 actual_fasteners = line.split(":", 1)[-1].strip().split("+")[0].strip()

    #         if nominal_code and actual_fasteners:
    #             extracted_data.append({
    #                 "HOLE CODE": nominal_code,
    #                 "FASTENERS": actual_fasteners
    #             })

    #     if extracted_data:
    #         print("\n✅ Checkpoint 7 - Extracted Nominal HOLE CODE & Actual FASTENERS:")
    #         for i, data in enumerate(extracted_data, start=1):
    #             print(f"\nNonConformity {i}:")
    #             print(f"  Nominal HOLE CODE: {data['HOLE CODE']}")
    #             print(f"  Actual FASTENERS: {data['FASTENERS']}")

    #         self.update_status_cell(6, "Extracted", (0, 1, 0, 0.6))
    #     else:
    #         self.update_status_cell(6, "Not Extracted", (1, 0, 0, 0.7))

    # def extract_filename_display_text(self, filepath):
    #     filename = os.path.basename(filepath)
    #     parts = filename.split("-")
    #     if len(parts) >= 2:
    #         return "-".join(parts[:2])
    #     return filename.split(".pdf")[0]

    # def switch_to_main(self):
    #     self.main_root.clear_widgets()
    #     self.main_root.add_widget(self.build_main_screen())

    #     # ✅ Restore previous results if any
    #     if self.last_uploaded_pdf_path:
    #         self.process_pdf_for_checkpoint_1(self.last_uploaded_pdf_path)
    #         self.process_pdf_for_checkpoint_2(self.last_uploaded_pdf_path)
    #         self.process_pdf_for_checkpoint_3(self.last_uploaded_pdf_path)
    #         self.process_pdf_for_checkpoint_4(self.last_uploaded_pdf_path)
    #         self.process_pdf_for_checkpoint_5(self.last_uploaded_pdf_path)
    #         self.process_pdf_for_checkpoint_6(self.last_uploaded_pdf_path)
    #         self.process_pdf_for_checkpoint_7(self.last_uploaded_pdf_path)

    #         self.file_label.text = self.extract_filename_display_text(self.last_uploaded_pdf_path)

    # def clear_results(self, instance):
    #     self.last_uploaded_pdf_path = None
    #     for status_cell in self.status_cells:
    #         status_cell.label.text = ""

    #         # ✅ Set background to fully transparent white
    #         if hasattr(status_cell, 'bg_color_instruction'):
    #             status_cell.bg_color_instruction.rgba = (1, 1, 1, 0)  # transparent

    #         # ✅ Also reset bg_rect size/position to trigger refresh
    #         if hasattr(status_cell, 'bg_rect'):
    #             status_cell.bg_rect.size = status_cell.size
    #             status_cell.bg_rect.pos = status_cell.pos

    #         status_cell.bind(pos=status_cell.update_graphics, size=status_cell.update_graphics)

    #     self.file_label.text = ""


    # Checkpoint 7
    def process_pdf_for_checkpoint_7(self, pdf_path):
        import sqlite3, fitz, re

        def clean_fastener_text(text):
            tokens = re.findall(r"[A-Z0-9\-]+", text.upper())
            blacklist = {"FASTENERS", "POSITION", "EDGE", "AND", "PITCH", "DISTANCE", "AFFECTED", "PARTS"}
            return [t for t in tokens if t not in blacklist and len(t) >= 3]

        try:
            doc = fitz.open(pdf_path)
        except Exception as e:
            print(f"❌ Failed to open PDF: {e}")
            self.update_status_cell(6, "Not Extracted", (1, 0, 0, 0.7))
            return

        db_path = getattr(self, "db_input", None)
        if db_path:
            db_path = db_path.text.strip()
        else:
            db_path = r"C:\Users\shashank.aswath\Documents\airbus_t100\details.db"

        nominal_data, actual_data = [], []

        for page in doc:
            lines = page.get_text().splitlines()
            current_section = None
            nominal_info = {"HOLE CODE": "", "FASTENERS": ""}
            actual_info = {"HOLE CODE": "", "FASTENERS": ""}

            for line in lines:
                line_upper = line.strip().upper()

                if line_upper.startswith("NONCONFORMITY") and (":" in line_upper):
                    if nominal_info["HOLE CODE"] and actual_info["HOLE CODE"]:
                        nominal_data.append(nominal_info.copy())
                        actual_data.append(actual_info.copy())
                        nominal_info = {"HOLE CODE": "", "FASTENERS": ""}
                        actual_info = {"HOLE CODE": "", "FASTENERS": ""}

                # Handle both NOMINAL/ACTUAL and SHOULD BE / IS
                if line_upper.startswith("NOMINAL") or line_upper.startswith("SHOULD BE"):
                    current_section = "NOMINAL"
                elif line_upper.startswith("ACTUAL") or line_upper.startswith("IS"):
                    current_section = "ACTUAL"

                elif "HOLE CODE" in line_upper:
                    code = line.split(":", 1)[-1].strip().upper()
                    if current_section == "NOMINAL":
                        nominal_info["HOLE CODE"] = code
                    elif current_section == "ACTUAL":
                        actual_info["HOLE CODE"] = code

                elif line_upper.startswith("FASTENERS:"):
                    # take only first line after "FASTENERS:"
                    fasteners_full = line.split(":", 1)[-1].strip()

                    # split by + and clean each part
                    cleaned = []
                    for part in fasteners_full.split("+"):
                        token = part.strip().upper()
                        if token and token not in {"FASTENERS", "POSITION", "EDGE", "AND", "PITCH", "DISTANCE", "AFFECTED", "PARTS", "OVERSIZE", "INSTALLED"}:
                            cleaned.append(token)

                    fasteners = " + ".join(cleaned)

                    if current_section == "NOMINAL" and not nominal_info["FASTENERS"]:
                        nominal_info["FASTENERS"] = fasteners
                    elif current_section == "ACTUAL" and not actual_info["FASTENERS"]:
                        actual_info["FASTENERS"] = fasteners

            if nominal_info["HOLE CODE"] and actual_info["HOLE CODE"]:
                nominal_data.append(nominal_info.copy())
                actual_data.append(actual_info.copy())

        if not nominal_data:
            print("❌ No Nominal/Actual data found.")
            self.update_status_cell(6, "Not Extracted", (1, 0, 0, 0.7))
            return

        print("\n✅ Checkpoint 7 - Extracted Nominal and Actual Info:")
        for i, (n, a) in enumerate(zip(nominal_data, actual_data), 1):
            print(f"\nNonConformity {i}:")
            print(f"  Nominal HOLE CODE: {n['HOLE CODE']}")
            print(f"  Nominal FASTENERS: {n['FASTENERS']}")
            print(f"  Actual HOLE CODE: {a['HOLE CODE']}")
            print(f"  Actual FASTENERS: {a['FASTENERS']}")

        self.update_status_cell(6, "Extracted", (0, 1, 0, 0.6))

        # ✅ DB Matching (Nominal and Actual must each exist in DB with matching fasteners)
        matched = False
        try:
            conn = sqlite3.connect(db_path)
            cur = conn.cursor()

            for n, a in zip(nominal_data, actual_data):
                nominal_hole = n['HOLE CODE'].strip().upper()
                actual_hole = a['HOLE CODE'].strip().upper()

                # 👉 Take only the FIRST fastener
                nominal_first = n['FASTENERS'].split("+")[0].strip().upper() if n['FASTENERS'] else ""
                actual_first = a['FASTENERS'].split("+")[0].strip().upper() if a['FASTENERS'] else ""

                # --- 🔍 Check Nominal ---
                cur.execute("SELECT [Fastener 1], [Fastener 2] FROM Sheet1 WHERE [Hole code1] = ?", (nominal_hole,))
                row_nominal = cur.fetchone()
                nominal_ok = False
                if row_nominal:
                    db_fasteners_nominal = {f.strip().upper() for f in row_nominal if f}
                    nominal_ok = nominal_first in db_fasteners_nominal

                # --- 🔍 Check Actual ---
                cur.execute("SELECT [Fastener 1], [Fastener 2] FROM Sheet1 WHERE [Hole code1] = ?", (actual_hole,))
                row_actual = cur.fetchone()
                actual_ok = False
                if row_actual:
                    db_fasteners_actual = {f.strip().upper() for f in row_actual if f}
                    actual_ok = actual_first in db_fasteners_actual

                # ✅ Final Decision
                if nominal_ok and actual_ok:
                    print(f"   ✅ Matched → "
                        f"Nominal({nominal_hole}, {nominal_first}, DB={db_fasteners_nominal}) | "
                        f"Actual({actual_hole}, {actual_first}, DB={db_fasteners_actual})")
                    matched = True
                else:
                    print(f"   ❌ Not Matched → "
                        f"Nominal({nominal_hole}, {nominal_first}, DB={row_nominal}) | "
                        f"Actual({actual_hole}, {actual_first}, DB={row_actual})")

            conn.close()
        except Exception as e:
            print(f"❌ DB Error: {e}")

        self.update_status_cell(
            6,
            "Matched" if matched else "Not Matched",
            (0, 1, 0, 0.6) if matched else (1, 0, 0, 0.7)
        )

    def extract_filename_display_text(self, filepath):
        filename = os.path.basename(filepath)
        parts = filename.split("-")
        if len(parts) >= 2:
            return "-".join(parts[:2])
        return filename.split(".pdf")[0]

    def switch_to_main(self):
        self.main_root.clear_widgets()
        self.main_root.add_widget(self.build_main_screen())

        if self.last_uploaded_pdf_path:
            self.process_pdf_for_checkpoint_1(self.last_uploaded_pdf_path)
            self.process_pdf_for_checkpoint_2(self.last_uploaded_pdf_path)
            self.process_pdf_for_checkpoint_3(self.last_uploaded_pdf_path)
            self.process_pdf_for_checkpoint_4(self.last_uploaded_pdf_path)
            self.process_pdf_for_checkpoint_5(self.last_uploaded_pdf_path)
            self.process_pdf_for_checkpoint_6(self.last_uploaded_pdf_path)
            self.process_pdf_for_checkpoint_7(self.last_uploaded_pdf_path)
            self.process_pdf_for_checkpoint_8(self.last_uploaded_pdf_path)
            self.process_pdf_for_checkpoint_9(self.last_uploaded_pdf_path)

            self.file_label.text = self.extract_filename_display_text(self.last_uploaded_pdf_path)

    def clear_results(self, instance):
        self.last_uploaded_pdf_path = None
        for status_cell in self.status_cells:
            status_cell.label.text = ""

            if hasattr(status_cell, 'bg_color_instruction'):
                status_cell.bg_color_instruction.rgba = (1, 1, 1, 0)

            if hasattr(status_cell, 'bg_rect'):
                status_cell.bg_rect.size = status_cell.size
                status_cell.bg_rect.pos = status_cell.pos

            status_cell.bind(pos=status_cell.update_graphics, size=status_cell.update_graphics)

        self.file_label.text = ""





    def process_pdf_for_checkpoint_8(self, pdf_path):
        import sqlite3, fitz, re

        def extract_type(code):
            """
            Extract base type of fastener/collar by removing numeric suffixes.
            Examples:
                EN6115B3E -> EN6115B
                ABS1738K3 -> ABS1738K
            """
            if not code:
                return ""
            match = re.match(r"([A-Z]+\d+[A-Z]?)", code.strip().upper())
            return match.group(1) if match else code.strip().upper()

        try:
            doc = fitz.open(pdf_path)
        except Exception as e:
            print(f"❌ Failed to open PDF: {e}")
            self.update_status_cell(8, "Not Extracted", (1, 0, 0, 0.7))
            return

        db_path = getattr(self, "db_input", None)
        if db_path:
            db_path = db_path.text.strip()
        else:
            db_path =  r"C:\Users\shashank.aswath\Documents\airbus_t100\details.db"

        nominal_data, actual_data = [], []

        # --- Step 1: Extract data from PDF ---
        for page in doc:
            lines = page.get_text().splitlines()
            current_section = None
            nominal_info = {"HOLE CODE": "", "FASTENER": "", "COLLAR": ""}
            actual_info = {"HOLE CODE": "", "FASTENER": "", "COLLAR": ""}

            for i, line in enumerate(lines):
                line_upper = line.strip().upper()

                if line_upper.startswith("NONCONFORMITY") and ":" in line_upper:
                    if nominal_info["HOLE CODE"] and actual_info["HOLE CODE"]:
                        nominal_data.append(nominal_info.copy())
                        actual_data.append(actual_info.copy())
                        nominal_info = {"HOLE CODE": "", "FASTENER": "", "COLLAR": ""}
                        actual_info = {"HOLE CODE": "", "FASTENER": "", "COLLAR": ""}

                # Determine current section
                if line_upper.startswith("NOMINAL") or line_upper.startswith("SHOULD BE"):
                    current_section = "NOMINAL"
                elif line_upper.startswith("ACTUAL") or line_upper.startswith("IS"):
                    current_section = "ACTUAL"

                # Extract hole codes
                elif "HOLE CODE" in line_upper:
                    code = line.split(":", 1)[-1].strip().upper()
                    if current_section == "NOMINAL":
                        nominal_info["HOLE CODE"] = code
                    elif current_section == "ACTUAL":
                        actual_info["HOLE CODE"] = code

                # Extract first fastener
                elif "FASTENERS" in line_upper:
                    fastener = line.split(":", 1)[-1].split("+")[0].strip().upper()
                    if current_section == "NOMINAL":
                        nominal_info["FASTENER"] = fastener
                    elif current_section == "ACTUAL":
                        actual_info["FASTENER"] = fastener

                # Extract first collar (robust)
                elif "COLLAR" in line_upper:
                    if ":" in line:
                        collar_text = line.split(":", 1)[-1].strip()
                    else:
                        # fallback: take next line if exists
                        if i + 1 < len(lines):
                            collar_text = lines[i + 1].strip()
                        else:
                            collar_text = line.strip()
                    collar = collar_text.split("+")[0].upper()
                    if current_section == "NOMINAL":
                        nominal_info["COLLAR"] = collar
                    elif current_section == "ACTUAL":
                        actual_info["COLLAR"] = collar

            if nominal_info["HOLE CODE"] and actual_info["HOLE CODE"]:
                nominal_data.append(nominal_info.copy())
                actual_data.append(actual_info.copy())

        if not nominal_data:
            print("❌ No Nominal/Actual data found.")
            self.update_status_cell(8, "Not Extracted", (1, 0, 0, 0.7))
            return

        print("\n✅ Checkpoint 8 - Extracted Nominal and Actual Info:")

        # --- Step 2: Match with DB ---
        matched = False
        try:
            conn = sqlite3.connect(db_path)
            cur = conn.cursor()

            for idx, (n, a) in enumerate(zip(nominal_data, actual_data), 1):
                nominal_fastener_type = extract_type(n['FASTENER'])
                nominal_collar_type = extract_type(n['COLLAR'])
                actual_fastener_type = extract_type(a['FASTENER'])
                actual_collar_type = extract_type(a['COLLAR'])

                # Fetch DB row for nominal
                cur.execute("""SELECT [Fastener 1], [Fastener 2],
                                    [Collar 1], [Collar 2], [Collar 3],
                                    [Collar 4], [Collar 5], [Collar 6]
                            FROM Sheet1 WHERE [Hole code1] = ?""", (n['HOLE CODE'],))
                row_nominal = cur.fetchone()
                db_fasteners_nominal = set(extract_type(f) for f in row_nominal[:2] if f) if row_nominal else set()
                db_collars_nominal = set(extract_type(c) for c in row_nominal[2:] if c) if row_nominal else set()
                nominal_ok = (nominal_fastener_type in db_fasteners_nominal and
                            (nominal_collar_type in db_collars_nominal if nominal_collar_type else True))

                # Fetch DB row for actual
                cur.execute("""SELECT [Fastener 1], [Fastener 2],
                                    [Collar 1], [Collar 2], [Collar 3],
                                    [Collar 4], [Collar 5], [Collar 6]
                            FROM Sheet1 WHERE [Hole code1] = ?""", (a['HOLE CODE'],))
                row_actual = cur.fetchone()
                db_fasteners_actual = set(extract_type(f) for f in row_actual[:2] if f) if row_actual else set()
                db_collars_actual = set(extract_type(c) for c in row_actual[2:] if c) if row_actual else set()
                actual_ok = (actual_fastener_type in db_fasteners_actual and
                            (actual_collar_type in db_collars_actual if actual_collar_type else True))

                # Print info like Checkpoint 7
                print(f"\nNonConformity {idx}:")
                print(f"  Nominal HOLE CODE: {n['HOLE CODE']}")
                print(f"  Nominal FASTENER: {n['FASTENER']} (Type={nominal_fastener_type})")
                print(f"  Nominal COLLAR: {n['COLLAR']} (Type={nominal_collar_type})")
                print(f"  Actual HOLE CODE: {a['HOLE CODE']}")
                print(f"  Actual FASTENER: {a['FASTENER']} (Type={actual_fastener_type})")
                print(f"  Actual COLLAR: {a['COLLAR']} (Type={actual_collar_type})")

                if nominal_ok and actual_ok:
                    print(f"   ✅ Matched → Nominal({n['HOLE CODE']}, {nominal_fastener_type}, DB={db_fasteners_nominal}) "
                        f"| Actual({a['HOLE CODE']}, {actual_fastener_type}, DB={db_fasteners_actual})")
                    matched = True
                else:
                    print(f"   ❌ Not Matched → Nominal({n['HOLE CODE']}, {nominal_fastener_type}, DB={db_fasteners_nominal}) "
                        f"| Actual({a['HOLE CODE']}, {actual_fastener_type}, DB={db_fasteners_actual})")

            conn.close()
        except Exception as e:
            print(f"❌ DB Error: {e}")

        self.update_status_cell(
            8,
            "Matched" if matched else "Not Matched",
            (0, 1, 0, 0.6) if matched else (1, 0, 0, 0.7)
        )




    # Checkpoint 9 
    def process_pdf_for_checkpoint_9(self, pdf_path):
        import sqlite3
        import fitz
        import os
        import re

        def clean_fastener_text(text):
            """Extract only valid fastener/collar codes, ignoring unwanted words."""
            tokens = re.findall(r"[A-Z0-9\-]+", text.upper())
            blacklist = {"FASTENERS", "POSITION", "EDGE", "AND", "PITCH", "DISTANCE",
                        "AFFECTED", "PARTS", "OVERSIZE", "INSTALLED"}
            return [t for t in tokens if t not in blacklist and len(t) >= 3]

        # 1️⃣ Open PDF
        try:
            doc = fitz.open(pdf_path)
        except Exception as e:
            print(f"❌ Failed to open PDF: {e}")
            return

        # 2️⃣ Get DB path
        db_path = getattr(self, "db_input", None)
        if db_path:
            db_path = db_path.text.strip()
        else:
            db_path = r"C:\Users\shashank.aswath\Downloads\temp\temp\Hole codes and fasteners.db"

        nominal_data = []
        actual_data = []

        # 3️⃣ Extract text from PDF
        for page_num, page in enumerate(doc):
            lines = page.get_text().splitlines()
            current_section = None
            nominal_info = {"HOLE CODE": "", "FASTENERS_COLLARS": ""}
            actual_info = {"HOLE CODE": "", "FASTENERS_COLLARS": ""}

            for line in lines:
                line_upper = line.strip().upper()

                # ✅ New nonconformity
                if line_upper.startswith("NONCONFORMITY") and (":" in line_upper):
                    if nominal_info["HOLE CODE"] and actual_info["HOLE CODE"]:
                        nominal_data.append(nominal_info.copy())
                        actual_data.append(actual_info.copy())
                        nominal_info = {"HOLE CODE": "", "FASTENERS_COLLARS": ""}
                        actual_info = {"HOLE CODE": "", "FASTENERS_COLLARS": ""}

                # Detect current section
                if line_upper.startswith("NOMINAL") or line_upper.startswith("SHOULD BE"):
                    current_section = "NOMINAL"
                elif line_upper.startswith("ACTUAL") or line_upper.startswith("IS"):
                    current_section = "ACTUAL"
                elif "HOLE CODE" in line_upper:
                    code = line.split(":", 1)[-1].strip().upper()
                    if current_section == "NOMINAL":
                        nominal_info["HOLE CODE"] = code
                    elif current_section == "ACTUAL":
                        actual_info["HOLE CODE"] = code
                elif line_upper.startswith("FASTENERS:"):
                    raw = line.split(":", 1)[-1].strip()
                    cleaned = clean_fastener_text(raw)
                    fasteners = " + ".join(cleaned)
                    if current_section == "NOMINAL" and not nominal_info["FASTENERS_COLLARS"]:
                        nominal_info["FASTENERS_COLLARS"] = fasteners
                    elif current_section == "ACTUAL" and not actual_info["FASTENERS_COLLARS"]:
                        actual_info["FASTENERS_COLLARS"] = fasteners

            # Append last nonconformity
            if nominal_info["HOLE CODE"] and actual_info["HOLE CODE"]:
                nominal_data.append(nominal_info.copy())
                actual_data.append(actual_info.copy())

        # 4️⃣ Print extracted
        if not nominal_data:
            print("❌ No Nominal/Actual data found.")
            return

        print("\n✅ Checkpoint 9 - Extracted Nominal and Actual Info:")
        for i, (n, a) in enumerate(zip(nominal_data, actual_data), start=1):
            print(f"\nNonConformity {i}:")
            print(f"  Nominal HOLE CODE: {n['HOLE CODE']}")
            print(f"  Nominal FASTENERS+COLLARS: {n['FASTENERS_COLLARS']}")
            print(f"  Actual HOLE CODE: {a['HOLE CODE']}")
            print(f"  Actual FASTENERS+COLLARS: {a['FASTENERS_COLLARS']}")

        self.update_status_cell(8, "Extracted", (0, 1, 0, 0.6))

        # 5️⃣ DB Matching (check both nominal and actual pairs against nominal DB entry)
        matched = False
        try:
            conn = sqlite3.connect(db_path)
            cur = conn.cursor()

            for n, a in zip(nominal_data, actual_data):
                cur.execute("""
                    SELECT [Fastener 1], [Fastener 2],
                           [Collar 1], [Collar 2], [Collar3], [Collar4], [Collar 5], [Collar 6]
                    FROM Sheet1
                    WHERE [Hole code1] = ?
                """, (n['HOLE CODE'],))
                row = cur.fetchone()
                print(f"\n🔍 Query for Nominal HOLE CODE {n['HOLE CODE']} returned: {row}")

                if row:
                    fasteners_db = [row[0], row[1]]
                    collars_db = [c for c in row[2:] if c]  # remove None values

                    # ---- Split nominal ----
                    n_parts = [p.strip() for p in n['FASTENERS_COLLARS'].split("+")]
                    if len(n_parts) == 2:
                        nominal_fastener, nominal_collar = n_parts
                    else:
                        print("   ❌ Invalid nominal fastener+collar format")
                        continue

                    # ---- Split actual ----
                    a_parts = [p.strip() for p in a['FASTENERS_COLLARS'].split("+")]
                    if len(a_parts) == 2:
                        actual_fastener, actual_collar = a_parts
                    else:
                        print("   ❌ Invalid actual fastener+collar format")
                        continue

                    print(f"   Checking Nominal Fastener: {nominal_fastener}, Nominal Collar: {nominal_collar}")
                    print(f"   Checking Actual Fastener: {actual_fastener}, Actual Collar: {actual_collar}")
                    print(f"   DB Fasteners: {fasteners_db}")
                    print(f"   DB Collars: {collars_db}")

                    # ---- Final check: both nominal & actual must exist ----
                    if (nominal_fastener in fasteners_db and nominal_collar in collars_db and
                        actual_fastener in fasteners_db and actual_collar in collars_db):
                        print("   ✅ Both Nominal and Actual Matched in DB")
                        matched = True
                    else:
                        print("   ❌ Not Matched in DB")
                else:
                    print("   ❌ HOLE CODE not found in DB")

            conn.close()
        except Exception as e:
            print(f"❌ DB Error: {e}")


        # 6️⃣ Update UI
        self.update_status_cell(8, "Matched" if matched else "Not Matched",
                                (0, 1, 0, 0.6) if matched else (1, 0, 0, 0.7))

if __name__ == "__main__":
    try:
        FlexibleKivyApp().run()
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
