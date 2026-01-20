import os
import time
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.uix.scrollview import ScrollView
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.filechooser import FileChooserIconView
from kivy.uix.popup import Popup
from kivy.core.window import Window
from kivy.metrics import dp
from kivy.utils import platform

# ✅ ANDROID STORAGE (SAFE)
from androidstorage4kivy import SharedStorage

# --- ANDROID PERMISSIONS ---
def request_android_permissions():
    if platform == "android":
        from android.permissions import request_permissions, Permission
        request_permissions([
            Permission.READ_EXTERNAL_STORAGE,
            Permission.WRITE_EXTERNAL_STORAGE,
            Permission.READ_MEDIA_DOCUMENTS
        ])

request_android_permissions()

# --- STORAGE PATHS ---
shared_storage = SharedStorage()
DOCUMENTS_DIR = shared_storage.get_documents_dir()
DOWNLOADS_DIR = shared_storage.get_downloads_dir()

# --- VISUAL THEME ---
COLOR_BG = (0.1, 0.1, 0.1, 1)
COLOR_CARD = (0.18, 0.18, 0.18, 1)
COLOR_ACCENT = (0, 0.8, 0.4, 1)
COLOR_TEXT = (1, 1, 1, 1)
COLOR_SUBTEXT = (0.7, 0.7, 0.7, 1)

# --- CONFIGURATION ---
GLOBAL_SETTINGS = {
    "exchange_rate": 36.0,
    "shipping_rate": 50000.0,
    "margin_percent": 0.0
}

SESSION_STATE = {
    "data": [],
    "filepath": None,
    "header_row": 1,
    "col_map": {},
    "total_investment": 0
}

# --- LOGIC ENGINE ---
def process_excel_preserve_images(filepath):
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        sheet = wb.active

        header_row_index = -1
        col_map = {}

        for r in range(1, 20):
            row_values = [str(cell.value) if cell.value else "" for cell in sheet[r]]
            if "ITEM" in row_values or "Price(RMB)" in row_values:
                header_row_index = r
                for idx, val in enumerate(row_values):
                    col_map[val] = idx + 1
                break

        if header_row_index == -1:
            return None, "Header not found"

        SESSION_STATE["filepath"] = filepath
        SESSION_STATE["header_row"] = header_row_index
        SESSION_STATE["col_map"] = col_map

        results = []
        grand_total = 0

        for r in range(header_row_index + 1, sheet.max_row + 1):
            try:
                def get_val(name):
                    return sheet.cell(row=r, column=col_map.get(name, 0)).value or 0

                item_name = str(get_val("ITEM") or "Unknown")
                rmb_price = float(get_val("Price(RMB)"))
                boxes = float(get_val("Ctn"))
                qty = float(get_val("Qty") or 1)
                cbm = float(get_val("CBM"))

                if boxes == 0:
                    continue

                ship_box = cbm * GLOBAL_SETTINGS["shipping_rate"]
                ship_unit = ship_box / qty
                base_dzd = rmb_price * GLOBAL_SETTINGS["exchange_rate"]
                unit_cost = base_dzd + ship_unit

                total_line = unit_cost * boxes * qty
                grand_total += total_line

                results.append({
                    "row_index": r,
                    "name": item_name,
                    "unit_cost": round(unit_cost, 2),
                    "total_line": round(total_line, 2),
                    "rmb_price": rmb_price,
                    "qty": int(boxes * qty)
                })

            except:
                continue

        SESSION_STATE["total_investment"] = grand_total
        return results, "OK"

    except Exception as e:
        return None, str(e)

# --- EXPORT TO DOWNLOADS ---
def export_results_smart():
    try:
        wb = openpyxl.load_workbook(SESSION_STATE["filepath"])
        sheet = wb.active

        target_col = SESSION_STATE["col_map"].get("Price(RMB)", 5) + 1
        header = sheet.cell(row=SESSION_STATE["header_row"], column=target_col)
        header.value = "Unit Cost (DZD)"
        header.font = Font(bold=True)

        for item in SESSION_STATE["data"]:
            sheet.cell(row=item["row_index"], column=target_col).value = item["unit_cost"]

        filename = f"CostSheet_{int(time.time())}.xlsx"
        output_path = os.path.join(DOWNLOADS_DIR, filename)
        wb.save(output_path)

        return True, output_path

    except Exception as e:
        return False, str(e)

# --- UI SCREENS (UNCHANGED LOGIC) ---
# Only file chooser path changed

class HomeScreen(Screen):
    def show_file_chooser(self, instance):
        content = BoxLayout(orientation='vertical')
        chooser = FileChooserIconView(path=DOCUMENTS_DIR, filters=['*.xlsx'])

        btn_load = Button(text="Load")
        btn_cancel = Button(text="Cancel")

        box = BoxLayout(size_hint_y=0.1)
        box.add_widget(btn_cancel)
        box.add_widget(btn_load)

        content.add_widget(chooser)
        content.add_widget(box)

        popup = Popup(title="Select Excel", content=content, size_hint=(0.9, 0.9))

        def load(_):
            if chooser.selection:
                data, _ = process_excel_preserve_images(chooser.selection[0])
                SESSION_STATE["data"] = data
                self.manager.get_screen("results").load_data()
                self.manager.current = "results"
                popup.dismiss()

        btn_load.bind(on_press=load)
        btn_cancel.bind(on_press=popup.dismiss)
        popup.open()

# (ResultsScreen, SettingsScreen, InfoCard etc stay EXACTLY the same as yours)

class ImportApp(App):
    def build(self):
        Window.clearcolor = COLOR_BG
        sm = ScreenManager()
        sm.add_widget(HomeScreen(name="home"))
        sm.add_widget(SettingsScreen(name="settings"))
        sm.add_widget(ResultsScreen(name="results"))
        return sm

if __name__ == "__main__":
    ImportApp().run()
