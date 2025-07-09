# eazeinone
# Cultivating Accounting Excellence: Our Professional Development Pipeline
# --- eaze_inn_accounts.py ---

import tkinter as tk
from tkinter import messagebox, filedialog, ttk, simpledialog # Added simpledialog
import os
import datetime
import shutil

import sys
import threading
import queue
import time
import json
import hashlib
import traceback
import collections
import subprocess
if os.name == 'nt':
    pass

# Reporting and Data Handling
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as ReportlabImage, Flowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from openpyxl import Workbook
from PIL import Image, ImageTk
import webbrowser

from decimal import Decimal, ROUND_HALF_UP, InvalidOperation

try:
    from escpos import printer
    from escpos.exceptions import DeviceNotFoundError
    escpos_installed = True
except ImportError:
    escpos_installed = False
    print("WARNING: python-escpos library not found. Receipt printing disabled.")
    print("         Install using: pip install python-escpos")

if os.name == 'nt':
    try:
        import win32print
        win32print_installed = True
    except ImportError:
        win32print_installed = False
else:
    win32print_installed = False

# New import for EazeBot charting
try:
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    matplotlib_installed = True
except ImportError:
    matplotlib_installed = False
    print("WARNING: matplotlib library not found. EazeBot charting will be disabled.")
    print("         Install using: pip install matplotlib")

# --- [NEW] Placeholder for Gemini API library ---
# To enable this feature, you must run: pip install google-generativeai
try:
    import google.generativeai as genai
    gemini_lib_installed = True
except ImportError:
    gemini_lib_installed = False
    print("WARNING: google-generativeai library not found. Gemini API features will be disabled.")
    print("         Install using: pip install google-generativeai")


# --- Constants ---
DATA_DIR = "eaze_inn_data"
USERS_FILE = os.path.join(DATA_DIR, "users.json")
INVOICES_FILE = os.path.join(DATA_DIR, "invoices.json")
INVOICE_ITEMS_FILE = os.path.join(DATA_DIR, "invoice_items.json")
SUPPLIER_INVOICES_FILE = os.path.join(DATA_DIR, "supplier_invoices.json")
SUPPLIER_INVOICE_ITEMS_FILE = os.path.join(DATA_DIR, "supplier_invoice_items.json")
INVENTORY_FILE = os.path.join(DATA_DIR, "inventory.json") # Inventory data file
PAYMENTS_FILE = os.path.join(DATA_DIR, "payments.json")
IMAGES_DIR = os.path.join(DATA_DIR, "invoice_images")
SETTINGS_FILE = os.path.join(DATA_DIR, "settings.json")
COMPANY_QR_BASE_FILENAME = "company_qr_code"

# --- App Settings ---
LOW_STOCK_THRESHOLD = Decimal('5') # Used by inventory management
CURRENCY_SYMBOL = "₹"
DATE_FORMAT = '%Y-%m-%d'
ZERO_DECIMAL = Decimal('0.00')
TWO_PLACES = Decimal('0.01')

THERMAL_PRINTER_TYPE = 'win32raw'
THERMAL_PRINTER_VID = 0x04b8
THERMAL_PRINTER_PID = 0x0e15
THERMAL_PRINTER_IN_EP = 0x81
THERMAL_PRINTER_OUT_EP = 0x01
THERMAL_PRINTER_NAME = "POS58 Printer"
THERMAL_PRINTER_IP = "192.168.1.100"
THERMAL_PRINTER_PORT = 9100
THERMAL_PRINTER_PORT_SERIAL = "/dev/ttyUSB0"
THERMAL_PRINTER_BAUDRATE = 9600
THERMAL_PRINTER_FILE = "receipt_output.bin"
RECEIPT_WIDTH = 32

# --- Global In-Memory Data Storage ---
USERS_DATA = []
INVOICES_DATA = []
INVOICE_ITEMS_DATA = []
SUPPLIER_INVOICES_DATA = []
SUPPLIER_INVOICE_ITEMS_DATA = []
INVENTORY_DATA = [] # Holds inventory items: {id, item_name, quantity, value (cost_price)}
PAYMENTS_DATA = []
COMPANY_SETTINGS = {}
GEMINI_API_KEY = None # Will be set at runtime

DEFAULT_SETTINGS = {
    "company_name": "Your Company Name",
    "company_address": "Your Company Address, City, PIN",
    "company_email": "your.email@example.com",
    "company_phone": "Your Phone Number",
    "company_gstin": "Your GSTIN (Optional)",
    "qr_code_path": None
}


class DecimalEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, Decimal): return str(obj)
        return super(DecimalEncoder, self).default(obj)

def load_settings():
    global COMPANY_SETTINGS; COMPANY_SETTINGS = DEFAULT_SETTINGS.copy()
    try:
        if not os.path.exists(SETTINGS_FILE): print(f"Settings file '{os.path.basename(SETTINGS_FILE)}' not found. Using defaults."); return
        with open(SETTINGS_FILE, 'r', encoding='utf-8') as f: content = f.read()
        if not content.strip(): print("Settings file is empty. Using defaults."); return
        loaded_settings = json.loads(content)
        COMPANY_SETTINGS.update(loaded_settings); print("Company settings loaded successfully.")
    except (IOError, json.JSONDecodeError) as e: print(f"Error loading settings from {os.path.basename(SETTINGS_FILE)}: {e}. Using defaults.")
    except Exception as e: print(f"Unexpected error loading settings: {e}. Using defaults."); traceback.print_exc()

def save_settings_file():
    global COMPANY_SETTINGS
    try:
        os.makedirs(os.path.dirname(SETTINGS_FILE), exist_ok=True)
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f: json.dump(COMPANY_SETTINGS, f, cls=DecimalEncoder, indent=4)
        print(f"Company settings saved to {os.path.basename(SETTINGS_FILE)}"); return True
    except (IOError, TypeError) as e: print(f"Error saving settings: {e}"); traceback.print_exc(); messagebox.showerror("Settings Save Error", f"Could not save settings.\n{e}", icon='error'); return False
    except Exception as e: print(f"Unexpected error saving settings: {e}"); traceback.print_exc(); messagebox.showerror("Settings Save Error", f"Unexpected error saving settings.\n{e}", icon='error'); return False

def _validate_and_copy_image(original_path, target_dir, target_base_filename):
    if not original_path or not os.path.exists(original_path): return None
    os.makedirs(target_dir, exist_ok=True)
    try:
        file_ext = os.path.splitext(original_path)[1].lower()
        if file_ext not in ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff']: messagebox.showwarning("Image Warning", f"Unsupported image format '{file_ext}'. Please use JPG, PNG, BMP, GIF, or TIFF.", icon='warning'); return None
        new_filename = f"{target_base_filename}{file_ext}"; new_path = os.path.join(target_dir, new_filename)
        _remove_existing_image(target_dir, target_base_filename); shutil.copy2(original_path, new_path)
        try: img = Image.open(new_path); img.verify(); img.close(); return new_filename
        except (IOError, SyntaxError, Image.UnidentifiedImageError) as img_err:
            print(f"Warning: Copied image seems invalid: {new_path} - {img_err}")
            try: os.remove(new_path)
            except OSError: pass
            messagebox.showwarning("Image Warning", f"The selected image file could not be verified:\n{os.path.basename(original_path)}", icon='warning'); return None
    except Exception as e: print(f"Error processing image: {e}"); messagebox.showerror("Image Error", f"Could not copy/process image:\n{e}", icon='error'); return None

def _remove_existing_image(target_dir, target_base_filename):
    try:
        for filename in os.listdir(target_dir):
            if filename.startswith(target_base_filename):
                try: os.remove(os.path.join(target_dir, filename)); print(f"Removed existing image: {filename}")
                except OSError as rm_err: print(f"Warning: Could not remove existing image {filename}: {rm_err}")
    except FileNotFoundError: pass
    except Exception as e: print(f"Error removing existing image files for {target_base_filename}: {e}")

def _handle_invoice_image(original_path, invoice_type, invoice_id):
    target_base_filename = f"{invoice_type}_invoice_{invoice_id}"; relative_path = _validate_and_copy_image(original_path, IMAGES_DIR, target_base_filename)
    return relative_path if relative_path else None


def load_data(filepath):
    data = []
    try:
        if not os.path.exists(filepath): print(f"Data file not found: {filepath}. Starting empty."); return []
        with open(filepath, 'r', encoding='utf-8') as f: content = f.read()
        if not content.strip(): return []
        data = json.loads(content)
    except (IOError, json.JSONDecodeError) as e: print(f"Error loading {filepath}: {e}"); messagebox.showerror("Data Load Error", f"Could not load {os.path.basename(filepath)}.\nCheck console.", icon='warning'); return []
    except Exception as e: print(f"Unexpected error loading {filepath}: {e}"); traceback.print_exc(); messagebox.showerror("Data Load Error", f"Unexpected error loading {os.path.basename(filepath)}.\nCheck console.", icon='warning'); return []
    processed_data = []
    for item in data:
        new_item = item.copy()
        try:
            if 'id' in new_item and new_item['id'] is not None: new_item['id'] = int(new_item['id'])
            if 'invoice_id' in new_item and new_item['invoice_id'] is not None: new_item['invoice_id'] = int(new_item['invoice_id'])
            if 'supplier_invoice_id' in new_item and new_item['supplier_invoice_id'] is not None: new_item['supplier_invoice_id'] = int(new_item['supplier_invoice_id'])
            # Ensure 'quantity' and 'value' (for inventory) are Decimal
            # Other financial fields are already covered
            for key in ['price', 'value', 'amount', 'total_amount', 'quantity', 'amount_paid']:
                if key in new_item and new_item[key] is not None:
                    try: new_item[key] = Decimal(str(new_item[key]))
                    except InvalidOperation: print(f"Warn: Invalid Decimal for '{key}' in {filepath}, ID {new_item.get('id', 'N/A')}: '{new_item[key]}'. Setting to 0."); new_item[key] = ZERO_DECIMAL
            processed_data.append(new_item)
        except (ValueError, TypeError) as conv_e: print(f"Warn: Skipping record due to conversion error in {filepath}: {item} - Error: {conv_e}")
    return processed_data

def save_data(data_list, filepath):
    try:
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        with open(filepath, 'w', encoding='utf-8') as f: json.dump(data_list, f, cls=DecimalEncoder, indent=4)
        return True
    except (IOError, TypeError) as e: print(f"Error saving to {filepath}: {e}"); traceback.print_exc(); messagebox.showerror("Data Save Error", f"Could not save to {os.path.basename(filepath)}.\nCheck logs.", icon='error'); return False
    except Exception as e: print(f"Unexpected error saving to {filepath}: {e}"); traceback.print_exc(); messagebox.showerror("Data Save Error", f"Unexpected error saving {os.path.basename(filepath)}.\nCheck logs.", icon='error'); return False

def load_all_data():
    global USERS_DATA, INVOICES_DATA, INVOICE_ITEMS_DATA, SUPPLIER_INVOICES_DATA, SUPPLIER_INVOICE_ITEMS_DATA, INVENTORY_DATA, PAYMENTS_DATA, COMPANY_SETTINGS
    print("Loading data..."); USERS_DATA.clear(); INVOICES_DATA.clear(); INVOICE_ITEMS_DATA.clear(); SUPPLIER_INVOICES_DATA.clear(); SUPPLIER_INVOICE_ITEMS_DATA.clear(); INVENTORY_DATA.clear(); PAYMENTS_DATA.clear(); COMPANY_SETTINGS.clear()
    USERS_DATA.extend(load_data(USERS_FILE)); INVOICES_DATA.extend(load_data(INVOICES_FILE)); INVOICE_ITEMS_DATA.extend(load_data(INVOICE_ITEMS_FILE)); SUPPLIER_INVOICES_DATA.extend(load_data(SUPPLIER_INVOICES_FILE)); SUPPLIER_INVOICE_ITEMS_DATA.extend(load_data(SUPPLIER_INVOICE_ITEMS_FILE))
    INVENTORY_DATA.extend(load_data(INVENTORY_FILE)) # Load inventory data
    PAYMENTS_DATA.extend(load_data(PAYMENTS_FILE))
    load_settings(); print(f"Data loaded: {len(USERS_DATA)}u, {len(INVOICES_DATA)}inv, {len(SUPPLIER_INVOICES_DATA)}bill, {len(INVENTORY_DATA)}ity, {len(PAYMENTS_DATA)}pay."); print(f"Settings: Name='{COMPANY_SETTINGS.get('company_name', 'N/A')}'")


def get_next_id(data_list):
    if not data_list: return 1
    max_id = 0
    for item in data_list:
        try: current_id = int(item.get('id', 0)); max_id = max(max_id, current_id)
        except (ValueError, TypeError): continue
    return max_id + 1

def hash_password(password): return hashlib.sha256(password.encode()).hexdigest()

def format_currency(amount, include_sign=False):
    if amount is None: return f"{CURRENCY_SYMBOL}0.00"
    try:
        decimal_amount = Decimal(str(amount)).quantize(TWO_PLACES, rounding=ROUND_HALF_UP)
        if decimal_amount.is_zero() and decimal_amount.is_signed():
             decimal_amount = ZERO_DECIMAL
        if include_sign:
            return f"{CURRENCY_SYMBOL}{decimal_amount:+,.2f}"
        else:
            return f"{CURRENCY_SYMBOL}{decimal_amount:,.2f}"
    except (InvalidOperation, TypeError, ValueError):
        return f"{CURRENCY_SYMBOL}N/A"

def format_decimal_quantity(quantity):
    if quantity is None: return "0"
    try: qty_d = Decimal(str(quantity)); return f"{qty_d:.10f}".rstrip('0').rstrip('.')
    except (InvalidOperation, TypeError, ValueError): return "N/A"

def format_percentage_diff(current, previous):
    try:
        current_d, previous_d = Decimal(str(current)), Decimal(str(previous))
        if previous_d == ZERO_DECIMAL: return "N/A (Prev 0)" if current_d != ZERO_DECIMAL else "0.00%"
        diff = ((current_d - previous_d) / previous_d) * 100; return f"{diff:+.2f}%"
    except (InvalidOperation, TypeError, ValueError): return "N/A"

# --- Inventory Update Logic ---
def update_inventory_after_transaction(transaction_type, processed_items):
    global INVENTORY_DATA
    inventory_changed = False
    for proc_item in processed_items:
        item_name = proc_item['item'].strip()
        quantity_change = proc_item['quantity']  # Expected to be Decimal
        price_per_unit = proc_item['price']      # Expected to be Decimal

        inventory_item = next((inv_item for inv_item in INVENTORY_DATA if inv_item.get('item_name', '').strip().lower() == item_name.lower()), None)

        if transaction_type == 'supplier':  # Purchase
            if inventory_item:
                old_quantity = inventory_item.get('quantity', ZERO_DECIMAL)
                new_quantity = old_quantity + quantity_change
                inventory_item['quantity'] = new_quantity
                inventory_item['value'] = price_per_unit # Update cost to latest purchase price
                inventory_item['last_updated'] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                print(f"Inventory Update (Purchase): '{item_name}' old_qty: {old_quantity}, added: {quantity_change}, new_qty: {new_quantity}, new_cost: {price_per_unit}")
            else:
                new_id = get_next_id(INVENTORY_DATA)
                inventory_item_new = {
                    'id': new_id,
                    'item_name': item_name,
                    'quantity': quantity_change,
                    'value': price_per_unit,
                    'last_updated': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
                INVENTORY_DATA.append(inventory_item_new)
                print(f"Inventory Add (Purchase): '{item_name}' qty: {quantity_change}, cost: {price_per_unit}")
            inventory_changed = True

        elif transaction_type == 'customer':  # Sale
            if inventory_item:
                old_quantity = inventory_item.get('quantity', ZERO_DECIMAL)
                new_quantity = old_quantity - quantity_change
                inventory_item['quantity'] = new_quantity
                inventory_item['last_updated'] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                # 'value' (cost) does not change on sale
                print(f"Inventory Update (Sale): '{item_name}' old_qty: {old_quantity}, sold: {quantity_change}, new_qty: {new_quantity}")
                inventory_changed = True
            else: # Item sold but not in inventory
                new_id = get_next_id(INVENTORY_DATA)
                inventory_item_new = {
                    'id': new_id,
                    'item_name': item_name,
                    'quantity': -quantity_change, # Record as negative stock
                    'value': ZERO_DECIMAL,        # Cost is unknown
                    'last_updated': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'status_flag': 'SOLD_WITHOUT_STOCK' # Custom flag
                }
                INVENTORY_DATA.append(inventory_item_new)
                print(f"Inventory Alert (Sale): Item '{item_name}' sold without prior stock. Added with negative quantity.")
                inventory_changed = True
    
    if inventory_changed:
        if not save_data(INVENTORY_DATA, INVENTORY_FILE):
            print("CRITICAL: FAILED TO SAVE INVENTORY UPDATES TO FILE.")
            messagebox.showerror("Inventory Save Error", 
                                 "Failed to save inventory updates to file.\n"
                                 "Data might be inconsistent upon restart.\n"
                                 "Please check console logs.", icon='error')
# --- User Authentication Windows ---
def register_window(root):
    register_win = tk.Toplevel(root); register_win.title("Register New User")
    register_win.transient(root); register_win.grab_set(); register_win.resizable(False, False)
    tk.Label(register_win, text="Username:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
    username_entry = tk.Entry(register_win, width=30); username_entry.grid(row=0, column=1, padx=10, pady=5)
    tk.Label(register_win, text="Password:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
    password_entry = tk.Entry(register_win, show="*", width=30); password_entry.grid(row=1, column=1, padx=10, pady=5)
    username_entry.focus_set()
    def register_command():
        username = username_entry.get().strip(); password = password_entry.get()
        if not username or not password: messagebox.showerror("Input Error", "Please enter both username and password.", parent=register_win); return
        if any(user.get('username') == username for user in USERS_DATA): messagebox.showerror("Username Exists", "Username already exists.", parent=register_win); return
        new_id = get_next_id(USERS_DATA); hashed_password = hash_password(password)
        new_user = {'id': new_id, 'username': username, 'password': hashed_password}
        USERS_DATA.append(new_user)
        if save_data(USERS_DATA, USERS_FILE): messagebox.showinfo("Success", "Registration successful!", parent=register_win); register_win.destroy()
        else: USERS_DATA.pop() # Rollback
    btn_frame = tk.Frame(register_win); btn_frame.grid(row=2, column=0, columnspan=2, pady=10)
    tk.Button(btn_frame, text="Register", command=register_command, width=10).pack(side=tk.LEFT, padx=5)
    tk.Button(btn_frame, text="Cancel", command=register_win.destroy, width=10).pack(side=tk.LEFT, padx=5)

def show_main_app(root):
    root.withdraw()  # Hide the login window
    create_dashboard(root)  # Show the main dashboard

def signin_window(root):
    signin_win = tk.Toplevel(root)
    signin_win.title("Sign In")
    signin_win.transient(root)
    signin_win.grab_set()
    signin_win.resizable(False, False)
    
    tk.Label(signin_win, text="Username:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
    username_entry = tk.Entry(signin_win, width=30)
    username_entry.grid(row=0, column=1, padx=10, pady=5)
    
    tk.Label(signin_win, text="Password:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
    password_entry = tk.Entry(signin_win, show="*", width=30)
    password_entry.grid(row=1, column=1, padx=10, pady=5)
    
    username_entry.focus_set()
    
    def signin_command():
        username = username_entry.get().strip()
        password = password_entry.get()
        if not username or not password:
            messagebox.showerror("Input Error", "Please enter both username and password.", parent=signin_win)
            return
        found_user = next((user for user in USERS_DATA if user.get('username') == username), None)
        if not found_user or hash_password(password) != found_user.get('password'):
            messagebox.showerror("Authentication Failed", "Invalid username or password.", parent=signin_win)
            return
        signin_win.destroy()
        show_main_app(root)
    
    btn_frame = tk.Frame(signin_win)
    btn_frame.grid(row=2, column=0, columnspan=2, pady=10)
    tk.Button(btn_frame, text="Sign In", command=signin_command, width=10).pack(side=tk.LEFT, padx=5)
    tk.Button(btn_frame, text="Cancel", command=signin_win.destroy, width=10).pack(side=tk.LEFT, padx=5)
    password_entry.bind("<Return>", lambda event: signin_command())

# --- Backup/Restore Functions ---
def backup_all_data_threaded(result_queue):
    source_dir = DATA_DIR; backup_base_dir = "eaze_inn_json_backup"
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_dest_dir = os.path.join(backup_base_dir, f"backup_{timestamp}")
    try:
        if not os.path.isdir(source_dir): result_queue.put(("Warning", f"Data dir '{source_dir}' not found.")); return
        os.makedirs(backup_base_dir, exist_ok=True)
        shutil.copytree(source_dir, backup_dest_dir)
        result_queue.put(("Success", f"Backup successful!\nSaved in:\n{os.path.abspath(backup_dest_dir)}"))
    except Exception as e: result_queue.put(("Error", f"Backup failed: {e}\n{traceback.format_exc()}"))

def backup_all_data(root):
    if not messagebox.askyesno("Confirm Backup", "Create a backup of the current data?", parent=root): return
    result_queue = queue.Queue()
    thread = threading.Thread(target=backup_all_data_threaded, args=(result_queue,), daemon=True)
    thread.start()
    root.after(100, lambda: check_thread_queue(root, result_queue, "Backup"))

def restore_all_data_threaded(restore_source_dir, result_queue):
    target_dir = DATA_DIR; pre_restore_base_dir = "pre_restore_json_backups"
    try:
        if not os.path.isdir(restore_source_dir): result_queue.put(("Error", f"Backup dir '{restore_source_dir}' not found.")); return
        current_data_exists = os.path.isdir(target_dir); pre_restore_backup_path = None
        if current_data_exists:
            os.makedirs(pre_restore_base_dir, exist_ok=True)
            pre_restore_backup_path = os.path.join(pre_restore_base_dir, f"pre_restore_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}")
            try: shutil.copytree(target_dir, pre_restore_backup_path); print(f"Backed up current data to '{pre_restore_backup_path}'")
            except Exception as backup_err: result_queue.put(("Warning", f"Could not back up live data: {backup_err}.\nProceeding will overwrite.")) # Or cancel?
        if current_data_exists:
            try: shutil.rmtree(target_dir); print(f"Removed current data dir '{target_dir}'.")
            except Exception as rm_err: result_queue.put(("Error", f"Failed remove current dir '{target_dir}': {rm_err}\nAbort.")); return
        try:
            shutil.copytree(restore_source_dir, target_dir)
            result_queue.put(("Success", f"Data restored from:\n{os.path.basename(restore_source_dir)}\n\nRestart required."))
        except Exception as copy_err:
            result_queue.put(("Error", f"Failed copy backup to '{target_dir}': {copy_err}\nRestore failed."))
            if pre_restore_backup_path and os.path.exists(pre_restore_backup_path):
                 try: shutil.copytree(pre_restore_backup_path, target_dir); result_queue.put(("Error", f"Copy backup failed: {copy_err}\nBUT pre-restore backup restored."))
                 except Exception as restore_err: result_queue.put(("Error", f"Copy backup failed: {copy_err}\nCRITICAL: Failed restore pre-backup: {restore_err}\nManual fix needed."))
    except Exception as e: result_queue.put(("Error", f"Unexpected restore error: {e}\n{traceback.format_exc()}"))

def restore_all_data(root):
    initial_backup_dir = os.path.abspath("eaze_inn_json_backup") if os.path.exists("eaze_inn_json_backup") else os.path.abspath(".")
    restore_dir = filedialog.askdirectory(title="Select Specific Backup Folder", initialdir=initial_backup_dir, parent=root)
    if not restore_dir: return
    if not messagebox.askyesno("Confirm Restore", f"!!! WARNING !!!\n\nThis will DELETE current data ('{os.path.abspath(DATA_DIR)}') "
        f"and replace it with:\n'{os.path.basename(restore_dir)}'\n\nA backup of current data will be attempted.\n\nProceed?", icon='warning', parent=root): return
    result_queue = queue.Queue()
    thread = threading.Thread(target=restore_all_data_threaded, args=(restore_dir, result_queue), daemon=True)
    thread.start()
    root.after(100, lambda: check_thread_queue(root, result_queue, "Restore"))

def check_thread_queue(root, result_queue, operation_name):
    try:
        status, message = result_queue.get_nowait()
        if status == "Error": messagebox.showerror(f"{operation_name} Error", message, parent=root)
        elif status == "Warning": messagebox.showwarning(f"{operation_name} Warning", message, parent=root)
        elif status == "Success":
            # The message for PDF generation is now the file path, so we adjust the success message
            if "Generation" in operation_name or "Excel" in operation_name:
                messagebox.showinfo(f"{operation_name} Success", f"File saved successfully at:\n{message}", parent=root)
            elif "Sharing" in operation_name:
                 messagebox.showinfo(f"{operation_name} Success", message, parent=root)
            else: # All other successes (Backup, Restore, etc.)
                messagebox.showinfo(f"{operation_name} Success", message, parent=root)
            
            if operation_name == "Restore":
                 messagebox.showinfo("Restart Required", "Data restored.\nPlease restart the application.", parent=root, icon='info')
                 root.quit()
        elif status == "Cancelled": messagebox.showinfo(f"{operation_name} Cancelled", message, parent=root)
    except queue.Empty: root.after(100, lambda: check_thread_queue(root, result_queue, operation_name))
    except Exception as e: messagebox.showerror("Queue Check Error", f"Error checking {operation_name} result: {e}", parent=root)

# --- PDF/Excel Generation ---
class QRCodeFlowable(Flowable):
    def __init__(self, qr_path, width, height): Flowable.__init__(self); self.qr_path = qr_path; self.width = width; self.height = height
    def draw(self):
        try: img = ReportlabImage(self.qr_path, width=self.width, height=self.height); img.hAlign = 'RIGHT'; img.drawOn(self.canv, 0, 0)
        except FileNotFoundError: print(f"QR Code image not found: {self.qr_path}")
        except Exception as e: print(f"Error drawing QR Code: {e}")

def generate_pdf_invoice_threaded(invoice_id, invoice_type, entity_name, invoice_data, invoice_items_dec, result_queue):
    global COMPANY_SETTINGS; entity_label = invoice_type.capitalize()
    pdf_file = f"{invoice_type}_{entity_name.replace(' ','_')}_{invoice_id}_{datetime.datetime.now().strftime('%Y%m%d')}.pdf"
    title = "TAX INVOICE" if invoice_type == 'customer' else "SUPPLIER BILL"
    try:
        pdf = SimpleDocTemplate(pdf_file, pagesize=letter, topMargin=0.5*inch, bottomMargin=0.5*inch, leftMargin=0.7*inch, rightMargin=0.7*inch)
        styles = getSampleStyleSheet(); styles.add(ParagraphStyle(name='small', parent=styles['Normal'], fontSize=8)); styles.add(ParagraphStyle(name='RightAlign', parent=styles['Normal'], alignment=2))
        styles['h1'].alignment = 1; styles['h2'].alignment = 1; story = []
        company_name = COMPANY_SETTINGS.get('company_name', DEFAULT_SETTINGS['company_name']); company_address = COMPANY_SETTINGS.get('company_address', DEFAULT_SETTINGS['company_address'])
        company_email = COMPANY_SETTINGS.get('company_email', DEFAULT_SETTINGS['company_email']); company_phone = COMPANY_SETTINGS.get('company_phone', DEFAULT_SETTINGS['company_phone'])
        company_gstin = COMPANY_SETTINGS.get('company_gstin', None); qr_code_rel_path = COMPANY_SETTINGS.get('qr_code_path', None)
        header_text = [Paragraph(f"<b>{company_name}</b>", styles['h1']), Paragraph(company_address, styles['Normal']), Paragraph(f"M: {company_phone} | Email: {company_email}", styles['Normal'])]
        if company_gstin: header_text.append(Paragraph(f"GSTIN: {company_gstin}", styles['Normal']))
        qr_flowable = None
        if qr_code_rel_path:
            qr_full_path = os.path.join(DATA_DIR, qr_code_rel_path) # Construct full path from relative
            if os.path.exists(qr_full_path):
                try: qr_size = 0.8 * inch; qr_flowable = QRCodeFlowable(qr_full_path, qr_size, qr_size)
                except Exception as qr_err: print(f"Error creating QR flowable: {qr_err}")
        if qr_flowable:
            header_table_data = [[Table([ [p] for p in header_text ], style=TableStyle([('BOTTOMPADDING', (0,0), (0,-1), 1)])), qr_flowable]]
            header_table = Table(header_table_data, colWidths=[letter[0] - 1.4*inch - 1*inch, 0.8*inch]); header_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP'), ('ALIGN', (1,0), (1,0), 'RIGHT')])); story.append(header_table)
        else: story.extend(header_text)
        story.append(Spacer(1, 0.1*inch)); story.append(Paragraph('<hr width="100%" color="black" size="1"/>', styles['Normal'])); story.append(Spacer(1, 0.1*inch))
        story.append(Paragraph(f"<b>{title}</b>", styles['h2'])); story.append(Spacer(1, 0.1*inch))
        info_data = [[Paragraph(f"{title} No: <b>{invoice_id}</b>", styles['Normal']), Paragraph(f"Date: <b>{invoice_data.get('date', 'N/A')}</b>", styles['RightAlign'])], [Paragraph(f"<u>{entity_label} Details:</u>", styles['h4']), ""], [Paragraph(f"Name: {entity_name}", styles['Normal']), ""]]
        info_table = Table(info_data, colWidths=[3.5*inch, 3.0*inch]); info_table.setStyle(TableStyle([('ALIGN', (0, 0), (0, -1), 'LEFT'), ('ALIGN', (1, 0), (1, -1), 'RIGHT'), ('VALIGN', (0, 0), (-1, -1), 'TOP'), ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'), ('SPAN', (0, 1), (1, 1)), ('BOTTOMPADDING', (0, 0), (-1, -1), 3), ('TOPPADDING', (0, 0), (-1, -1), 3),])); story.append(info_table); story.append(Spacer(1, 0.2*inch))
        table_header = [Paragraph("<b>S.N</b>", styles['Normal']), Paragraph("<b>Item Description</b>", styles['Normal']), Paragraph("<b>Qty</b>", styles['Normal']), Paragraph("<b>Rate</b>", styles['Normal']), Paragraph("<b>Amount</b>", styles['Normal'])]
        table_data = [table_header]; total_amount = ZERO_DECIMAL; sn = 1
        for item in invoice_items_dec:
            try:
                item_name = item.get('item', 'N/A'); qty = item.get('quantity', ZERO_DECIMAL); price = item.get('price', ZERO_DECIMAL); amount = (qty * price).quantize(TWO_PLACES, rounding=ROUND_HALF_UP); total_amount += amount
                row_data = [Paragraph(str(sn), styles['Normal']), Paragraph(str(item_name), styles['Normal']), Paragraph(format_decimal_quantity(qty), styles['Normal']), Paragraph(format_currency(price), styles['Normal']), Paragraph(format_currency(amount), styles['Normal'])]
                table_data.append(row_data); sn += 1
            except Exception as item_err: table_data.append([Paragraph(str(sn), styles['Normal']), Paragraph(f"Err: {item_err}", styles['small']), "", "", ""]); sn += 1
        items_table = Table(table_data, colWidths=[0.5*inch, 3.0*inch, 0.7*inch, 1.0*inch, 1.1*inch]); items_table.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.grey), ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke), ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('ALIGN', (1, 1), (1, -1), 'LEFT'), ('ALIGN', (2, 1), (-1, -1), 'RIGHT'), ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'), ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'), ('FONTSIZE', (0, 0), (-1, -1), 9), ('BOTTOMPADDING', (0, 0), (-1, 0), 8), ('TOPPADDING', (0, 0), (-1, 0), 4), ('BOTTOMPADDING', (0, 1), (-1, -1), 4), ('TOPPADDING', (0, 1), (-1, -1), 4), ('GRID', (0, 0), (-1, -1), 1, colors.black), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')])); story.append(items_table); story.append(Spacer(1, 0.1*inch))
        totals_data = [['', '', Paragraph('<b>Total Amount:</b>', styles['Normal']), Paragraph(f"<b>{format_currency(total_amount)}</b>", styles['Normal'])]]; totals_table = Table(totals_data, colWidths=[3.5*inch + 0.7*inch, 1.0*inch, 1.1*inch]); totals_table.setStyle(TableStyle([('ALIGN', (0, 0), (-1, -1), 'RIGHT'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'), ('FONTSIZE', (0, 0), (-1, -1), 10), ('BOTTOMPADDING', (0, 0), (-1, -1), 5), ('TOPPADDING', (0, 0), (-1, -1), 5)])); story.append(totals_table); story.append(Spacer(1, 0.3*inch))
        story.append(Paragraph("<u>Terms & Conditions:</u>", styles['h4'])); story.append(Paragraph("1. Goods once sold will not be taken back.", styles['small'])); story.append(Paragraph("2. Interest @18% p.a. charged if bill not paid within 30 days.", styles['small'])); story.append(Paragraph("3. Subject to Jalandhar jurisdiction only.", styles['small'])); story.append(Spacer(1, 0.5*inch)); story.append(Paragraph(f"For {company_name}", styles['Normal'])); story.append(Spacer(1, 0.5*inch)); story.append(Paragraph("Authorised Signatory", styles['Normal']))
        pdf.build(story)
        # On success, put the FULL PATH into the queue
        result_queue.put(("Success", os.path.abspath(pdf_file)))
    except Exception as e: result_queue.put(("Error", f"PDF generation failed: {e}\n{traceback.format_exc()}"))

def generate_excel_invoice_threaded(invoice_id, invoice_type, entity_name, invoice_data, invoice_items_dec, result_queue):
    global COMPANY_SETTINGS; entity_label = invoice_type.capitalize()
    excel_file = f"{invoice_type}_{entity_name.replace(' ','_')}_{invoice_id}_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx"; title = "Invoice" if invoice_type == 'customer' else "Bill"
    try:
        wb = Workbook(); ws = wb.active; ws.title = f"{title} {invoice_id}"
        company_name = COMPANY_SETTINGS.get('company_name', DEFAULT_SETTINGS['company_name']); company_address = COMPANY_SETTINGS.get('company_address', DEFAULT_SETTINGS['company_address']); company_email = COMPANY_SETTINGS.get('company_email', DEFAULT_SETTINGS['company_email']); company_phone = COMPANY_SETTINGS.get('company_phone', DEFAULT_SETTINGS['company_phone']); company_gstin = COMPANY_SETTINGS.get('company_gstin', None)
        header_rows = [[company_name], [company_address], [f"M: {company_phone} | Email: {company_email}"]];
        if company_gstin: header_rows.append([f"GSTIN: {company_gstin}"])
        header_rows.extend([[], [f"{title.upper()} No:", invoice_id, "", "Date:", invoice_data.get('date', 'N/A')], [f"{entity_label} Name:", entity_name], []])
        for row_data in header_rows: ws.append(row_data)
        current_row_idx = len(header_rows)
        ws.append(["S.N", "Item Description", "Quantity", "Rate", "Amount"]); current_row_idx += 1; items_header_row_idx = current_row_idx
        total_amount = ZERO_DECIMAL; sn = 1
        for item in invoice_items_dec:
             try:
                 item_name = item.get('item', 'N/A'); qty = item.get('quantity', ZERO_DECIMAL); price = item.get('price', ZERO_DECIMAL); amount = (qty * price).quantize(TWO_PLACES, rounding=ROUND_HALF_UP); total_amount += amount
                 ws.append([sn, str(item_name), qty, price, amount]); sn += 1; current_row_idx += 1
             except Exception as item_err: print(f"Excel Gen Error: Skipping item {item}: {item_err}"); ws.append([sn, "Error", "", "", ""]); sn += 1; current_row_idx += 1
        ws.append([]); current_row_idx += 1; ws.append(["", "", "", "Total Amount:", total_amount]); current_row_idx += 1
        currency_format = f'"{CURRENCY_SYMBOL}" #,##0.00'; qty_format = '#,##0.###'
        for row_idx in range(items_header_row_idx + 1, current_row_idx):
            if ws.max_row < row_idx: continue
            for col_letter in ['D', 'E']:
                cell = ws[f"{col_letter}{row_idx}"]
                if isinstance(cell.value, (int, float, Decimal)): cell.number_format = currency_format
            cell_qty = ws[f"C{row_idx}"]
            if isinstance(cell_qty.value, (int, float, Decimal)): cell_qty.number_format = qty_format
        total_amount_cell = ws[f"E{current_row_idx}"]
        if isinstance(total_amount_cell.value, (int, float, Decimal)):
             total_amount_cell.number_format = currency_format
        for col in ws.columns:
             max_length = 0; column_letter = col[0].column_letter
             for cell in col:
                  try:
                      if cell.value is not None: max_length = max(max_length, len(str(cell.value)))
                  except: pass
             adjusted_width = min((max_length + 2) * 1.1, 60); ws.column_dimensions[column_letter].width = max(adjusted_width, 8)
        wb.save(excel_file); result_queue.put(("Success", os.path.abspath(excel_file)))
    except Exception as e: result_queue.put(("Error", f"Excel generation failed: {e}\n{traceback.format_exc()}"))

def calculate_invoice_total(invoice_id, invoice_type):
    """
    Calculates the total amount for a given invoice or supplier bill.
    invoice_type: 'customer' or 'supplier'
    """
    items_list = INVOICE_ITEMS_DATA if invoice_type == 'customer' else SUPPLIER_INVOICE_ITEMS_DATA
    id_key = 'invoice_id' if invoice_type == 'customer' else 'supplier_invoice_id'
    total = ZERO_DECIMAL
    for item in items_list:
        if item.get(id_key) == invoice_id:
            qty = item.get('quantity', ZERO_DECIMAL)
            price = item.get('price', ZERO_DECIMAL)
            try:
                total += (qty * price).quantize(TWO_PLACES, rounding=ROUND_HALF_UP)
            except Exception:
                continue
    return total

def calculate_trial_balance():
    """
    Calculates the trial balance based on current data.
    Returns a tuple: (list of (account, debit, credit), total_debit, total_credit)
    """
    # Example accounts: AR: CustomerName, AP: SupplierName
    ar_balances = {}
    ap_balances = {}

    # Calculate AR (Accounts Receivable) for customers
    for inv in INVOICES_DATA:
        customer = inv.get('customer_name', 'Unknown')
        amount = calculate_invoice_total(inv.get('id'), 'customer')
        ar_balances[customer] = ar_balances.get(customer, ZERO_DECIMAL) + amount
        # Subtract payments
    for payment in PAYMENTS_DATA:
        if payment.get('payment_for') == 'customer':
            customer = payment.get('customer_name', 'Unknown')
            amount = payment.get('amount', ZERO_DECIMAL)
            ar_balances[customer] = ar_balances.get(customer, ZERO_DECIMAL) - amount

    # Calculate AP (Accounts Payable) for suppliers
    for bill in SUPPLIER_INVOICES_DATA:
        supplier = bill.get('supplier_name', 'Unknown')
        amount = calculate_invoice_total(bill.get('id'), 'supplier')
        ap_balances[supplier] = ap_balances.get(supplier, ZERO_DECIMAL) + amount
    for payment in PAYMENTS_DATA:
        if payment.get('payment_for') == 'supplier':
            supplier = payment.get('supplier_name', 'Unknown')
            amount = payment.get('amount', ZERO_DECIMAL)
            ap_balances[supplier] = ap_balances.get(supplier, ZERO_DECIMAL) - amount

    # Inventory value
    inventory_value = sum((item.get('quantity', ZERO_DECIMAL) * item.get('value', ZERO_DECIMAL)) for item in INVENTORY_DATA)

    tb_data = []
    total_debit = ZERO_DECIMAL
    total_credit = ZERO_DECIMAL

    # AR accounts (debit balances)
    for customer, balance in ar_balances.items():
        if balance > ZERO_DECIMAL:
            tb_data.append((f"AR: {customer}", balance, ZERO_DECIMAL))
            total_debit += balance
        elif balance < ZERO_DECIMAL:
            tb_data.append((f"AR: {customer}", ZERO_DECIMAL, -balance))
            total_credit += -balance

    # AP accounts (credit balances)
    for supplier, balance in ap_balances.items():
        if balance > ZERO_DECIMAL:
            tb_data.append((f"AP: {supplier}", ZERO_DECIMAL, balance))
            total_credit += balance
        elif balance < ZERO_DECIMAL:
            tb_data.append((f"AP: {supplier}", -balance, ZERO_DECIMAL))
            total_debit += -balance

    # Inventory (asset, debit)
    if inventory_value > ZERO_DECIMAL:
        tb_data.append(("Inventory", inventory_value, ZERO_DECIMAL))
        total_debit += inventory_value

    return tb_data, total_debit, total_credit

def generate_pdf_trial_balance_threaded(result_queue):
    """Generates a PDF of the complete Trial Balance."""
    pdf_file = f"trial_balance_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    try:
        tb_data, total_debit, total_credit = calculate_trial_balance()

        pdf = SimpleDocTemplate(pdf_file, pagesize=letter, topMargin=0.5*inch, bottomMargin=0.5*inch, leftMargin=0.7*inch, rightMargin=0.7*inch)
        styles = getSampleStyleSheet()
        styles.add(ParagraphStyle(name='RightAlign', parent=styles['Normal'], alignment=2))
        styles['h1'].alignment = 1
        story = []

        # Company Header
        company_name = COMPANY_SETTINGS.get('company_name', DEFAULT_SETTINGS['company_name'])
        company_address = COMPANY_SETTINGS.get('company_address', DEFAULT_SETTINGS['company_address'])
        header_text = [Paragraph(f"<b>{company_name}</b>", styles['h1']), Paragraph(company_address, styles['Normal'])]
        story.extend(header_text)
        story.append(Spacer(1, 0.1*inch))
        story.append(Paragraph('<hr width="100%" color="black" size="1"/>', styles['Normal']))
        story.append(Spacer(1, 0.1*inch))

        # Title
        now_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        story.append(Paragraph("<b>Trial Balance</b>", styles['h2']))
        story.append(Paragraph(f"<i>As at: {now_str}</i>", styles['Normal']))
        story.append(Spacer(1, 0.2*inch))

        # Table Data
        table_header = [Paragraph("<b>Account Title</b>"), Paragraph("<b>Debit</b>"), Paragraph("<b>Credit</b>")]
        table_data = [table_header]
        for account, debit_val, credit_val in tb_data:
            debit_str = format_currency(debit_val) if debit_val > ZERO_DECIMAL else ""
            credit_str = format_currency(credit_val) if credit_val > ZERO_DECIMAL else ""
            table_data.append([Paragraph(account, styles['Normal']), Paragraph(debit_str, styles['RightAlign']), Paragraph(credit_str, styles['RightAlign'])])
        
        # Totals Row
        table_data.append([Paragraph("<b>TOTALS</b>", styles['Normal']), 
                           Paragraph(f"<b>{format_currency(total_debit)}</b>", styles['RightAlign']), 
                           Paragraph(f"<b>{format_currency(total_credit)}</b>", styles['RightAlign'])])

        # Table Style
        tb_table = Table(table_data, colWidths=[4*inch, 1.5*inch, 1.5*inch])
        tb_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),  # First column left-aligned
            ('ALIGN', (1, 0), (-1, -1), 'RIGHT'), # Other columns right-aligned
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'), # Header bold
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'), # Total row bold
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        story.append(tb_table)
        pdf.build(story)
        result_queue.put(("Success", os.path.abspath(pdf_file)))
    except Exception as e:
        result_queue.put(("Error", f"PDF generation failed: {e}\n{traceback.format_exc()}"))

def generate_pdf_detailed_ledger_threaded(account_title, result_queue):
    """Generates a PDF of a detailed account ledger/statement."""
    try:
        is_customer = account_title.startswith("AR: ")
        entity_name = account_title[4:] if account_title.startswith(("AR: ", "AP: ")) else account_title
        entity_type = 'customer' if is_customer else 'supplier'
        
        pdf_file = f"account_statement_{entity_name.replace(' ', '_')}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"

        # --- Data Gathering (Corrected self-contained logic) ---
        transactions = []
        if is_customer:
            # Customer Invoices (Debit AR)
            for inv in INVOICES_DATA:
                if inv.get('customer_name') == entity_name:
                    inv_id = inv.get('id')
                    if inv_id is not None:
                        amount = calculate_invoice_total(inv.get('id'), 'customer')
                        transactions.append({'date': inv.get('date', 'N/A'), 'type': 'Invoice', 'reference': f'INV#{inv_id}', 'debit': amount, 'credit': ZERO_DECIMAL})
            # Customer Payments (Credit AR)
            for payment in PAYMENTS_DATA:
                if payment.get('customer_name') == entity_name and payment.get('payment_for') == 'customer':
                    amount = payment.get('amount', ZERO_DECIMAL)
                    transactions.append({'date': payment.get('date', 'N/A'), 'type': 'Payment Received', 'reference': f'PMT#{payment.get("id", "N/A")}', 'debit': ZERO_DECIMAL, 'credit': amount})
        else: # Supplier (AP - Liability)
            # Supplier Bills (Credit AP)
            for bill in SUPPLIER_INVOICES_DATA:
                if bill.get('supplier_name') == entity_name:
                    bill_id = bill.get('id')
                    if bill_id is not None:
                        amount = calculate_invoice_total(bill_id, 'supplier')
                        transactions.append({'date': bill.get('date', 'N/A'), 'type': 'Bill', 'reference': f'BILL#{bill_id}', 'debit': ZERO_DECIMAL, 'credit': amount})
            # Supplier Payments (Debit AP)
            for payment in PAYMENTS_DATA:
                if payment.get('supplier_name') == entity_name and payment.get('payment_for') == 'supplier':
                    amount = payment.get('amount', ZERO_DECIMAL)
                    transactions.append({'date': payment.get('date', 'N/A'), 'type': 'Payment Made', 'reference': f'PMT#{payment.get("id", "N/A")}', 'debit': amount, 'credit': ZERO_DECIMAL})
        
        # Sort transactions by date and calculate running balance
        transactions.sort(key=lambda x: x.get('date', ''))
        running_balance = ZERO_DECIMAL
        for trans in transactions:
            running_balance += (trans.get('debit', ZERO_DECIMAL) - trans.get('credit', ZERO_DECIMAL))
            trans['balance'] = running_balance

        # --- PDF Generation ---
        pdf = SimpleDocTemplate(pdf_file, pagesize=letter, topMargin=0.5*inch, bottomMargin=0.5*inch, leftMargin=0.7*inch, rightMargin=0.7*inch)
        styles = getSampleStyleSheet()
        styles.add(ParagraphStyle(name='RightAlign', parent=styles['Normal'], alignment=2))
        styles.add(ParagraphStyle(name='small', parent=styles['Normal'], fontSize=8))
        styles['h1'].alignment = 1
        styles['h2'].alignment = 1
        story = []

        # Company Header
        company_name = COMPANY_SETTINGS.get('company_name', DEFAULT_SETTINGS['company_name'])
        company_address = COMPANY_SETTINGS.get('company_address', DEFAULT_SETTINGS['company_address'])
        header_text = [Paragraph(f"<b>{company_name}</b>", styles['h1']), Paragraph(company_address, styles['Normal'])]
        story.extend(header_text)
        story.append(Spacer(1, 0.1*inch))
        story.append(Paragraph('<hr width="100%" color="black" size="1"/>', styles['Normal']))
        story.append(Spacer(1, 0.1*inch))
        
        # Title and Entity Info
        story.append(Paragraph("<b>Account Statement</b>", styles['h2']))
        story.append(Spacer(1, 0.1*inch))
        entity_phone = get_entity_phone_number(entity_name, entity_type)
        phone_info = f" | Phone: {entity_phone}" if entity_phone else ""
        info_data = [
            [Paragraph(f"For: <b>{entity_name}</b>{phone_info}", styles['Normal']), 
             Paragraph(f"Statement Date: <b>{datetime.datetime.now().strftime(DATE_FORMAT)}</b>", styles['RightAlign'])]
        ]
        info_table = Table(info_data, colWidths=[4.5*inch, 2.0*inch])
        info_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
        story.append(info_table)
        story.append(Spacer(1, 0.2*inch))

        # Table of Transactions
        table_header = [Paragraph("<b>Date</b>"), Paragraph("<b>Type</b>"), Paragraph("<b>Reference</b>"), Paragraph("<b>Debit</b>"), Paragraph("<b>Credit</b>"), Paragraph("<b>Balance</b>")]
        table_data = [table_header]
        
        for trans in transactions:
            debit_str = format_currency(trans['debit']) if trans['debit'] > ZERO_DECIMAL else ""
            credit_str = format_currency(trans['credit']) if trans['credit'] > ZERO_DECIMAL else ""
            table_data.append([
                Paragraph(trans['date'], styles['Normal']), Paragraph(trans['type'], styles['Normal']), Paragraph(trans['reference'], styles['Normal']),
                Paragraph(debit_str, styles['RightAlign']), Paragraph(credit_str, styles['RightAlign']), Paragraph(format_currency(trans['balance']), styles['RightAlign'])
            ])
        
        # Table Style
        trans_table = Table(table_data, colWidths=[0.8*inch, 1.5*inch, 1.2*inch, 1*inch, 1*inch, 1*inch])
        trans_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (2, -1), 'LEFT'),  # Left-aligned columns
            ('ALIGN', (3, 0), (-1, -1), 'RIGHT'), # Right-aligned columns
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        story.append(trans_table)
        story.append(Spacer(1, 0.2*inch))

        # Final Balance
        final_balance_text = f"<b>Closing Balance: {format_currency(running_balance)}</b>"
        story.append(Paragraph(final_balance_text, styles['RightAlign']))
        
        # Add a small note to explain the balance
        note_style = styles['small']
        note_style.alignment = 2 # Right align
        if is_customer:
            story.append(Paragraph("<i>(A positive balance is due from the customer)</i>", note_style))
        else:
            story.append(Paragraph("<i>(A positive balance is an advance paid to the supplier)</i>", note_style))

        pdf.build(story)
        # On success, put the FULL PATH into the queue
        result_queue.put(("Success", os.path.abspath(pdf_file)))
    except Exception as e:
        result_queue.put(("Error", f"PDF generation failed: {e}\n{traceback.format_exc()}"))


# --- Helper function to get entity phone number ---
def get_entity_phone_number(entity_name, entity_type):
    """
    Returns the phone number for a customer or supplier by name and type.
    entity_type: 'customer' or 'supplier'
    """
    if entity_type == 'customer':
        for inv in INVOICES_DATA:
            if inv.get('customer_name') == entity_name:
                return inv.get('customer_phone', '')
    elif entity_type == 'supplier':
        for bill in SUPPLIER_INVOICES_DATA:
            if bill.get('supplier_name') == entity_name:
                return bill.get('supplier_phone', '')
    return ""

# --- EazeBot Query Processing ---
def process_eazebot_query(query, context_data=None):
    """
    Processes a user query using Gemini API (if available) and returns the response.
    If Gemini API is not available, returns a fallback message.
    context_data: Optional dict with additional business data to provide context.
    """
    global GEMINI_API_KEY, gemini_lib_installed
    if not gemini_lib_installed:
        return "Gemini API is not installed. Please install google-generativeai to use EazeBot features."
    if not GEMINI_API_KEY:
        return "Gemini API key is not set. Please set your Gemini API key in the application settings."
    try:
        import google.generativeai as genai
        genai.configure(api_key=GEMINI_API_KEY)
        prompt = query
        if context_data:
            prompt += "\nContext Data: " + json.dumps(context_data, cls=DecimalEncoder)
        model = genai.GenerativeModel('gemini-pro')
        response = model.generate_content(prompt)
        if hasattr(response, 'text'):
            return response.text
        elif hasattr(response, 'candidates') and response.candidates:
            return response.candidates[0].text
        else:
            return str(response)
    except Exception as e:
        return f"EazeBot error: {e}"

# EazeBot: Chart Generation (No change)
def eazebot_create_daily_sales_chart(days=7):
    if not matplotlib_installed: return None
    today = datetime.date.today()
    date_range = [today - datetime.timedelta(days=i) for i in range(days)]
    sales_by_day = {d.strftime(DATE_FORMAT): ZERO_DECIMAL for d in date_range}
    
    start_date_str = (today - datetime.timedelta(days=days-1)).strftime(DATE_FORMAT)
    
    for inv in INVOICES_DATA:
        inv_date = inv.get('date', '')
        if inv_date >= start_date_str:
            if inv_date in sales_by_day:
                sales_by_day[inv_date] += calculate_invoice_total(inv.get('id'), 'customer')
                
    dates = sorted(sales_by_day.keys())
    sales_values = [sales_by_day[d] for d in dates]

    fig, ax = plt.subplots(figsize=(6, 4), dpi=96)
    ax.bar(dates, sales_values, color='cornflowerblue')
    ax.set_title(f'Sales for the Last {days} Days', fontsize=12)
    ax.set_ylabel(f'Total Sales ({CURRENCY_SYMBOL})', fontsize=10)
    ax.set_xlabel('Date', fontsize=10)
    ax.tick_params(axis='x', rotation=45, labelsize=8)
    ax.tick_params(axis='y', labelsize=8)
    fig.tight_layout()
    
    chart_path = os.path.join(DATA_DIR, "eazebot_chart.png")
    fig.savefig(chart_path)
    plt.close(fig)
    
    img = Image.open(chart_path)
    img.thumbnail((450, 350))
    photo_image = ImageTk.PhotoImage(img)
    return photo_image


# --- Constants ---
DATA_DIR = "eaze_inn_data"
USERS_FILE = os.path.join(DATA_DIR, "users.json")
INVOICES_FILE = os.path.join(DATA_DIR, "invoices.json")
INVOICE_ITEMS_FILE = os.path.join(DATA_DIR, "invoice_items.json")
SUPPLIER_INVOICES_FILE = os.path.join(DATA_DIR, "supplier_invoices.json")
SUPPLIER_INVOICE_ITEMS_FILE = os.path.join(DATA_DIR, "supplier_invoice_items.json")
INVENTORY_FILE = os.path.join(DATA_DIR, "inventory.json") # Inventory data file
PAYMENTS_FILE = os.path.join(DATA_DIR, "payments.json")
IMAGES_DIR = os.path.join(DATA_DIR, "invoice_images")
SETTINGS_FILE = os.path.join(DATA_DIR, "settings.json")
COMPANY_QR_BASE_FILENAME = "company_qr_code"

# --- App Settings ---
LOW_STOCK_THRESHOLD = Decimal('5') # Used by inventory management
CURRENCY_SYMBOL = "₹"
DATE_FORMAT = '%Y-%m-%d'
ZERO_DECIMAL = Decimal('0.00')
TWO_PLACES = Decimal('0.01')

THERMAL_PRINTER_TYPE = 'win32raw'
THERMAL_PRINTER_VID = 0x04b8
THERMAL_PRINTER_PID = 0x0e15
THERMAL_PRINTER_IN_EP = 0x81
THERMAL_PRINTER_OUT_EP = 0x01
THERMAL_PRINTER_NAME = "POS58 Printer"
THERMAL_PRINTER_IP = "192.168.1.100"
THERMAL_PRINTER_PORT = 9100
THERMAL_PRINTER_PORT_SERIAL = "/dev/ttyUSB0"
THERMAL_PRINTER_BAUDRATE = 9600
THERMAL_PRINTER_FILE = "receipt_output.bin"
RECEIPT_WIDTH = 32

# --- Global In-Memory Data Storage ---
USERS_DATA = []
INVOICES_DATA = []
INVOICE_ITEMS_DATA = []
SUPPLIER_INVOICES_DATA = []
SUPPLIER_INVOICE_ITEMS_DATA = []
INVENTORY_DATA = [] # Holds inventory items: {id, item_name, quantity, value (cost_price)}
PAYMENTS_DATA = []
COMPANY_SETTINGS = {}
GEMINI_API_KEY = None # Will be set at runtime

DEFAULT_SETTINGS = {
    "company_name": "Your Company Name",
    "company_address": "Your Company Address, City, PIN",
    "company_email": "your.email@example.com",
    "company_phone": "Your Phone Number",
    "company_gstin": "Your GSTIN (Optional)",
    "qr_code_path": None
}


class DecimalEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, Decimal): return str(obj)
        return super(DecimalEncoder, self).default(obj)

def load_settings():
    global COMPANY_SETTINGS; COMPANY_SETTINGS = DEFAULT_SETTINGS.copy()
    try:
        if not os.path.exists(SETTINGS_FILE): print(f"Settings file '{os.path.basename(SETTINGS_FILE)}' not found. Using defaults."); return
        with open(SETTINGS_FILE, 'r', encoding='utf-8') as f: content = f.read()
        if not content.strip(): print("Settings file is empty. Using defaults."); return
        loaded_settings = json.loads(content)
        COMPANY_SETTINGS.update(loaded_settings); print("Company settings loaded successfully.")
    except (IOError, json.JSONDecodeError) as e: print(f"Error loading settings from {os.path.basename(SETTINGS_FILE)}: {e}. Using defaults.")
    except Exception as e: print(f"Unexpected error loading settings: {e}. Using defaults."); traceback.print_exc()

def save_settings_file():
    global COMPANY_SETTINGS
    try:
        os.makedirs(os.path.dirname(SETTINGS_FILE), exist_ok=True)
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f: json.dump(COMPANY_SETTINGS, f, cls=DecimalEncoder, indent=4)
        print(f"Company settings saved to {os.path.basename(SETTINGS_FILE)}"); return True
    except (IOError, TypeError) as e: print(f"Error saving settings: {e}"); traceback.print_exc(); messagebox.showerror("Settings Save Error", f"Could not save settings.\n{e}", icon='error'); return False
    except Exception as e: print(f"Unexpected error saving settings: {e}"); traceback.print_exc(); messagebox.showerror("Settings Save Error", f"Unexpected error saving settings.\n{e}", icon='error'); return False

def _validate_and_copy_image(original_path, target_dir, target_base_filename):
    if not original_path or not os.path.exists(original_path): return None
    os.makedirs(target_dir, exist_ok=True)
    try:
        file_ext = os.path.splitext(original_path)[1].lower()
        if file_ext not in ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff']: messagebox.showwarning("Image Warning", f"Unsupported image format '{file_ext}'. Please use JPG, PNG, BMP, GIF, or TIFF.", icon='warning'); return None
        new_filename = f"{target_base_filename}{file_ext}"; new_path = os.path.join(target_dir, new_filename)
        _remove_existing_image(target_dir, target_base_filename); shutil.copy2(original_path, new_path)
        try: img = Image.open(new_path); img.verify(); img.close(); return new_filename
        except (IOError, SyntaxError, Image.UnidentifiedImageError) as img_err:
            print(f"Warning: Copied image seems invalid: {new_path} - {img_err}")
            try: os.remove(new_path)
            except OSError: pass
            messagebox.showwarning("Image Warning", f"The selected image file could not be verified:\n{os.path.basename(original_path)}", icon='warning'); return None
    except Exception as e: print(f"Error processing image: {e}"); messagebox.showerror("Image Error", f"Could not copy/process image:\n{e}", icon='error'); return None

def _remove_existing_image(target_dir, target_base_filename):
    try:
        for filename in os.listdir(target_dir):
            if filename.startswith(target_base_filename):
                try: os.remove(os.path.join(target_dir, filename)); print(f"Removed existing image: {filename}")
                except OSError as rm_err: print(f"Warning: Could not remove existing image {filename}: {rm_err}")
    except FileNotFoundError: pass
    except Exception as e: print(f"Error removing existing image files for {target_base_filename}: {e}")

def _handle_invoice_image(original_path, invoice_type, invoice_id):
    target_base_filename = f"{invoice_type}_invoice_{invoice_id}"; relative_path = _validate_and_copy_image(original_path, IMAGES_DIR, target_base_filename)
    return relative_path if relative_path else None


def load_data(filepath):
    data = []
    try:
        if not os.path.exists(filepath): print(f"Data file not found: {filepath}. Starting empty."); return []
        with open(filepath, 'r', encoding='utf-8') as f: content = f.read()
        if not content.strip(): return []
        data = json.loads(content)
    except (IOError, json.JSONDecodeError) as e: print(f"Error loading {filepath}: {e}"); messagebox.showerror("Data Load Error", f"Could not load {os.path.basename(filepath)}.\nCheck console.", icon='warning'); return []
    except Exception as e: print(f"Unexpected error loading {filepath}: {e}"); traceback.print_exc(); messagebox.showerror("Data Load Error", f"Unexpected error loading {os.path.basename(filepath)}.\nCheck console.", icon='warning'); return []
    processed_data = []
    for item in data:
        new_item = item.copy()
        try:
            if 'id' in new_item and new_item['id'] is not None: new_item['id'] = int(new_item['id'])
            if 'invoice_id' in new_item and new_item['invoice_id'] is not None: new_item['invoice_id'] = int(new_item['invoice_id'])
            if 'supplier_invoice_id' in new_item and new_item['supplier_invoice_id'] is not None: new_item['supplier_invoice_id'] = int(new_item['supplier_invoice_id'])
            # Ensure 'quantity' and 'value' (for inventory) are Decimal
            # Other financial fields are already covered
            for key in ['price', 'value', 'amount', 'total_amount', 'quantity', 'amount_paid']:
                if key in new_item and new_item[key] is not None:
                    try: new_item[key] = Decimal(str(new_item[key]))
                    except InvalidOperation: print(f"Warn: Invalid Decimal for '{key}' in {filepath}, ID {new_item.get('id', 'N/A')}: '{new_item[key]}'. Setting to 0."); new_item[key] = ZERO_DECIMAL
            processed_data.append(new_item)
        except (ValueError, TypeError) as conv_e: print(f"Warn: Skipping record due to conversion error in {filepath}: {item} - Error: {conv_e}")
    return processed_data

def save_data(data_list, filepath):
    try:
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        with open(filepath, 'w', encoding='utf-8') as f: json.dump(data_list, f, cls=DecimalEncoder, indent=4)
        return True
    except (IOError, TypeError) as e: print(f"Error saving to {filepath}: {e}"); traceback.print_exc(); messagebox.showerror("Data Save Error", f"Could not save to {os.path.basename(filepath)}.\nCheck logs.", icon='error'); return False
    except Exception as e: print(f"Unexpected error saving to {filepath}: {e}"); traceback.print_exc(); messagebox.showerror("Data Save Error", f"Unexpected error saving {os.path.basename(filepath)}.\nCheck logs.", icon='error'); return False

def load_all_data():
    global USERS_DATA, INVOICES_DATA, INVOICE_ITEMS_DATA, SUPPLIER_INVOICES_DATA, SUPPLIER_INVOICE_ITEMS_DATA, INVENTORY_DATA, PAYMENTS_DATA, COMPANY_SETTINGS
    print("Loading data..."); USERS_DATA.clear(); INVOICES_DATA.clear(); INVOICE_ITEMS_DATA.clear(); SUPPLIER_INVOICES_DATA.clear(); SUPPLIER_INVOICE_ITEMS_DATA.clear(); INVENTORY_DATA.clear(); PAYMENTS_DATA.clear(); COMPANY_SETTINGS.clear()
    USERS_DATA.extend(load_data(USERS_FILE)); INVOICES_DATA.extend(load_data(INVOICES_FILE)); INVOICE_ITEMS_DATA.extend(load_data(INVOICE_ITEMS_FILE)); SUPPLIER_INVOICES_DATA.extend(load_data(SUPPLIER_INVOICES_FILE)); SUPPLIER_INVOICE_ITEMS_DATA.extend(load_data(SUPPLIER_INVOICE_ITEMS_FILE))
    INVENTORY_DATA.extend(load_data(INVENTORY_FILE)) # Load inventory data
    PAYMENTS_DATA.extend(load_data(PAYMENTS_FILE))
    load_settings(); print(f"Data loaded: {len(USERS_DATA)}u, {len(INVOICES_DATA)}inv, {len(SUPPLIER_INVOICES_DATA)}bill, {len(INVENTORY_DATA)}ity, {len(PAYMENTS_DATA)}pay."); print(f"Settings: Name='{COMPANY_SETTINGS.get('company_name', 'N/A')}'")


def get_next_id(data_list):
    if not data_list: return 1
    max_id = 0
    for item in data_list:
        try: current_id = int(item.get('id', 0)); max_id = max(max_id, current_id)
        except (ValueError, TypeError): continue
    return max_id + 1

def hash_password(password): return hashlib.sha256(password.encode()).hexdigest()

def format_currency(amount, include_sign=False):
    if amount is None: return f"{CURRENCY_SYMBOL}0.00"
    try:
        decimal_amount = Decimal(str(amount)).quantize(TWO_PLACES, rounding=ROUND_HALF_UP)
        if decimal_amount.is_zero() and decimal_amount.is_signed():
             decimal_amount = ZERO_DECIMAL
        if include_sign:
            return f"{CURRENCY_SYMBOL}{decimal_amount:+,.2f}"
        else:
            return f"{CURRENCY_SYMBOL}{decimal_amount:,.2f}"
    except (InvalidOperation, TypeError, ValueError):
        return f"{CURRENCY_SYMBOL}N/A"

def format_decimal_quantity(quantity):
    if quantity is None: return "0"
    try: qty_d = Decimal(str(quantity)); return f"{qty_d:.10f}".rstrip('0').rstrip('.')
    except (InvalidOperation, TypeError, ValueError): return "N/A"

def format_percentage_diff(current, previous):
    try:
        current_d, previous_d = Decimal(str(current)), Decimal(str(previous))
        if previous_d == ZERO_DECIMAL: return "N/A (Prev 0)" if current_d != ZERO_DECIMAL else "0.00%"
        diff = ((current_d - previous_d) / previous_d) * 100; return f"{diff:+.2f}%"
    except (InvalidOperation, TypeError, ValueError): return "N/A"

# --- Inventory Update Logic ---
def update_inventory_after_transaction(transaction_type, processed_items):
    global INVENTORY_DATA
    inventory_changed = False
    for proc_item in processed_items:
        item_name = proc_item['item'].strip()
        quantity_change = proc_item['quantity']  # Expected to be Decimal
        price_per_unit = proc_item['price']      # Expected to be Decimal

        inventory_item = next((inv_item for inv_item in INVENTORY_DATA if inv_item.get('item_name', '').strip().lower() == item_name.lower()), None)

        if transaction_type == 'supplier':  # Purchase
            if inventory_item:
                old_quantity = inventory_item.get('quantity', ZERO_DECIMAL)
                new_quantity = old_quantity + quantity_change
                inventory_item['quantity'] = new_quantity
                inventory_item['value'] = price_per_unit # Update cost to latest purchase price
                inventory_item['last_updated'] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                print(f"Inventory Update (Purchase): '{item_name}' old_qty: {old_quantity}, added: {quantity_change}, new_qty: {new_quantity}, new_cost: {price_per_unit}")
            else:
                new_id = get_next_id(INVENTORY_DATA)
                inventory_item_new = {
                    'id': new_id,
                    'item_name': item_name,
                    'quantity': quantity_change,
                    'value': price_per_unit,
                    'last_updated': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
                INVENTORY_DATA.append(inventory_item_new)
                print(f"Inventory Add (Purchase): '{item_name}' qty: {quantity_change}, cost: {price_per_unit}")
            inventory_changed = True

        elif transaction_type == 'customer':  # Sale
            if inventory_item:
                old_quantity = inventory_item.get('quantity', ZERO_DECIMAL)
                new_quantity = old_quantity - quantity_change
                inventory_item['quantity'] = new_quantity
                inventory_item['last_updated'] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                # 'value' (cost) does not change on sale
                print(f"Inventory Update (Sale): '{item_name}' old_qty: {old_quantity}, sold: {quantity_change}, new_qty: {new_quantity}")
                inventory_changed = True
            else: # Item sold but not in inventory
                new_id = get_next_id(INVENTORY_DATA)
                inventory_item_new = {
                    'id': new_id,
                    'item_name': item_name,
                    'quantity': -quantity_change, # Record as negative stock
                    'value': ZERO_DECIMAL,        # Cost is unknown
                    'last_updated': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'status_flag': 'SOLD_WITHOUT_STOCK' # Custom flag
                }
                INVENTORY_DATA.append(inventory_item_new)
                print(f"Inventory Alert (Sale): Item '{item_name}' sold without prior stock. Added with negative quantity.")
                inventory_changed = True
    
    if inventory_changed:
        if not save_data(INVENTORY_DATA, INVENTORY_FILE):
            print("CRITICAL: FAILED TO SAVE INVENTORY UPDATES TO FILE.")
            messagebox.showerror("Inventory Save Error", 
                                 "Failed to save inventory updates to file.\n"
                                 "Data might be inconsistent upon restart.\n"
                                 "Please check console logs.", icon='error')
# --- User Authentication Windows ---
def register_window(root):
    register_win = tk.Toplevel(root); register_win.title("Register New User")
    register_win.transient(root); register_win.grab_set(); register_win.resizable(False, False)
    tk.Label(register_win, text="Username:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
    username_entry = tk.Entry(register_win, width=30); username_entry.grid(row=0, column=1, padx=10, pady=5)
    tk.Label(register_win, text="Password:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
    password_entry = tk.Entry(register_win, show="*", width=30); password_entry.grid(row=1, column=1, padx=10, pady=5)
    username_entry.focus_set()
    def register_command():
        username = username_entry.get().strip(); password = password_entry.get()
        if not username or not password: messagebox.showerror("Input Error", "Please enter both username and password.", parent=register_win); return
        if any(user.get('username') == username for user in USERS_DATA): messagebox.showerror("Username Exists", "Username already exists.", parent=register_win); return
        new_id = get_next_id(USERS_DATA); hashed_password = hash_password(password)
        new_user = {'id': new_id, 'username': username, 'password': hashed_password}
        USERS_DATA.append(new_user)
        if save_data(USERS_DATA, USERS_FILE): messagebox.showinfo("Success", "Registration successful!", parent=register_win); register_win.destroy()
        else: USERS_DATA.pop() # Rollback
    btn_frame = tk.Frame(register_win); btn_frame.grid(row=2, column=0, columnspan=2, pady=10)
    tk.Button(btn_frame, text="Register", command=register_command, width=10).pack(side=tk.LEFT, padx=5)
    tk.Button(btn_frame, text="Cancel", command=register_win.destroy, width=10).pack(side=tk.LEFT, padx=5)

def show_main_app(root):
    root.withdraw()  # Hide the login window
    create_dashboard(root)  # Show the main dashboard

def signin_window(root):
    signin_win = tk.Toplevel(root)
    signin_win.title("Sign In")
    signin_win.transient(root)
    signin_win.grab_set()
    signin_win.resizable(False, False)
    
    tk.Label(signin_win, text="Username:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
    username_entry = tk.Entry(signin_win, width=30)
    username_entry.grid(row=0, column=1, padx=10, pady=5)
    
    tk.Label(signin_win, text="Password:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
    password_entry = tk.Entry(signin_win, show="*", width=30)
    password_entry.grid(row=1, column=1, padx=10, pady=5)
    
    username_entry.focus_set()
    
    def signin_command():
        username = username_entry.get().strip()
        password = password_entry.get()
        if not username or not password:
            messagebox.showerror("Input Error", "Please enter both username and password.", parent=signin_win)
            return
        found_user = next((user for user in USERS_DATA if user.get('username') == username), None)
        if not found_user or hash_password(password) != found_user.get('password'):
            messagebox.showerror("Authentication Failed", "Invalid username or password.", parent=signin_win)
            return
        signin_win.destroy()
        show_main_app(root)
    
    btn_frame = tk.Frame(signin_win)
    btn_frame.grid(row=2, column=0, columnspan=2, pady=10)
    tk.Button(btn_frame, text="Sign In", command=signin_command, width=10).pack(side=tk.LEFT, padx=5)
    tk.Button(btn_frame, text="Cancel", command=signin_win.destroy, width=10).pack(side=tk.LEFT, padx=5)
    password_entry.bind("<Return>", lambda event: signin_command())

# --- Backup/Restore Functions ---
def backup_all_data_threaded(result_queue):
    source_dir = DATA_DIR; backup_base_dir = "eaze_inn_json_backup"
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_dest_dir = os.path.join(backup_base_dir, f"backup_{timestamp}")
    try:
        if not os.path.isdir(source_dir): result_queue.put(("Warning", f"Data dir '{source_dir}' not found.")); return
        os.makedirs(backup_base_dir, exist_ok=True)
        shutil.copytree(source_dir, backup_dest_dir)
        result_queue.put(("Success", f"Backup successful!\nSaved in:\n{os.path.abspath(backup_dest_dir)}"))
    except Exception as e: result_queue.put(("Error", f"Backup failed: {e}\n{traceback.format_exc()}"))

def backup_all_data(root):
    if not messagebox.askyesno("Confirm Backup", "Create a backup of the current data?", parent=root): return
    result_queue = queue.Queue()
    thread = threading.Thread(target=backup_all_data_threaded, args=(result_queue,), daemon=True)
    thread.start()
    root.after(100, lambda: check_thread_queue(root, result_queue, "Backup"))

def restore_all_data_threaded(restore_source_dir, result_queue):
    target_dir = DATA_DIR; pre_restore_base_dir = "pre_restore_json_backups"
    try:
        if not os.path.isdir(restore_source_dir): result_queue.put(("Error", f"Backup dir '{restore_source_dir}' not found.")); return
        current_data_exists = os.path.isdir(target_dir); pre_restore_backup_path = None
        if current_data_exists:
            os.makedirs(pre_restore_base_dir, exist_ok=True)
            pre_restore_backup_path = os.path.join(pre_restore_base_dir, f"pre_restore_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}")
            try: shutil.copytree(target_dir, pre_restore_backup_path); print(f"Backed up current data to '{pre_restore_backup_path}'")
            except Exception as backup_err: result_queue.put(("Warning", f"Could not back up live data: {backup_err}.\nProceeding will overwrite.")) # Or cancel?
        if current_data_exists:
            try: shutil.rmtree(target_dir); print(f"Removed current data dir '{target_dir}'.")
            except Exception as rm_err: result_queue.put(("Error", f"Failed remove current dir '{target_dir}': {rm_err}\nAbort.")); return
        try:
            shutil.copytree(restore_source_dir, target_dir)
            result_queue.put(("Success", f"Data restored from:\n{os.path.basename(restore_source_dir)}\n\nRestart required."))
        except Exception as copy_err:
            result_queue.put(("Error", f"Failed copy backup to '{target_dir}': {copy_err}\nRestore failed."))
            if pre_restore_backup_path and os.path.exists(pre_restore_backup_path):
                 try: shutil.copytree(pre_restore_backup_path, target_dir); result_queue.put(("Error", f"Copy backup failed: {copy_err}\nBUT pre-restore backup restored."))
                 except Exception as restore_err: result_queue.put(("Error", f"Copy backup failed: {copy_err}\nCRITICAL: Failed restore pre-backup: {restore_err}\nManual fix needed."))
    except Exception as e: result_queue.put(("Error", f"Unexpected restore error: {e}\n{traceback.format_exc()}"))

def restore_all_data(root):
    initial_backup_dir = os.path.abspath("eaze_inn_json_backup") if os.path.exists("eaze_inn_json_backup") else os.path.abspath(".")
    restore_dir = filedialog.askdirectory(title="Select Specific Backup Folder", initialdir=initial_backup_dir, parent=root)
    if not restore_dir: return
    if not messagebox.askyesno("Confirm Restore", f"!!! WARNING !!!\n\nThis will DELETE current data ('{os.path.abspath(DATA_DIR)}') "
        f"and replace it with:\n'{os.path.basename(restore_dir)}'\n\nA backup of current data will be attempted.\n\nProceed?", icon='warning', parent=root): return
    result_queue = queue.Queue()
    thread = threading.Thread(target=restore_all_data_threaded, args=(restore_dir, result_queue), daemon=True)
    thread.start()
    root.after(100, lambda: check_thread_queue(root, result_queue, "Restore"))

def check_thread_queue(root, result_queue, operation_name):
    try:
        status, message = result_queue.get_nowait()
        if status == "Error": messagebox.showerror(f"{operation_name} Error", message, parent=root)
        elif status == "Warning": messagebox.showwarning(f"{operation_name} Warning", message, parent=root)
        elif status == "Success":
            # The message for PDF generation is now the file path, so we adjust the success message
            if "Generation" in operation_name or "Excel" in operation_name:
                messagebox.showinfo(f"{operation_name} Success", f"File saved successfully at:\n{message}", parent=root)
            elif "Sharing" in operation_name:
                 messagebox.showinfo(f"{operation_name} Success", message, parent=root)
            else: # All other successes (Backup, Restore, etc.)
                messagebox.showinfo(f"{operation_name} Success", message, parent=root)
            
            if operation_name == "Restore":
                 messagebox.showinfo("Restart Required", "Data restored.\nPlease restart the application.", parent=root, icon='info')
                 root.quit()
        elif status == "Cancelled": messagebox.showinfo(f"{operation_name} Cancelled", message, parent=root)
    except queue.Empty: root.after(100, lambda: check_thread_queue(root, result_queue, operation_name))
    except Exception as e: messagebox.showerror("Queue Check Error", f"Error checking {operation_name} result: {e}", parent=root)

# --- PDF/Excel Generation ---
class QRCodeFlowable(Flowable):
    def __init__(self, qr_path, width, height): Flowable.__init__(self); self.qr_path = qr_path; self.width = width; self.height = height
    def draw(self):
        try: img = ReportlabImage(self.qr_path, width=self.width, height=self.height); img.hAlign = 'RIGHT'; img.drawOn(self.canv, 0, 0)
        except FileNotFoundError: print(f"QR Code image not found: {self.qr_path}")
        except Exception as e: print(f"Error drawing QR Code: {e}")

def generate_pdf_invoice_threaded(invoice_id, invoice_type, entity_name, invoice_data, invoice_items_dec, result_queue):
    global COMPANY_SETTINGS; entity_label = invoice_type.capitalize()
    pdf_file = f"{invoice_type}_{entity_name.replace(' ','_')}_{invoice_id}_{datetime.datetime.now().strftime('%Y%m%d')}.pdf"
    title = "TAX INVOICE" if invoice_type == 'customer' else "SUPPLIER BILL"
    try:
        pdf = SimpleDocTemplate(pdf_file, pagesize=letter, topMargin=0.5*inch, bottomMargin=0.5*inch, leftMargin=0.7*inch, rightMargin=0.7*inch)
        styles = getSampleStyleSheet(); styles.add(ParagraphStyle(name='small', parent=styles['Normal'], fontSize=8)); styles.add(ParagraphStyle(name='RightAlign', parent=styles['Normal'], alignment=2))
        styles['h1'].alignment = 1; styles['h2'].alignment = 1; story = []
        company_name = COMPANY_SETTINGS.get('company_name', DEFAULT_SETTINGS['company_name']); company_address = COMPANY_SETTINGS.get('company_address', DEFAULT_SETTINGS['company_address'])
        company_email = COMPANY_SETTINGS.get('company_email', DEFAULT_SETTINGS['company_email']); company_phone = COMPANY_SETTINGS.get('company_phone', DEFAULT_SETTINGS['company_phone'])
        company_gstin = COMPANY_SETTINGS.get('company_gstin', None); qr_code_rel_path = COMPANY_SETTINGS.get('qr_code_path', None)
        header_text = [Paragraph(f"<b>{company_name}</b>", styles['h1']), Paragraph(company_address, styles['Normal']), Paragraph(f"M: {company_phone} | Email: {company_email}", styles['Normal'])]
        if company_gstin: header_text.append(Paragraph(f"GSTIN: {company_gstin}", styles['Normal']))
        qr_flowable = None
        if qr_code_rel_path:
            qr_full_path = os.path.join(DATA_DIR, qr_code_rel_path) # Construct full path from relative
            if os.path.exists(qr_full_path):
                try: qr_size = 0.8 * inch; qr_flowable = QRCodeFlowable(qr_full_path, qr_size, qr_size)
                except Exception as qr_err: print(f"Error creating QR flowable: {qr_err}")
        if qr_flowable:
            header_table_data = [[Table([ [p] for p in header_text ], style=TableStyle([('BOTTOMPADDING', (0,0), (0,-1), 1)])), qr_flowable]]
            header_table = Table(header_table_data, colWidths=[letter[0] - 1.4*inch - 1*inch, 0.8*inch]); header_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP'), ('ALIGN', (1,0), (1,0), 'RIGHT')])); story.append(header_table)
        else: story.extend(header_text)
        story.append(Spacer(1, 0.1*inch)); story.append(Paragraph('<hr width="100%" color="black" size="1"/>', styles['Normal'])); story.append(Spacer(1, 0.1*inch))
        story.append(Paragraph(f"<b>{title}</b>", styles['h2'])); story.append(Spacer(1, 0.1*inch))
        info_data = [[Paragraph(f"{title} No: <b>{invoice_id}</b>", styles['Normal']), Paragraph(f"Date: <b>{invoice_data.get('date', 'N/A')}</b>", styles['RightAlign'])], [Paragraph(f"<u>{entity_label} Details:</u>", styles['h4']), ""], [Paragraph(f"Name: {entity_name}", styles['Normal']), ""]]
        info_table = Table(info_data, colWidths=[3.5*inch, 3.0*inch]); info_table.setStyle(TableStyle([('ALIGN', (0, 0), (0, -1), 'LEFT'), ('ALIGN', (1, 0), (1, -1), 'RIGHT'), ('VALIGN', (0, 0), (-1, -1), 'TOP'), ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'), ('SPAN', (0, 1), (1, 1)), ('BOTTOMPADDING', (0, 0), (-1, -1), 3), ('TOPPADDING', (0, 0), (-1, -1), 3),])); story.append(info_table); story.append(Spacer(1, 0.2*inch))
        table_header = [Paragraph("<b>S.N</b>", styles['Normal']), Paragraph("<b>Item Description</b>", styles['Normal']), Paragraph("<b>Qty</b>", styles['Normal']), Paragraph("<b>Rate</b>", styles['Normal']), Paragraph("<b>Amount</b>", styles['Normal'])]
        table_data = [table_header]; total_amount = ZERO_DECIMAL; sn = 1
        for item in invoice_items_dec:
            try:
                item_name = item.get('item', 'N/A'); qty = item.get('quantity', ZERO_DECIMAL); price = item.get('price', ZERO_DECIMAL); amount = (qty * price).quantize(TWO_PLACES, rounding=ROUND_HALF_UP); total_amount += amount
                row_data = [Paragraph(str(sn), styles['Normal']), Paragraph(str(item_name), styles['Normal']), Paragraph(format_decimal_quantity(qty), styles['Normal']), Paragraph(format_currency(price), styles['Normal']), Paragraph(format_currency(amount), styles['Normal'])]
                table_data.append(row_data); sn += 1
            except Exception as item_err: table_data.append([Paragraph(str(sn), styles['Normal']), Paragraph(f"Err: {item_err}", styles['small']), "", "", ""]); sn += 1
        items_table = Table(table_data, colWidths=[0.5*inch, 3.0*inch, 0.7*inch, 1.0*inch, 1.1*inch]); items_table.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.grey), ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke), ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('ALIGN', (1, 1), (1, -1), 'LEFT'), ('ALIGN', (2, 1), (-1, -1), 'RIGHT'), ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'), ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'), ('FONTSIZE', (0, 0), (-1, -1), 9), ('BOTTOMPADDING', (0, 0), (-1, 0), 8), ('TOPPADDING', (0, 0), (-1, 0), 4), ('BOTTOMPADDING', (0, 1), (-1, -1), 4), ('TOPPADDING', (0, 1), (-1, -1), 4), ('GRID', (0, 0), (-1, -1), 1, colors.black), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')])); story.append(items_table); story.append(Spacer(1, 0.1*inch))
        totals_data = [['', '', Paragraph('<b>Total Amount:</b>', styles['Normal']), Paragraph(f"<b>{format_currency(total_amount)}</b>", styles['Normal'])]]; totals_table = Table(totals_data, colWidths=[3.5*inch + 0.7*inch, 1.0*inch, 1.1*inch]); totals_table.setStyle(TableStyle([('ALIGN', (0, 0), (-1, -1), 'RIGHT'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'), ('FONTSIZE', (0, 0), (-1, -1), 10), ('BOTTOMPADDING', (0, 0), (-1, -1), 5), ('TOPPADDING', (0, 0), (-1, -1), 5)])); story.append(totals_table); story.append(Spacer(1, 0.3*inch))
        story.append(Paragraph("<u>Terms & Conditions:</u>", styles['h4'])); story.append(Paragraph("1. Goods once sold will not be taken back.", styles['small'])); story.append(Paragraph("2. Interest @18% p.a. charged if bill not paid within 30 days.", styles['small'])); story.append(Paragraph("3. Subject to Jalandhar jurisdiction only.", styles['small'])); story.append(Spacer(1, 0.5*inch)); story.append(Paragraph(f"For {company_name}", styles['Normal'])); story.append(Spacer(1, 0.5*inch)); story.append(Paragraph("Authorised Signatory", styles['Normal']))
        pdf.build(story)
        # On success, put the FULL PATH into the queue
        result_queue.put(("Success", os.path.abspath(pdf_file)))
    except Exception as e: result_queue.put(("Error", f"PDF generation failed: {e}\n{traceback.format_exc()}"))

def create_invoice_window(title_text, entity_label_text, invoice_type, parent):
    invoice_window = tk.Toplevel(parent)
    invoice_window.title(title_text)
    invoice_window.geometry("750x600")
    invoice_window.transient(parent)
    invoice_window.grab_set()
    
    main_frame = ttk.Frame(invoice_window, padding="10")
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # Basic Information Section
    info_frame = ttk.LabelFrame(main_frame, text="Basic Information", padding="10")
    info_frame.pack(fill=tk.X, padx=5, pady=5)
    
    ttk.Label(info_frame, text=entity_label_text).grid(row=0, column=0, sticky="w", padx=5, pady=3)
    entity_var = tk.StringVar()
    if invoice_type == 'customer':
        entity_list = sorted(list(set(inv['customer_name'] for inv in INVOICES_DATA if 'customer_name' in inv)))
    else:
        entity_list = sorted(list(set(bill['supplier_name'] for bill in SUPPLIER_INVOICES_DATA if 'supplier_name' in bill)))
    
    entity_combo = ttk.Combobox(info_frame, textvariable=entity_var, values=entity_list)
    entity_combo.grid(row=0, column=1, sticky="ew", padx=5, pady=3)
    
    ttk.Label(info_frame, text="Date:").grid(row=1, column=0, sticky="w", padx=5, pady=3)
    date_var = tk.StringVar(value=datetime.datetime.now().strftime(DATE_FORMAT))
    date_entry = ttk.Entry(info_frame, textvariable=date_var)
    date_entry.grid(row=1, column=1, sticky="w", padx=5, pady=3)
    
    # Items Section
    items_frame = ttk.LabelFrame(main_frame, text="Items", padding="10")
    items_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    
    # Create treeview for items
    columns = ("Item", "Quantity", "Price", "Total")
    tree = ttk.Treeview(items_frame, columns=columns, show="headings")
    
    # Configure column widths and headings
    tree.column("Item", width=300)
    tree.column("Quantity", width=100)
    tree.column("Price", width=100)
    tree.column("Total", width=100)
    
    for col in columns:
        tree.heading(col, text=col)
    
    tree_scroll = ttk.Scrollbar(items_frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=tree_scroll.set)
    
    tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    
    # Add Item Section
    add_item_frame = ttk.Frame(main_frame)
    add_item_frame.pack(fill=tk.X, padx=5, pady=5)
    
    ttk.Label(add_item_frame, text="Item:").pack(side=tk.LEFT, padx=2)
    item_var = tk.StringVar()
    ttk.Entry(add_item_frame, textvariable=item_var).pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
    
    ttk.Label(add_item_frame, text="Qty:").pack(side=tk.LEFT, padx=2)
    qty_var = tk.StringVar()
    ttk.Entry(add_item_frame, textvariable=qty_var, width=8).pack(side=tk.LEFT, padx=2)
    
    ttk.Label(add_item_frame, text="Price:").pack(side=tk.LEFT, padx=2)
    price_var = tk.StringVar()
    ttk.Entry(add_item_frame, textvariable=price_var, width=10).pack(side=tk.LEFT, padx=2)
    
    def add_item():
        item = item_var.get().strip()
        if not item:
            messagebox.showerror("Error", "Please enter an item name", parent=invoice_window)
            return
            
        try:
            qty = Decimal(qty_var.get())
            price = Decimal(price_var.get())
            if qty <= 0:
                raise ValueError("Quantity must be positive")
            if price < 0:
                raise ValueError("Price cannot be negative")
                
            total = qty * price
            tree.insert("", tk.END, values=(
                item,
                format_decimal_quantity(qty),
                format_currency(price),
                format_currency(total)
            ))
            
            # Clear inputs
            item_var.set("")
            qty_var.set("")
            price_var.set("")
            
        except (InvalidOperation, ValueError) as e:
            messagebox.showerror("Error", str(e), parent=invoice_window)
    
    ttk.Button(add_item_frame, text="Add Item", command=add_item).pack(side=tk.LEFT, padx=5)
    
    # Bottom Buttons
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill=tk.X, pady=10)
    
    def save_invoice():
        # Get entity name
        entity_name = entity_var.get().strip()
        if not entity_name:
            messagebox.showerror("Error", f"Please select or enter {entity_label_text.lower().replace(':', '')}", 
                               parent=invoice_window)
            return
            
        # Get date
        try:
            invoice_date = datetime.datetime.strptime(date_var.get(), DATE_FORMAT)
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid date (YYYY-MM-DD)", parent=invoice_window)
            return
            
        # Get items
        if not tree.get_children():
            messagebox.showerror("Error", "Please add at least one item", parent=invoice_window)
            return
            
        # Create new invoice ID
        new_id = get_next_id(INVOICES_DATA if invoice_type == 'customer' else SUPPLIER_INVOICES_DATA)
        
        # Process items
        items = []
        for item_id in tree.get_children():
            item_values = tree.item(item_id)['values']
            items.append({
                'item': item_values[0],
                'quantity': Decimal(str(item_values[1])),
                'price': Decimal(str(item_values[2]).replace(CURRENCY_SYMBOL, '')),
            })
            
        # Save invoice
        invoice_data = {
            'id': new_id,
            'date': invoice_date.strftime(DATE_FORMAT),
            'customer_name' if invoice_type == 'customer' else 'supplier_name': entity_name,
            'payment_status': 'P'  # Default to Pending
        }
        
        if invoice_type == 'customer':
            INVOICES_DATA.append(invoice_data)
            save_data(INVOICES_DATA, INVOICES_FILE)
            
            for item in items:
                item_id = get_next_id(INVOICE_ITEMS_DATA)
                INVOICE_ITEMS_DATA.append({
                    'id': item_id,
                    'invoice_id': new_id,
                    **item
                })
            save_data(INVOICE_ITEMS_DATA, INVOICE_ITEMS_FILE)
        else:
            SUPPLIER_INVOICES_DATA.append(invoice_data)
            save_data(SUPPLIER_INVOICES_DATA, SUPPLIER_INVOICES_FILE)
            
            for item in items:
                item_id = get_next_id(SUPPLIER_INVOICE_ITEMS_DATA)
                SUPPLIER_INVOICE_ITEMS_DATA.append({
                    'id': item_id,
                    'supplier_invoice_id': new_id,
                    **item
                })
            save_data(SUPPLIER_INVOICE_ITEMS_DATA, SUPPLIER_INVOICE_ITEMS_FILE)
            
        messagebox.showinfo("Success", 
                           f"{'Invoice' if invoice_type == 'customer' else 'Bill'} #{new_id} created successfully",
                           parent=invoice_window)
        invoice_window.destroy()
    
    ttk.Button(button_frame, text="Save", command=save_invoice).pack(side=tk.RIGHT, padx=5)
    ttk.Button(button_frame, text="Cancel", command=invoice_window.destroy).pack(side=tk.RIGHT, padx=5)
    
    entity_combo.focus_set()

def create_dashboard(root):
    dashboard_window = tk.Toplevel(root)
    dashboard_window.title("Dashboard")
    dashboard_window.geometry("800x600")
    dashboard_window.transient(root)
    dashboard_window.grab_set()
    
    main_frame = ttk.Frame(dashboard_window, padding="10")
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    title_label = ttk.Label(main_frame, text="Eaze Inn Accounts Dashboard", font=('TkDefaultFont', 16, 'bold'))
    title_label.pack(pady=(0, 20))
    
    summary_frame = ttk.Frame(main_frame)
    summary_frame.pack(fill=tk.X, pady=(0, 20))
    
    # Calculate values
    total_receivables = ZERO_DECIMAL
    for inv in INVOICES_DATA:
        if inv.get('payment_status', 'P') == 'P':
            total_receivables += calculate_invoice_total(inv['id'], 'customer')
    
    total_payables = ZERO_DECIMAL
    for bill in SUPPLIER_INVOICES_DATA:
        if bill.get('payment_status', 'P') == 'P':
            total_payables += calculate_invoice_total(bill['id'], 'supplier')
    
    inventory_value = sum((item.get('quantity', ZERO_DECIMAL) * item.get('value', ZERO_DECIMAL)) 
                         for item in INVENTORY_DATA if item.get('quantity', ZERO_DECIMAL) > ZERO_DECIMAL)
    
    summary_cards_data = [
        ("Pending Receivables", format_currency(total_receivables), "blue"),
        ("Pending Payables", format_currency(total_payables), "red"),
        ("Inventory Value (Est.)", format_currency(inventory_value), "green")
    ]
    
    # Configure grid columns for equal spacing
    for i in range(3):
        summary_frame.grid_columnconfigure(i, weight=1)
    
    # Create simple cards without custom styles
    for i, (title, value, color) in enumerate(summary_cards_data):
        card_frame = ttk.LabelFrame(summary_frame, text=title, padding="15")
        card_frame.grid(row=0, column=i, padx=10, pady=5, sticky="news")
        
        value_label = ttk.Label(card_frame, text=value, font=('TkDefaultFont', 14, 'bold'), foreground=color)
        value_label.pack(pady=10, expand=True)
    
    actions_frame = ttk.LabelFrame(main_frame, text="Quick Actions", padding="10")
    actions_frame.pack(fill=tk.X, pady=20)
    
    def open_enter_payment(root):
        messagebox.showinfo("Not Implemented", "The 'Enter Payment' feature is not yet implemented.", parent=root)
    
    # Create buttons in a 2x2 grid
    ttk.Button(actions_frame, text="New Customer Invoice",
               command=lambda: create_invoice_window("Create Customer Invoice", "Customer Name:", "customer", dashboard_window)) \
        .grid(row=0, column=0, padx=5, pady=5, sticky="ew")
    
    ttk.Button(actions_frame, text="New Supplier Bill",
               command=lambda: create_invoice_window("Create Supplier Bill", "Supplier Name:", "supplier", dashboard_window)) \
        .grid(row=0, column=1, padx=5, pady=5, sticky="ew")
    
    ttk.Button(actions_frame, text="Enter Payment",
               command=lambda: open_enter_payment(dashboard_window)) \
        .grid(row=0, column=2, padx=5, pady=5, sticky="ew")
    
    for i in range(3):
        actions_frame.grid_columnconfigure(i, weight=1)
    
    def on_dashboard_closing():
        if messagebox.askokcancel("Quit", "Do you want to exit the application?", parent=dashboard_window):
            root.quit()
        
    dashboard_window.protocol("WM_DELETE_WINDOW", on_dashboard_closing)
    
    # Bottom buttons
    ttk.Button(main_frame, text="Close Application", 
               command=on_dashboard_closing) \
        .pack(pady=20, side=tk.BOTTOM)

# --- Main Function ---
def main():
    root = tk.Tk(); root.title("Eaze Inn Accounts - Login"); root.geometry("350x250"); root.resizable(False, False);
    root.update_idletasks(); width = root.winfo_width(); height = root.winfo_height(); x_pos = (root.winfo_screenwidth() // 2) - (width // 2); y_pos = (root.winfo_screenheight() // 2) - (height // 2); root.geometry(f'{width}x{height}+{x_pos}+{y_pos}')
    style = ttk.Style(root)
    try:
        available_themes = style.theme_names()
        if 'clam' in available_themes: style.theme_use('clam')
        elif 'vista' in available_themes and os.name == 'nt': style.theme_use('vista')
        elif 'aqua' in available_themes and sys.platform == "darwin": style.theme_use('aqua')
        else: style.theme_use(style.theme_use())
    except tk.TclError: print("Theming engine error or selected theme not available. Using default.")
    style.configure("TLabel", padding=2); style.configure("TButton", padding=5, font=('TkDefaultFont', 10)); style.configure("Accent.TButton", font=('TkDefaultFont', 10, 'bold')); style.configure("TLabelframe.Label", font=('TkDefaultFont', 10, 'bold'))
    login_outer_frame = ttk.Frame(root, padding="20"); login_outer_frame.pack(expand=True, fill=tk.BOTH)
    login_frame = ttk.LabelFrame(login_outer_frame, text="Login or Register", padding=20); login_frame.pack(expand=True)
    ttk.Button(login_frame, text="Sign In", command=lambda: signin_window(root), width=20, style="Accent.TButton").pack(pady=10, ipady=5)
    ttk.Button(login_frame, text="Register New User", command=lambda: register_window(root), width=20).pack(pady=10, ipady=5)
    def on_closing_main_app():
        if messagebox.askokcancel("Quit", "Are you sure you want to exit Eaze Inn Accounts?", parent=root, icon=messagebox.WARNING): print("Exit confirmed by user."); root.quit()
    root.protocol("WM_DELETE_WINDOW", on_closing_main_app); root.mainloop(); print("Application main loop finished.")

if __name__ == "__main__":
    try:
        print(f"--- Starting Eaze Inn Accounts (JSON Version) [{datetime.datetime.now()}] ---")
        os.makedirs(DATA_DIR, exist_ok=True); print(f"Data directory: '{os.path.abspath(DATA_DIR)}'"); os.makedirs(IMAGES_DIR, exist_ok=True)
        if THERMAL_PRINTER_TYPE == 'win32raw' and not win32print_installed and os.name == 'nt': print("\nWARNING: pywin32 library not found, but required for 'win32raw' printer type.\n         Install using: pip install pywin32\n")
        # New check for matplotlib
        if not matplotlib_installed: print("\nWARNING: matplotlib not found. EazeBot charting will be disabled.\n         Install using: pip install matplotlib\n")
        load_all_data(); print("Starting main application UI..."); main()
    except Exception as e_global:
         print(f"\n--- FATAL APPLICATION ERROR ---"); print(f"Error Type: {type(e_global).__name__}"); print(f"Error: {e_global}"); print(traceback.format_exc()); print("-------------------------------")
         try:
             root_err_popup = tk.Tk(); root_err_popup.withdraw()
             messagebox.showerror("Critical Application Error", f"A critical error occurred:\n\n{type(e_global).__name__}: {e_global}\n\nPlease check the console log (terminal) for detailed information.\nThe application will now close.", icon='error')
             root_err_popup.destroy()
         except Exception as tk_err_popup: print(f"Could not display Tkinter error message box: {tk_err_popup}")
         finally: print("\nApplication encountered a fatal error. Exiting.")
    finally: print(f"--- Eaze Inn Accounts finished [{datetime.datetime.now()}] ---")
