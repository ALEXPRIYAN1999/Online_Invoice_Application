import tkinter as tk
from tkinter import ttk, messagebox, PhotoImage, filedialog
from PIL import Image, ImageTk
from datetime import datetime, timedelta
from fpdf import FPDF
from num2words import num2words
import fitz 
import json
import os
import webbrowser
import subprocess
import re
import sys
import glob
from PyPDF2 import PdfMerger
import firebase_admin
from firebase_admin import credentials, db


# ---- Flexible date parser (used in multiple places) ----
def parse_date_flexible(date_str):
    """
    Try several date formats and return a datetime object.
    Works with 28/01/2025, 28.01.2025, 28-01-2025, 2025-01-28, etc.
    """
    if not date_str:
        return None

    s = str(date_str).strip()
    formats = ("%d/%m/%Y", "%d.%m.%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y.%m.%d")

    for fmt in formats:
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            continue

    try:
        return datetime.fromisoformat(s)
    except Exception:
        pass

    print(f"DEBUG: Unrecognized date format: {s}")
    return None
# ---------------------------------------------------------


class ModernInvoiceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🚀  Invoice Pro - Modern Edition")
        self.root.geometry("{0}x{1}+0+0".format(root.winfo_screenwidth(), root.winfo_screenheight()))
        
        # Modern color scheme
        self.colors = {
            'primary': "#C62828",
            'secondary': '#283593',
            'accent': '#3949ab',
            'success': '#2e7d32',
            'warning': '#d32f2f',
            'info': '#0288d1',
            'light_bg': '#f5f5f5',
            'dark_bg': '#121212',
            'card_bg': '#ffffff',
            'text_light': '#ffffff',
            'text_dark': '#212121',
            'text_muted': '#757575',
            'focus_border': '#ff4081',
            'hover_effect': '#e3f2fd'
        }

        self.current_theme = "light"  # Default theme
        try:
            # Try to load saved theme
            home_dir = os.path.expanduser("~")
            settings_path = os.path.join(home_dir, "settings.json")
            if os.path.exists(settings_path):
                with open(settings_path, 'r') as f:
                    settings = json.load(f)
                    if "theme" in settings:
                        self.current_theme = settings["theme"]
        except Exception as e:
            print(f"Error loading theme settings: {e}")
        
        # Initialize color scheme based on current theme
        self.initialize_theme_colors()
        
        self.root.configure(bg=self.colors['light_bg'])

        # Initialize company selection with default value
        self.selected_office = "A1"  # Default to Angel Pyrotech
        
        # Dictionary to track the highest suffix for each base bill number
        self.bill_suffix_tracker = {}
        
        # Global counter to track the suffix for all bills
        self.bill_suffix_counter = 0

        #self.show_login_page()
        self.bill_no_edited = False 

        self.firebase_connected = False
        
        # Initialize product_frame
        self.product_frame = tk.Frame(self.root)
        self.product_frame.pack(fill=tk.X, pady=10, padx=20)

        # -------------------- FIREBASE CONNECTION (START) --------------------
        try:
            # Auto-detect user home folder (works in all 3 systems)
            FIREBASE_KEY_PATH = os.path.join(os.path.expanduser("~"), "billing_key_Invoice.json")
            
            print(f"Looking for key at: {FIREBASE_KEY_PATH}")
            print(f"File exists: {os.path.exists(FIREBASE_KEY_PATH)}")

            # Your Firebase Realtime Database URL
            DATABASE_URL = "https://onlineinvoiceapplication-default-rtdb.firebaseio.com/"

            # Initialize Firebase only once
            if not firebase_admin._apps:
                cred = credentials.Certificate(FIREBASE_KEY_PATH)
                firebase_admin.initialize_app(cred, {
                    'databaseURL': DATABASE_URL
                })

            # Database table references
            self.party_ref = db.reference('party_data')
            self.product_ref = db.reference('product_data')
            self.bills_ref = db.reference('bills')

            self.party_data = self.load_data(self.party_ref)
            self.product_data = self.clean_product_keys(self.load_data(self.product_ref))  # CLEAN HERE
            self.bills_data = self.load_data(self.bills_ref)

            # Add this in your __init__ method after loading product_data
            self.clean_all_product_keys()


            print("🔥 Firebase connected successfully.")
            self.firebase_connected = True


        except Exception as e:
            messagebox.showerror("Error", f"Firebase Connection Failed:\n{e}")
            self.root.destroy()
        # -------------------- FIREBASE CONNECTION (END) --------------------


        # In your __init__ method, replace the customer_names initialization:
        self.customer_names = []

        for key, pdata in self.party_data.items():
            # Handle multiple possible key variations
            name = (
                pdata.get('Customer Name') or 
                pdata.get('Customer_Name') or 
                pdata.get('customer_name') or
                pdata.get('Customer  Name') or
                pdata.get('Customer name') or
                pdata.get(' Customer Name') or
                ""
            )
            if name and name not in self.customer_names:
                self.customer_names.append(name)

        
        # NEW CODE:
        self.product_names = []
        for key in self.product_data.keys():
            product = self.product_data[key]
            # Handle both old and new field names
            if 'Product_Name' in product:
                self.product_names.append(product['Product_Name'])
            elif 'Product Name' in product:
                self.product_names.append(product['Product Name'])
            else:
                print(f"⚠️ WARNING: Product {key} has no name field")
        

        # Migrate absolute paths to relative paths (one-time operation) - FROM ORIGINAL
        self.migrate_absolute_paths_to_relative()
        
        
        #self.set_window_icon()
        self.screen_stack = []  # Stack to manage screen navigation
               
        self.bill_data = []  # Initialize bill data list

        # Variables - FROM ORIGINAL APPLICATION
        self.bill_no = self.get_next_bill_number()
        
        self.bill_date = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        self.lr_number = tk.StringVar()
        self.to_address = tk.StringVar()
        self.to_gstin = tk.StringVar()
        self.to_name = tk.StringVar()
        self.agent_name = tk.StringVar()
        self.document_through = tk.StringVar()
        self.from_ = tk.StringVar()
        self.to_ = tk.StringVar()
        self.product_code_2 = tk.StringVar()
        self.rate = tk.DoubleVar()  # Or tk.StringVar() depending on the data type
        self.product_name = tk.StringVar()
        self.type = tk.StringVar()  # Ensure this is defined
        self.case_details = tk.StringVar()
        self.quantity = tk.StringVar()
        self.discount = tk.StringVar()
        self.amount = tk.StringVar()
        self.region = tk.StringVar(value="South")  # Default region
        self.gst_percentage = tk.DoubleVar(value=18.0)  # Default GST percentage
        self.cgst = tk.DoubleVar()
        self.sgst = tk.DoubleVar()
        self.igst = tk.DoubleVar()
        self.total_amount =  tk.DoubleVar()
        self.cgst_amount_1 = tk.DoubleVar()
        self.sgst_amount_1 = tk.DoubleVar()
        self.igst_amount_1 = tk.DoubleVar()
        self.discount_percentage =  tk.DoubleVar()
        self.packing_charge = tk.DoubleVar()
        self.discount_amount_field = tk.DoubleVar()
        self.before_discount_amount_field = tk.DoubleVar()
        self.amount_before_discount = tk.DoubleVar()
        self.After_discount_amount_field = tk.DoubleVar()
        self.After_Discount_Total_Amount = tk.DoubleVar()
        self.Packing_Amount = tk.DoubleVar()
        self.per = tk.StringVar()
        per_field = tk.StringVar()
        self.total_no_of_cases = tk.StringVar()
        self.rate = tk.DoubleVar(value=0.0)  # Initialize with 0.0 instead of empty
        self.quantity = tk.IntVar(value=0)   # Initialize with 0 instead of empty string
        self.discount = tk.DoubleVar(value=0.0)  # Initialize with 0.0

        # Save data on exit
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # Keyboard navigation variables
        self.current_focus_index = 0
        self.focusable_widgets = []
        self.undo_stack = []
        self.redo_stack = []
        self.max_undo_steps = 50
        self.current_screen = None
        
        
        # Bind global keyboard shortcuts
        self.setup_global_keyboard_shortcuts()
        
        # Show modern login
        self.show_modern_login()
        
        # Save data on exit
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ========== ORIGINAL APPLICATION METHODS - DATA MANAGEMENT ==========
    
    def load_data(self, ref):
        """Load data from Firebase and handle different data structures"""
        try:
            # If ref is a string (old filename), return empty dict
            if isinstance(ref, str):
                return {}

            data = ref.get()  # Get from Firebase

            # If nothing in Firebase → return empty dict
            if data is None:
                return {}

            # Handle LIST data (Firebase array)
            if isinstance(data, list):
                data_dict = {}
                for i, item in enumerate(data):
                    if item is not None:  # Skip None values
                        # Use index as key, or try to find a better key
                        key = str(i)
                        # If item has a code/ID field, use that instead
                        if isinstance(item, dict) and 'Product_Name' in item:
                            # Try to use product name or create a key
                            key = f"product_{i}"
                        data_dict[key] = item
                return data_dict
            # If already dictionary → return as-is
            if isinstance(data, dict):
                return data

            
            return {}

        except Exception as e:
            print(f"❌ load_data error: {e}")
            return {}


    def save_data(self, ref, data):
        """
        Save data to Firebase Realtime Database.
        'ref' is a Firebase Database reference.
        """
        try:
            ref.set(data)
            print("✅ Firebase save successful")
            return True
        except Exception as e:
            print(f"❌ Firebase Save Error: {e}")
            return False

    def on_close(self):
        """Simple application shutdown with data saving - FROM ORIGINAL"""
        # Save all data
        self.save_data(self.party_ref, self.party_data)
        self.save_data(self.product_ref, self.product_data)
        self.save_data(self.bills_ref, self.bills_data)
        
        # Close application
        self.root.destroy()
        sys.exit(0)

    def get_next_bill_number(self):
        """Simplified bill number generation - FROM ORIGINAL"""
        try:
            office = getattr(self, 'selected_office', 'A1')
            prefix_map = {"A1": "AP", "A2": "AFI", "A3": "AFF"}
            prefix = prefix_map.get(office, "AP")
            
            # Find existing bills with this prefix
            prefix_bills = [b for b in self.bills_data.keys() if b.startswith(prefix)]
            
            if not prefix_bills:
                return f"{prefix}001"
            
            # Extract numbers safely
            numbers = []
            for bill in prefix_bills:
                try:
                    num_part = bill[len(prefix):].split('_')[0]
                    if num_part.isdigit():
                        numbers.append(int(num_part))
                except:
                    continue
            
            next_num = max(numbers) + 1 if numbers else 1
            return f"{prefix}{next_num:03d}"
            
        except Exception as e:
            print(f"Bill number error: {e}")
        return "AP001"  # Simple fallback

    def migrate_absolute_paths_to_relative(self):
        """
        Convert existing absolute paths in bills_data to relative paths
        This should be called once when the application starts - FROM ORIGINAL
        """
        documents_dir = os.path.join(os.path.expanduser("~"), "Documents")
        invoice_app_base = os.path.join(documents_dir, "InvoiceApp")
        
        migrated_count = 0
        for bill_no, bill_data in self.bills_data.items():
            pdf_path = bill_data.get("pdf_file_name", "")
            if pdf_path and os.path.isabs(pdf_path):
                # Check if this is an InvoiceApp path
                if "InvoiceApp" in pdf_path:
                    # Convert to relative path
                    try:
                        # Extract the part after "InvoiceApp"
                        parts = pdf_path.split("InvoiceApp")
                        if len(parts) > 1:
                            relative_path = "InvoiceApp" + parts[1]
                            # Normalize the path (remove any leading/trailing issues)
                            relative_path = os.path.normpath(relative_path)
                            # Update the bill data
                            self.bills_data[bill_no]["pdf_file_name"] = relative_path
                            migrated_count += 1
                    except Exception as e:
                        print(f"Error migrating path for bill {bill_no}: {e}")
        
        if migrated_count > 0:
            # Save the updated data
            self.save_data(self.bills_ref, self.bills_data)
            print(f"Migrated {migrated_count} bill paths from absolute to relative")

    def get_pdf_path(self, relative_path):
        """
        Convert relative PDF path to absolute path based on current user's Documents folder
        FROM ORIGINAL
        """
        try:
            documents_dir = os.path.join(os.path.expanduser("~"), "Documents")
            absolute_path = os.path.join(documents_dir, relative_path)
            
            # Check if file exists at the resolved path
            if os.path.exists(absolute_path):
                return absolute_path
            else:
                # Try to find the file by filename only (fallback)
                filename = os.path.basename(relative_path)
                # Search in the expected directory structure
                expected_dir = os.path.join(documents_dir, "InvoiceApp", "Invoice_Bill")
                for root, dirs, files in os.walk(expected_dir):
                    if filename in files:
                        return os.path.join(root, filename)
                
                # If still not found, return the expected path (file might be missing)
                return absolute_path
                
        except Exception as e:
            print(f"Error resolving PDF path: {e}")
            # Fallback: return the path as-is (might be absolute path from old data)
            return relative_path

    def set_window_icon(self):
        """Sets the window icon to logo.png - FROM ORIGINAL"""
        try:
            # Load the image using PhotoImage for Tkinter or ImageTk for PIL
            logo_image = Image.open("logo.png")
            logo_image = logo_image.resize((64, 64))  # Resize if needed to match the icon size
            logo_photo = ImageTk.PhotoImage(logo_image)
            
            # Set the icon for the main window
            self.root.iconphoto(False, logo_photo)
        except Exception as e:
            print("Error setting logo:", e)

    # ========== ORIGINAL APPLICATION METHODS - CUSTOMER/PRODUCT AUTOCOMPLETE ==========
    
    def update_customer_suggestions(self, event):
        """Update customer name suggestions without interrupting typing. - FROM ORIGINAL"""
        # Only process certain key events to avoid interrupting typing
        if event.keysym in ['Return', 'Up', 'Down', 'Escape']:
            return
        
        typed = self.customer_combobox.get()
        
        # Don't update if the dropdown is open or during navigation
        if self.customer_combobox.focus_get() != self.customer_combobox:
            return
        
        # Update suggestions in the background without forcing dropdown
        suggestions = []
        if typed == '':
            suggestions = self.customer_names
        else:
            suggestions = [name for name in self.customer_names if typed.lower() in name.lower()]
        
        # Update values but don't force dropdown open
        self.customer_combobox['values'] = suggestions

    def fill_customer_details(self, event=None):
        """Fill GST, Phone, Address automatically when a customer name is selected."""
        
        selected_customer_name = self.customer_combobox.get().strip()
        if not selected_customer_name:
            return

        # Accepted key variations
        possible_name_keys = [
            "Customer Name", "Customer_Name", "customer_name",
            "Customer  Name", "Customer name", " Customer Name", 
            "Customer Name ", "Customer_Name ", "customer name"
        ]

        for key, pdata in self.party_data.items():

            # Find the correct customer name key dynamically
            customer_name_value = None
            for nk in possible_name_keys:
                if nk in pdata:
                    customer_name_value = pdata[nk]
                    break

            if not customer_name_value:
                continue

            # Compare selected name
            if customer_name_value.strip().lower() == selected_customer_name.lower():

                # Auto-fill all fields
                self.to_name.set(customer_name_value)
                self.to_address.set(
                    pdata.get("Address") or 
                    pdata.get("address") or 
                    ""
                )
                self.to_gstin.set(
                    pdata.get("GST Number") or 
                    pdata.get("gst") or 
                    pdata.get("GST_Number") or 
                    ""
                )
                self.agent_name.set(
                    pdata.get("Agent Name") or 
                    pdata.get("agent") or 
                    pdata.get("Agent_Name") or 
                    ""
                )
                return

    def handle_manual_customer_entry(self):
        """Handle when user types a customer name manually. - FROM ORIGINAL"""
        typed_name = self.customer_combobox.get().strip()
        if typed_name and typed_name not in self.customer_names:
            # User typed a new customer name manually
            self.to_name.set(typed_name)
            # Clear other fields for new customer
            self.to_address.set("")
            self.to_gstin.set("")
            self.agent_name.set("")

    def update_product_suggestions(self, event):
        """Update product name suggestions as the user types - IMPROVED VERSION. - FROM ORIGINAL"""
        # Only process certain key events to avoid interrupting typing
        if event.keysym in ['Return', 'Up', 'Down', 'Escape']:
            return
        
        typed = self.product_name_combobox.get()
        
        # Don't update if the dropdown is open or during navigation
        if self.product_name_combobox.focus_get() != self.product_name_combobox:
            return
        
        # Update suggestions in the background without forcing dropdown
        suggestions = []
        if typed == '':
            suggestions = self.product_names
        else:
            suggestions = [name for name in self.product_names if typed.lower() in name.lower()]
        
        # Update values but don't force dropdown open
        self.product_name_combobox['values'] = suggestions

    def handle_manual_product_entry(self, event=None):
        """Handle manual product entry - FROM ORIGINAL"""
        typed_name = self.product_name_combobox.get().strip()
        if not typed_name:
            return

        # Existing product - fill details
        for key, value in self.product_data.items():
            if value.get('Product Name', '').lower() == typed_name.lower():
                self.no_of_case_entry.delete(0, tk.END)
                self.no_of_case_entry.insert(0, value.get('No. of Case', ''))
                self.per_case_entry.delete(0, tk.END)
                self.per_case_entry.insert(0, value.get('Per Case', ''))
                self.unit_type_combobox.set(value.get('Unit Type', ''))
                self.rate.set(value.get('Selling Price', ''))
                self.per_entry.delete(0, tk.END)
                self.per_entry.insert(0, value.get('Per', 0))
                self.quantity_entry.delete(0, tk.END)
                self.quantity_entry.insert(0, value.get('Quantity', ''))
                self.discount.set(value.get('Discount', ''))
                return  # stop

        # New product - clear
        self.no_of_case_entry.delete(0, tk.END)
        self.per_case_entry.delete(0, tk.END)
        self.unit_type_combobox.set('')
        self.rate.set('')
        self.per_entry.delete(0, tk.END)
        self.quantity_entry.delete(0, tk.END)
        self.discount.set('')

    # ========== MODERN UI METHODS ==========
    
    def setup_global_keyboard_shortcuts(self):
        """Setup global keyboard shortcuts for the entire application"""
        # Global bindings
        self.root.bind('<Control-z>', self.undo_action)
        self.root.bind('<Control-y>', self.redo_action)
        self.root.bind('<Control-s>', self.quick_save)
        self.root.bind('<Control-n>', self.new_bill_shortcut)
        self.root.bind('<Control-q>', self.quit_application)
        self.root.bind('<F1>', self.show_help)
        self.root.bind('<Escape>', self.focus_escape)
        self.root.bind('<F5>', self.refresh_data)
        
        # Navigation bindings
        self.root.bind('<Tab>', self.focus_next_widget)
        self.root.bind('<Shift-Tab>', self.focus_previous_widget)
        self.root.bind('<Return>', self.enter_key_action)
        self.root.bind('<Up>', self.arrow_key_navigation)
        self.root.bind('<Down>', self.arrow_key_navigation)
        self.root.bind('<Left>', self.arrow_key_navigation)
        self.root.bind('<Right>', self.arrow_key_navigation)
        
        # Function keys
        for i in range(1, 13):
            self.root.bind(f'<F{i}>', self.function_key_handler)

    def create_modern_button(self, parent, text, command, bg_color=None, width=20, height=2, style="primary"):
        """Create a modern 3D-style button"""
        if bg_color is None:
            color_map = {
                "primary": self.colors['accent'],
                "success": self.colors['success'],
                "warning": self.colors['warning'],
                "info": self.colors['info'],
                "secondary": self.colors['secondary']
            }
            bg_color = color_map.get(style, self.colors['accent'])
            
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            font=("Segoe UI", 10, "bold"),
            bg=bg_color,
            fg=self.colors['text_light'],
            relief="raised",
            bd=2,
            width=width,
            height=height,
            cursor="hand2",
            padx=10
        )

        # Add hover effects
        def on_enter(e):
            btn.config(bg=self.darken_color(bg_color, 15), relief="sunken")
            
        def on_leave(e):
            btn.config(bg=bg_color, relief="raised")
            
        def on_press(e):
            btn.config(relief="sunken")
            
        def on_release(e):
            btn.config(relief="raised")
            
        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
        btn.bind("<ButtonPress-1>", on_press)
        btn.bind("<ButtonRelease-1>", on_release)
        
        return btn

    def darken_color(self, color, percent):
        """Darken a hex color by given percent"""
        color = color.lstrip('#')
        rgb = tuple(int(color[i:i+2], 16) for i in (0, 2, 4))
        darkened = tuple(max(0, int(c * (100 - percent) / 100)) for c in rgb)
        return f'#{darkened[0]:02x}{darkened[1]:02x}{darkened[2]:02x}'

    def create_modern_frame(self, parent, title=None, padding=10, bg=None):
        """Create a modern frame with shadow effect"""
        if bg is None:
            bg = self.colors['card_bg']
            
        frame = tk.Frame(
            parent,
            bg=bg,
            relief="flat",
            bd=0,
            highlightbackground=self.colors['accent'],
            highlightthickness=0
        )
        
        if title:
            title_frame = tk.Frame(frame, bg=self.colors['primary'])
            title_frame.pack(fill=tk.X, padx=0, pady=(0, padding))
            
            title_label = tk.Label(
                title_frame,
                text=title,
                font=("Segoe UI", 12, "bold"),
                bg=self.colors['primary'],
                fg=self.colors['text_light'],
                pady=8
            )
            title_label.pack(fill=tk.X)
            
        return frame

    def create_focusable_widget(self, widget, widget_type="entry"):
        """Make a widget focusable and track it for keyboard navigation - FIXED VERSION"""
        # Remove existing bindings to avoid duplicates
        widget.unbind('<Tab>')
        widget.unbind('<Shift-Tab>')
        
        # Mark widget as focusable
        widget._is_focusable = True
        
        # Add focus in/out effects
        def on_focus_in(event):
            self.highlight_focused_widget(event.widget)
            
        def on_focus_out(event):
            self.remove_highlight(event.widget)
        
        widget.bind('<FocusIn>', on_focus_in)
        widget.bind('<FocusOut>', on_focus_out)
        
        # Add tab navigation for all focusable widgets
        widget.bind('<Tab>', self.focus_next_widget)
        widget.bind('<Shift-Tab>', self.focus_previous_widget)
        
        return widget

    def focus_next_widget(self, event=None):
        """Move focus to next widget (Tab key) - FIXED VERSION"""
        # Rebuild focusable widgets list with only existing widgets
        self.rebuild_focusable_widgets()
        
        if not self.focusable_widgets:
            return "break"
        
        # Get current focus
        current_widget = self.root.focus_get()
        
        # Find current index
        if current_widget in self.focusable_widgets:
            self.current_focus_index = self.focusable_widgets.index(current_widget)
        else:
            self.current_focus_index = 0
        
        # Move to next widget
        self.current_focus_index = (self.current_focus_index + 1) % len(self.focusable_widgets)
        next_widget = self.focusable_widgets[self.current_focus_index]
        
        try:
            next_widget.focus_set()
            # For entry widgets, move cursor to end
            if hasattr(next_widget, 'icursor'):
                next_widget.icursor(tk.END)
        except tk.TclError:
            # Widget might be destroyed, try next one
            self.current_focus_index = (self.current_focus_index + 1) % len(self.focusable_widgets)
            if self.focusable_widgets:
                try:
                    self.focusable_widgets[self.current_focus_index].focus_set()
                except tk.TclError:
                    pass
        
        return "break"

    def focus_previous_widget(self, event=None):
        """Move focus to previous widget (Shift+Tab) - FIXED VERSION"""
        # Rebuild focusable widgets list with only existing widgets
        self.rebuild_focusable_widgets()
        
        if not self.focusable_widgets:
            return "break"
        
        # Get current focus
        current_widget = self.root.focus_get()
        
        # Find current index
        if current_widget in self.focusable_widgets:
            self.current_focus_index = self.focusable_widgets.index(current_widget)
        else:
            self.current_focus_index = 0
        
        # Move to previous widget
        self.current_focus_index = (self.current_focus_index - 1) % len(self.focusable_widgets)
        prev_widget = self.focusable_widgets[self.current_focus_index]
        
        try:
            prev_widget.focus_set()
            # For entry widgets, move cursor to end
            if hasattr(prev_widget, 'icursor'):
                prev_widget.icursor(tk.END)
        except tk.TclError:
            # Widget might be destroyed, try previous one
            self.current_focus_index = (self.current_focus_index - 1) % len(self.focusable_widgets)
            if self.focusable_widgets:
                try:
                    self.focusable_widgets[self.current_focus_index].focus_set()
                except tk.TclError:
                    pass
        
        return "break"

    def rebuild_focusable_widgets(self):
        """Rebuild the focusable widgets list with only existing widgets"""
        # Clear the current list
        self.focusable_widgets = []
        
        # Function to recursively find all focusable widgets
        def find_focusable_widgets(widget):
            try:
                # Check if widget exists and is focusable
                if widget.winfo_exists():
                    # Check if it's a focusable widget type
                    if (isinstance(widget, (tk.Entry, tk.Button, ttk.Combobox, ttk.Button)) or 
                        hasattr(widget, '_is_focusable')):
                        self.focusable_widgets.append(widget)
                    
                    # Recursively check children
                    for child in widget.winfo_children():
                        find_focusable_widgets(child)
            except tk.TclError:
                # Widget no longer exists, skip it
                pass
        
        # Start from root and find all focusable widgets
        find_focusable_widgets(self.root)



    def highlight_focused_widget(self, widget):
        """Highlight the currently focused widget"""
        try:
            widget.config(highlightbackground=self.colors['focus_border'], 
                         highlightcolor=self.colors['focus_border'],
                         highlightthickness=2)
        except:
            pass

    def remove_highlight(self, widget):
        """Remove highlight from widget"""
        try:
            widget.config(highlightbackground='SystemWindowFrame', 
                         highlightcolor='SystemWindowFrame',
                         highlightthickness=1)
        except:
            pass



    def arrow_key_navigation(self, event):
        """Handle arrow key navigation in tables and lists"""
        focused_widget = self.root.focus_get()
        
        # Handle table navigation
        if hasattr(focused_widget, 'identify_row'):
            current_selection = focused_widget.selection()
            if current_selection:
                current_item = current_selection[0]
                all_items = focused_widget.get_children()
                
                if event.keysym == 'Down':
                    current_index = all_items.index(current_item)
                    if current_index < len(all_items) - 1:
                        next_item = all_items[current_index + 1]
                        focused_widget.selection_set(next_item)
                        focused_widget.focus(next_item)
                elif event.keysym == 'Up':
                    current_index = all_items.index(current_item)
                    if current_index > 0:
                        next_item = all_items[current_index - 1]
                        focused_widget.selection_set(next_item)
                        focused_widget.focus(next_item)
                        
            return "break"
        
        # Handle combobox navigation
        if isinstance(focused_widget, ttk.Combobox):
            if event.keysym in ['Up', 'Down']:
                focused_widget.event_generate('<Button-1>')
                return "break"
                
        return None

    def enter_key_action(self, event):
        """Handle Enter key actions based on context"""
        focused_widget = self.root.focus_get()
        
        # If focused on a button, trigger it
        if isinstance(focused_widget, tk.Button):
            focused_widget.invoke()
            return "break"
            
        # If in table, edit mode
        elif hasattr(focused_widget, 'identify_row'):
            # Double-click simulation for tables
            return "break"
            
        # If in combobox, select and move to next
        elif isinstance(focused_widget, ttk.Combobox):
            self.focus_next_widget()
            return "break"
            
        # Default: move to next widget
        else:
            self.focus_next_widget()
            return "break"

    def focus_escape(self, event=None):
        """Handle Escape key - clear focus or go back"""
        # Clear current focus
        self.root.focus_set()
        return "break"

    def undo_action(self, event=None):
        """Undo last action (Ctrl+Z)"""
        if self.undo_stack:
            action = self.undo_stack.pop()
            self.redo_stack.append(action)
            
            # Execute undo based on action type
            if action['type'] == 'field_change':
                widget = action['widget']
                old_value = action['old_value']
                try:
                    if isinstance(widget, tk.Entry):
                        widget.delete(0, tk.END)
                        widget.insert(0, old_value)
                    elif isinstance(widget, ttk.Combobox):
                        widget.set(old_value)
                except:
                    pass
                    
            self.show_status_message(f"↶ Undo: {action['description']}")
            
        return "break"

    def redo_action(self, event=None):
        """Redo last undone action (Ctrl+Y)"""
        if self.redo_stack:
            action = self.redo_stack.pop()
            self.undo_stack.append(action)
            
            # Execute redo based on action type
            if action['type'] == 'field_change':
                widget = action['widget']
                new_value = action['new_value']
                try:
                    if isinstance(widget, tk.Entry):
                        widget.delete(0, tk.END)
                        widget.insert(0, new_value)
                    elif isinstance(widget, ttk.Combobox):
                        widget.set(new_value)
                except:
                    pass
                    
            self.show_status_message(f"↷ Redo: {action['description']}")
            
        return "break"

    def quick_save(self, event=None):
        """Quick save current data (Ctrl+S)"""
        try:
            self.save_data(self.party_ref, self.party_data)
            self.save_data(self.product_ref, self.product_data)
            self.save_data(self.bills_ref, self.bills_data)  # Original uses bills.json
            self.show_status_message("💾 All data saved successfully!")
        except Exception as e:
            self.show_status_message(f"❌ Save failed: {str(e)}", error=True)
            
        return "break"

    def new_bill_shortcut(self, event=None):
        """Create new bill (Ctrl+N)"""
        if hasattr(self, 'show_billing_dashboard'):
            self.show_billing_dashboard()
            self.show_status_message("🧾 New bill created")
        return "break"

    def quit_application(self, event=None):
        """Quit application (Ctrl+Q)"""
        self.on_close()
        return "break"

    def show_help(self, event=None):
        """Show keyboard shortcuts help (F1)"""
        help_text = """
🎯 KEYBOARD SHORTCUTS - ANGEL INVOICE PRO

📋 NAVIGATION:
• TAB → Next field
• SHIFT+TAB → Previous field  
• ENTER → Confirm / Move next
• ESC → Clear focus
• ↑↓←→ → Navigate tables

⚡ ACTIONS:
• CTRL+Z → Undo
• CTRL+Y → Redo  
• CTRL+S → Quick Save
• CTRL+N → New Bill
• CTRL+Q → Quit
• ALT+A → Add Item
• ALT+R → Reset Form

🔧 FUNCTION KEYS:
• F1 → This help
• F2 → Party Management
• F3 → Product Management  
• F4 → Billing Center
• F5 → Refresh Data
• F6 → Reports
• F7 → Settings
• F12 → Dashboard

💡 TIPS:
• Use SPACEBAR to click focused buttons
• Press ENTER in tables to edit
• Use arrow keys in dropdowns
• CTRL+Z works for text fields
        """
        messagebox.showinfo("🎮 Keyboard Shortcuts", help_text)
        return "break"

    def function_key_handler(self, event):
        """Handle function keys F1-F12"""
        key_map = {
            'F2': self.show_party_management,
            'F3': self.show_product_management,
            'F4': self.show_billing_dashboard,
            'F5': self.refresh_data,
            'F6': self.show_stock_report,
            'F7': self.show_settings,
            'F12': self.show_modern_dashboard
        }
        
        key = event.keysym
        if key in key_map:
            key_map[key]()
            return "break"

    def track_field_change(self, widget, old_value, new_value, description=""):
        """Track field changes for undo/redo functionality"""
        if old_value != new_value:
            action = {
                'type': 'field_change',
                'widget': widget,
                'old_value': old_value,
                'new_value': new_value,
                'description': description,
                'timestamp': datetime.now()
            }
            
            self.undo_stack.append(action)
            # Limit undo stack size
            if len(self.undo_stack) > self.max_undo_steps:
                self.undo_stack.pop(0)
                
            # Clear redo stack when new action is performed
            self.redo_stack.clear()

    def create_modern_entry(self, parent, textvariable=None, width=20, font_size=10, **kwargs):
        """Create a modern entry widget with keyboard enhancements - FIXED VERSION"""
        # Remove font_size from kwargs and handle font separately
        font_kwarg = {"font": ("Segoe UI", font_size)}
        
        entry = tk.Entry(
            parent, 
            textvariable=textvariable, 
            width=width,
            **font_kwarg,
            **kwargs
        )
        entry = self.create_focusable_widget(entry)
        
        # Track changes for undo/redo - FIXED VERSION
        if textvariable is not None:
            def on_change(*args):
                try:
                    current_value = textvariable.get()
                    # Handle empty values gracefully
                    if current_value == "":
                        current_value = 0 if isinstance(textvariable, (tk.DoubleVar, tk.IntVar)) else ""
                    
                    if hasattr(entry, '_last_value') and entry._last_value != current_value:
                        self.track_field_change(entry, entry._last_value, current_value, "Field modified")
                    entry._last_value = current_value
                except (ValueError, tk.TclError):
                    # Handle conversion errors gracefully
                    pass
                    
            # Store initial value
            try:
                entry._last_value = textvariable.get()
            except (ValueError, tk.TclError):
                entry._last_value = 0 if isinstance(textvariable, (tk.DoubleVar, tk.IntVar)) else ""
                
            textvariable.trace_add('write', on_change)
                
        return entry

    def create_modern_combobox(self, parent, values=None, textvariable=None, width=20, font_size=10, **kwargs):
        """Create a modern combobox with keyboard enhancements"""
        # Remove font_size from kwargs and handle font separately
        font_kwarg = {"font": ("Segoe UI", font_size)}
        
        combobox = ttk.Combobox(
            parent, 
            values=values, 
            textvariable=textvariable, 
            width=width,
            **font_kwarg,
            **kwargs
        )
        combobox = self.create_focusable_widget(combobox)
        
        # Enhanced keyboard navigation for combobox
        combobox.bind('<Down>', lambda e: combobox.event_generate('<Button-1>'))
        combobox.bind('<Up>', lambda e: combobox.event_generate('<Button-1>'))
        
        return combobox

    def show_status_message(self, message, error=False):
        """Show status message in the status bar - FIXED VERSION"""
        try:
            if hasattr(self, 'status_label') and self.status_label.winfo_exists():
                color = self.colors['warning'] if error else self.colors['success']
                self.status_label.config(text=message, fg=color)
        except (tk.TclError, AttributeError):
            # Status label doesn't exist yet, just print to console
            print(f"Status: {message}")

    def refresh_data(self, event=None):
        """Refresh all data (F5)"""
        self.party_data = self.load_data("party_data.json")
        self.product_data = self.load_data("product_data.json")
        self.bills_data = self.load_data("bills.json")  # Original uses bills.json
        self.show_status_message("🔄 Data refreshed successfully!")
        return "break"

    def create_navigation_bar(self):
        """Create modern top navigation bar"""
        nav_frame = tk.Frame(self.root, bg=self.colors['primary'], height=50)
        nav_frame.pack(fill=tk.X, side=tk.TOP)
        nav_frame.pack_propagate(False)

        # Navigation container
        nav_container = tk.Frame(nav_frame, bg=self.colors['primary'])
        nav_container.pack(expand=True, fill=tk.BOTH, padx=20)

        # App logo/title
        title_label = tk.Label(
            nav_container,
            text="🚀 ANGEL INVOICE PRO",
            font=("Segoe UI", 16, "bold"),
            bg=self.colors['primary'],
            fg=self.colors['text_light']
        )
        title_label.pack(side=tk.LEFT, padx=10)

        # Navigation buttons
        nav_buttons_frame = tk.Frame(nav_container, bg=self.colors['primary'])
        nav_buttons_frame.pack(side=tk.LEFT, padx=50)

        nav_buttons = [
            ("🏠 Dashboard", self.show_modern_dashboard),
            ("👥 Parties", self.show_party_management),
            ("📦 Products", self.show_product_management),
            ("🧾 Billing", self.show_billing_dashboard),
            ("📊 Reports", self.show_stock_report),
        ]

        for text, command in nav_buttons:
            btn = self.create_modern_button(
                nav_buttons_frame, text, command,
                style="secondary", width=12, height=1
            )
            btn.pack(side=tk.LEFT, padx=2)

        # Right side buttons
        right_frame = tk.Frame(nav_container, bg=self.colors['primary'])
        right_frame.pack(side=tk.RIGHT)

        # Firebase status on TOP BAR
        self.firebase_status = tk.Label(
            right_frame,
            text="●  Online" if self.firebase_connected else "●  Offline",
            font=("Segoe UI", 9, "bold"),
            bg=self.colors['primary'],
            fg="lime" if self.firebase_connected else "red"
        )
        self.firebase_status.pack(side=tk.LEFT, padx=15)

        # Help button
        help_btn = self.create_modern_button(
            right_frame, "❓ Help", self.show_help,
            style="info", width=8, height=1
        )
        help_btn.pack(side=tk.LEFT, padx=2)

        # Logout button
        logout_btn = self.create_modern_button(
            right_frame, "🚪 Logout", self.show_modern_login,
            style="warning", width=8, height=1
        )
        logout_btn.pack(side=tk.LEFT, padx=2)

    def create_status_bar(self):
        """Create modern status bar"""
        status_frame = tk.Frame(self.root, bg=self.colors['secondary'], height=30)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)
        status_frame.pack_propagate(False)

        # Status label
        self.status_label = tk.Label(
            status_frame,
            text="🚀 Ready - Press F1 for keyboard shortcuts",
            font=("Segoe UI", 9),
            bg=self.colors['secondary'],
            fg=self.colors['text_light']
        )
        self.status_label.pack(side=tk.LEFT, padx=10, pady=5)

        # Keyboard hints
        hints_label = tk.Label(
            status_frame,
            text="TAB: Navigate • ENTER: Confirm • CTRL+Z: Undo • CTRL+S: Save",
            font=("Segoe UI", 8),
            bg=self.colors['secondary'],
            fg=self.colors['text_light']
        )
        hints_label.pack(side=tk.RIGHT, padx=10, pady=5)

    # ========== MODERN LOGIN AND OFFICE SELECTION ==========

    def show_modern_login(self):
        """Show modern login screen - FIXED TAB NAVIGATION"""
        self.clear_screen()
        self.current_screen = "login"
        
        # Main container with gradient background
        main_container = tk.Frame(self.root, bg=self.colors['primary'])
        main_container.pack(fill=tk.BOTH, expand=True)

        # Center frame
        center_frame = tk.Frame(main_container, bg=self.colors['primary'])
        center_frame.place(relx=0.5, rely=0.5, anchor="center")

        # Login card
        login_card = self.create_modern_frame(center_frame, "🔐 SYSTEM LOGIN", padding=20, bg=self.colors['card_bg'])
        login_card.pack(padx=20, pady=20)

        # Login form
        form_frame = tk.Frame(login_card, bg=self.colors['card_bg'])
        form_frame.pack(padx=40, pady=30)

        # App title
        title_label = tk.Label(
            form_frame,
            text="ANGEL INVOICE PRO",
            font=("Segoe UI", 20, "bold"),
            bg=self.colors['card_bg'],
            fg=self.colors['primary']
        )
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))

        subtitle_label = tk.Label(
            form_frame,
            text="Modern Billing Solution",
            font=("Segoe UI", 12),
            bg=self.colors['card_bg'],
            fg=self.colors['text_muted']
        )
        subtitle_label.grid(row=1, column=0, columnspan=2, pady=(0, 30))

        # Username
        tk.Label(form_frame, text="👤 Username:", font=("Segoe UI", 10, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=2, column=0, sticky="w", pady=10)
        
        self.username_entry = self.create_modern_entry(form_frame, width=25)
        self.username_entry.grid(row=2, column=1, padx=15, pady=10)

        # Password
        tk.Label(form_frame, text="🔒 Password:", font=("Segoe UI", 10, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=3, column=0, sticky="w", pady=10)
        
        self.password_entry = self.create_modern_entry(form_frame, width=25, show="•")
        self.password_entry.grid(row=3, column=1, padx=15, pady=10)

        # Login button
        login_btn = self.create_modern_button(
            form_frame, 
            "🚀 LOGIN TO DASHBOARD", 
            self.attempt_login,
            style="success",
            width=25,
            height=2
        )
        login_btn.grid(row=4, column=0, columnspan=2, pady=30)

        # Bind Enter key to login
        self.password_entry.bind('<Return>', lambda e: self.attempt_login())
        
        # FIX: Set up proper tab navigation for login screen
        self.focusable_widgets.clear()
        self.create_focusable_widget(self.username_entry)
        self.create_focusable_widget(self.password_entry)
        self.create_focusable_widget(login_btn, "button")
        
        # Set initial focus
        self.username_entry.focus_set()

            # Rebuild focusable widgets list
        self.rebuild_focusable_widgets()
        if self.focusable_widgets:
            self.current_focus_index = self.focusable_widgets.index(self.username_entry)

        # Footer
        footer_label = tk.Label(
            main_container,
            text="© 2024 Angel Invoice Pro - Press F1 for Keyboard Shortcuts",
            font=("Segoe UI", 9),
            bg=self.colors['primary'],
            fg=self.colors['text_light']
        )
        footer_label.pack(side=tk.BOTTOM, pady=10)

    def attempt_login(self):
        """Attempt user login with original credentials"""
        username = self.username_entry.get()
        password = self.password_entry.get()

        if not username or not password:
            messagebox.showerror("Error", "❌ Please enter both username and password!")
            return

        # Use original authentication logic
        if username == "Angel" and password == "2010":
            self.show_modern_office_selection()
        else:
            messagebox.showerror("Error", "❌ Invalid credentials!")

    def show_modern_office_selection(self):
        """Show modern office selection screen"""
        self.clear_screen()
        self.current_screen = "office_selection"
        
        # Main container
        main_container = tk.Frame(self.root, bg=self.colors['light_bg'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Header
        header_frame = self.create_modern_frame(main_container, "🏢 SELECT YOUR OFFICE")
        header_frame.pack(fill=tk.X, pady=(0, 20))

        header_label = tk.Label(
            header_frame,
            text="Choose your office location to continue",
            font=("Segoe UI", 14),
            bg=self.colors['card_bg'],
            fg=self.colors['text_dark'],
            pady=20
        )
        header_label.pack()

        # Office selection cards
        offices_frame = tk.Frame(main_container, bg=self.colors['light_bg'])
        offices_frame.pack(fill=tk.BOTH, expand=True, pady=20)

        # Office options with enhanced styling
        office_options = [
            {
                "title": "🏭 Angel Pyrotech",
                "description": "Main Office - ONDIPULINAIKANOOR",
                "value": "A1",
                "color": self.colors['info'],
                "details": "D NO 3/89 3/89/1 TO 3/89/11 ONDIPULINAIKANOOR\nONDIPULINAIKANOOR VILLAGE TAMILNADU 626119\nGSTIN: 33ABRFA4846J1Z3"
            },
            {
                "title": "🏢 Angel Fireworks Industries", 
                "description": "Factory - SIVAKASI",
                "value": "A2",
                "color": self.colors['success'],
                "details": "FACTORY AT:O.KOVILPATTI,2/2204/W,DEVINAGAR\nSIVAKASI-626123\nGSTIN: 33AARFA9673N2ZL"
            },
            {
                "title": "🏭 Angel Fireworks Factory",
                "description": "Factory - VIRUTHUNAGAR", 
                "value": "A3",
                "color": self.colors['warning'],
                "details": "FACTORY AT:O.KOVILPATTI,2/2204/X,DEVINAGAR\nVIRUTHUNAGAR-626123\nGSTIN: 33ABKFA4066F1ZN"
            }
        ]

        # Create office selection cards
        self.office_var = tk.StringVar(value="A1")  # Default selection
        
        for i, office in enumerate(office_options):
            card = self.create_office_card(offices_frame, office, i)
            card.pack(fill=tk.X, pady=10, padx=50)

        # Continue button
        button_frame = tk.Frame(main_container, bg=self.colors['light_bg'])
        button_frame.pack(fill=tk.X, pady=20)

        continue_btn = self.create_modern_button(
            button_frame,
            "🚀 CONTINUE TO DASHBOARD",
            self.proceed_with_office_selection,
            style="success",
            width=25,
            height=2
        )
        continue_btn.pack()

        # Back button
        back_btn = self.create_modern_button(
            button_frame,
            "↶ BACK TO LOGIN",
            self.show_modern_login,
            style="secondary",
            width=15,
            height=1
        )
        back_btn.pack(pady=10)

    def create_office_card(self, parent, office_data, index):
        """Create a modern office selection card"""
        card = tk.Frame(
            parent,
            bg=self.colors['card_bg'],
            relief="raised",
            bd=2,
            highlightbackground=self.colors['accent'],
            highlightthickness=1
        )
        
        # Card content
        content_frame = tk.Frame(card, bg=self.colors['card_bg'])
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        # Radio button and main info
        top_frame = tk.Frame(content_frame, bg=self.colors['card_bg'])
        top_frame.pack(fill=tk.X)

        # Radio button
        radio_btn = tk.Radiobutton(
            top_frame,
            variable=self.office_var,
            value=office_data["value"],
            bg=self.colors['card_bg'],
            fg=self.colors['text_dark'],
            font=("Segoe UI", 12, "bold"),
            cursor="hand2"
        )
        radio_btn.pack(side=tk.LEFT)

        # Office info
        info_frame = tk.Frame(top_frame, bg=self.colors['card_bg'])
        info_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)

        # Title
        title_label = tk.Label(
            info_frame,
            text=office_data["title"],
            font=("Segoe UI", 14, "bold"),
            bg=self.colors['card_bg'],
            fg=office_data["color"],
            anchor="w"
        )
        title_label.pack(fill=tk.X)

        # Description
        desc_label = tk.Label(
            info_frame,
            text=office_data["description"],
            font=("Segoe UI", 11),
            bg=self.colors['card_bg'],
            fg=self.colors['text_muted'],
            anchor="w"
        )
        desc_label.pack(fill=tk.X)

        # Details (initially hidden)
        details_frame = tk.Frame(content_frame, bg=self.colors['card_bg'])
        details_label = tk.Label(
            details_frame,
            text=office_data["details"],
            font=("Segoe UI", 9),
            bg=self.colors['card_bg'],
            fg=self.colors['text_muted'],
            justify=tk.LEFT
        )
        details_label.pack(fill=tk.X, pady=(5, 0))

        # Show details on hover
        def show_details(e):
            details_frame.pack(fill=tk.X, pady=(5, 0))
            
        def hide_details(e):
            if self.office_var.get() != office_data["value"]:
                details_frame.pack_forget()

        # Bind hover events
        for widget in [card, content_frame, title_label, desc_label]:
            widget.bind("<Enter>", show_details)
            widget.bind("<Leave>", hide_details)

        # Show details if this office is selected
        if self.office_var.get() == office_data["value"]:
            details_frame.pack(fill=tk.X, pady=(5, 0))

        return card

    def proceed_with_office_selection(self):
        """Proceed with the selected office - FIXED VERSION"""
        self.selected_office = self.office_var.get()
        office_names = {
            "A1": "Angel Pyrotech",
            "A2": "Angel Fireworks Industries", 
            "A3": "Angel Fireworks Factory"
        }
        
        office_name = office_names.get(self.selected_office, "Unknown Office")
        
        # Use safe status message
        try:
            self.show_status_message(f"✅ Selected: {office_name}")
        except:
            print(f"Selected: {office_name}")
        
        self.show_modern_dashboard()

    def show_modern_dashboard(self):
        """Show modern dashboard"""
        self.clear_screen()
        self.current_screen = "dashboard"
        
        # Create navigation and status bars
        self.create_navigation_bar()
        self.create_status_bar()

        self.show_status_message("✅ Dashboard loaded - Use function keys for quick navigation!")

        self.check_firebase_connection()


        if hasattr(self, 'firebase_status'):
            self.firebase_status.config(
                text="● Online" if self.firebase_connected else "●  Offline",
                fg="lime" if self.firebase_connected else "red"
            )

        # Main content area
        content_frame = tk.Frame(self.root, bg=self.colors['light_bg'])
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Welcome header with office info
        office_names = {
            "A1": "Angel Pyrotech",
            "A2": "Angel Fireworks Industries",
            "A3": "Angel Fireworks Factory"
        }
        current_office = office_names.get(self.selected_office, "Unknown Office")
        
        welcome_frame = self.create_modern_frame(content_frame, f"🎯 DASHBOARD - {current_office}")
        welcome_frame.pack(fill=tk.X, pady=(0, 20))

        welcome_text = tk.Label(
            welcome_frame,
            text=f"Welcome to {current_office} - Your Complete Billing Solution",
            font=("Segoe UI", 14),
            bg=self.colors['card_bg'],
            fg=self.colors['text_dark'],
            pady=20
        )
        welcome_text.pack()

        # Stats cards
        stats_frame = tk.Frame(content_frame, bg=self.colors['light_bg'])
        stats_frame.pack(fill=tk.X, pady=10)

        # Calculate office-specific stats
        office_prefix = {"A1": "AP", "A2": "AFI", "A3": "AFF"}.get(self.selected_office, "AP")
        office_bills = {k: v for k, v in self.bills_data.items() if k.startswith(office_prefix)}
        
        stats_data = [
            ("👥 Total Customers", len(self.party_data), self.colors['info'], "F2"),
            ("📦 Total Products", len(self.product_data), self.colors['success'], "F3"),
            ("🧾 Office Invoices", len(office_bills), self.colors['warning'], "F4"),
            ("💰 Revenue Today", "₹0", self.colors['accent'], "F6")
        ]

        for title, value, color, shortcut in stats_data:
            card = self.create_stat_card(stats_frame, title, value, color, shortcut)
            card.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=5, pady=5)

        # Quick actions
        actions_frame = self.create_modern_frame(content_frame, "⚡ QUICK ACTIONS")
        actions_frame.pack(fill=tk.X, pady=20)

        actions_container = tk.Frame(actions_frame, bg=self.colors['card_bg'])
        actions_container.pack(padx=20, pady=20)

        quick_actions = [
            ("🧾 Create New Invoice", self.show_billing_dashboard, "F4", "Create a new invoice"),
            ("👥 Add New Customer", self.show_party_management, "F2", "Manage customer database"),
            ("📦 Update Products", self.show_product_management, "F3", "Manage product catalog"),
            ("📊 View Reports", self.show_stock_report, "F6", "Generate business reports"),
            ("⚙️ System Settings", self.show_settings, "F7", "Configure application settings"),
            ("🏢 Switch Office", self.show_modern_office_selection, "F12", "Change current office")
        ]

        for i, (text, command, shortcut, tooltip) in enumerate(quick_actions):
            btn = self.create_modern_button(
                actions_container, text, command,
                style="primary", width=25, height=2
            )
            btn.grid(row=i//3, column=i%3, padx=10, pady=10, sticky="ew")

        # Set up focusable widgets for keyboard navigation
        self.focusable_widgets.clear()
        for widget in actions_container.winfo_children():
            if isinstance(widget, tk.Button):
                self.create_focusable_widget(widget, "button")

        if self.focusable_widgets:
            self.focusable_widgets[0].focus_set()

        self.show_status_message(f"✅ {current_office} Dashboard loaded - Use function keys for quick navigation!")

    def create_stat_card(self, parent, title, value, color, shortcut):
        """Create a modern stat card"""
        card = tk.Frame(
            parent,
            bg=self.colors['card_bg'],
            relief="raised",
            bd=1,
            width=200,
            height=120
        )
        card.pack_propagate(False)

        # Card content
        content_frame = tk.Frame(card, bg=self.colors['card_bg'])
        content_frame.pack(expand=True, fill=tk.BOTH, padx=15, pady=15)

        # Title
        title_label = tk.Label(
            content_frame,
            text=title,
            font=("Segoe UI", 10, "bold"),
            bg=self.colors['card_bg'],
            fg=self.colors['text_muted']
        )
        title_label.pack(anchor="w")

        # Value
        value_label = tk.Label(
            content_frame,
            text=str(value),
            font=("Segoe UI", 24, "bold"),
            bg=self.colors['card_bg'],
            fg=color
        )
        value_label.pack(expand=True)

        # Shortcut hint
        shortcut_label = tk.Label(
            content_frame,
            text=f"Press {shortcut}",
            font=("Segoe UI", 8),
            bg=self.colors['card_bg'],
            fg=self.colors['text_muted']
        )
        shortcut_label.pack(anchor="e")

        return card
    
    def check_firebase_connection(self):
        """Check Firebase connection every 5 seconds and updates status bar label safely."""
        # ⚠️ CRITICAL FIX: Check if the widget exists before trying to configure it
        if not hasattr(self, 'firebase_status') or not self.firebase_status.winfo_exists():
            # If the widget is not available (e.g., screen switched), stop the loop 
            # or just reschedule if you want the check to continue in the background.
            # In a multi-screen app, stopping the loop is usually safer.
            # We'll just return and let the new screen's initialization handle the status.
            return 
        
        try:
            # Try reading a small value from Firebase
            # NOTE: For a real production app, only checking database connectivity 
            # once or upon major failure is more efficient than polling every 5s.
            db.reference("/").get()

            # Connected
            self.firebase_connected = True
            self.firebase_status.config(
                text="● Online",
                fg="lime"
            )
        except Exception:
            # Not connected
            self.firebase_connected = False
            self.firebase_status.config(
                text="● Offline",
                fg="red"
            )

        # Run again after 5 seconds - this is the source of the repeated error
        # It's called after the widget is destroyed. The check above prevents the crash.
        self.root.after(5000, self.check_firebase_connection)

    

    def clear_screen(self):
        """Clear the current screen and reset focus tracking"""
        
        # --- Recommended Cleanup (Only unbinds the specific global binding if used) ---
        # Since you should have changed 'bind_all' to 'bind' in show_settings, 
        # this entire complex block is unnecessary.

        # If you MUST have the global binding (NOT RECOMMENDED):
        if self.current_screen == "settings":
            self.root.unbind_all("<MouseWheel>") 
            self.setup_global_keyboard_shortcuts() # Re-bind the necessary globals
        # ----------------------------------------------------------------------------

        for widget in self.root.winfo_children():
            widget.destroy()
        
        # Clear focus tracking
        self.focusable_widgets = []
        self.current_focus_index = 0

    # ========== PLACEHOLDER METHODS FOR FUTURE IMPLEMENTATION ==========
    
    def show_party_management(self):
        """Show party management screen"""
        self.clear_screen()
        self.current_screen = "party_management"
        self.create_navigation_bar()
        self.create_status_bar()

        content_frame = tk.Frame(self.root, bg=self.colors['light_bg'])
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Header
        header_frame = self.create_modern_frame(content_frame, "👥 PARTY MANAGEMENT")
        header_frame.pack(fill=tk.X, pady=(0, 20))

        # Quick actions
        actions_frame = tk.Frame(header_frame, bg=self.colors['card_bg'])
        actions_frame.pack(padx=20, pady=20)

        actions = [
            ("➕ Add New Party", self.show_party_entry, "Add new customer"),
            ("✏️ Modify Party", self.show_party_modify, "Edit existing customer"),
            ("📋 Party List", self.show_party_list, "View all customers"),
            ("📊 Party Statement", self.show_party_statement, "Generate statements")
        ]

        for i, (text, command, tooltip) in enumerate(actions):
            btn = self.create_modern_button(
                actions_frame, text, command,
                style="primary", width=20, height=2
            )
            btn.grid(row=i//2, column=i%2, padx=10, pady=10)

        # Set up keyboard navigation
        self.focusable_widgets.clear()
        for widget in actions_frame.winfo_children():
            if isinstance(widget, tk.Button):
                self.create_focusable_widget(widget, "button")

        if self.focusable_widgets:
            self.focusable_widgets[0].focus_set()

        self.show_status_message("👥 Party Management - Use TAB to navigate between options")

    def show_party_entry(self):
        """Modern enhanced version of show_entry_page for adding new parties"""
        self.clear_screen()
        self.current_screen = "party_entry"
        self.create_navigation_bar()
        self.create_status_bar()
        
    # ensure attribute exists
        if not hasattr(self, "focusable_widgets"):
            self.focusable_widgets = []
        else:
            self.focusable_widgets.clear()

        # Main content area with centered layout
        main_container = tk.Frame(self.root, bg=self.colors['light_bg'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Center container to hold everything
        center_container = tk.Frame(main_container, bg=self.colors['light_bg'])
        center_container.place(relx=0.5, rely=0.5, anchor="center", width=800, height=600)

        # Header frame
        header_frame = self.create_modern_frame(center_container, "👥 ADD NEW PARTY")
        header_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))

        # Form container with scrollbar for better mobile experience
        form_main = tk.Frame(header_frame, bg=self.colors['card_bg'])
        form_main.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Create a canvas and scrollbar for the form
        canvas = tk.Canvas(form_main, bg=self.colors['card_bg'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(form_main, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.colors['card_bg'])

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Form container - using grid for better alignment
        form_container = scrollable_frame
        form_container.grid_columnconfigure(1, weight=1)

        # Auto-increment party code
        party_codes = [int(code) for code in self.party_data.keys() if code.isdigit()]
        next_party_code = str(max(party_codes) + 1) if party_codes else "1"

        # Party Code - Row 0
        tk.Label(form_container, text="🔢 Party Code:", font=("Segoe UI", 11, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=0, column=0, sticky="w", padx=10, pady=12)
        
        self.party_code_entry = self.create_modern_entry(form_container, width=55)
        self.party_code_entry.insert(0, next_party_code)
        self.party_code_entry.grid(row=0, column=1, padx=10, pady=12, sticky="ew")
        self.party_code_entry.config(state='readonly')

        # Customer Name - Row 1
        tk.Label(form_container, text="👤 Customer Name *:", font=("Segoe UI", 11, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=1, column=0, sticky="w", padx=10, pady=12)
        
        self.customer_name_entry = self.create_modern_entry(form_container, width=25)
        self.customer_name_entry.grid(row=1, column=1, padx=10, pady=12, sticky="ew")

        # Address - Row 2 (Fixed Tab Navigation)
        tk.Label(form_container, text="🏠 Address *:", font=("Segoe UI", 11, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=2, column=0, sticky="nw", padx=10, pady=12)
        
        # Create a frame for the address entry with proper tab handling
        address_frame = tk.Frame(form_container, bg=self.colors['card_bg'])
        address_frame.grid(row=2, column=1, padx=10, pady=12, sticky="nsew")
        
        # Create Text widget with proper tab handling
        self.address_entry = tk.Text(
            address_frame, 
            font=("Segoe UI", 10), 
            width=25, 
            height=3,
            relief="solid", 
            bd=1,
            wrap=tk.WORD,
            tabs=('0.5c')  # Set tab stops
        )
        
        # Scrollbar for address field
        address_scrollbar = ttk.Scrollbar(address_frame, orient="vertical", command=self.address_entry.yview)
        self.address_entry.configure(yscrollcommand=address_scrollbar.set)
        
        self.address_entry.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        address_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Fix Tab navigation for Text widget - bind Tab to move to next field
        def handle_text_tab(event):
            if event.keysym == 'Tab':
                self.focus_next_widget()
                return "break"  # Prevent default tab behavior
            return None
        
        self.address_entry.bind('<Tab>', lambda e: (self.focus_next_widget(), "break"))
        self.address_entry.bind('<Shift-Tab>', lambda e: (self.focus_previous_widget(), "break"))

        self.address_entry.bind('<FocusIn>', lambda e: self.highlight_focused_widget(e.widget))
        self.address_entry.bind('<FocusOut>', lambda e: self.remove_highlight(e.widget))

        self.create_focusable_widget(self.address_entry)


        # GST Number - Row 3
        tk.Label(form_container, text="📊 GST Number *:", font=("Segoe UI", 11, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=3, column=0, sticky="w", padx=10, pady=12)
        
        self.gst_number_entry = self.create_modern_entry(form_container, width=25)
        self.gst_number_entry.grid(row=3, column=1, padx=10, pady=12, sticky="ew")

        # Phone Number - Row 4
        tk.Label(form_container, text="📞 Phone Number *:", font=("Segoe UI", 11, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=4, column=0, sticky="w", padx=10, pady=12)
        
        self.phone_number_entry = self.create_modern_entry(form_container, width=25)
        self.phone_number_entry.grid(row=4, column=1, padx=10, pady=12, sticky="ew")

        # Agent Name - Row 5
        tk.Label(form_container, text="🤵 Agent Name *:", font=("Segoe UI", 11, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=5, column=0, sticky="w", padx=10, pady=12)
        
        self.agent_name_entry = self.create_modern_entry(form_container, width=25)
        self.agent_name_entry.grid(row=5, column=1, padx=10, pady=12, sticky="ew")

        # Required fields note
        note_frame = tk.Frame(form_container, bg=self.colors['card_bg'])
        note_frame.grid(row=6, column=0, columnspan=2, pady=15, sticky="w")
        
        note_label = tk.Label(
            note_frame,
            text="* Fields marked with asterisk are required",
            font=("Segoe UI", 9, "italic"),
            bg=self.colors['card_bg'],
            fg=self.colors['text_muted']
        )
        note_label.pack()

        # Pack the canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Action buttons frame - Centered at bottom
        action_frame = tk.Frame(center_container, bg=self.colors['light_bg'])
        action_frame.pack(fill=tk.X, pady=10)

        # Center the buttons
        button_container = tk.Frame(action_frame, bg=self.colors['light_bg'])
        button_container.pack(expand=True)

        # Save button
        save_btn = self.create_modern_button(
            button_container,
            "💾 Save Party (Ctrl+S)",
            self.save_party_details_enhanced,
            style="success",
            width=18,
            height=2
        )
        save_btn.pack(side=tk.LEFT, padx=8)

        # Reset button
        reset_btn = self.create_modern_button(
            button_container,
            "🔄 Reset Form (Alt+R)",
            self.reset_party_form,
            style="warning",
            width=15,
            height=2
        )
        reset_btn.pack(side=tk.LEFT, padx=8)

        # Back button
        back_btn = self.create_modern_button(
            button_container,
            "← Back (F2)",
            self.show_party_management,
            style="secondary",
            width=12,
            height=2
        )
        back_btn.pack(side=tk.LEFT, padx=8)

        # Quick Add another button
        quick_add_btn = self.create_modern_button(
            button_container,
            "⚡ Save & New (Ctrl+N)",
            self.save_and_new_party,
            style="info",
            width=15,
            height=2
        )
        quick_add_btn.pack(side=tk.LEFT, padx=8)

        # Configure row weights for better spacing
        for i in range(6):
            form_container.grid_rowconfigure(i, weight=1)

        # Bind keyboard shortcuts
        self.root.bind('<Control-s>', lambda e: self.save_party_details_enhanced())
        self.root.bind('<Alt-r>', lambda e: self.reset_party_form())
        self.root.bind('<Control-n>', lambda e: self.save_and_new_party())
        self.root.bind('<F2>', lambda e: self.show_party_management())

        # Update the scroll region when the window is resized
        def update_scroll_region(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        canvas.bind("<Configure>", update_scroll_region)

        # --- Prepare focusable widgets list and focus tracking ---
        # Order here decides Tab order
        self.focusable_widgets = [
            # include read-only code too if you want it focusable; else comment out
            # self.party_code_entry,
            self.customer_name_entry,
            self.address_entry,
            self.gst_number_entry,
            self.phone_number_entry,
            self.agent_name_entry,
            save_btn,
            reset_btn,
            quick_add_btn,
            back_btn
        ]

        # helper: update current focus index when user clicks into a widget
        def update_current_focus(event):
            widget = event.widget
            try:
                if widget in self.focusable_widgets:
                    self.current_focus_index = self.focusable_widgets.index(widget)
            except Exception:
                # ignore if widget not in list
                pass

        # bind focus-in to each registered widget so clicks update the index
        for w in self.focusable_widgets:
            try:
                w.bind("<FocusIn>", update_current_focus)
            except Exception:
                # some widgets might not support bind directly; skip
                pass

        # initialize index
        self.current_focus_index = 0

        # Ensure the global Tab/Shift-Tab handlers are bound to the root (one time)
        # If other code already binds_all, overwriting is fine; these must exist
        self.root.bind_all('<Tab>', self.focus_next_widget)
        self.root.bind_all('<Shift-Tab>', self.focus_previous_widget)

        # Set focus to first editable field
        self.customer_name_entry.focus_set()

        # Make the window responsive
        center_container.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        self.show_status_message("📝 Adding new party - Fill in all required fields and press Ctrl+S to save")

    def save_party_details_enhanced(self, silent=False):
        """Enhanced save function with Firebase-safe validation and flexible key handling"""
        print("🔄 DEBUG: Starting save_party_details_enhanced")

        # Validate fields
        party_code = self.party_code_entry.get().strip()
        customer_name = self.customer_name_entry.get().strip()
        address = self.address_entry.get("1.0", "end-1c").strip()
        gst_number = self.gst_number_entry.get().strip()
        phone_number = self.phone_number_entry.get().strip()
        agent_name = self.agent_name_entry.get().strip()

        if not all([party_code, customer_name, address, gst_number, phone_number, agent_name]):
            print("❌ DEBUG: Validation failed - missing fields")
            if not silent:
                messagebox.showerror(
                    "❌ Validation Error",
                    "Please fill out all fields before saving.\n\n"
                    "All fields marked with * are required."
                )
            return False

        # Firebase key validation
        import re
        if not re.match(r'^[A-Za-z0-9_-]+$', party_code):
            print(f"❌ DEBUG: Invalid party code '{party_code}'")
            if not silent:
                messagebox.showerror(
                    "❌ Invalid Code",
                    "Party Code cannot contain spaces or special characters.\n"
                    "Allowed characters: A–Z, 0–9, dash (-), underscore (_)."
                )
            return False

        # Check if party code already exists
        if party_code in self.party_data:
            print(f"❌ DEBUG: Party code {party_code} already exists")
            if not silent:
                messagebox.showerror(
                    "❌ Duplicate Error",
                    f"Party code '{party_code}' already exists!\n"
                    "Please use a different party code."
                )
            return False

        # Prepare entry with consistent key naming
        self.party_data[party_code] = {
            "Customer Name": customer_name,      # Standardized key name
            "Address": address,                  # Standardized key name
            "GST Number": gst_number,            # Standardized key name
            "Phone Number": phone_number,        # Standardized key name
            "Agent Name": agent_name             # Standardized key name
        }

        # SAVE TO FIREBASE
        print("🔄 DEBUG: Saving to Firebase…")
        save_result = self.save_data(self.party_ref, self.party_data)
        print(f"🔄 DEBUG: save_data returned: {save_result}")

        if save_result:
            # Update local dropdown list with flexible key retrieval
            self.customer_names = []
            for key, party in self.party_data.items():
                # Try multiple possible key names for customer name
                customer = (
                    party.get("Customer Name") or 
                    party.get("Customer_Name") or 
                    party.get("customer_name") or
                    party.get("Customer  Name") or
                    party.get("Customer name") or
                    party.get(" Customer Name") or
                    ""
                )
                if customer:
                    self.customer_names.append(customer)
            
            # Remove duplicates while preserving order
            seen = set()
            self.customer_names = [x for x in self.customer_names if not (x in seen or seen.add(x))]

            if not silent:
                messagebox.showinfo(
                    "✅ Success",
                    f"Party details saved successfully!\n\n"
                    f"🔢 Party Code: {party_code}\n"
                    f"👤 Customer: {customer_name}\n"
                    f"🤵 Agent: {agent_name}"
                )

            self.show_status_message(
                f"✅ Party '{customer_name}' saved successfully!"
            )

            return True

        else:
            print("❌ DEBUG: Firebase save failed")
            if not silent:
                messagebox.showerror(
                    "❌ Save Error",
                    "Failed to save party details.\n"
                    "Please check Firebase connection and try again."
                )
            return False

    def reset_party_form(self):
        """Reset the party form to initial state"""
        try:
            # Get all existing party codes
            existing_codes = []
            for key in self.party_data.keys():
                if key.isdigit():
                    try:
                        existing_codes.append(int(key))
                    except ValueError:
                        continue
            
            # Find the next available code
            if existing_codes:
                next_code = str(max(existing_codes) + 1)
            else:
                next_code = "1"
            
            # Clear and set party code
            self.party_code_entry.config(state='normal')
            self.party_code_entry.delete(0, tk.END)
            self.party_code_entry.insert(0, next_code)
            self.party_code_entry.config(state='readonly')
            
            # Clear other fields
            self.customer_name_entry.delete(0, tk.END)
            self.address_entry.delete("1.0", tk.END)
            self.gst_number_entry.delete(0, tk.END)
            self.phone_number_entry.delete(0, tk.END)
            self.agent_name_entry.delete(0, tk.END)
            
            # Set focus to customer name field
            self.customer_name_entry.focus_set()
            
            self.show_status_message("🔄 Form reset - Ready for new party entry")
            
        except Exception as e:
            print(f"Error in reset_party_form: {e}")
            # Simple fallback
            self.party_code_entry.config(state='normal')
            self.party_code_entry.delete(0, tk.END)
            self.party_code_entry.insert(0, "1")
            self.party_code_entry.config(state='readonly')
            self.customer_name_entry.delete(0, tk.END)
            self.customer_name_entry.focus_set()

    
    def save_and_new_party(self):
        """Save current party and reset form for new entry"""
        # Save the party first
        if self.save_party_details_enhanced(silent=True):
            # Clear form for new entry
            self.reset_party_form()
            self.show_status_message("✅ Party saved! Ready for new entry...")

    


    def show_party_modify(self):
        """Modern enhanced version of show_modify_page for modifying party details"""
        self.clear_screen()
        self.current_screen = "party_modify"
        self.create_navigation_bar()
        self.create_status_bar()
        
        # Ensure focusable_widgets list exists
        if not hasattr(self, "focusable_widgets"):
            self.focusable_widgets = []
        else:
            self.focusable_widgets.clear()

        # Main content area with centered layout
        main_container = tk.Frame(self.root, bg=self.colors['light_bg'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Center container to hold everything
        center_container = tk.Frame(main_container, bg=self.colors['light_bg'])
        center_container.place(relx=0.5, rely=0.5, anchor="center", width=900, height=650)

        # Header frame
        header_frame = self.create_modern_frame(center_container, "✏️ MODIFY PARTY DETAILS")
        header_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))

        # Form container
        form_main = tk.Frame(header_frame, bg=self.colors['card_bg'])
        form_main.pack(fill=tk.BOTH, expand=True, padx=25, pady=25)

        # Create a canvas and scrollbar for the form
        canvas = tk.Canvas(form_main, bg=self.colors['card_bg'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(form_main, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.colors['card_bg'])

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Form container - using grid for better alignment
        form_container = scrollable_frame
        form_container.grid_columnconfigure(1, weight=1)

        # Party Selection Section
        selection_frame = tk.Frame(form_container, bg=self.colors['card_bg'], relief="solid", bd=1, padx=10, pady=10)
        selection_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=15)
        
        tk.Label(selection_frame, text="🔍 Select Party to Modify:", font=("Segoe UI", 12, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).pack(side=tk.LEFT, padx=5)
        
        # Party Code dropdown with search functionality
        self.party_code_combobox = self.create_modern_combobox(
            selection_frame, 
            values=list(self.party_data.keys()),
            width=25,
            state="readonly"
        )
        self.party_code_combobox.pack(side=tk.LEFT, padx=10, pady=5)
        self.party_code_combobox.set("")  # Start with empty selection

        # Quick load button
        load_btn = self.create_modern_button(
            selection_frame,
            "📥 Load Party Details (Enter)",
            self.load_selected_party_details_enhanced,
            style="info",
            width=20,
            height=1
        )
        load_btn.pack(side=tk.LEFT, padx=10)

        # Bind Enter key to load party details
        self.party_code_combobox.bind('<<ComboboxSelected>>', lambda e: self.load_selected_party_details_enhanced())
        self.party_code_combobox.bind('<Return>', lambda e: self.load_selected_party_details_enhanced())

        # Form Fields Section
        fields_frame = tk.Frame(form_container, bg=self.colors['card_bg'])
        fields_frame.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=10, pady=10)

        # Customer Name - Row 0
        tk.Label(fields_frame, text="👤 Customer Name *:", font=("Segoe UI", 11, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=0, column=0, sticky="w", padx=10, pady=12)
        
        self.customer_name_entry_modify = self.create_modern_entry(fields_frame, width=30)
        self.customer_name_entry_modify.grid(row=0, column=1, padx=10, pady=12, sticky="ew")

        # Address - Row 1 (Fixed Tab Navigation)
        tk.Label(fields_frame, text="🏠 Address *:", font=("Segoe UI", 11, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=1, column=0, sticky="nw", padx=10, pady=12)
        
        # Create a frame for the address entry with proper tab handling
        address_frame = tk.Frame(fields_frame, bg=self.colors['card_bg'])
        address_frame.grid(row=1, column=1, padx=10, pady=12, sticky="nsew")
        
        # Create Text widget with proper tab handling
        self.address_entry_modify = tk.Text(
            address_frame, 
            font=("Segoe UI", 10), 
            width=30, 
            height=3,
            relief="solid", 
            bd=1,
            wrap=tk.WORD,
            tabs=('0.5c')
        )
        
        # Scrollbar for address field
        address_scrollbar = ttk.Scrollbar(address_frame, orient="vertical", command=self.address_entry_modify.yview)
        self.address_entry_modify.configure(yscrollcommand=address_scrollbar.set)
        
        self.address_entry_modify.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        address_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.address_entry_modify.bind('<Tab>', lambda e: (self.focus_next_widget(), "break"))
        self.address_entry_modify.bind('<Shift-Tab>', lambda e: (self.focus_previous_widget(), "break"))


        self.address_entry_modify.bind('<FocusIn>', lambda e: self.highlight_focused_widget(e.widget))
        self.address_entry_modify.bind('<FocusOut>', lambda e: self.remove_highlight(e.widget))

        self.create_focusable_widget(self.address_entry_modify)


        # GST Number - Row 2
        tk.Label(fields_frame, text="📊 GST Number *:", font=("Segoe UI", 11, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=2, column=0, sticky="w", padx=10, pady=12)
        
        self.gst_number_entry_modify = self.create_modern_entry(fields_frame, width=30)
        self.gst_number_entry_modify.grid(row=2, column=1, padx=10, pady=12, sticky="ew")

        # Phone Number - Row 3
        tk.Label(fields_frame, text="📞 Phone Number *:", font=("Segoe UI", 11, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=3, column=0, sticky="w", padx=10, pady=12)
        
        self.phone_number_entry_modify = self.create_modern_entry(fields_frame, width=30)
        self.phone_number_entry_modify.grid(row=3, column=1, padx=10, pady=12, sticky="ew")

        # Agent Name - Row 4
        tk.Label(fields_frame, text="🤵 Agent Name *:", font=("Segoe UI", 11, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=4, column=0, sticky="w", padx=10, pady=12)
        
        self.agent_name_entry_modify = self.create_modern_entry(fields_frame, width=30)
        self.agent_name_entry_modify.grid(row=4, column=1, padx=10, pady=12, sticky="ew")

        # Current Party Info Section
        info_frame = tk.Frame(form_container, bg=self.colors['light_bg'], relief="solid", bd=1, padx=15, pady=10)
        info_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=15)
        
        self.current_party_info = tk.Label(
            info_frame,
            text="ℹ️  Select a party code to view and modify details",
            font=("Segoe UI", 10),
            bg=self.colors['light_bg'],
            fg=self.colors['text_muted'],
            justify=tk.LEFT
        )
        self.current_party_info.pack(anchor="w")

        # Required fields note
        note_frame = tk.Frame(form_container, bg=self.colors['card_bg'])
        note_frame.grid(row=3, column=0, columnspan=2, pady=10, sticky="w")
        
        note_label = tk.Label(
            note_frame,
            text="* Fields marked with asterisk are required • Use Tab to navigate between fields",
            font=("Segoe UI", 9, "italic"),
            bg=self.colors['card_bg'],
            fg=self.colors['text_muted']
        )
        note_label.pack()

        # Pack the canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Action buttons frame - Centered at bottom
        action_frame = tk.Frame(center_container, bg=self.colors['light_bg'])
        action_frame.pack(fill=tk.X, pady=15)

        # Center the buttons
        button_container = tk.Frame(action_frame, bg=self.colors['light_bg'])
        button_container.pack(expand=True)

        # Save button
        save_btn = self.create_modern_button(
            button_container,
            "💾 Save Changes (Ctrl+S)",
            self.save_modified_party_details_enhanced,
            style="success",
            width=18,
            height=2
        )
        save_btn.pack(side=tk.LEFT, padx=8)

        # Reset button
        reset_btn = self.create_modern_button(
            button_container,
            "🔄 Reset Form (Alt+R)",
            self.reset_modify_form,
            style="warning",
            width=15,
            height=2
        )
        reset_btn.pack(side=tk.LEFT, padx=8)

        # Back button
        back_btn = self.create_modern_button(
            button_container,
            "← Back to Parties (F2)",
            self.show_party_management,
            style="secondary",
            width=16,
            height=2
        )
        back_btn.pack(side=tk.LEFT, padx=8)

        # Delete button
        delete_btn = self.create_modern_button(
            button_container,
            "🗑️ Delete Party (Del)",
            self.delete_party_confirm,
            style="warning",
            width=15,
            height=2
        )
        delete_btn.pack(side=tk.LEFT, padx=8)

        # Configure grid weights for responsive layout
        fields_frame.grid_columnconfigure(1, weight=1)
        for i in range(5):
            fields_frame.grid_rowconfigure(i, weight=1)

        # Bind keyboard shortcuts
        self.root.bind('<Control-s>', lambda e: self.save_modified_party_details_enhanced())
        self.root.bind('<Alt-r>', lambda e: self.reset_modify_form())
        self.root.bind('<F2>', lambda e: self.show_party_management())
        self.root.bind('<Delete>', lambda e: self.delete_party_confirm())

        # Update the scroll region
        def update_scroll_region(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        canvas.bind("<Configure>", update_scroll_region)

        # Set tab navigation order
        self.focusable_widgets = [
            self.party_code_combobox,
            self.customer_name_entry_modify,
            self.address_entry_modify,
            self.gst_number_entry_modify,
            self.phone_number_entry_modify,
            self.agent_name_entry_modify,
            save_btn,
            reset_btn,
            delete_btn,
            back_btn
        ]

        # Track mouse click focus index
        def update_current_focus(event):
            widget = event.widget
            if widget in self.focusable_widgets:
                self.current_focus_index = self.focusable_widgets.index(widget)

        for w in self.focusable_widgets:
            try:
                w.bind("<FocusIn>", update_current_focus)
            except Exception:
                pass

        # Initialize focus system
        self.current_focus_index = 0
        self.root.bind_all('<Tab>', self.focus_next_widget)
        self.root.bind_all('<Shift-Tab>', self.focus_previous_widget)

        # Default focus
        self.party_code_combobox.focus_set()
        self.show_status_message("✏️ Select a party code to modify details - Use Enter to load party data")

    def load_selected_party_details_enhanced(self):
        """
        Enhanced load function with better feedback and flexible key retrieval (FIXED)
        Fixes syncing by checking for multiple possible key variations (e.g., 'Customer Name' and 'Customer_Name').
        """
        party_code = self.party_code_combobox.get().strip()
        
        if not party_code:
            messagebox.showwarning("⚠️ Selection Required", "Please select a party code first!")
            return
        
        # Helper function for flexible retrieval of data keys
        def _get_flexible_value(data, *keys):
            """Tries multiple keys for retrieval, returns stripped string or empty string."""
            for key in keys:
                value = data.get(key)
                if value is not None:
                    # Return the stripped string value, or the original value if not a string
                    if isinstance(value, str):
                        return value.strip()
                    return value
            return ""

        if party_code in self.party_data:
            details = self.party_data[party_code]
            
            # --- Clear and populate fields using flexible retrieval ---
            
            # 1. Customer Name (Check for "Customer Name", "Customer_Name", "customer_name")
            customer_name = _get_flexible_value(details, "Customer Name", "Customer_Name", "customer_name")
            self.customer_name_entry_modify.delete(0, tk.END)
            self.customer_name_entry_modify.insert(0, customer_name)
            
            # 2. Address (Check for "Address", "Address_Line", "address")
            address = _get_flexible_value(details, "Address", "Address_Line", "address")
            self.address_entry_modify.delete("1.0", tk.END)
            self.address_entry_modify.insert("1.0", address)
            
            # 3. GST Number (Check for "GST Number", "GST_Number", "gstin")
            gst_number = _get_flexible_value(details, "GST Number", "GST_Number", "gstin")
            self.gst_number_entry_modify.delete(0, tk.END)
            self.gst_number_entry_modify.insert(0, gst_number)
            
            # 4. Phone Number (Check for "Phone Number", "Phone_Number", "phone")
            phone_number = _get_flexible_value(details, "Phone Number", "Phone_Number", "phone")
            self.phone_number_entry_modify.delete(0, tk.END)
            self.phone_number_entry_modify.insert(0, phone_number)
            
            # 5. Agent Name (Check for "Agent Name" and "Agent_Name")
            agent_name = _get_flexible_value(details, "Agent Name", "Agent_Name")
            self.agent_name_entry_modify.delete(0, tk.END)
            self.agent_name_entry_modify.insert(0, agent_name)
            
            # --- Update UI ---
            
            # Update info display
            self.current_party_info.config(
                text=f"ℹ️ Currently editing: {customer_name} (Code: {party_code})",
                fg=self.colors['success']
            )
            
            # Set focus to first editable field
            self.customer_name_entry_modify.focus_set()
            
            self.show_status_message(f"📥 Loaded party: {customer_name}")
        else:
            messagebox.showerror("❌ Error", f"Party code {party_code} not found in database!")

    def save_modified_party_details_enhanced(self):
        """Enhanced save function with validation"""
        party_code = self.party_code_combobox.get()
        
        if not party_code:
            messagebox.showerror("❌ Error", "Please select a party code to modify!")
            return
            
        # Validate fields
        if not all([
            self.customer_name_entry_modify.get().strip(),
            self.address_entry_modify.get("1.0", "end-1c").strip(),
            self.gst_number_entry_modify.get().strip(),
            self.phone_number_entry_modify.get().strip(),
            self.agent_name_entry_modify.get().strip()
        ]):
            messagebox.showerror("❌ Validation Error", 
                            "Please fill out all required fields before saving!")
            return

        # Save the modified party details
        self.party_data[party_code] = {
            "Customer Name": self.customer_name_entry_modify.get().strip(),
            "Address": self.address_entry_modify.get("1.0", "end-1c").strip(),
            "GST Number": self.gst_number_entry_modify.get().strip(),
            "Phone Number": self.phone_number_entry_modify.get().strip(),
            "Agent Name": self.agent_name_entry_modify.get().strip()
        }

        # Save to file
        if self.save_data(self.party_ref, self.party_data):
            # Reload the data to ensure consistency
            self.party_ref = db.reference('party_data')
            
            messagebox.showinfo("✅ Success", 
                            f"Party details modified and saved successfully!\n\n"
                            f"Party: {self.customer_name_entry_modify.get().strip()}\n"
                            f"Code: {party_code}")
            
            self.show_party_management()
        else:
            messagebox.showerror("❌ Save Error", 
                            "Failed to save party details.\n"
                            "Please check file permissions and try again.")

    def reset_modify_form(self):
        """Reset the modify form"""
        self.party_code_combobox.set("")
        self.customer_name_entry_modify.delete(0, tk.END)
        self.address_entry_modify.delete("1.0", tk.END)
        self.gst_number_entry_modify.delete(0, tk.END)
        self.phone_number_entry_modify.delete(0, tk.END)
        self.agent_name_entry_modify.delete(0, tk.END)
        
        self.current_party_info.config(
            text="ℹ️  Select a party code to view and modify details",
            fg=self.colors['text_muted']
        )
        
        # Set focus back to party code combobox
        self.party_code_combobox.focus_set()
        
        self.show_status_message("🔄 Form reset - Select a party code to continue")

    def delete_party_confirm(self):
        """Confirm and delete selected party"""
        party_code = self.party_code_combobox.get()
        
        if not party_code:
            messagebox.showwarning("⚠️ Selection Required", "Please select a party code to delete!")
            return
            
        if party_code in self.party_data:
            party_name = self.party_data[party_code].get("Customer Name", "Unknown")
            
            confirm = messagebox.askyesno(
                "🗑️ Confirm Deletion",
                f"Are you sure you want to delete this party?\n\n"
                f"Party Code: {party_code}\n"
                f"Customer Name: {party_name}\n\n"
                f"This action cannot be undone!",
                icon='warning'
            )
            
            if confirm:
                del self.party_data[party_code]
                if self.save_data(self.party_ref, self.party_data):
                    messagebox.showinfo("✅ Success", f"Party '{party_name}' deleted successfully!")
                    self.show_party_management()
                else:
                    messagebox.showerror("❌ Error", "Failed to delete party. Please try again.")


    def show_party_list(self):
        """Modern enhanced version of show_list_page for displaying party list"""
        self.clear_screen()
        self.current_screen = "party_list"
        self.create_navigation_bar()
        self.create_status_bar()
        
        # Reset focusable widgets
        self.focusable_widgets.clear()

        # Main content area
        main_container = tk.Frame(self.root, bg=self.colors['light_bg'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Header frame with search and actions
        header_frame = self.create_modern_frame(main_container, "📋 PARTY LIST - ALL CUSTOMERS")
        header_frame.pack(fill=tk.X, pady=(0, 15))

        # Search and filter section
        search_frame = tk.Frame(header_frame, bg=self.colors['card_bg'])
        search_frame.pack(fill=tk.X, padx=20, pady=15)

        # Search entry
        tk.Label(search_frame, text="🔍 Search:", font=("Segoe UI", 10, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).pack(side=tk.LEFT, padx=5)
        
        self.party_search_entry = self.create_modern_entry(search_frame, width=30)
        self.party_search_entry.pack(side=tk.LEFT, padx=10, pady=5)
        self.party_search_entry.bind('<KeyRelease>', self.filter_party_list)

        # Action buttons
        action_frame = tk.Frame(search_frame, bg=self.colors['card_bg'])
        action_frame.pack(side=tk.RIGHT, padx=10)

        refresh_btn = self.create_modern_button(
            action_frame, "🔄 Refresh (F5)", self.refresh_party_list,
            style="info", width=15, height=1
        )
        refresh_btn.pack(side=tk.LEFT, padx=5)

        export_btn = self.create_modern_button(
            action_frame, "📤 Export CSV", self.export_party_list,
            style="success", width=12, height=1
        )
        export_btn.pack(side=tk.LEFT, padx=5)

        # Table container
        table_container = self.create_modern_frame(main_container, "")
        table_container.pack(fill=tk.BOTH, expand=True, pady=10)

        # Create a frame for table and scrollbars
        table_main = tk.Frame(table_container, bg=self.colors['card_bg'])
        table_main.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        # ✅ Create Treeview with modern styling (fixed invisible heading issue)
        style = ttk.Style()
        style.theme_use("default")

        # --- Fix invisible heading text layout ---
        style.layout("Modern.Treeview.Heading", [
            ("Treeheading.cell", {"sticky": "nswe"}),
            ("Treeheading.border", {"sticky": "nswe", "children": [
                ("Treeheading.padding", {"sticky": "nswe", "children": [
                    ("Treeheading.image", {"side": "right", "sticky": ""}),
                    ("Treeheading.text", {"sticky": "we"})
                ]})
            ]}),
        ])

        # --- Table rows styling ---
        style.configure("Modern.Treeview", 
            font=("Segoe UI", 10),
            rowheight=25,
            background=self.colors['card_bg'],
            fieldbackground=self.colors['card_bg'],
            foreground=self.colors['text_dark']
        )

        # --- Header styling (always visible now) ---
        style.configure("Modern.Treeview.Heading", 
            font=("Segoe UI", 11, "bold"),
            background=self.colors['primary'],
            foreground=self.colors['text_light'],
            relief="flat"
        )

        # --- Hover / pressed effects for headers ---
        style.map("Modern.Treeview.Heading", 
            background=[('active', self.colors['accent']), ('pressed', self.colors['secondary'])],
            relief=[('pressed', 'groove'), ('active', 'ridge')]
        )

        # ✅ Create Treeview - store as instance variable
        self.party_table = ttk.Treeview(
            table_main, 
            columns=("Party Code", "Customer Name", "Address", "GST Number", "Phone Number", "Agent Name"), 
            show="headings",
            style="Modern.Treeview",
            selectmode="extended"
        )


        # Define column headings with larger address column
        columns = {
            "Party Code": {"width": 100, "anchor": "center"},
            "Customer Name": {"width": 200, "anchor": "w"},
            "Address": {"width": 400, "anchor": "w"},
            "GST Number": {"width": 150, "anchor": "center"},
            "Phone Number": {"width": 150, "anchor": "center"},
            "Agent Name": {"width": 150, "anchor": "w"}
        }

        for col, settings in columns.items():
            self.party_table.heading(col, text=col)
            self.party_table.column(col, width=settings["width"], anchor=settings["anchor"])

        # Add scrollbars
        v_scrollbar = ttk.Scrollbar(table_main, orient="vertical", command=self.party_table.yview)
        h_scrollbar = ttk.Scrollbar(table_main, orient="horizontal", command=self.party_table.xview)
        self.party_table.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        # Grid layout for table and scrollbars
        self.party_table.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")

        # Configure grid weights
        table_main.grid_rowconfigure(0, weight=1)
        table_main.grid_columnconfigure(0, weight=1)

        # Status bar for table info
        status_frame = tk.Frame(main_container, bg=self.colors['light_bg'])
        status_frame.pack(fill=tk.X, pady=10)

        self.table_status = tk.Label(
            status_frame,
            text="📊 Total Parties: 0",
            font=("Segoe UI", 10),
            bg=self.colors['light_bg'],
            fg=self.colors['text_dark']
        )
        self.table_status.pack(side=tk.LEFT, padx=10)

        # Action buttons frame
        action_buttons_frame = tk.Frame(main_container, bg=self.colors['light_bg'])
        action_buttons_frame.pack(fill=tk.X, pady=10)

        # Center the buttons
        button_container = tk.Frame(action_buttons_frame, bg=self.colors['light_bg'])
        button_container.pack(expand=True)

        # Edit selected button
        edit_btn = self.create_modern_button(
            button_container,
            "✏️ Edit Selected",
            self.edit_selected_party,
            style="primary",
            width=18,
            height=2
        )
        edit_btn.pack(side=tk.LEFT, padx=8)

        # Delete selected button
        delete_btn = self.create_modern_button(
            button_container,
            "🗑️ Delete Selected",
            self.delete_selected_party,
            style="warning",
            width=18,
            height=2
        )
        delete_btn.pack(side=tk.LEFT, padx=8)

        # View statement button
        statement_btn = self.create_modern_button(
            button_container,
            "📊 View Statement",
            self.show_party_statement_from_list,
            style="info",
            width=16,
            height=2
        )
        statement_btn.pack(side=tk.LEFT, padx=8)

        # Back button
        back_btn = self.create_modern_button(
            button_container,
            "← Back to Parties",
            self.show_party_management,
            style="secondary",
            width=16,
            height=2
        )
        back_btn.pack(side=tk.LEFT, padx=8)

        # NOW populate the table after creating table_status
        self.populate_party_table()

        # Bind keyboard shortcuts - REMOVE global bindings that cause issues
        self.root.bind('<F5>', lambda e: self.refresh_party_list())
        self.root.bind('<F2>', lambda e: self.show_party_management())

        # Bind events directly to the table
        self.party_table.bind('<Double-1>', self.edit_selected_party_event)
        self.party_table.bind('<Return>', self.edit_selected_party_event)
        self.party_table.bind('<Delete>', self.delete_selected_party_event)
        self.party_table.bind('<Button-3>', self.show_party_context_menu)

        # Set focus to search field
        self.party_search_entry.focus_set()

        self.show_status_message("📋 Party list loaded - Double-click or use context menu to edit")

    def filter_party_list(self, event=None):
        """Filter party list based on search term"""
        search_term = self.party_search_entry.get().lower()
        
        if not search_term:
            self.populate_party_table()
            return
        
        filtered_data = {}
        for party_code, details in self.party_data.items():
            if (search_term in party_code.lower() or
                search_term in details.get("Customer Name", "").lower() or
                search_term in details.get("Address", "").lower() or
                search_term in details.get("GST Number", "").lower() or
                search_term in details.get("Phone Number", "").lower() or
                search_term in details.get("Agent Name", "").lower()):
                filtered_data[party_code] = details
        
        self.populate_party_table(filtered_data)

    def refresh_party_list(self, event=None):
        """Refresh the party list"""
        self.party_ref = db.reference('party_data')
        if hasattr(self, 'party_search_entry'):
            self.party_search_entry.delete(0, tk.END)
        self.populate_party_table()
        self.show_status_message("🔄 Party list refreshed")
        return "break"
    
    def export_party_list(self):
        """Export party list to CSV"""
        try:
            from datetime import datetime
            import csv
            
            filename = f"party_list_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            file_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                initialfile=filename
            )
            
            if file_path:
                with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(["Party Code", "Customer Name", "Address", "GST Number", "Phone Number", "Agent Name"])
                    for party_code, details in self.party_data.items():
                        writer.writerow([
                            party_code,
                            details.get("Customer Name", ""),
                            details.get("Address", ""),
                            details.get("GST Number", ""),
                            details.get("Phone Number", ""),
                            details.get("Agent Name", "")
                        ])
                
                messagebox.showinfo("✅ Export Successful", f"Party list exported to:\n{file_path}")
                self.show_status_message("📤 Party list exported successfully")
                
        except Exception as e:
            messagebox.showerror("❌ Export Failed", f"Failed to export party list:\n{str(e)}")

    def populate_party_table(self, data=None):
        """Populate the party table with flexible key handling (GST/Phone/Agent fix)."""

        # Clear old rows
        for item in self.party_table.get_children():
            self.party_table.delete(item)

        # Use provided data or full DB
        party_data = data if data is not None else self.party_data

        # Helper: safe getter that checks many key variations
        def g(d, keys):
            for k in keys:
                if k in d:
                    return d[k]
            return ""

        # Insert rows
        for index, (party_code, details) in enumerate(party_data.items()):
            
            # Row color
            tags = ('evenrow',) if index % 2 == 0 else ('oddrow',)

            # Customer name key variations
            customer_name = g(details, [
                "Customer Name", "Customer_Name", "customer_name",
                "Customer  Name", "Customer name", " Customer Name"
            ])

            # Address
            address = g(details, ["Address", "address"])
            if len(address) > 100:
                address = address[:100] + "..."

            # GST key variations
            gst_number = g(details, [
                "GST Number", "GST_Number", "GSTNumber", "gst_number", "gst"
            ])

            # Phone key variations
            phone_number = g(details, [
                "Phone Number", "Phone", "phone", "phone_number"
            ])

            # Agent Name key variations
            agent_name = g(details, [
                "Agent Name", "Agent_Name", "agent_name", "Agent", "agent"
            ])

            # Insert row
            self.party_table.insert(
                "", "end",
                values=(party_code, customer_name, address, gst_number, phone_number, agent_name),
                tags=tags
            )

        # Row colors
        self.party_table.tag_configure('evenrow', background='#ffffff')
        self.party_table.tag_configure('oddrow', background='#f0f8ff')

        # Update total label
        if hasattr(self, 'table_status'):
            self.table_status.config(text=f"📊 Total Parties: {len(party_data)}")

    def edit_selected_party(self):
        """Edit the selected party - called from button click"""
        selected = self.party_table.selection()
        if not selected:
            messagebox.showwarning("⚠️ Selection Required", "Please select a party to edit!")
            return
        
        party_code = self.party_table.item(selected[0], "values")[0]
        self.selected_party_code = party_code
        self.show_party_modify_with_selection()

    def delete_selected_party(self):
        """Delete selected party/parties - called from button click"""
        selected = self.party_table.selection()
        if not selected:
            messagebox.showwarning("⚠️ Selection Required", "Please select at least one party to delete!")
            return
        
        if len(selected) == 1:
            party_code = self.party_table.item(selected[0], "values")[0]
            party_name = self.party_table.item(selected[0], "values")[1]
            self.delete_party_confirmation(party_code, party_name)
        else:
            self.delete_multiple_parties(selected)

    def edit_selected_party_event(self, event=None):
        """Handle edit event from table bindings"""
        selected = self.party_table.selection()
        if selected:
            party_code = self.party_table.item(selected[0], "values")[0]
            self.selected_party_code = party_code
            self.show_party_modify_with_selection()
        return "break"  # Prevent event propagation

    def delete_selected_party_event(self, event=None):
        """Handle delete event from table bindings"""
        selected = self.party_table.selection()
        if selected:
            if len(selected) == 1:
                party_code = self.party_table.item(selected[0], "values")[0]
                party_name = self.party_table.item(selected[0], "values")[1]
                self.delete_party_confirmation(party_code, party_name)
            else:
                self.delete_multiple_parties(selected)
        return "break"  # Prevent event propagation
    
    

    def show_party_context_menu(self, event):
        """Show right-click context menu for party table"""
        item = self.party_table.identify_row(event.y)
        if item:
            self.party_table.selection_set(item)
            
            context_menu = tk.Menu(self.root, tearoff=0, font=("Segoe UI", 10))
            context_menu.add_command(label="✏️ Edit Party", command=self.edit_selected_party)
            context_menu.add_command(label="📊 View Statement", command=self.show_party_statement_from_list)
            context_menu.add_separator()
            context_menu.add_command(label="🗑️ Delete Party", command=self.delete_selected_party)
            context_menu.add_separator()
            context_menu.add_command(label="📋 Copy Party Code", 
                                command=lambda: self.copy_to_clipboard(self.party_table.item(item, "values")[0]))
            context_menu.add_command(label="👤 Copy Customer Name", 
                                command=lambda: self.copy_to_clipboard(self.party_table.item(item, "values")[1]))
            context_menu.add_command(label="🏠 View Full Address", 
                                command=lambda: self.view_full_address(item))
            
            try:
                context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                context_menu.grab_release()

    def view_full_address(self, item):
        """Show full address in a message box"""
        party_code = self.party_table.item(item, "values")[0]
        customer_name = self.party_table.item(item, "values")[1]
        
        if party_code in self.party_data:
            full_address = self.party_data[party_code].get("Address", "No address available")
            messagebox.showinfo(
                f"🏠 Full Address - {customer_name}",
                f"Party Code: {party_code}\n"
                f"Customer: {customer_name}\n\n"
                f"Full Address:\n{full_address}"
            )

    def copy_to_clipboard(self, text):
        """Copy text to clipboard"""
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.show_status_message(f"📋 Copied to clipboard: {text}")

    def show_party_statement_from_list(self):
        """Show party statement from list selection"""
        selected = self.party_table.selection()
        if not selected:
            messagebox.showwarning("⚠️ Selection Required", "Please select a party first!")
            return
        
        party_code = self.party_table.item(selected[0], "values")[0]
        party_name = self.party_table.item(selected[0], "values")[1]
        
        self.selected_statement_party = f"{party_name} ({party_code})"
        self.show_party_statement()
    
        if hasattr(self, 'party_statement_combobox'):
            self.party_statement_combobox.set(self.selected_statement_party)

    def show_party_modify_with_selection(self):
        """Show modify screen with pre-selected party"""
        # Store the selection before clearing screen
        selected_party_code = getattr(self, 'selected_party_code', None)
        
        self.show_party_modify()
    
        # Auto-select the party in the combobox after screen is loaded
        if selected_party_code and hasattr(self, 'party_code_combobox'):
            self.party_code_combobox.set(selected_party_code)
            # Use after to ensure the combobox is fully loaded
            self.root.after(100, self.load_selected_party_details_enhanced)

    def delete_party_confirmation(self, party_code, party_name):
        """Show confirmation dialog for party deletion"""
        confirm = messagebox.askyesno(
            "🗑️ Confirm Deletion",
            f"Are you sure you want to delete this party?\n\n"
            f"Party Code: {party_code}\n"
            f"Customer Name: {party_name}\n\n"
            f"This action cannot be undone!",
            icon='warning'
        )
        
        if confirm:
            if self.delete_party(party_code):
                self.refresh_party_list()
                messagebox.showinfo("✅ Success", f"Party '{party_name}' deleted successfully!")

    def delete_multiple_parties(self, selected_items):
        """Delete multiple selected parties"""
        party_count = len(selected_items)
        confirm = messagebox.askyesno(
            "🗑️ Confirm Multiple Deletion",
            f"Are you sure you want to delete {party_count} parties?\n\n"
            f"This action cannot be undone!",
            icon='warning'
        )
        if confirm:
            deleted_count = 0
            for item in selected_items:
                party_code = self.party_table.item(item, "values")[0]
                if self.delete_party(party_code):
                    deleted_count += 1
            
            self.refresh_party_list()
            messagebox.showinfo("✅ Success", f"Successfully deleted {deleted_count} out of {party_count} parties!")

    def delete_party(self, party_code):
        """Delete a party from data"""
        if party_code in self.party_data:
            del self.party_data[party_code]
            return self.save_data(self.party_ref, self.party_data)
        return False
        
    def show_party_statement(self):
        messagebox.showinfo("Info", "Party Statement interface would open here")

        
    def show_product_management(self):
        """Show product management screen"""
        self.clear_screen()
        self.current_screen = "product_management"
        self.create_navigation_bar()
        self.create_status_bar()

        content_frame = tk.Frame(self.root, bg=self.colors['light_bg'])
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Header
        header_frame = self.create_modern_frame(content_frame, "📦 PRODUCT MANAGEMENT")
        header_frame.pack(fill=tk.X, pady=(0, 20))

        # Quick actions
        actions_frame = tk.Frame(header_frame, bg=self.colors['card_bg'])
        actions_frame.pack(padx=20, pady=20)

        actions = [
            ("➕ Add Product", self.show_product_entry, "Add new product"),
            ("✏️ Modify Product", self.show_product_modify, "Edit existing product"),
            ("📋 Product List", self.show_product_list, "View all products"),
            ("📊 View Report", self.show_stock_report, "View stock levels")
        ]

        for i, (text, command, tooltip) in enumerate(actions):
            btn = self.create_modern_button(
                actions_frame, text, command,
                style="success", width=20, height=2
            )
            btn.grid(row=i//2, column=i%2, padx=10, pady=10)

        # Set up keyboard navigation
        self.focusable_widgets.clear()
        for widget in actions_frame.winfo_children():
            if isinstance(widget, tk.Button):
                self.create_focusable_widget(widget, "button")

        if self.focusable_widgets:
            self.focusable_widgets[0].focus_set()

        self.show_status_message("📦 Product Management - Use TAB to navigate between options")

    def show_product_entry(self):
        """Modern enhanced version of show_product_entry_page for adding new products"""
        self.clear_screen()
        self.current_screen = "product_entry"
        self.create_navigation_bar()
        self.create_status_bar()
        
        # Reset focusable widgets
        self.focusable_widgets.clear()

        # Main content area with centered layout
        main_container = tk.Frame(self.root, bg=self.colors['light_bg'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Center container to hold everything
        center_container = tk.Frame(main_container, bg=self.colors['light_bg'])
        center_container.place(relx=0.5, rely=0.5, anchor="center", width=900, height=650)

        # Header frame
        header_frame = self.create_modern_frame(center_container, "📦 ADD NEW PRODUCT")
        header_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))

        # Form container with scrollbar
        form_main = tk.Frame(header_frame, bg=self.colors['card_bg'])
        form_main.pack(fill=tk.BOTH, expand=True, padx=25, pady=25)

        # Create a canvas and scrollbar for the form
        canvas = tk.Canvas(form_main, bg=self.colors['card_bg'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(form_main, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.colors['card_bg'])

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Form container - using grid for better alignment
        form_container = scrollable_frame
        form_container.grid_columnconfigure(1, weight=1)

        # Auto-increment product code
        product_codes = [int(code) for code in self.product_data.keys() if code.isdigit()]
        next_product_code = str(max(product_codes) + 1) if product_codes else "1"

        # Product Code - Row 0
        tk.Label(form_container, text="🔢 Product Code:", font=("Segoe UI", 11, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=0, column=0, sticky="w", padx=10, pady=12)
        
        self.product_code_entry = self.create_modern_entry(form_container, width=45)
        self.product_code_entry.insert(0, next_product_code)
        self.product_code_entry.grid(row=0, column=1, padx=10, pady=12, sticky="ew")
        self.product_code_entry.config(state='readonly')  # Auto-generated, read-only

        # Product Name - Row 1
        tk.Label(form_container, text="📝 Product Name *:", font=("Segoe UI", 11, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=1, column=0, sticky="w", padx=10, pady=12)
        
        self.product_name_entry = self.create_modern_entry(form_container, width=25)
        self.product_name_entry.grid(row=1, column=1, padx=10, pady=12, sticky="ew")

        # No. of Case - Row 2
        tk.Label(form_container, text="📦 No. of Case *:", font=("Segoe UI", 11, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=2, column=0, sticky="w", padx=10, pady=12)
        
        self.no_of_case_entry = self.create_modern_entry(form_container, width=25)
        self.no_of_case_entry.grid(row=2, column=1, padx=10, pady=12, sticky="ew")

        # Per Case - Row 3
        tk.Label(form_container, text="🔢 Per Case *:", font=("Segoe UI", 11, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=3, column=0, sticky="w", padx=10, pady=12)
        
        self.per_case_entry = self.create_modern_entry(form_container, width=25)
        self.per_case_entry.grid(row=3, column=1, padx=10, pady=12, sticky="ew")

        # Unit Type - Row 4
        tk.Label(form_container, text="📏 Unit Type (U, N, Box) *:", font=("Segoe UI", 11, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=4, column=0, sticky="w", padx=10, pady=12)
        
        self.unit_type_combobox = self.create_modern_combobox(
            form_container, 
            values=["U", "N", "Box"],
            width=23
        )
        self.unit_type_combobox.grid(row=4, column=1, padx=10, pady=12, sticky="w")

        # Selling Price - Row 5
        tk.Label(form_container, text="💰 Selling Price *:", font=("Segoe UI", 11, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=5, column=0, sticky="w", padx=10, pady=12)
        
        self.selling_price_entry = self.create_modern_entry(form_container, width=25)
        self.selling_price_entry.grid(row=5, column=1, padx=10, pady=12, sticky="ew")

        # Per - Row 6
        tk.Label(form_container, text="📊 Per *:", font=("Segoe UI", 11, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=6, column=0, sticky="w", padx=10, pady=12)
        
        self.per_entry = self.create_modern_entry(form_container, width=25)
        self.per_entry.grid(row=6, column=1, padx=10, pady=12, sticky="ew")

        # Quantity - Row 7
        tk.Label(form_container, text="📈 Quantity *:", font=("Segoe UI", 11, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=7, column=0, sticky="w", padx=10, pady=12)
        
        self.quantity_entry = self.create_modern_entry(form_container, width=25)
        self.quantity_entry.grid(row=7, column=1, padx=10, pady=12, sticky="ew")

        # Discount - Row 8
        tk.Label(form_container, text="🎯 Discount *:", font=("Segoe UI", 11, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=8, column=0, sticky="w", padx=10, pady=12)
        
        self.discount_entry = self.create_modern_entry(form_container, width=25)
        self.discount_entry.grid(row=8, column=1, padx=10, pady=12, sticky="ew")

        # Calculation Preview Section
        preview_frame = tk.Frame(form_container, bg=self.colors['light_bg'], relief="solid", bd=1, padx=15, pady=10)
        preview_frame.grid(row=9, column=0, columnspan=2, sticky="ew", padx=10, pady=20)
        
        self.calculation_preview = tk.Label(
            preview_frame,
            text="🧮 Enter values to see calculation preview",
            font=("Segoe UI", 10),
            bg=self.colors['light_bg'],
            fg=self.colors['text_muted'],
            justify=tk.LEFT
        )
        self.calculation_preview.pack(anchor="w")

        # Bind calculation preview to relevant fields
        calculation_fields = [self.selling_price_entry, self.per_entry, self.quantity_entry, self.discount_entry]
        for field in calculation_fields:
            field.bind('<KeyRelease>', self.update_calculation_preview)

        # Required fields note
        note_frame = tk.Frame(form_container, bg=self.colors['card_bg'])
        note_frame.grid(row=10, column=0, columnspan=2, pady=10, sticky="w")
        
        note_label = tk.Label(
            note_frame,
            text="* Fields marked with asterisk are required • Use Tab to navigate between fields",
            font=("Segoe UI", 9, "italic"),
            bg=self.colors['card_bg'],
            fg=self.colors['text_muted']
        )
        note_label.pack()

        # Pack the canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Action buttons frame - Centered at bottom
        action_frame = tk.Frame(center_container, bg=self.colors['light_bg'])
        action_frame.pack(fill=tk.X, pady=15)

        # Center the buttons
        button_container = tk.Frame(action_frame, bg=self.colors['light_bg'])
        button_container.pack(expand=True)

        # Save button
        save_btn = self.create_modern_button(
            button_container,
            "💾 Save Product (Ctrl+S)",
            self.save_product_details_enhanced,
            style="success",
            width=18,
            height=2
        )
        save_btn.pack(side=tk.LEFT, padx=8)

        # Reset button
        reset_btn = self.create_modern_button(
            button_container,
            "🔄 Reset Form (Alt+R)",
            self.reset_product_form,
            style="warning",
            width=15,
            height=2
        )
        reset_btn.pack(side=tk.LEFT, padx=8)

        # Back button
        back_btn = self.create_modern_button(
            button_container,
            "← Back to Products (F3)",
            self.show_product_management,
            style="secondary",
            width=16,
            height=2
        )
        back_btn.pack(side=tk.LEFT, padx=8)

        # Quick Add another button
        quick_add_btn = self.create_modern_button(
            button_container,
            "⚡ Save & New (Ctrl+N)",
            self.save_and_new_product,
            style="info",
            width=15,
            height=2
        )
        quick_add_btn.pack(side=tk.LEFT, padx=8)

        # Configure grid weights for responsive layout
        for i in range(9):
            form_container.grid_rowconfigure(i, weight=1)

        # Bind keyboard shortcuts
        self.root.bind('<Control-s>', lambda e: self.save_product_details_enhanced())
        self.root.bind('<Alt-r>', lambda e: self.reset_product_form())
        self.root.bind('<Control-n>', lambda e: self.save_and_new_product())
        self.root.bind('<F3>', lambda e: self.show_product_management())

        # Update the scroll region
        def update_scroll_region(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        canvas.bind("<Configure>", update_scroll_region)

        # Set focus to product name field (first editable field)
        self.product_name_entry.focus_set()

        self.show_status_message("📦 Adding new product - Fill in all required fields and press Ctrl+S to save")

    def update_calculation_preview(self, event=None):
        """Update calculation preview based on entered values"""
        try:
            selling_price = float(self.selling_price_entry.get() or 0)
            per = float(self.per_entry.get() or 1)
            quantity = float(self.quantity_entry.get() or 0)
            discount = float(self.discount_entry.get() or 0)
            
            # Calculate basic amounts
            unit_price = selling_price / per if per > 0 else 0
            amount = unit_price * quantity
            discount_amount = (discount / 100) * amount
            final_amount = amount - discount_amount
            
            # Update preview
            preview_text = (
                f"🧮 Calculation Preview:\n"
                f"• Unit Price: ₹{unit_price:.2f}\n"
                f"• Amount: ₹{amount:.2f}\n"
                f"• Discount: ₹{discount_amount:.2f}\n"
                f"• Final Amount: ₹{final_amount:.2f}"
            )
            self.calculation_preview.config(text=preview_text, fg=self.colors['success'])
            
        except (ValueError, ZeroDivisionError):
            self.calculation_preview.config(
                text="🧮 Enter valid numeric values to see calculation preview",
                fg=self.colors['text_muted']
            )

    def clean_all_product_keys(self):
        """Clean all existing product keys in Firebase"""
        try:
            cleaned_data = {}
            changes_made = False
            
            for old_key, value in self.product_data.items():
                # Clean the key
                import re
                new_key = re.sub(r'[\.\$\#\[\]\ /]', '', str(old_key))
                
                if new_key != old_key:
                    print(f"🔄 Cleaning: '{old_key}' → '{new_key}'")
                    changes_made = True
                
                cleaned_data[new_key] = value
            
            if changes_made:
                # Save cleaned data back to Firebase
                self.product_ref.set(cleaned_data)
                self.product_data = cleaned_data
                print("✅ All product keys cleaned successfully!")
                messagebox.showinfo("✅ Success", "All product keys have been cleaned for Firebase compatibility.")
            else:
                print("✅ No cleaning needed - all keys are valid.")
                
        except Exception as e:
            print(f"❌ Error cleaning product keys: {e}")

    def clean_product_keys(self, product_data):
        """
        Clean product keys to make them Firebase-compatible
        Removes: . $ # [ ] / and spaces
        """
        import re
        
        cleaned_data = {}
        invalid_keys = []
        
        for old_key, value in product_data.items():
            # Remove invalid characters: . $ # [ ] / and spaces
            new_key = re.sub(r'[\.\$\#\[\]\ /]', '', str(old_key))
            
            # If key was modified, track it
            if new_key != old_key:
                invalid_keys.append((old_key, new_key))
            
            cleaned_data[new_key] = value
        
        # Log any key changes
        if invalid_keys:
            print("🔄 Cleaned product keys:")
            for old_key, new_key in invalid_keys:
                print(f"   '{old_key}' → '{new_key}'")
        
        return cleaned_data

    def save_and_new_product(self):
        """Save current product and reset form for new entry"""
        if self.save_product_details_enhanced(silent=True):
            self.reset_product_form()
            self.show_status_message("✅ Product saved! Ready for new entry...")

    def save_product_details_enhanced(self, silent=False):
        """Enhanced save function with Firebase validation"""
        print("🔄 DEBUG: Starting save_product_details_enhanced")

        # Get product code
        product_code = self.product_code_entry.get().strip()
        
        # Validate required fields
        if not all([
            product_code,
            self.product_name_entry.get().strip(),
            self.no_of_case_entry.get().strip(),
            self.per_case_entry.get().strip(),
            self.unit_type_combobox.get().strip(),
            self.selling_price_entry.get().strip(),
            self.per_entry.get().strip(),
            self.quantity_entry.get().strip(),
            self.discount_entry.get().strip()
        ]):
            print("❌ DEBUG: Missing fields")
            if not silent:
                messagebox.showerror("❌ Validation Error",
                    "Please fill out all fields before saving.")
            return False

        # Firebase key validation and cleaning
        import re
        # Remove all invalid Firebase characters
        clean_product_code = re.sub(r'[\.\$\#\[\]\ /]', '', product_code)
        
        # Ensure code is not empty after cleaning
        if not clean_product_code:
            clean_product_code = "PRODUCT_" + str(int(datetime.now().timestamp()))
            print(f"🔄 DEBUG: Generated new product code: {clean_product_code}")
        
        # If code was cleaned, notify user
        if clean_product_code != product_code:
            print(f"🔄 DEBUG: Cleaned product code '{product_code}' → '{clean_product_code}'")
            if not silent:
                messagebox.showwarning(
                    "⚠️ Code Cleaned", 
                    f"Product code contained invalid characters and was cleaned:\n"
                    f"'{product_code}' → '{clean_product_code}'\n\n"
                    f"Invalid characters (. $ # [ ] / space) are not allowed in Firebase keys."
                )

        # Final Firebase key validation
        if not re.match(r'^[A-Za-z0-9_-]+$', clean_product_code):
            print(f"❌ DEBUG: Invalid product code '{clean_product_code}'")
            if not silent:
                messagebox.showerror(
                    "❌ Invalid Product Code",
                    "Product Code cannot contain special characters.\n"
                    "Allowed: A-Z, a-z, 0-9, dash (-), underscore (_)."
                )
            return False

        # Check duplicate product code (using cleaned code)
        if clean_product_code in self.product_data:
            print("❌ DEBUG: Duplicate product code")
            if not silent:
                messagebox.showerror("❌ Duplicate Error",
                    f"Product code '{clean_product_code}' already exists!")
            return False

        # Build product entry with simple string values
        product_entry = {
            "Product_Name": self.product_name_entry.get().strip(),
            "No_of_Case": self.no_of_case_entry.get().strip(),
            "Per_Case": self.per_case_entry.get().strip(),
            "Unit_Type": self.unit_type_combobox.get().strip(),
            "Selling_Price": self.selling_price_entry.get().strip(),
            "Per": self.per_entry.get().strip(),
            "Quantity": self.quantity_entry.get().strip(),
            "Discount": self.discount_entry.get().strip()
        }

        print(f"🔄 DEBUG: Saving product '{clean_product_code}' to Firebase…")
        print(f"🔄 DEBUG: Product data: {product_entry}")

        try:
            # METHOD 1: Try direct child set with simple structure
            print("🔄 DEBUG: Attempting direct child set...")
            self.product_ref.child(clean_product_code).set(product_entry)
            
            print("✅ DEBUG: Product saved successfully to Firebase using direct child set")

            # Update local data
            self.product_data[clean_product_code] = product_entry
            
            # Update product list
            self.product_names = [
                self.product_data[key]["Product_Name"]
                for key in self.product_data
            ]

            if not silent:
                messagebox.showinfo(
                    "✅ Success",
                    f"Product saved successfully!\n\n"
                    f"📦 Code: {clean_product_code}\n"
                    f"📝 Name: {product_entry['Product_Name']}"
                )

            self.show_status_message(f"✅ Product '{product_entry['Product_Name']}' saved successfully!")
            
            # Reset form for new entry
            self.reset_product_form()
            
            return True

        except Exception as e:
            print(f"❌ DEBUG: Direct child set failed: {e}")
            
            # METHOD 2: Try with update instead of set
            try:
                print("🔄 DEBUG: Trying update method...")
                update_data = {clean_product_code: product_entry}
                self.product_ref.update(update_data)
                
                print("✅ DEBUG: Product saved successfully using update method")
                
                # Update local data
                self.product_data[clean_product_code] = product_entry
                
                # Update product list
                self.product_names = [
                    self.product_data[key]["Product_Name"]
                    for key in self.product_data
                ]

                if not silent:
                    messagebox.showinfo(
                        "✅ Success",
                        f"Product saved successfully!\n\n"
                        f"📦 Code: {clean_product_code}\n"
                        f"📝 Name: {product_entry['Product_Name']}"
                    )

                self.show_status_message(f"✅ Product '{product_entry['Product_Name']}' saved successfully!")
                self.reset_product_form()
                return True
                
            except Exception as update_error:
                print(f"❌ DEBUG: Update method failed: {update_error}")
                
                # METHOD 3: Last resort - manual Firebase REST API call
                try:
                    print("🔄 DEBUG: Trying manual REST API approach...")
                    success = self.save_product_via_manual_update(clean_product_code, product_entry)
                    
                    if success:
                        print("✅ DEBUG: Product saved successfully using manual update")
                        
                        # Update local data
                        self.product_data[clean_product_code] = product_entry
                        
                        # Update product list
                        self.product_names = [
                            self.product_data[key]["Product_Name"]
                            for key in self.product_data
                        ]

                        if not silent:
                            messagebox.showinfo(
                                "✅ Success",
                                f"Product saved successfully!\n\n"
                                f"📦 Code: {clean_product_code}\n"
                                f"📝 Name: {product_entry['Product_Name']}"
                            )

                        self.show_status_message(f"✅ Product '{product_entry['Product_Name']}' saved successfully!")
                        self.reset_product_form()
                        return True
                    else:
                        raise Exception("Manual update failed")
                        
                except Exception as manual_error:
                    print(f"❌ DEBUG: All methods failed: {manual_error}")
                    if not silent:
                        messagebox.showerror("❌ Save Error",
                            f"Failed to save product details to Firebase after multiple attempts.\n\n"
                            f"Last Error: {str(manual_error)}\n\n"
                            f"Please check your Firebase configuration and try again.")
                    return False
                
    def save_product_via_manual_update(self, product_code, product_data):
        """Manual update method for Firebase using requests"""
        try:
            import requests
            import json
            
            # Get your Firebase URL and credentials
            database_url = "https://onlineinvoiceapplication-default-rtdb.firebaseio.com/"
            secret_key = "YOUR_FIREBASE_SECRET"  # You might need to set up app authentication
            
            # Construct the URL
            url = f"{database_url}/product_data/{product_code}.json"
            
            # Make the PUT request
            response = requests.put(url, json=product_data)
            
            if response.status_code == 200:
                print("✅ Manual Firebase update successful")
                return True
            else:
                print(f"❌ Manual Firebase update failed: {response.status_code} - {response.text}")
                return False
                
        except Exception as e:
            print(f"❌ Manual update error: {e}")
            return False

    def reset_product_form(self):
        """Reset the product form to initial state"""
        # Auto-increment product code
        product_codes = [int(code) for code in self.product_data.keys() if code.isdigit()]
        next_product_code = str(max(product_codes) + 1) if product_codes else "1"
        
        # Clear all fields except product code
        self.product_code_entry.config(state='normal')
        self.product_code_entry.delete(0, tk.END)
        self.product_code_entry.insert(0, next_product_code)
        self.product_code_entry.config(state='readonly')
        
        self.product_name_entry.delete(0, tk.END)
        self.no_of_case_entry.delete(0, tk.END)
        self.per_case_entry.delete(0, tk.END)
        self.unit_type_combobox.set('')
        self.selling_price_entry.delete(0, tk.END)
        self.per_entry.delete(0, tk.END)
        self.quantity_entry.delete(0, tk.END)
        self.discount_entry.delete(0, tk.END)
        
        # Reset calculation preview
        self.calculation_preview.config(
            text="🧮 Enter values to see calculation preview",
            fg=self.colors['text_muted']
        )
        
        # Set focus to product name field
        self.product_name_entry.focus_set()
        
        self.show_status_message("🔄 Form reset - Ready for new product entry")

    def show_product_modify(self):
        """Modern enhanced version of show_product_modify_page for modifying products"""
        self.clear_screen()
        self.current_screen = "product_modify"
        self.create_navigation_bar()
        self.create_status_bar()
        
        # Reset focusable widgets
        self.focusable_widgets.clear()

        # Main content area with centered layout
        main_container = tk.Frame(self.root, bg=self.colors['light_bg'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Center container to hold everything
        center_container = tk.Frame(main_container, bg=self.colors['light_bg'])
        center_container.place(relx=0.5, rely=0.5, anchor="center", width=1000, height=700)

        # Header frame
        header_frame = self.create_modern_frame(center_container, "✏️ MODIFY PRODUCT DETAILS")
        header_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))

        # Form container
        form_main = tk.Frame(header_frame, bg=self.colors['card_bg'])
        form_main.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)

        # Create a canvas and scrollbar for the form
        canvas = tk.Canvas(form_main, bg=self.colors['card_bg'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(form_main, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.colors['card_bg'])

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Form container - using grid for better alignment
        form_container = scrollable_frame
        form_container.grid_columnconfigure(1, weight=1)

        # Product Selection Section
        selection_frame = tk.Frame(form_container, bg=self.colors['card_bg'], relief="solid", bd=1, padx=15, pady=15)
        selection_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=15)
        
        tk.Label(selection_frame, text="🔍 Select Product to Modify:", font=("Segoe UI", 12, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).pack(side=tk.LEFT, padx=5)
        
        # Product Code dropdown with search functionality - FIXED: removed font_size parameter
        self.product_code_combobox_modify = self.create_modern_combobox(
            selection_frame, 
            values=list(self.product_data.keys()),
            width=25,
            state="readonly"
        )
        self.product_code_combobox_modify.pack(side=tk.LEFT, padx=10, pady=5)
        self.product_code_combobox_modify.set("")  # Start with empty selection

        # Quick load button
        load_btn = self.create_modern_button(
            selection_frame,
            "📥 Load Product Details (Enter)",
            self.load_selected_product_details_enhanced,
            style="info",
            width=20,
            height=1
        )
        load_btn.pack(side=tk.LEFT, padx=10)

        # Bind Enter key to load product details
        self.product_code_combobox_modify.bind('<<ComboboxSelected>>', lambda e: self.load_selected_product_details_enhanced())
        self.product_code_combobox_modify.bind('<Return>', lambda e: self.load_selected_product_details_enhanced())

        # Form Fields Section
        fields_frame = tk.Frame(form_container, bg=self.colors['card_bg'])
        fields_frame.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=10, pady=10)

        # Product Name - Row 0 - FIXED: removed font_size parameter
        tk.Label(fields_frame, text="📝 Product Name *:", font=("Segoe UI", 12, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=0, column=0, sticky="w", padx=15, pady=12)
        
        self.product_name_entry_modify = self.create_modern_entry(fields_frame, width=35)
        self.product_name_entry_modify.grid(row=0, column=1, padx=15, pady=12, sticky="ew")

        # No. of Case - Row 1 - FIXED: removed font_size parameter
        tk.Label(fields_frame, text="📦 No. of Case *:", font=("Segoe UI", 12, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=1, column=0, sticky="w", padx=15, pady=12)
        
        self.no_of_case_entry_modify = self.create_modern_entry(fields_frame, width=35)
        self.no_of_case_entry_modify.grid(row=1, column=1, padx=15, pady=12, sticky="ew")

        # Per Case - Row 2 - FIXED: removed font_size parameter
        tk.Label(fields_frame, text="🔢 Per Case *:", font=("Segoe UI", 12, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=2, column=0, sticky="w", padx=15, pady=12)
        
        self.per_case_entry_modify = self.create_modern_entry(fields_frame, width=35)
        self.per_case_entry_modify.grid(row=2, column=1, padx=15, pady=12, sticky="ew")

        # Unit Type - Row 3 - FIXED: removed font_size parameter
        tk.Label(fields_frame, text="📏 Unit Type (U, N, Box) *:", font=("Segoe UI", 12, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=3, column=0, sticky="w", padx=15, pady=12)
        
        self.unit_type_combobox_modify = self.create_modern_combobox(
            fields_frame, 
            values=["U", "N", "Box"],
            width=33
        )
        self.unit_type_combobox_modify.grid(row=3, column=1, padx=15, pady=12, sticky="w")

        # Selling Price - Row 4 - FIXED: removed font_size parameter
        tk.Label(fields_frame, text="💰 Selling Price *:", font=("Segoe UI", 12, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=4, column=0, sticky="w", padx=15, pady=12)
        
        self.selling_price_entry_modify = self.create_modern_entry(fields_frame, width=35)
        self.selling_price_entry_modify.grid(row=4, column=1, padx=15, pady=12, sticky="ew")

        # Per - Row 5 - FIXED: removed font_size parameter
        tk.Label(fields_frame, text="📊 Per *:", font=("Segoe UI", 12, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=5, column=0, sticky="w", padx=15, pady=12)
        
        self.per_entry_modify = self.create_modern_entry(fields_frame, width=35)
        self.per_entry_modify.grid(row=5, column=1, padx=15, pady=12, sticky="ew")

        # Quantity - Row 6 - FIXED: removed font_size parameter
        tk.Label(fields_frame, text="📈 Quantity *:", font=("Segoe UI", 12, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=6, column=0, sticky="w", padx=15, pady=12)
        
        self.quantity_entry_modify = self.create_modern_entry(fields_frame, width=35)
        self.quantity_entry_modify.grid(row=6, column=1, padx=15, pady=12, sticky="ew")

        # Discount - Row 7 - FIXED: removed font_size parameter
        tk.Label(fields_frame, text="🎯 Discount *:", font=("Segoe UI", 12, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).grid(row=7, column=0, sticky="w", padx=15, pady=12)
        
        self.discount_entry_modify = self.create_modern_entry(fields_frame, width=35)
        self.discount_entry_modify.grid(row=7, column=1, padx=15, pady=12, sticky="ew")

        # Calculation Preview Section
        preview_frame = tk.Frame(form_container, bg=self.colors['light_bg'], relief="solid", bd=1, padx=20, pady=15)
        preview_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=20)
        
        self.calculation_preview_modify = tk.Label(
            preview_frame,
            text="🧮 Select a product to see calculation preview",
            font=("Segoe UI", 11),
            bg=self.colors['light_bg'],
            fg=self.colors['text_muted'],
            justify=tk.LEFT
        )
        self.calculation_preview_modify.pack(anchor="w")

        # Current Product Info Section
        info_frame = tk.Frame(form_container, bg=self.colors['light_bg'], relief="solid", bd=1, padx=15, pady=10)
        info_frame.grid(row=3, column=0, columnspan=2, sticky="ew", padx=10, pady=15)
        
        self.current_product_info = tk.Label(
            info_frame,
            text="ℹ️  Select a product code to view and modify details",
            font=("Segoe UI", 10),
            bg=self.colors['light_bg'],
            fg=self.colors['text_muted'],
            justify=tk.LEFT
        )
        self.current_product_info.pack(anchor="w")

        # Required fields note
        note_frame = tk.Frame(form_container, bg=self.colors['card_bg'])
        note_frame.grid(row=4, column=0, columnspan=2, pady=10, sticky="w")
        
        note_label = tk.Label(
            note_frame,
            text="* Fields marked with asterisk are required • Use Tab to navigate between fields",
            font=("Segoe UI", 9, "italic"),
            bg=self.colors['card_bg'],
            fg=self.colors['text_muted']
        )
        note_label.pack()

        # Pack the canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Action buttons frame - Centered at bottom
        action_frame = tk.Frame(center_container, bg=self.colors['light_bg'])
        action_frame.pack(fill=tk.X, pady=20)

        # Center the buttons
        button_container = tk.Frame(action_frame, bg=self.colors['light_bg'])
        button_container.pack(expand=True)

        # Save button
        save_btn = self.create_modern_button(
            button_container,
            "💾 Save Changes (Ctrl+S)",
            self.save_modified_product_details_enhanced,
            style="success",
            width=20,
            height=2
        )
        save_btn.pack(side=tk.LEFT, padx=10)

        # Reset button
        reset_btn = self.create_modern_button(
            button_container,
            "🔄 Reset Form (Alt+R)",
            self.reset_modify_product_form,
            style="warning",
            width=16,
            height=2
        )
        reset_btn.pack(side=tk.LEFT, padx=10)

        # Back button
        back_btn = self.create_modern_button(
            button_container,
            "← Back to Products (F3)",
            self.show_product_management,
            style="secondary",
            width=18,
            height=2
        )
        back_btn.pack(side=tk.LEFT, padx=10)

        # Delete button
        delete_btn = self.create_modern_button(
            button_container,
            "🗑️ Delete Product (Del)",
            self.delete_product_confirm,
            style="warning",
            width=18,
            height=2
        )
        delete_btn.pack(side=tk.LEFT, padx=10)

        # Configure grid weights for responsive layout
        fields_frame.grid_columnconfigure(1, weight=1)
        for i in range(8):
            fields_frame.grid_rowconfigure(i, weight=1)

        # Bind keyboard shortcuts
        self.root.bind('<Control-s>', lambda e: self.save_modified_product_details_enhanced())
        self.root.bind('<Alt-r>', lambda e: self.reset_modify_product_form())
        self.root.bind('<F3>', lambda e: self.show_product_management())
        self.root.bind('<Delete>', lambda e: self.delete_product_confirm())

        # Bind calculation preview to relevant fields
        calculation_fields = [self.selling_price_entry_modify, self.per_entry_modify, self.quantity_entry_modify, self.discount_entry_modify]
        for field in calculation_fields:
            field.bind('<KeyRelease>', self.update_calculation_preview_modify)

        # Update the scroll region
        def update_scroll_region(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        canvas.bind("<Configure>", update_scroll_region)

        # Set focus to product code combobox
        self.product_code_combobox_modify.focus_set()

        # Configure proper tab order
        self.setup_tab_order_product_modify()

        self.show_status_message("✏️ Select a product code to modify details - Use Enter to load product data")


    def setup_tab_order_product_modify(self):
        """Setup proper tab order for product modify form"""
        # Clear existing focusable widgets
        self.focusable_widgets.clear()
        
        # Define tab order - include combobox first, then editable fields
        tab_order = [
            self.product_code_combobox_modify,
            self.product_name_entry_modify,
            self.no_of_case_entry_modify,
            self.per_case_entry_modify,
            self.unit_type_combobox_modify,
            self.selling_price_entry_modify,
            self.per_entry_modify,
            self.quantity_entry_modify,
            self.discount_entry_modify
        ]
        
        # Add widgets to focusable list in correct order
        for widget in tab_order:
            self.create_focusable_widget(widget)
        
        # Set current focus index
        self.current_focus_index = 0

    def load_selected_product_details_enhanced(self):
        """
        Enhanced load function with better feedback and flexible key retrieval (FIXED)
        Fixes syncing by checking for multiple possible key variations (e.g., 'Product Name' and 'Product_Name').
        """
        product_code = self.product_code_combobox_modify.get().strip()
        
        if not product_code:
            messagebox.showwarning("⚠️ Selection Required", "Please select a product code first!")
            return
        
        # Helper function for flexible retrieval of data keys
        def _get_flexible_value(data, *keys):
            """Tries multiple keys for retrieval, returns stripped string or empty string."""
            for key in keys:
                value = data.get(key)
                if value is not None:
                    # Return the stripped string value, or the original value if not a string
                    if isinstance(value, str):
                        return value.strip()
                    return value
            return ""
            
        if product_code in self.product_data:
            details = self.product_data[product_code]
            
            # --- Clear and populate fields using flexible retrieval ---
            
            # 1. Product Name (Checks for "Product Name" and "Product_Name")
            product_name = _get_flexible_value(details, "Product Name", "Product_Name")
            self.product_name_entry_modify.delete(0, tk.END)
            self.product_name_entry_modify.insert(0, product_name)
            
            # 2. No. of Case (Checks for "No. of Case" and "No_of_Case")
            no_of_case = _get_flexible_value(details, "No. of Case", "No_of_Case")
            self.no_of_case_entry_modify.delete(0, tk.END)
            self.no_of_case_entry_modify.insert(0, no_of_case)
            
            # 3. Per Case (Checks for "Per Case" and "Per_Case")
            per_case = _get_flexible_value(details, "Per Case", "Per_Case")
            self.per_case_entry_modify.delete(0, tk.END)
            self.per_case_entry_modify.insert(0, per_case)
            
            # 4. Unit Type (Checks for "Unit Type" and "Unit_Type")
            unit_type = _get_flexible_value(details, "Unit Type", "Unit_Type")
            self.unit_type_combobox_modify.set(unit_type)
            
            # 5. Selling Price (Checks for "Selling Price" and "Selling_Price")
            selling_price = _get_flexible_value(details, "Selling Price", "Selling_Price")
            self.selling_price_entry_modify.delete(0, tk.END)
            self.selling_price_entry_modify.insert(0, selling_price)
            
            # 6. Per (Likely consistent key)
            per = _get_flexible_value(details, "Per")
            self.per_entry_modify.delete(0, tk.END)
            self.per_entry_modify.insert(0, per)
            
            # 7. Quantity (Likely consistent key)
            quantity = _get_flexible_value(details, "Quantity")
            self.quantity_entry_modify.delete(0, tk.END)
            self.quantity_entry_modify.insert(0, quantity)
            
            # 8. Discount (Likely consistent key)
            discount = _get_flexible_value(details, "Discount")
            self.discount_entry_modify.delete(0, tk.END)
            self.discount_entry_modify.insert(0, discount)
            
            # --- Update UI ---
            self.current_product_info.config(
                text=f"ℹ️ Currently editing: {product_name} (Code: {product_code})",
                fg=self.colors['success']
            )
            
            # Update calculation preview
            self.update_calculation_preview_modify()
            
            # Set focus to first editable field
            self.product_name_entry_modify.focus_set()
            
            self.show_status_message(f"📥 Loaded product: {product_name}")
        else:
            messagebox.showerror("❌ Error", f"Product code {product_code} not found in database!")

    def update_calculation_preview_modify(self, event=None):
        """Update calculation preview based on entered values"""
        try:
            selling_price = float(self.selling_price_entry_modify.get() or 0)
            per = float(self.per_entry_modify.get() or 1)
            quantity = float(self.quantity_entry_modify.get() or 0)
            discount = float(self.discount_entry_modify.get() or 0)
            
            # Calculate basic amounts
            unit_price = selling_price / per if per > 0 else 0
            amount = unit_price * quantity
            discount_amount = (discount / 100) * amount
            final_amount = amount - discount_amount
            
            # Update preview
            preview_text = (
                f"🧮 Calculation Preview:\n"
                f"• Unit Price: ₹{unit_price:.2f}\n"
                f"• Amount: ₹{amount:.2f}\n"
                f"• Discount: ₹{discount_amount:.2f}\n"
                f"• Final Amount: ₹{final_amount:.2f}"
            )
            self.calculation_preview_modify.config(text=preview_text, fg=self.colors['success'])
            
        except (ValueError, ZeroDivisionError):
            self.calculation_preview_modify.config(
                text="🧮 Enter valid numeric values to see calculation preview",
                fg=self.colors['text_muted']
            )

    def save_modified_product_details_enhanced(self):
        """Enhanced save function for modified product details"""
        product_code = self.product_code_combobox_modify.get()
        
        if not product_code:
            messagebox.showerror("❌ Error", "Please select a product code to modify!")
            return
            
        # Validate fields
        if not all([
            self.product_name_entry_modify.get().strip(),
            self.no_of_case_entry_modify.get().strip(),
            self.per_case_entry_modify.get().strip(),
            self.unit_type_combobox_modify.get().strip(),
            self.selling_price_entry_modify.get().strip(),
            self.per_entry_modify.get().strip(),
            self.quantity_entry_modify.get().strip(),
            self.discount_entry_modify.get().strip()
        ]):
            messagebox.showerror("❌ Validation Error", 
                            "Please fill out all required fields before saving!")
            return

        # Build product entry
        product_entry = {
            "Product Name": self.product_name_entry_modify.get().strip(),
            "No. of Case": self.no_of_case_entry_modify.get().strip(),
            "Per Case": self.per_case_entry_modify.get().strip(),
            "Unit Type": self.unit_type_combobox_modify.get().strip(),
            "Selling Price": self.selling_price_entry_modify.get().strip(),
            "Per": self.per_entry_modify.get().strip(),
            "Quantity": self.quantity_entry_modify.get().strip(),
            "Discount": self.discount_entry_modify.get().strip()
        }

        try:
            # Save to Firebase
            self.product_ref.child(product_code).set(product_entry)
            
            # Update local data
            self.product_data[product_code] = product_entry
            
            # Update product list
            self.product_names = [
                self.product_data[key]["Product Name"]
                for key in self.product_data
            ]

            messagebox.showinfo("✅ Success", 
                            f"Product details modified and saved successfully!\n\n"
                            f"Product: {product_entry['Product Name']}\n"
                            f"Code: {product_code}")
            
            self.show_product_management()
            
        except Exception as e:
            messagebox.showerror("❌ Save Error", 
                            f"Failed to save product details to Firebase.\n\n"
                            f"Error: {str(e)}")

    def reset_modify_product_form(self):
        """Reset the product modify form"""
        self.product_code_combobox_modify.set("")
        self.product_name_entry_modify.delete(0, tk.END)
        self.no_of_case_entry_modify.delete(0, tk.END)
        self.per_case_entry_modify.delete(0, tk.END)
        self.unit_type_combobox_modify.set('')
        self.selling_price_entry_modify.delete(0, tk.END)
        self.per_entry_modify.delete(0, tk.END)
        self.quantity_entry_modify.delete(0, tk.END)
        self.discount_entry_modify.delete(0, tk.END)
        
        self.current_product_info.config(
            text="ℹ️  Select a product code to view and modify details",
            fg=self.colors['text_muted']
        )
        
        # Reset calculation preview
        self.calculation_preview_modify.config(
            text="🧮 Select a product to see calculation preview",
            fg=self.colors['text_muted']
        )
        
        # Set focus back to product code combobox
        self.product_code_combobox_modify.focus_set()
        
        self.show_status_message("🔄 Form reset - Select a product code to continue")

    def delete_product_confirm(self):
        """Confirm and delete selected product"""
        product_code = self.product_code_combobox_modify.get()
        
        if not product_code:
            messagebox.showwarning("⚠️ Selection Required", "Please select a product code to delete!")
            return
            
        if product_code in self.product_data:
            product_name = self.product_data[product_code].get("Product Name", "Unknown")
            
            confirm = messagebox.askyesno(
                "🗑️ Confirm Deletion",
                f"Are you sure you want to delete this product?\n\n"
                f"Product Code: {product_code}\n"
                f"Product Name: {product_name}\n\n"
                f"This action cannot be undone!",
                icon='warning'
            )
            
            if confirm:
                del self.product_data[product_code]
                if self.save_data(self.product_ref, self.product_data):
                    messagebox.showinfo("✅ Success", f"Product '{product_name}' deleted successfully!")
                    self.show_product_management()
                else:
                    messagebox.showerror("❌ Error", "Failed to delete product. Please try again.")

    def show_product_list(self):
        """Modern enhanced version of show_product_list_page for displaying product list"""
        self.clear_screen()
        self.current_screen = "product_list"
        self.create_navigation_bar()
        self.create_status_bar()
        
        # Reset focusable widgets
        self.focusable_widgets.clear()

        # Main content area
        main_container = tk.Frame(self.root, bg=self.colors['light_bg'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Header frame with search and actions
        header_frame = self.create_modern_frame(main_container, "📦 PRODUCT LIST - ALL PRODUCTS")
        header_frame.pack(fill=tk.X, pady=(0, 15))

        # Search and filter section
        search_frame = tk.Frame(header_frame, bg=self.colors['card_bg'])
        search_frame.pack(fill=tk.X, padx=20, pady=15)

        # Search entry
        tk.Label(search_frame, text="🔍 Search:", font=("Segoe UI", 10, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).pack(side=tk.LEFT, padx=5)
        
        self.product_search_entry = self.create_modern_entry(search_frame, width=30)
        self.product_search_entry.pack(side=tk.LEFT, padx=10, pady=5)
        self.product_search_entry.bind('<KeyRelease>', self.filter_product_list)

        # Action buttons
        action_frame = tk.Frame(search_frame, bg=self.colors['card_bg'])
        action_frame.pack(side=tk.RIGHT, padx=10)

        refresh_btn = self.create_modern_button(
            action_frame, "🔄 Refresh (F5)", self.refresh_product_list,
            style="info", width=15, height=1
        )
        refresh_btn.pack(side=tk.LEFT, padx=5)

        export_btn = self.create_modern_button(
            action_frame, "📤 Export CSV", self.export_product_list,
            style="success", width=12, height=1
        )
        export_btn.pack(side=tk.LEFT, padx=5)

        # Table container
        table_container = self.create_modern_frame(main_container, "")
        table_container.pack(fill=tk.BOTH, expand=True, pady=10)

        # Create a frame for table and scrollbars
        table_main = tk.Frame(table_container, bg=self.colors['card_bg'])
        table_main.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        # ✅ Create Treeview with modern styling (fixed heading visibility issue)
        style = ttk.Style()
        style.theme_use("default")

        # --- Fix invisible heading text ---
        style.layout("Modern.Treeview.Heading", [
            ("Treeheading.cell", {"sticky": "nswe"}),
            ("Treeheading.border", {"sticky": "nswe", "children": [
                ("Treeheading.padding", {"sticky": "nswe", "children": [
                    ("Treeheading.image", {"side": "right", "sticky": ""}),
                    ("Treeheading.text", {"sticky": "we"})
                ]})
            ]}),
        ])

        # --- Normal row styling ---
        style.configure("Modern.Treeview",
            font=("Segoe UI", 10),
            rowheight=25,
            background=self.colors['card_bg'],
            fieldbackground=self.colors['card_bg'],
            foreground=self.colors['text_dark']
        )

        # --- Heading styling (always visible now) ---
        style.configure("Modern.Treeview.Heading",
            font=("Segoe UI", 11, "bold"),
            background=self.colors['primary'],
            foreground=self.colors['text_light'],
            relief="flat"
        )

        # --- Hover / click effects ---
        style.map("Modern.Treeview.Heading",
            background=[('active', self.colors['accent']), ('pressed', self.colors['secondary'])],
            relief=[('pressed', 'groove'), ('active', 'ridge')]
        )

        # ✅ Create Treeview - store as instance variable
        self.product_table = ttk.Treeview(
            table_main, 
            columns=("Product Code", "Product Name", "No. of Case", "Per Case", 
                    "Unit Type", "Selling Price", "Per", "Quantity", "Discount"), 
            show="headings",
            style="Modern.Treeview",
            selectmode="extended"
        )


        # Define column headings with appropriate widths
        columns = {
            "Product Code": {"width": 120, "anchor": "center"},
            "Product Name": {"width": 250, "anchor": "w"},
            "No. of Case": {"width": 100, "anchor": "center"},
            "Per Case": {"width": 100, "anchor": "center"},
            "Unit Type": {"width": 80, "anchor": "center"},
            "Selling Price": {"width": 120, "anchor": "center"},
            "Per": {"width": 80, "anchor": "center"},
            "Quantity": {"width": 100, "anchor": "center"},
            "Discount": {"width": 100, "anchor": "center"}
        }

        for col, settings in columns.items():
            self.product_table.heading(col, text=col)
            self.product_table.column(col, width=settings["width"], anchor=settings["anchor"])

        # Add scrollbars
        v_scrollbar = ttk.Scrollbar(table_main, orient="vertical", command=self.product_table.yview)
        h_scrollbar = ttk.Scrollbar(table_main, orient="horizontal", command=self.product_table.xview)
        self.product_table.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        # Grid layout for table and scrollbars
        self.product_table.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")

        # Configure grid weights
        table_main.grid_rowconfigure(0, weight=1)
        table_main.grid_columnconfigure(0, weight=1)

        # Status bar for table info
        status_frame = tk.Frame(main_container, bg=self.colors['light_bg'])
        status_frame.pack(fill=tk.X, pady=10)

        self.product_table_status = tk.Label(
            status_frame,
            text="📊 Total Products: 0",
            font=("Segoe UI", 10),
            bg=self.colors['light_bg'],
            fg=self.colors['text_dark']
        )
        self.product_table_status.pack(side=tk.LEFT, padx=10)

        # Action buttons frame
        action_buttons_frame = tk.Frame(main_container, bg=self.colors['light_bg'])
        action_buttons_frame.pack(fill=tk.X, pady=10)

        # Center the buttons
        button_container = tk.Frame(action_buttons_frame, bg=self.colors['light_bg'])
        button_container.pack(expand=True)

        # Edit selected button
        edit_btn = self.create_modern_button(
            button_container,
            "✏️ Edit Selected",
            self.edit_selected_product,
            style="primary",
            width=18,
            height=2
        )
        edit_btn.pack(side=tk.LEFT, padx=8)

        # Delete selected button
        delete_btn = self.create_modern_button(
            button_container,
            "🗑️ Delete Selected",
            self.delete_selected_product,
            style="warning",
            width=18,
            height=2
        )
        delete_btn.pack(side=tk.LEFT, padx=8)

        # View details button
        details_btn = self.create_modern_button(
            button_container,
            "📋 View Details",
            self.show_product_details,
            style="info",
            width=16,
            height=2
        )
        details_btn.pack(side=tk.LEFT, padx=8)

        # Back button
        back_btn = self.create_modern_button(
            button_container,
            "← Back to Products",
            self.show_product_management,
            style="secondary",
            width=16,
            height=2
        )
        back_btn.pack(side=tk.LEFT, padx=8)

        # Populate the table
        self.populate_product_table()

        # Bind keyboard shortcuts
        self.root.bind('<F5>', lambda e: self.refresh_product_list())
        self.root.bind('<F3>', lambda e: self.show_product_management())

        # Bind events directly to the table
        self.product_table.bind('<Double-1>', self.edit_selected_product_event)
        self.product_table.bind('<Return>', self.edit_selected_product_event)
        self.product_table.bind('<Delete>', self.delete_selected_product_event)
        self.product_table.bind('<Button-3>', self.show_product_context_menu)

        # Set focus to search field
        self.product_search_entry.focus_set()

        self.show_status_message("📦 Product list loaded - Double-click or use context menu to edit")

    def populate_product_table(self, data=None):
        """Populate the product table with data"""
        # Clear existing data
        for item in self.product_table.get_children():
            self.product_table.delete(item)
        
        # Use provided data or all product data
        product_data = data if data is not None else self.product_data
        
        # Insert data with alternating row colors
        for index, (product_code, details) in enumerate(product_data.items()):
            tags = ('evenrow',) if index % 2 == 0 else ('oddrow',)
            
            # Handle both field name formats
            product_name = details.get('Product_Name') or details.get('Product Name', '')
            no_of_case = details.get('No_of_Case') or details.get('No. of Case', '')
            per_case = details.get('Per_Case') or details.get('Per Case', '')
            unit_type = details.get('Unit_Type') or details.get('Unit Type', '')
            selling_price = details.get('Selling_Price') or details.get('Selling Price', '')
            
            # Format the Per value with Unit Type
            per = details.get("Per", "")
            per_value = f"{per} {unit_type}" if per and unit_type else per or unit_type
            
            self.product_table.insert("", "end", values=(
                product_code,
                product_name,
                no_of_case,
                per_case,
                unit_type,
                f"₹{selling_price}",
                per_value,
                details.get("Quantity", ""),
                f"{details.get('Discount', '')}%"
            ), tags=tags)

        # Configure row colors with better contrast
        self.product_table.tag_configure('evenrow', background='#ffffff')
        self.product_table.tag_configure('oddrow', background='#f0f8ff')
        
        # Update status
        if hasattr(self, 'product_table_status'):
            self.product_table_status.config(text=f"📊 Total Products: {len(product_data)}")

    def filter_product_list(self, event=None):
        """Filter product list based on search term"""
        search_term = self.product_search_entry.get().lower()
        
        if not search_term:
            self.populate_product_table()
            return
        
        filtered_data = {}
        for product_code, details in self.product_data.items():
            if (search_term in product_code.lower() or
                search_term in details.get("Product Name", "").lower() or
                search_term in details.get("Unit Type", "").lower() or
                search_term in str(details.get("Selling Price", "")).lower() or
                search_term in str(details.get("Per", "")).lower()):
                filtered_data[product_code] = details
        
        self.populate_product_table(filtered_data)

    def refresh_product_list(self, event=None):
        """Refresh the product list"""
        self.product_ref = db.reference('product_data')
        if hasattr(self, 'product_search_entry'):
            self.product_search_entry.delete(0, tk.END)
        self.populate_product_table()
        self.show_status_message("🔄 Product list refreshed")
        return "break"

    def edit_selected_product_event(self, event=None):
        """Handle edit event from table bindings"""
        selected = self.product_table.selection()
        if selected:
            product_code = self.product_table.item(selected[0], "values")[0]
            self.selected_product_code = product_code
            self.show_product_modify_with_selection()
        return "break"

    def delete_selected_product_event(self, event=None):
        """Handle delete event from table bindings"""
        selected = self.product_table.selection()
        if selected:
            if len(selected) == 1:
                product_code = self.product_table.item(selected[0], "values")[0]
                product_name = self.product_table.item(selected[0], "values")[1]
                self.delete_product_confirmation(product_code, product_name)
            else:
                self.delete_multiple_products(selected)
        return "break"

    def edit_selected_product(self):
        """Edit the selected product - called from button click"""
        selected = self.product_table.selection()
        if not selected:
            messagebox.showwarning("⚠️ Selection Required", "Please select a product to edit!")
            return
        
        product_code = self.product_table.item(selected[0], "values")[0]
        self.selected_product_code = product_code
        self.show_product_modify_with_selection()

    def show_product_modify_with_selection(self):
        """Show modify screen with pre-selected product"""
        selected_product_code = getattr(self, 'selected_product_code', None)
        
        self.show_product_modify()
        
        # Auto-select the product in the combobox after screen is loaded
        if selected_product_code and hasattr(self, 'product_code_combobox_modify'):
            self.product_code_combobox_modify.set(selected_product_code)
            # Use after to ensure the combobox is fully loaded
            self.root.after(100, self.load_selected_product_details_enhanced)

    def delete_selected_product(self):
        """Delete selected product/products - called from button click"""
        selected = self.product_table.selection()
        if not selected:
            messagebox.showwarning("⚠️ Selection Required", "Please select at least one product to delete!")
            return
        
        if len(selected) == 1:
            product_code = self.product_table.item(selected[0], "values")[0]
            product_name = self.product_table.item(selected[0], "values")[1]
            self.delete_product_confirmation(product_code, product_name)
        else:
            self.delete_multiple_products(selected)

    def delete_product_confirmation(self, product_code, product_name):
        """Show confirmation dialog for product deletion"""
        confirm = messagebox.askyesno(
            "🗑️ Confirm Deletion",
            f"Are you sure you want to delete this product?\n\n"
            f"Product Code: {product_code}\n"
            f"Product Name: {product_name}\n\n"
            f"This action cannot be undone!",
            icon='warning'
        )
        
        if confirm:
            if self.delete_product_from_data(product_code):
                self.refresh_product_list()
                messagebox.showinfo("✅ Success", f"Product '{product_name}' deleted successfully!")

    def delete_multiple_products(self, selected_items):
        """Delete multiple selected products"""
        product_count = len(selected_items)
        confirm = messagebox.askyesno(
            "🗑️ Confirm Multiple Deletion",
            f"Are you sure you want to delete {product_count} products?\n\n"
            f"This action cannot be undone!",
            icon='warning'
        )
        
        if confirm:
            deleted_count = 0
            for item in selected_items:
                product_code = self.product_table.item(item, "values")[0]
                if self.delete_product_from_data(product_code):
                    deleted_count += 1
            
            self.refresh_product_list()
            messagebox.showinfo("✅ Success", f"Successfully deleted {deleted_count} out of {product_count} products!")

    def delete_product_from_data(self, product_code):
        """Delete a product from data"""
        if product_code in self.product_data:
            del self.product_data[product_code]
            return self.save_data(self.product_ref, self.product_data)
        return False

    def show_product_context_menu(self, event):
        """Show right-click context menu for product table"""
        item = self.product_table.identify_row(event.y)
        if item:
            self.product_table.selection_set(item)
            
            context_menu = tk.Menu(self.root, tearoff=0, font=("Segoe UI", 10))
            context_menu.add_command(label="✏️ Edit Product", command=self.edit_selected_product)
            context_menu.add_command(label="📋 View Details", command=self.show_product_details)
            context_menu.add_separator()
            context_menu.add_command(label="🗑️ Delete Product", command=self.delete_selected_product)
            context_menu.add_separator()
            context_menu.add_command(label="📦 Copy Product Code", 
                                command=lambda: self.copy_to_clipboard(self.product_table.item(item, "values")[0]))
            context_menu.add_command(label="📝 Copy Product Name", 
                                command=lambda: self.copy_to_clipboard(self.product_table.item(item, "values")[1]))
            context_menu.add_command(label="💰 View Pricing", 
                                command=lambda: self.view_product_pricing(item))
            
            try:
                context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                context_menu.grab_release()

    def view_product_pricing(self, item):
        """Show product pricing details in a message box"""
        product_code = self.product_table.item(item, "values")[0]
        product_name = self.product_table.item(item, "values")[1]
        
        if product_code in self.product_data:
            details = self.product_data[product_code]
            messagebox.showinfo(
                f"💰 Pricing Details - {product_name}",
                f"Product Code: {product_code}\n"
                f"Product: {product_name}\n\n"
                f"📊 Pricing Information:\n"
                f"• Selling Price: ₹{details.get('Selling Price', 'N/A')}\n"
                f"• Per: {details.get('Per', 'N/A')} {details.get('Unit Type', '')}\n"
                f"• Quantity: {details.get('Quantity', 'N/A')}\n"
                f"• Discount: {details.get('Discount', 'N/A')}%\n"
                f"• No. of Case: {details.get('No. of Case', 'N/A')}\n"
                f"• Per Case: {details.get('Per Case', 'N/A')}"
            )

    def show_product_details(self):
        """Show detailed product information"""
        selected = self.product_table.selection()
        if not selected:
            messagebox.showwarning("⚠️ Selection Required", "Please select a product first!")
            return
        
        product_code = self.product_table.item(selected[0], "values")[0]
        product_name = self.product_table.item(selected[0], "values")[1]
        
        if product_code in self.product_data:
            details = self.product_data[product_code]
            
            # Create detailed view window
            detail_window = tk.Toplevel(self.root)
            detail_window.title(f"📦 Product Details - {product_name}")
            detail_window.geometry("500x400")
            detail_window.configure(bg=self.colors['light_bg'])
            detail_window.transient(self.root)
            detail_window.grab_set()
            
            # Center the window
            detail_window.update_idletasks()
            x = (detail_window.winfo_screenwidth() // 2) - (500 // 2)
            y = (detail_window.winfo_screenheight() // 2) - (400 // 2)
            detail_window.geometry(f"500x400+{x}+{y}")
            
            # Header
            header_frame = tk.Frame(detail_window, bg=self.colors['primary'], height=60)
            header_frame.pack(fill=tk.X)
            header_frame.pack_propagate(False)
            
            title_label = tk.Label(
                header_frame,
                text=f"📦 {product_name}",
                font=("Segoe UI", 16, "bold"),
                bg=self.colors['primary'],
                fg=self.colors['text_light'],
                pady=15
            )
            title_label.pack()
            
            # Content
            content_frame = tk.Frame(detail_window, bg=self.colors['light_bg'], padx=20, pady=20)
            content_frame.pack(fill=tk.BOTH, expand=True)
            
            # Product details
            details_text = f"""
    🔢 Product Code: {product_code}
    📝 Product Name: {product_name}
    📦 No. of Case: {details.get('No. of Case', 'N/A')}
    🔢 Per Case: {details.get('Per Case', 'N/A')}
    📏 Unit Type: {details.get('Unit Type', 'N/A')}
    💰 Selling Price: ₹{details.get('Selling Price', 'N/A')}
    📊 Per: {details.get('Per', 'N/A')} {details.get('Unit Type', '')}
    📈 Quantity: {details.get('Quantity', 'N/A')}
    🎯 Discount: {details.get('Discount', 'N/A')}%
            """
            
            details_label = tk.Label(
                content_frame,
                text=details_text.strip(),
                font=("Segoe UI", 12),
                bg=self.colors['light_bg'],
                fg=self.colors['text_dark'],
                justify=tk.LEFT
            )
            details_label.pack(anchor="w", pady=10)
            
            # Close button
            close_btn = self.create_modern_button(
                content_frame,
                "✅ Close",
                detail_window.destroy,
                style="success",
                width=15,
                height=2
            )
            close_btn.pack(pady=20)

    def export_product_list(self):
        """Export product list to CSV"""
        try:
            from datetime import datetime
            import csv
            
            filename = f"product_list_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            file_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                initialfile=filename
            )
            
            if file_path:
                with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(["Product Code", "Product Name", "No. of Case", "Per Case", "Unit Type", "Selling Price", "Per", "Quantity", "Discount"])
                    for product_code, details in self.product_data.items():
                        writer.writerow([
                            product_code,
                            details.get("Product Name", ""),
                            details.get("No. of Case", ""),
                            details.get("Per Case", ""),
                            details.get("Unit Type", ""),
                            details.get("Selling Price", ""),
                            details.get("Per", ""),
                            details.get("Quantity", ""),
                            details.get("Discount", "")
                        ])
                
                messagebox.showinfo("✅ Export Successful", f"Product list exported to:\n{file_path}")
                self.show_status_message("📤 Product list exported successfully")
                
        except Exception as e:
            messagebox.showerror("❌ Export Failed", f"Failed to export product list:\n{str(e)}")

    def show_stock_report(self):
        """Enhanced Stock Report - Product Delivery Counts with Right-Click Details"""
        self.clear_screen()
        self.current_screen = "stock_report"
        self.create_navigation_bar()
        self.create_status_bar()
        
        # Main content area
        main_container = tk.Frame(self.root, bg=self.colors['light_bg'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Header frame
        header_frame = self.create_modern_frame(main_container, "📊 PRODUCT DELIVERY SUMMARY")
        header_frame.pack(fill=tk.X, pady=(0, 20))

        # Search section
        search_frame = tk.Frame(header_frame, bg=self.colors['card_bg'])
        search_frame.pack(fill=tk.X, padx=20, pady=15)

        # Search entry
        tk.Label(search_frame, text="🔍 Search Product:", font=("Segoe UI", 10, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).pack(side=tk.LEFT, padx=5)
        
        self.stock_search_entry = self.create_modern_entry(search_frame, width=30)
        self.stock_search_entry.pack(side=tk.LEFT, padx=10, pady=5)
        self.stock_search_entry.bind('<KeyRelease>', self.filter_stock_report)

        # Refresh button
        refresh_btn = self.create_modern_button(
            search_frame, "🔄 Refresh", self.refresh_stock_report,
            style="info", width=12, height=1
        )
        refresh_btn.pack(side=tk.RIGHT, padx=5)

        # Summary Cards
        summary_frame = tk.Frame(main_container, bg=self.colors['light_bg'])
        summary_frame.pack(fill=tk.X, pady=(0, 15))

        # Create summary cards
        self.summary_cards = {}
        summary_data = [
            ("📦 Total Products", "total_products", self.colors['info']),
            ("📋 Total Deliveries", "total_deliveries", self.colors['success']),
            ("📊 Total Cases", "total_cases", self.colors['warning']),
            ("🔢 Total Quantity", "total_quantity", self.colors['accent'])
        ]

        for title, key, color in summary_data:
            card = self.create_summary_card(summary_frame, title, "0", color)
            card.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=5, pady=5)
            self.summary_cards[key] = card

        # Table container
        table_container = self.create_modern_frame(main_container, "📋 PRODUCT DELIVERY COUNTS")
        table_container.pack(fill=tk.BOTH, expand=True, pady=10)

        # Create a frame for table and scrollbars
        table_main = tk.Frame(table_container, bg=self.colors['card_bg'])
        table_main.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        # Create Treeview with modern styling
        style = ttk.Style()
        style.theme_use("default")

        # Configure styles
        style.configure("Modern.Treeview", 
            font=("Segoe UI", 9),
            rowheight=25,
            background=self.colors['card_bg'],
            fieldbackground=self.colors['card_bg'],
            foreground=self.colors['text_dark']
        )

        style.configure("Modern.Treeview.Heading", 
            font=("Segoe UI", 10, "bold"),
            background=self.colors['primary'],
            foreground=self.colors['text_light'],
            relief="flat"
        )

        # Create Treeview - Only product delivery counts
        self.stock_table = ttk.Treeview(
            table_main, 
            columns=("Product Name", "No. of Case", "Per Case", "Quantity", "Unit Type", "Delivery Count"), 
            show="headings",
            style="Modern.Treeview",
            selectmode="extended",
            height=15
        )

        # Define column headings
        columns = {
            "Product Name": {"width": 250, "anchor": "w"},
            "No. of Case": {"width": 100, "anchor": "center"},
            "Per Case": {"width": 100, "anchor": "center"},
            "Quantity": {"width": 120, "anchor": "center"},
            "Unit Type": {"width": 100, "anchor": "center"},
            "Delivery Count": {"width": 120, "anchor": "center"}
        }

        for col, settings in columns.items():
            self.stock_table.heading(col, text=col)
            self.stock_table.column(col, width=settings["width"], anchor=settings["anchor"])

        # Add scrollbars
        v_scrollbar = ttk.Scrollbar(table_main, orient="vertical", command=self.stock_table.yview)
        h_scrollbar = ttk.Scrollbar(table_main, orient="horizontal", command=self.stock_table.xview)
        self.stock_table.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        # Grid layout for table and scrollbars
        self.stock_table.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")

        # Configure grid weights
        table_main.grid_rowconfigure(0, weight=1)
        table_main.grid_columnconfigure(0, weight=1)

        # Status bar
        status_frame = tk.Frame(main_container, bg=self.colors['light_bg'])
        status_frame.pack(fill=tk.X, pady=10)

        self.stock_table_status = tk.Label(
            status_frame,
            text="📊 Loading product delivery counts...",
            font=("Segoe UI", 10),
            bg=self.colors['light_bg'],
            fg=self.colors['text_dark']
        )
        self.stock_table_status.pack(side=tk.LEFT, padx=10)

        # Action buttons
        action_buttons_frame = tk.Frame(main_container, bg=self.colors['light_bg'])
        action_buttons_frame.pack(fill=tk.X, pady=10)

        button_container = tk.Frame(action_buttons_frame, bg=self.colors['light_bg'])
        button_container.pack(expand=True)

        # Back button
        back_btn = self.create_modern_button(
            button_container,
            "← Back to Products",
            self.show_product_management,
            style="secondary",
            width=18,
            height=2
        )
        back_btn.pack(side=tk.LEFT, padx=8)

        # Export button
        export_btn = self.create_modern_button(
            button_container,
            "📤 Export to CSV",
            self.export_stock_report,
            style="success",
            width=15,
            height=2
        )
        export_btn.pack(side=tk.LEFT, padx=8)

        # 🔥 NEW: View Delivery Details button
        details_btn = self.create_modern_button(
            button_container,
            "📋 View Delivery Details",
            self.show_delivery_details,
            style="info",
            width=18,
            height=2
        )
        details_btn.pack(side=tk.LEFT, padx=8)

        # Load the data
        self.populate_stock_report()

        # 🔥 NEW: Bind right-click context menu to stock table
        self.stock_table.bind('<Button-3>', self.show_stock_context_menu)

        # Set focus to search field
        self.stock_search_entry.focus_set()

        self.show_status_message("📊 Product delivery counts loaded - Right-click any row to view delivery details")


    def show_stock_context_menu(self, event):
        """Show right-click context menu for stock table with delivery details option"""
        item = self.stock_table.identify_row(event.y)
        if item:
            self.stock_table.selection_set(item)
            
            context_menu = tk.Menu(self.root, tearoff=0, font=("Segoe UI", 10))
            context_menu.add_command(
                label="📋 View Delivery Details", 
                command=self.show_delivery_details
            )
            context_menu.add_separator()
            context_menu.add_command(
                label="📊 Refresh Data", 
                command=self.refresh_stock_report
            )
            
            try:
                context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                context_menu.grab_release()

    def show_delivery_details(self):
        """Show detailed delivery records for the selected product"""
        selected = self.stock_table.selection()
        if not selected:
            messagebox.showwarning("⚠️ Selection Required", "Please select a product first!")
            return
        
        # Get the selected product name
        product_name = self.stock_table.item(selected[0], "values")[0]
        delivery_count = self.stock_table.item(selected[0], "values")[5]
        
        # Get all delivery records for this product
        delivery_records = self.get_product_delivery_records(product_name)
        
        if not delivery_records:
            messagebox.showinfo("ℹ️ No Records", f"No delivery records found for '{product_name}'")
            return
        
        # Create detailed view window
        detail_window = tk.Toplevel(self.root)
        detail_window.title(f"📋 Delivery Details - {product_name}")
        detail_window.geometry("900x500")
        detail_window.configure(bg=self.colors['light_bg'])
        detail_window.transient(self.root)
        detail_window.grab_set()
        
        # Center the window
        detail_window.update_idletasks()
        x = (detail_window.winfo_screenwidth() // 2) - (900 // 2)
        y = (detail_window.winfo_screenheight() // 2) - (500 // 2)
        detail_window.geometry(f"900x500+{x}+{y}")
        
        # Header
        header_frame = tk.Frame(detail_window, bg=self.colors['primary'], height=70)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(
            header_frame,
            text=f"📋 Delivery Details - {product_name}",
            font=("Segoe UI", 16, "bold"),
            bg=self.colors['primary'],
            fg=self.colors['text_light'],
            pady=20
        )
        title_label.pack()
        
        subtitle_label = tk.Label(
            header_frame,
            text=f"Total Deliveries: {delivery_count} | Records Found: {len(delivery_records)}",
            font=("Segoe UI", 11),
            bg=self.colors['primary'],
            fg=self.colors['text_light']
        )
        subtitle_label.pack()
        
        # Content frame
        content_frame = tk.Frame(detail_window, bg=self.colors['light_bg'])
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Create delivery details table
        table_frame = tk.Frame(content_frame, bg=self.colors['card_bg'])
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create Treeview for delivery details
        delivery_table = ttk.Treeview(
            table_frame, 
            columns=("Bill No", "Bill Date", "Customer Name", "Agent Name", "Quantity", "No. of Cases"), 
            show="headings",
            height=15
        )
        
        # Define column headings
        columns = {
            "Bill No": {"width": 100, "anchor": "center"},
            "Bill Date": {"width": 100, "anchor": "center"},
            "Customer Name": {"width": 200, "anchor": "w"},
            "Agent Name": {"width": 150, "anchor": "w"},
            "Quantity": {"width": 100, "anchor": "center"},
            "No. of Cases": {"width": 100, "anchor": "center"}
        }
        
        for col, settings in columns.items():
            delivery_table.heading(col, text=col)
            delivery_table.column(col, width=settings["width"], anchor=settings["anchor"])
        
        # Add scrollbars
        v_scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=delivery_table.yview)
        h_scrollbar = ttk.Scrollbar(table_frame, orient="horizontal", command=delivery_table.xview)
        delivery_table.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Grid layout
        delivery_table.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        
        # Populate delivery details table
        for record in delivery_records:
            delivery_table.insert("", "end", values=(
                record['bill_no'],
                record['bill_date'],
                record['customer_name'],
                record['agent_name'],
                record['quantity'],
                record['no_of_cases']
            ))
        
        # Configure row colors
        delivery_table.tag_configure('evenrow', background='#ffffff')
        delivery_table.tag_configure('oddrow', background='#f0f8ff')
        
        # Apply alternating row colors
        for i, item in enumerate(delivery_table.get_children()):
            tags = ('evenrow',) if i % 2 == 0 else ('oddrow',)
            delivery_table.item(item, tags=tags)
        
        # Action buttons frame
        action_frame = tk.Frame(content_frame, bg=self.colors['light_bg'])
        action_frame.pack(fill=tk.X, pady=10)
        
        # Close button
        close_btn = self.create_modern_button(
            action_frame,
            "✅ Close",
            detail_window.destroy,
            style="success",
            width=15,
            height=1
        )
        close_btn.pack(side=tk.RIGHT, padx=10)
        
        # Export button for delivery details
        export_delivery_btn = self.create_modern_button(
            action_frame,
            "📤 Export Delivery Data",
            lambda: self.export_delivery_details(delivery_records, product_name),
            style="info",
            width=18,
            height=1
        )
        export_delivery_btn.pack(side=tk.RIGHT, padx=10)

    def get_product_delivery_records(self, product_name):
        """Get all delivery records for a specific product"""
        delivery_records = []
        
        try:
            # Load bills data
            self.bills_ref = db.reference('bills')
            
            if not self.bills_data:
                return delivery_records
            
            # Search through all bills for this product
            for bill_no, bill_info in self.bills_data.items():
                items = bill_info.get('items', [])
                
                for item in items:
                    if len(item) >= 7:  # Ensure item has required data
                        current_product_name = item[1]  # Product Name is at index 1
                        
                        if current_product_name == product_name:
                            # Extract quantity from the quantity field (e.g., "100 Box" -> 100)
                            quantity_str = item[4]  # Quantity is at index 4
                            try:
                                quantity = int(quantity_str.split()[0]) if quantity_str else 0
                            except (ValueError, IndexError):
                                quantity = 0
                            
                            # Extract number of cases
                            no_of_cases = item[2] if len(item) > 2 else 0  # No. of Case is at index 2
                            
                            # Create delivery record
                            record = {
                                'bill_no': bill_no,
                                'bill_date': bill_info.get('bill_date', 'N/A'),
                                'customer_name': bill_info.get('customer_name', 'N/A'),
                                'agent_name': bill_info.get('agent_name', 'N/A'),
                                'quantity': quantity,
                                'no_of_cases': no_of_cases
                            }
                            delivery_records.append(record)
            
            # Sort by bill date (most recent first)
            
            delivery_records.sort(key=lambda x: parse_date_flexible(x['bill_date']) or datetime.min, reverse=True)
            
        except Exception as e:
            print(f"Error getting delivery records: {e}")
        
        return delivery_records

    def export_delivery_details(self, delivery_records, product_name):
        """Export delivery details to CSV"""
        try:
            import csv
            from datetime import datetime
            
            # Create filename
            clean_product_name = re.sub(r'[^\w\s-]', '', product_name).strip().replace(' ', '_')
            filename = f"delivery_details_{clean_product_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            
            file_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                initialfile=filename
            )
            
            if file_path:
                with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(["Product Name", "Bill No", "Bill Date", "Customer Name", "Agent Name", "Quantity", "No. of Cases"])
                    
                    for record in delivery_records:
                        writer.writerow([
                            product_name,
                            record['bill_no'],
                            record['bill_date'],
                            record['customer_name'],
                            record['agent_name'],
                            record['quantity'],
                            record['no_of_cases']
                        ])
                
                messagebox.showinfo("✅ Export Successful", f"Delivery details exported to:\n{file_path}")
                self.show_status_message("📤 Delivery details exported successfully")
                
        except Exception as e:
            messagebox.showerror("❌ Export Failed", f"Failed to export delivery details:\n{str(e)}")

    def create_summary_card(self, parent, title, value, color):
        """Create a summary card for stock report"""
        card = tk.Frame(
            parent,
            bg=self.colors['card_bg'],
            relief="raised",
            bd=1,
            width=200,
            height=80
        )
        card.pack_propagate(False)

        # Card content
        content_frame = tk.Frame(card, bg=self.colors['card_bg'])
        content_frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

        # Title
        title_label = tk.Label(
            content_frame,
            text=title,
            font=("Segoe UI", 10, "bold"),
            bg=self.colors['card_bg'],
            fg=self.colors['text_muted']
        )
        title_label.pack(anchor="w")

        # Value
        value_label = tk.Label(
            content_frame,
            text=str(value),
            font=("Segoe UI", 16, "bold"),
            bg=self.colors['card_bg'],
            fg=color
        )
        value_label.pack(expand=True)

        # Store reference to update value later
        card.value_label = value_label

        return card

    def populate_stock_report(self, filtered_data=None):
        """Populate stock report with product delivery counts"""
        try:
            # Clear existing data
            for item in self.stock_table.get_children():
                self.stock_table.delete(item)

            # Load bills data
            self.bills_ref = db.reference('bills')

            if not self.bills_data:
                self.stock_table_status.config(text="❌ No bills data found!")
                return

            # Calculate product delivery counts
            product_stats = self.calculate_product_delivery_counts()

            # Use filtered data or all data
            display_data = filtered_data if filtered_data is not None else product_stats

            # Insert data into table
            total_products = 0
            total_deliveries = 0
            total_cases = 0
            total_quantity = 0

            for index, (product_name, stats) in enumerate(display_data.items()):
                tags = ('evenrow',) if index % 2 == 0 else ('oddrow',)
                
                self.stock_table.insert("", "end", values=(
                    product_name,
                    stats['total_cases'],
                    stats['per_case'],
                    stats['total_quantity'],
                    stats['unit_type'],
                    stats['delivery_count']
                ), tags=tags)

                # Update totals
                total_products += 1
                total_deliveries += stats['delivery_count']
                total_cases += stats['total_cases']
                total_quantity += stats['total_quantity']

            # Configure row colors
            self.stock_table.tag_configure('evenrow', background='#ffffff')
            self.stock_table.tag_configure('oddrow', background='#f0f8ff')

            # Update summary cards
            self.update_summary_cards(total_products, total_deliveries, total_cases, total_quantity)

            # Update status
            self.stock_table_status.config(text=f"📊 Displaying {total_products} products • {total_deliveries} deliveries • {total_quantity} total quantity")

        except Exception as e:
            self.stock_table_status.config(text=f"❌ Error loading stock report: {str(e)}")

    def calculate_product_delivery_counts(self):
        """Calculate delivery counts for all products"""
        product_stats = {}

        for bill_no, bill_info in self.bills_data.items():
            items = bill_info.get('items', [])
            
            for item in items:
                if len(item) >= 7:  # Ensure item has required data
                    product_name = item[1]  # Product Name
                    no_of_case = item[2]    # No. of Case
                    per_case = item[3]      # Per Case
                    quantity = item[4]      # Quantity
                    unit_type = item[6]     # Unit Type

                    # Initialize product in stats if not exists
                    if product_name not in product_stats:
                        product_stats[product_name] = {
                            'total_cases': 0,
                            'per_case': per_case,
                            'total_quantity': 0,
                            'unit_type': unit_type,
                            'delivery_count': 0
                        }

                    # Update statistics
                    try:
                        product_stats[product_name]['total_cases'] += int(no_of_case) if no_of_case else 0
                        product_stats[product_name]['total_quantity'] += int(quantity.split()[0]) if quantity else 0
                        product_stats[product_name]['delivery_count'] += 1
                    except (ValueError, IndexError):
                        continue

        return product_stats

    def update_summary_cards(self, total_products, total_deliveries, total_cases, total_quantity):
        """Update the summary cards with current statistics"""
        if hasattr(self, 'summary_cards'):
            # Update total products
            if 'total_products' in self.summary_cards:
                self.summary_cards['total_products'].value_label.config(text=f"{total_products:,}")
            
            # Update total deliveries
            if 'total_deliveries' in self.summary_cards:
                self.summary_cards['total_deliveries'].value_label.config(text=f"{total_deliveries:,}")
            
            # Update total cases
            if 'total_cases' in self.summary_cards:
                self.summary_cards['total_cases'].value_label.config(text=f"{total_cases:,}")
            
            # Update total quantity
            if 'total_quantity' in self.summary_cards:
                self.summary_cards['total_quantity'].value_label.config(text=f"{total_quantity:,}")

    def filter_stock_report(self, event=None):
        """Filter stock report based on search term"""
        search_term = self.stock_search_entry.get().lower()
        
        if not search_term:
            self.populate_stock_report()
            return
        
        # Calculate fresh statistics
        product_stats = self.calculate_product_delivery_counts()
        
        filtered_data = {}
        for product_name, stats in product_stats.items():
            if search_term in product_name.lower():
                filtered_data[product_name] = stats
        
        self.populate_stock_report(filtered_data)

    def refresh_stock_report(self, event=None):
        """Refresh the stock report"""
        self.bills_ref = db.reference('bills')
        
        if hasattr(self, 'stock_search_entry'):
            self.stock_search_entry.delete(0, tk.END)
        
        self.populate_stock_report()
        self.show_status_message("🔄 Stock report refreshed")
        return "break"

    def export_stock_report(self):
        """Export stock report to CSV"""
        try:
            import csv
            from datetime import datetime
            
            # Create directory for exports
            documents_dir = os.path.join(os.path.expanduser("~"), "Documents")
            exports_dir = os.path.join(documents_dir, "InvoiceApp", "Stock_Exports")
            os.makedirs(exports_dir, exist_ok=True)
            
            filename = f"product_delivery_counts_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            file_path = os.path.join(exports_dir, filename)
            
            # Calculate delivery counts
            product_stats = self.calculate_product_delivery_counts()
            
            with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(["Product Name", "No. of Case", "Per Case", "Quantity", "Unit Type", "Delivery Count"])
                
                for product_name, stats in product_stats.items():
                    writer.writerow([
                        product_name,
                        stats['total_cases'],
                        stats['per_case'],
                        stats['total_quantity'],
                        stats['unit_type'],
                        stats['delivery_count']
                    ])
            
            messagebox.showinfo("✅ Export Successful", f"Product delivery counts exported to:\n{file_path}")
            self.show_status_message("📤 Product delivery counts exported successfully")
            
        except Exception as e:
            messagebox.showerror("❌ Export Failed", f"Failed to export delivery counts:\n{str(e)}")
    def show_billing_dashboard(self):
        """Show billing dashboard"""
        self.clear_screen()
        self.current_screen = "billing"
        self.create_navigation_bar()
        self.create_status_bar()

        content_frame = tk.Frame(self.root, bg=self.colors['light_bg'])
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Header
        header_frame = self.create_modern_frame(content_frame, "🧾 BILLING CENTER")
        header_frame.pack(fill=tk.X, pady=(0, 20))

        # Quick actions
        actions_frame = tk.Frame(header_frame, bg=self.colors['card_bg'])
        actions_frame.pack(padx=20, pady=20)

        actions = [
            ("🆕 Create New Bill", self.create_new_bill, "CTRL+N - Create new invoice"),
            ("✏️ Edit Existing Bill", self.show_edit_bill, "Edit previous invoices"),
            ("👀 View All Bills", self.show_view_bill, "Browse all invoices"),
            ("📄 Generate Statements", self.show_statement_options, "Create customer statements")  # UPDATED THIS LINE
        ]

        for i, (text, command, tooltip) in enumerate(actions):
            btn = self.create_modern_button(
                actions_frame, text, command,
                style="warning", width=22, height=2
            )
            btn.grid(row=i//2, column=i%2, padx=10, pady=10)

        # ... rest of the method remains the same

        # Recent bills preview
        recent_frame = self.create_modern_frame(content_frame, "📋 RECENT INVOICES")
        recent_frame.pack(fill=tk.BOTH, expand=True, pady=20)

        # Add a simple table or list here for recent bills
        info_label = tk.Label(
            recent_frame,
            text="Recent invoices will be displayed here...",
            font=("Segoe UI", 11),
            bg=self.colors['card_bg'],
            fg=self.colors['text_muted'],
            pady=50
        )
        info_label.pack()

        # Set up keyboard navigation
        self.focusable_widgets.clear()
        for widget in actions_frame.winfo_children():
            if isinstance(widget, tk.Button):
                self.create_focusable_widget(widget, "button")

        if self.focusable_widgets:
            self.focusable_widgets[0].focus_set()

        self.show_status_message("🧾 Billing Center - Press F4 to return here anytime")
        
    def create_new_bill(self, bill_no=None):
        """Modern enhanced version of create_new_bill with full-screen layout (no scrolling)"""
        self.clear_screen()
        self.current_screen = "billing_entry"
        self.create_navigation_bar()
        self.create_status_bar()
        
        # Reset focusable widgets
        self.focusable_widgets.clear()

        # Main content area - FULL SCREEN
        main_container = tk.Frame(self.root, bg=self.colors['light_bg'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)  # Reduced padding

        # Header frame
        header_frame = self.create_modern_frame(main_container, "🧾 CREATE NEW INVOICE")
        header_frame.pack(fill=tk.X, pady=(0, 10))  # Reduced padding

        # Create the billing GUI with COMPACT layout
        self.create_gui_compact(main_container)
        
        # Reset the GUI
        self.reset_gui()

        # If a bill number is provided (for editing), load the saved bill details COMPLETELY
        if bill_no:
            # Set the flag to indicate the bill is being edited
            self.bill_no_edited = True
            
            # Load the bill details from the JSON file
            bill_details = self.bills_data.get(bill_no, {})
            if bill_details:
                print(f"DEBUG: Loading bill {bill_no} for editing")
                
                # Set the bill details in the GUI fields
                self.bill_no = bill_no
                self.bill_no_entry.delete(0, tk.END)
                self.bill_no_entry.insert(0, self.bill_no)

                # Set ALL bill details COMPLETELY
                self.bill_date.set(bill_details.get("bill_date", ""))
                self.customer_combobox.set(bill_details.get("customer_name", ""))
                self.to_address.set(bill_details.get("address", ""))
                self.agent_name.set(bill_details.get("agent_name", ""))
                self.to_gstin.set(bill_details.get("gstin", ""))
                self.lr_number.set(bill_details.get("lr_number", ""))
                self.from_.set(bill_details.get("from_", ""))
                self.to_.set(bill_details.get("to_", ""))
                self.document_through.set(bill_details.get("document_through", ""))
                self.region.set(bill_details.get("region", "South"))
                self.gst_percentage.set(float(bill_details.get("gst_percentage", 18.0)))
                self.packing_charge.set(float(bill_details.get("packing_charge", 0.0)))
                
                # Load ALL GST values
                self.cgst.set(float(bill_details.get("cgst_amount", 0.0)))
                self.sgst.set(float(bill_details.get("sgst_amount", 0.0)))
                self.igst.set(float(bill_details.get("igst_amount", 0.0)))
                
                # Update GST fields visibility based on region
                self.update_gst_fields()

                # Clear existing items in table
                for item in self.table.get_children():
                    self.table.delete(item)

                # Load ALL product items into the product table with COMPLETE data
                items = bill_details.get("items", [])
                print(f"DEBUG: Loading {len(items)} items")
                for item in items:
                    if len(item) >= 10:  # Ensure item has all columns
                        self.table.insert("", "end", values=item)

                # Load ALL calculated amounts
                self.before_discount_amount_field.set(float(bill_details.get("goods_value", 0.0)))
                self.discount_amount_field.set(float(bill_details.get("special_discount", 0.0)))
                self.After_discount_amount_field.set(float(bill_details.get("sub_total", 0.0)))
                self.Packing_Amount.set(float(bill_details.get("packing_charges", 0.0)))
                self.total_amount.set(float(bill_details.get("net_amount", 0.0)))

                print(f"DEBUG: Bill data loaded completely")
                
                # Force update of customer details
                self.fill_customer_details()

        # Action buttons frame at bottom
        action_frame = tk.Frame(main_container, bg=self.colors['light_bg'])
        action_frame.pack(fill=tk.X, pady=5)  # Reduced padding

        # Center the buttons
        button_container = tk.Frame(action_frame, bg=self.colors['light_bg'])
        button_container.pack(expand=True)

        # Back button
        back_btn = self.create_modern_button(
            button_container,
            "← Back to Billing (F4)",
            self.show_billing_dashboard,
            style="secondary",
            width=16,
            height=1  # Reduced height
        )
        back_btn.pack(side=tk.LEFT, padx=5)

        # Reset button
        reset_btn = self.create_modern_button(
            button_container,
            "🔄 Reset Form",
            self.reset_gui,
            style="warning",
            width=12,
            height=1  # Reduced height
        )
        reset_btn.pack(side=tk.LEFT, padx=5)

        # Save Bill button
        save_btn = self.create_modern_button(
            button_container,
            "💾 Save Bill (Ctrl+S)",
            self.generate_pdf,
            style="success",
            width=14,
            height=1  # Reduced height
        )
        save_btn.pack(side=tk.LEFT, padx=5)

        # View Bills button
        view_btn = self.create_modern_button(
            button_container,
            "📋 View Bills",
            self.show_view_bill,
            style="info",
            width=12,
            height=1  # Reduced height
        )
        view_btn.pack(side=tk.LEFT, padx=5)

        # Bind keyboard shortcuts
        self.root.bind('<Control-s>', lambda e: self.generate_pdf())
        self.root.bind('<F4>', lambda e: self.show_billing_dashboard())

        self.show_status_message("🧾 Creating new invoice - Fill in all details and press Ctrl+S to save")

    def create_gui_compact(self, parent):
        """Create compact billing GUI that matches the user-friendly layout"""
        # Ensure customer_names and product_names are initialized
        if not hasattr(self, 'customer_names') or not self.customer_names:
            self.customer_names = []
            for key, pdata in self.party_data.items():
                name = (
                    pdata.get('Customer Name') or 
                    pdata.get('Customer_Name') or 
                    pdata.get('customer_name') or
                    pdata.get('Customer  Name') or
                    pdata.get('Customer name')
                )
                if name:
                    self.customer_names.append(name)
                else:
                    print(f"⚠️ Missing Customer Name in party_data entry: {key}")

        
        if not hasattr(self, 'product_names') or not self.product_names:
            self.product_names = [self.product_data[key]['Product Name'] for key in self.product_data.keys()]

        # Header - Centered "Estimate"
        header_label = tk.Label(parent, text="Tax Invoice", font=("Arial", 16, "bold"), 
                            bg=self.colors['light_bg'], fg=self.colors['primary'])
        header_label.pack(fill=tk.X, pady=5)

        # Top Section - Customer and Bill Details
        top_frame = tk.Frame(parent, bg=self.colors['light_bg'])
        top_frame.pack(fill=tk.X, pady=5, padx=10)

        # Left side - Receiver details
        left_frame = tk.Frame(top_frame, bg=self.colors['light_bg'])
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10)

        # "To" Label
        tk.Label(left_frame, text="To:", font=("Arial", 12, "bold"), 
                bg=self.colors['light_bg']).grid(row=0, column=0, sticky=tk.W, pady=2)

        # Account Name (with Auto-Suggestions) - FIXED: Use StringVar for manual editing
        tk.Label(left_frame, text="Account Name", font=("Arial", 10),
                bg=self.colors['light_bg']).grid(row=1, column=0, sticky=tk.W, pady=2)

        # Use StringVar for customer name to allow manual editing
        self.to_name = tk.StringVar()  # ADD THIS LINE
        self.customer_combobox = ttk.Combobox(left_frame, textvariable=self.to_name, width=25, font=("Arial", 10))  # CHANGED to to_name
        self.customer_combobox.grid(row=1, column=1, padx=5, pady=2, sticky=tk.W)
        self.customer_combobox['values'] = self.customer_names

        # Bind events for customer combobox - REMOVE FocusOut binding
        self.customer_combobox.bind("<KeyRelease>", self.update_customer_suggestions)
        self.customer_combobox.bind("<<ComboboxSelected>>", self.fill_customer_details)

        # Address Field
        tk.Label(left_frame, text="Address", font=("Arial", 10),
                bg=self.colors['light_bg']).grid(row=1, column=2, sticky=tk.W, padx=10, pady=2)
        
        self.to_address = tk.StringVar()
        address_entry = tk.Entry(left_frame, textvariable=self.to_address, width=40, font=("Arial", 10))
        address_entry.grid(row=1, column=3, padx=5, pady=2, sticky=tk.W)

        # GSTIN Field
        tk.Label(left_frame, text="GSTIN", font=("Arial", 10),
                bg=self.colors['light_bg']).grid(row=2, column=0, sticky=tk.W, pady=2)
        
        self.to_gstin = tk.StringVar()
        gstin_entry = tk.Entry(left_frame, textvariable=self.to_gstin, width=25, font=("Arial", 10))
        gstin_entry.grid(row=2, column=1, padx=5, pady=2, sticky=tk.W)

        # Agent Name Field
        tk.Label(left_frame, text="Agent Name", font=("Arial", 10),
                bg=self.colors['light_bg']).grid(row=2, column=2, sticky=tk.W, padx=10, pady=2)
        
        self.agent_name = tk.StringVar()
        agent_entry = tk.Entry(left_frame, textvariable=self.agent_name, width=25, font=("Arial", 10))
        agent_entry.grid(row=2, column=3, padx=5, pady=2, sticky=tk.W)

        # Right side - Bill details
        right_frame = tk.Frame(top_frame, bg=self.colors['light_bg'])
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=10)

        # Bill Number
        tk.Label(right_frame, text="Bill No:", font=("Arial", 10),
                bg=self.colors['light_bg']).grid(row=0, column=0, sticky=tk.W, pady=2)
        
        self.bill_no_entry = tk.Entry(right_frame, font=("Arial", 10), width=20)
        self.bill_no_entry.insert(0, self.bill_no)
        self.bill_no_entry.grid(row=0, column=1, sticky=tk.W, pady=2)
        self.bill_no_entry.bind("<KeyRelease>", self.on_bill_no_edit)

        # Bill Date
        tk.Label(right_frame, text="Bill Date:", font=("Arial", 10),
                bg=self.colors['light_bg']).grid(row=1, column=0, sticky=tk.W, pady=2)
        
        self.bill_date = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        date_entry = tk.Entry(right_frame, textvariable=self.bill_date, width=20, font=("Arial", 10))
        date_entry.grid(row=1, column=1, pady=2)

        # LR Number
        tk.Label(right_frame, text="L.R. Number", font=("Arial", 10),
                bg=self.colors['light_bg']).grid(row=2, column=0, sticky=tk.W, pady=2)
        
        self.lr_number = tk.StringVar()
        lr_entry = tk.Entry(right_frame, textvariable=self.lr_number, width=20, font=("Arial", 10))
        lr_entry.grid(row=2, column=1, pady=2)

        # From Field
        tk.Label(right_frame, text="From", font=("Arial", 10),
                bg=self.colors['light_bg']).grid(row=3, column=0, sticky=tk.W, pady=2)

        self.from_ = tk.StringVar(value="ONDIPULINAIKANOOR")
        from_entry = tk.Entry(right_frame, textvariable=self.from_, width=20, font=("Arial", 10))
        from_entry.grid(row=3, column=1, pady=2)

        # To Field
        tk.Label(right_frame, text="To", font=("Arial", 10),
                bg=self.colors['light_bg']).grid(row=4, column=0, sticky=tk.W, pady=2)
        
        self.to_ = tk.StringVar()
        to_entry = tk.Entry(right_frame, textvariable=self.to_, width=20, font=("Arial", 10))
        to_entry.grid(row=4, column=1, pady=2)

        # Document Through
        tk.Label(right_frame, text="Document Through", font=("Arial", 10),
                bg=self.colors['light_bg']).grid(row=5, column=0, sticky=tk.W, pady=2)
        
        self.document_through = tk.StringVar()
        doc_entry = tk.Entry(right_frame, textvariable=self.document_through, width=20, font=("Arial", 10))
        doc_entry.grid(row=5, column=1, pady=2)

        # Region and Tax Section
        region_frame = tk.Frame(parent, bg=self.colors['light_bg'])
        region_frame.pack(fill=tk.X, pady=10, padx=10)

        # Region Selection
        tk.Label(region_frame, text="Region:", font=("Arial", 10),
                bg=self.colors['light_bg']).grid(row=0, column=0, sticky=tk.W, pady=5)
        
        tk.Radiobutton(region_frame, text="South", variable=self.region, value="South", 
                    command=self.update_gst_fields, font=("Arial", 10),
                    bg=self.colors['light_bg']).grid(row=0, column=1, sticky=tk.W, padx=5)
        
        tk.Radiobutton(region_frame, text="North", variable=self.region, value="North", 
                    command=self.update_gst_fields, font=("Arial", 10),
                    bg=self.colors['light_bg']).grid(row=0, column=2, sticky=tk.W, padx=5)

        # GST Percentage
        tk.Label(region_frame, text="GST Percentage:", font=("Arial", 10),
                bg=self.colors['light_bg']).grid(row=1, column=0, sticky=tk.W, pady=5)
        
        self.gst_percentage = tk.DoubleVar(value=18.0)
        gst_entry = tk.Entry(region_frame, textvariable=self.gst_percentage, width=10, font=("Arial", 10))
        gst_entry.grid(row=1, column=1, sticky=tk.W, pady=5)
        self.gst_percentage.trace_add('write', self.on_gst_percentage_change)

        # GST Fields Frame
        gst_fields_frame = tk.Frame(region_frame, bg=self.colors['light_bg'])
        gst_fields_frame.grid(row=2, column=0, columnspan=4, sticky=tk.W, pady=5)

        # CGST and SGST Fields (Visible for South) - UPDATED LABELS
        self.cgst_label = tk.Label(gst_fields_frame, text="CGST Amount:", font=("Arial", 10),
                                bg=self.colors['light_bg'])
        self.cgst_label.grid(row=0, column=0, sticky=tk.W, padx=5)

        self.cgst = tk.DoubleVar(value=0.0)
        self.cgst_entry = tk.Entry(gst_fields_frame, textvariable=self.cgst, width=15, font=("Arial", 10))
        self.cgst_entry.grid(row=0, column=1, sticky=tk.W, padx=5)
        self.cgst_entry.config(state="readonly")  # Make it read-only since it's auto-calculated

        self.sgst_label = tk.Label(gst_fields_frame, text="SGST Amount:", font=("Arial", 10),
                                bg=self.colors['light_bg'])
        self.sgst_label.grid(row=0, column=2, sticky=tk.W, padx=5)

        self.sgst = tk.DoubleVar(value=0.0)
        self.sgst_entry = tk.Entry(gst_fields_frame, textvariable=self.sgst, width=15, font=("Arial", 10))
        self.sgst_entry.grid(row=0, column=3, sticky=tk.W, padx=5)
        self.sgst_entry.config(state="readonly")  # Make it read-only since it's auto-calculated

        # IGST Field (Visible for North) - UPDATED LABEL
        self.igst_label = tk.Label(gst_fields_frame, text="IGST Amount:", font=("Arial", 10),
                                bg=self.colors['light_bg'])
        self.igst_label.grid(row=0, column=4, sticky=tk.W, padx=5)

        self.igst = tk.DoubleVar(value=0.0)
        self.igst_entry = tk.Entry(gst_fields_frame, textvariable=self.igst, width=15, font=("Arial", 10))
        self.igst_entry.grid(row=0, column=5, sticky=tk.W, padx=5)
        self.igst_entry.config(state="readonly")  # Make it read-only since it's auto-calculated

        # ✅ Add bindings for GST field changes
        self.cgst_entry.bind('<FocusOut>', lambda e: self.recalculate_gst_from_fields())
        self.sgst_entry.bind('<FocusOut>', lambda e: self.recalculate_gst_from_fields())
        self.sgst_entry.bind('<FocusOut>', lambda e: self.recalculate_gst_from_fields())
        self.igst_entry.bind('<FocusOut>', lambda e: self.recalculate_gst_from_fields())

        # Reset Button on the right
        reset_button = tk.Button(region_frame, text="Reset Frame", command=self.reset_gui,
                            font=("Arial", 10), bg=self.colors['warning'], fg="white")
        reset_button.grid(row=0, column=4, padx=20)

        # Initially update GST fields
        self.update_gst_fields()

        # Product Section
        product_frame = tk.Frame(parent, bg=self.colors['light_bg'])
        product_frame.pack(fill=tk.X, pady=10, padx=10)

    # Product Name - FIXED: Use StringVar for manual editing
        tk.Label(product_frame, text="Product Name", font=("Arial", 10),
                bg=self.colors['light_bg']).grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)

        self.product_name_var = tk.StringVar()  # ADD THIS LINE
        self.product_name_combobox = ttk.Combobox(product_frame, textvariable=self.product_name_var, width=25, font=("Arial", 10))  # CHANGED
        self.product_name_combobox.grid(row=0, column=1, padx=5, pady=5)
        self.product_name_combobox['values'] = self.product_names

        # Bind product events - REMOVE any FocusOut bindings that override manual changes
        self.product_name_combobox.bind("<KeyRelease>", self.update_product_suggestions)
        self.product_name_combobox.bind("<<ComboboxSelected>>", self.fill_product_details_modern)
        self.product_name_combobox.bind("<Return>", self.handle_manual_product_entry)

        # No. of Case
        tk.Label(product_frame, text="No. of Case", font=("Arial", 10),
                bg=self.colors['light_bg']).grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        
        self.no_of_case_entry = tk.Entry(product_frame, width=10, font=("Arial", 10))
        self.no_of_case_entry.grid(row=0, column=3, padx=5, pady=5)

        # Per Case
        tk.Label(product_frame, text="Per Case", font=("Arial", 10),
                bg=self.colors['light_bg']).grid(row=0, column=4, padx=5, pady=5, sticky=tk.W)
        
        self.per_case_entry = tk.Entry(product_frame, width=10, font=("Arial", 10))
        self.per_case_entry.grid(row=0, column=5, padx=5, pady=5)

        # Unit Type
        tk.Label(product_frame, text="Unit Type", font=("Arial", 10),
                bg=self.colors['light_bg']).grid(row=0, column=6, padx=5, pady=5, sticky=tk.W)
        
        self.unit_type_combobox = ttk.Combobox(product_frame, values=["U", "N", "Box"], width=10, font=("Arial", 10))
        self.unit_type_combobox.grid(row=0, column=7, padx=5, pady=5)

        # Rate
        tk.Label(product_frame, text="Rate", font=("Arial", 10),
                bg=self.colors['light_bg']).grid(row=0, column=8, padx=5, pady=5, sticky=tk.W)
        
        self.rate = tk.DoubleVar(value=0.0)
        self.rate_entry = tk.Entry(product_frame, textvariable=self.rate, width=10, font=("Arial", 10))
        self.rate_entry.grid(row=0, column=9, padx=5, pady=5)

        # Per
        tk.Label(product_frame, text="Per", font=("Arial", 10),
                bg=self.colors['light_bg']).grid(row=0, column=10, padx=5, pady=5, sticky=tk.W)
        
        self.per_entry = tk.Entry(product_frame, width=10, font=("Arial", 10))
        self.per_entry.grid(row=0, column=11, padx=5, pady=5)

        # Quantity
        tk.Label(product_frame, text="Quantity", font=("Arial", 10),
                bg=self.colors['light_bg']).grid(row=0, column=12, padx=5, pady=5, sticky=tk.W)
        
        self.quantity = tk.IntVar(value=0)
        self.quantity_entry = tk.Entry(product_frame, textvariable=self.quantity, width=10, font=("Arial", 10))
        self.quantity_entry.grid(row=0, column=13, padx=5, pady=5)

        # Discount
        tk.Label(product_frame, text="Discount %", font=("Arial", 10),
                bg=self.colors['light_bg']).grid(row=0, column=14, padx=5, pady=5, sticky=tk.W)
        
        self.discount = tk.DoubleVar(value=0.0)
        self.discount_entry = tk.Entry(product_frame, textvariable=self.discount, width=10, font=("Arial", 10))
        self.discount_entry.grid(row=0, column=15, padx=5, pady=5)

        # Action Buttons
        add_button = tk.Button(product_frame, text="Add Item", command=self.add_item,
                            font=("Arial", 10), bg=self.colors['success'], fg="white")
        add_button.grid(row=0, column=16, padx=5, pady=5)

        remove_button = tk.Button(product_frame, text="Remove Item", command=self.remove_item,
                                font=("Arial", 10), bg=self.colors['warning'], fg="white")
        remove_button.grid(row=0, column=17, padx=5, pady=5)

        update_button = tk.Button(product_frame, text="Update Item", command=self.update_item,
                                font=("Arial", 10), bg=self.colors['info'], fg="white")
        update_button.grid(row=0, column=18, padx=5, pady=5)

        # Products Table
        table_frame = tk.Frame(parent, bg=self.colors['light_bg'])
        table_frame.pack(fill=tk.BOTH, expand=True, pady=10, padx=10)

        self.table = ttk.Treeview(table_frame, 
                                columns=("S.No", "Product Name", "No. of Case", "Per Case", "Quantity", "Rate", "Unit Type", "Per", "Discount", "Amount"), 
                                show="headings",
                                height=8)

        # Configure columns
        columns_config = [
            ("S.No", 50, "center"),
            ("Product Name", 150, "center"),
            ("No. of Case", 100, "center"),
            ("Per Case", 100, "center"),
            ("Quantity", 100, "center"),
            ("Rate", 100, "center"),
            ("Unit Type", 100, "center"),
            ("Per", 100, "center"),
            ("Discount", 120, "center"),
            ("Amount", 120, "center")
        ]

        for col, width, anchor in columns_config:
            self.table.heading(col, text=col, anchor=anchor)
            self.table.column(col, width=width, anchor=anchor)

        # Add scrollbars
        v_scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.table.yview)
        h_scrollbar = ttk.Scrollbar(table_frame, orient="horizontal", command=self.table.xview)
        self.table.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        self.table.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")

        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        # Configure row colors
        self.table.tag_configure('oddrow', background='#f0f0f0')
        self.table.tag_configure('evenrow', background='#ffffff')

        # Bind double-click event
        self.table.bind("<Double-1>", self.load_selected_item)

        # Total Amount Section
        total_frame = tk.Frame(parent, bg=self.colors['light_bg'])
        total_frame.pack(fill=tk.X, pady=10, padx=10)

        # Amount fields
        amount_fields = [
            ("Total Amount", self.before_discount_amount_field, 0, 0),
            ("Discount Amount", self.discount_amount_field, 0, 2),
            ("After Discount Total Amount", self.After_discount_amount_field, 0, 4),
            ("Total (Before GST & Mahamai):", self.total_amount, 0, 6)
        ]

        for label, var, row, col in amount_fields:
            tk.Label(total_frame, text=label, font=("Arial", 10),
                    bg=self.colors['light_bg']).grid(row=row, column=col, sticky=tk.W, pady=2, padx=5)
            
            entry = tk.Entry(total_frame, textvariable=var, width=20, state="readonly", font=("Arial", 10))
            entry.grid(row=row, column=col+1, sticky=tk.W, pady=2, padx=5)

        # Packing Charge Section
        tk.Label(total_frame, text="Packing Charge Percentage %:", font=("Arial", 10),
                bg=self.colors['light_bg']).grid(row=1, column=0, sticky=tk.W, pady=2, padx=5)
        
        self.packing_charge = tk.DoubleVar(value=0.0)
        packing_entry = tk.Entry(total_frame, textvariable=self.packing_charge, width=20, font=("Arial", 10))
        packing_entry.grid(row=1, column=1, sticky=tk.W, pady=2, padx=5)

        tk.Label(total_frame, text="Packing charge", font=("Arial", 10),
                bg=self.colors['light_bg']).grid(row=1, column=2, sticky=tk.W, pady=2, padx=5)
        
        self.Packing_Amount = tk.DoubleVar(value=0.0)
        packing_amount_entry = tk.Entry(total_frame, textvariable=self.Packing_Amount, width=20, 
                                    state="readonly", font=("Arial", 10))
        packing_amount_entry.grid(row=1, column=3, sticky=tk.W, pady=2, padx=5)

        # Calculate Total Button
        calc_button = tk.Button(total_frame, text="Calculate Total", command=self.calculate_total,
                            font=("Arial", 10), bg=self.colors['primary'], fg="white")
        calc_button.grid(row=0, column=8, rowspan=2, padx=10, pady=5)

        # ✅ Focus navigation setup
            # --- Tab Navigation System ---
        self.focusable_widgets = [
            self.customer_combobox, address_entry, gstin_entry, agent_entry,
            self.bill_no_entry, date_entry, lr_entry, from_entry, to_entry, doc_entry,
            self.product_name_combobox, self.no_of_case_entry, self.per_case_entry,
            self.unit_type_combobox, self.rate_entry, self.per_entry, self.quantity_entry,
            self.discount_entry, packing_entry, calc_button
        ]
        self.current_focus_index = 0

        # Helper to update focus index
        def update_current_focus(event):
            widget = event.widget
            if widget in self.focusable_widgets:
                self.current_focus_index = self.focusable_widgets.index(widget)

        for w in self.focusable_widgets:
            w.bind("<FocusIn>", update_current_focus)

        self.root.bind_all('<Tab>', self.focus_next_widget)
        self.root.bind_all('<Shift-Tab>', self.focus_previous_widget)

        # Set focus to first field
        self.customer_combobox.focus_set()

    def recalculate_gst_from_fields(self, event=None):
        """Recalculate GST percentage when individual GST fields are manually edited"""
        try:
            if self.region.get() == "South":
                # South: CGST + SGST = Total GST
                cgst_val = float(self.cgst.get() or 0)
                sgst_val = float(self.sgst.get() or 0)
                total_gst = cgst_val + sgst_val
                
                # Update main GST percentage field
                self.gst_percentage.set(total_gst)
            else:
                # North: IGST = Total GST
                igst_val = float(self.igst.get() or 0)
                self.gst_percentage.set(igst_val)
            
            # Recalculate totals
            if hasattr(self, 'table') and self.table.get_children():
                self.calculate_total()
                
        except (ValueError, tk.TclError):
            # Handle invalid input
            pass


    def fill_product_details_modern(self, event=None):
        """Modern version of fill_product_details"""
        selected_product_name = self.product_name_combobox.get()
        for key, value in self.product_data.items():
            # Handle both field name formats
            product_name = value.get('Product_Name') or value.get('Product Name')
            if product_name == selected_product_name:
                self.no_of_case_entry.delete(0, tk.END)
                self.no_of_case_entry.insert(0, value.get('No_of_Case') or value.get('No. of Case', ''))
                self.per_case_entry.delete(0, tk.END)
                self.per_case_entry.insert(0, value.get('Per_Case') or value.get('Per Case', ''))
                self.unit_type_combobox.set(value.get('Unit_Type') or value.get('Unit Type', ''))
                self.rate.set(value.get('Selling_Price') or value.get('Selling Price', ''))
                self.per_entry.delete(0, tk.END)
                self.per_entry.insert(0, value.get('Per', 0))
                self.quantity_entry.delete(0, tk.END)
                self.quantity_entry.insert(0, value.get('Quantity', ''))
                self.discount.set(value.get('Discount', ''))
                break

    def load_selected_item(self, event):
        # Get the selected item from the table
        selected_item = self.table.selection()
        if selected_item:
            # Get the values of the selected row
            values = self.table.item(selected_item, "values")

            # Load the values into the input fields
            self.product_name_combobox.set(values[1])  # Product Name
            self.no_of_case_entry.delete(0, tk.END)
            self.no_of_case_entry.insert(0, values[2])  # No. of Case
            self.per_case_entry.delete(0, tk.END)
            self.per_case_entry.insert(0, values[3])  # Per Case
            self.unit_type_combobox.set(values[6])  # Unit Type
            self.rate.set(float(values[5]))  # Rate
            self.quantity.set(int(values[4].split()[0]))  # Quantity
            self.discount.set(float(values[8].split('(')[1].strip('%)')))  # Discount
            self.per_entry.delete(0, tk.END)
            self.per_entry.insert(0, values[7].split()[0])  # Per

    def update_item(self):
        # Get the selected item from the table
        selected_item = self.table.selection()
        if selected_item:
            # Get the updated values from the input fields
            product_name = self.product_name_combobox.get()
            no_of_case = self.no_of_case_entry.get()
            per_case = self.per_case_entry.get()
            unit_type = self.unit_type_combobox.get()
            rate = float(self.rate.get())
            quantity = int(self.quantity.get())
            discount = float(self.discount.get())
            

            # Calculate the updated amount
            amount = (rate / int(self.per_entry.get())) * quantity
            discount_amount = (discount / 100) * amount
            final_amount = amount 

            # Update the table row
            self.table.item(selected_item, values=(
                self.table.item(selected_item, "values")[0],  # S.No remains the same
                product_name,
                no_of_case,
                per_case,
                f"{quantity} {unit_type}",
                rate,
                unit_type,
                f"{self.per_entry.get()} {unit_type}",
                f"{int(discount_amount)} ({int(discount)}%)",
                round(final_amount, 2)
            ))
            
        # Clear the product frame after updating the item
        self.clear_product_frame()


    def reset_gui(self):
        # Reset the flag
        self.bill_no_edited = False

        # Reset the Bill No. entry
        self.bill_no = self.get_next_bill_number()
        self.bill_no_entry.delete(0, tk.END)
        self.bill_no_entry.insert(0, self.bill_no)

        # Reset other fields
        self.bill_date.set(datetime.now().strftime("%d/%m/%Y"))
        self.lr_number.set("")
        self.to_address.set("")
        self.to_gstin.set("")
        self.customer_combobox.set("") 
        self.to_name.set("")
        self.agent_name.set("")
        self.document_through.set("")
        self.from_.set("")  # Always reset the 'From' field
        self.to_.set("")
        self.product_code_2.set("")
        self.rate.set("")
        self.product_name.set("")
        self.type.set("")
        self.case_details.set("")
        self.quantity.set("")
        self.per.set("")  # Clear the 'per' field
        self.discount.set("")
        self.amount.set("")
        self.region.set("South")
        self.gst_percentage.set(18.0)
        self.cgst.set(0.0)
        self.sgst.set(0.0)
        self.igst.set(0.0)
        self.total_amount.set(0.0)
        self.cgst_amount_1.set(0.0)
        self.sgst_amount_1.set(0.0)
        self.igst_amount_1.set(0.0)
        self.discount_percentage.set(0.0)
        self.packing_charge.set(0.0)
        self.discount_amount_field.set(0.0)
        self.before_discount_amount_field.set(0.0)
        self.amount_before_discount.set(0.0)
        self.After_discount_amount_field.set(0.0)
        self.After_Discount_Total_Amount.set(0.0)
        self.Packing_Amount.set(0.0)

        # Clear the product name combobox and other related fields
        self.product_name_combobox.set('')  # Reset product name
        self.no_of_case_entry.delete(0, tk.END)  # Reset no. of case
        self.per_case_entry.delete(0, tk.END)  # Reset per case
        self.unit_type_combobox.set('')  # Reset unit type
        self.rate.set('')  # Reset rate
        self.quantity.set('')  # Reset quantity
        self.discount.set('')  # Reset discount percentage
        self.per_entry.delete(0, tk.END)  # Reset per entry

        # Clear the table
        for item in self.table.get_children():
            self.table.delete(item)

    def generate_pdf(self):
        
        """Enhanced PDF generation with modern UI feedback and error handling"""
        try:
            # Show loading status
            self.show_status_message("📄 Generating PDF invoice...")
            
            # Define the path to the InvoiceApp directory in the user's Documents folder
            documents_dir = os.path.join(os.path.expanduser("~"), "Documents")
            
            # Determine folder name based on bill number prefix
            if self.bill_no.startswith('AP'):
                office_folder = "AP"
            elif self.bill_no.startswith('AFI'):
                office_folder = "AFI" 
            elif self.bill_no.startswith('AFF'):
                office_folder = "AFF"
            else:
                office_folder = {
                    "A1": "AP",
                    "A2": "AFI", 
                    "A3": "AFF"
                }.get(self.selected_office, "AP")

            # Get agent name and create a safe folder name
            agent_name = self.agent_name.get().strip()
            if not agent_name:
                agent_name = "Unknown_Agent"
            
            # Clean agent name for folder
            clean_agent_name = re.sub(r'[^\w\-_]', '', agent_name.replace(" ", "_"))
            
            # 🆕 ADD THIS: Get the year from bill date
            try:
                # Parse the bill date to get year
                bill_date_obj = parse_date_flexible(self.bill_date.get())
                if bill_date_obj:
                    year = str(bill_date_obj.year)
                else:
                    year = str(datetime.now().year)
            except:
                year = str(datetime.now().year)
            
            # 🆕 MODIFIED: Create year-based main folder name
            invoice_app_dir = os.path.join(documents_dir, "InvoiceApp")
            
            # Create year-based main folder: Invoice_Bill_2025, Invoice_Bill_2026, etc.
            main_invoice_folder = f"Invoice_Bill_{year}"
            
            # Create office-specific directory path within year folder
            invoice_bill_dir = os.path.join(
                invoice_app_dir, 
                main_invoice_folder,  # Year-based main folder
                office_folder, 
                clean_agent_name
            )

            # Ensure the year-based directory structure exists
            if not os.path.exists(invoice_bill_dir):
                os.makedirs(invoice_bill_dir, exist_ok=True)

            # Generate the PDF file name
            clean_customer_name = re.sub(r'[^\w\-_]', '', self.to_name.get().replace(" ", "_"))
            padded_bill_no = self.bill_no.zfill(3)
            timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
            
            # 🆕 UPDATE: Include year in relative path
            relative_pdf_path = os.path.join(
                "InvoiceApp", 
                main_invoice_folder,  # Year-based main folder
                office_folder, 
                clean_agent_name, 
                f"{clean_customer_name}_{padded_bill_no}_{timestamp}.pdf"
            )
            absolute_pdf_path = os.path.join(documents_dir, relative_pdf_path)

            # Validate required fields
            if not self.to_name.get().strip():
                messagebox.showerror("❌ Validation Error", "Customer name is required!")
                self.show_status_message("❌ PDF generation failed - Missing customer name")
                return

            if not self.table.get_children():
                messagebox.showerror("❌ Validation Error", "Please add at least one product item!")
                self.show_status_message("❌ PDF generation failed - No products added")
                return

            # If the user has not edited the Bill No., auto-increment it
            if not self.bill_no_edited:
                self.bill_no = self.get_next_bill_number()
                self.bill_no_entry.delete(0, tk.END)
                self.bill_no_entry.insert(0, self.bill_no)
            else:
                self.bill_no = self.bill_no_entry.get()
                
            # Generate the PDF file name with new format: CustomerName_BillNo_Date
            pdf_file_name = os.path.join(invoice_bill_dir, f"{clean_customer_name}_{padded_bill_no}_{timestamp}.pdf")

            # Create PDF object
            pdf = FPDF()
            pdf.add_page()
            pdf.set_left_margin(15)
            pdf.set_right_margin(15)

            # Set font for the entire document
            pdf.set_font("Arial", "B", 9)
            row_height = 5
            
            # Function to add header to every page
            def add_header():
                try:
                    # Add logo on the left
                    pdf.image("logo.png", x=6, y=1, w=27)
                except:
                    # If logo not found, continue without it
                    pass

                # Move to the right for the title
                pdf.set_xy(35, 10)

                if self.selected_office == "A1":
                    pdf.cell(0, row_height, "ANGEL PYROTECH                             ", ln=True, align="C")
                    # Address details
                    pdf.set_x(35)
                    pdf.cell(0, row_height, "D NO 3/89 3/89/1 TO 3/89/11 ONDIPULINAIKANOOR                         ", ln=True, align="C")
                    pdf.set_x(40)
                    pdf.cell(0, row_height, "ONDIPULINAIKANOOR VILLAGE TAMILNADU 626119                                ", ln=True, align="C")

                    # Add Title.png to the top-right corner
                    try:
                        title_img_width = 50
                        title_img_height = 20
                        page_width = pdf.w
                        image_x = page_width - title_img_width - 6
                        image_y = 1
                        pdf.image("Title.png", x=image_x, y=image_y, w=title_img_width, h=title_img_height)
                    except:
                        pass

                    # Add the "Glory To God" slogan at the top-center
                    pdf.set_font("Arial", "I", 8)
                    slogan_text = "Glory To God                                                                           "
                    text_width = pdf.get_string_width(slogan_text)
                    center_x = (page_width - text_width) / 2
                    pdf.set_xy(center_x, 0)
                    pdf.cell(0, 10, slogan_text, ln=True, align='C')

                    # Move down and add "TAX INVOICE" centered
                    pdf.set_font("Arial", "B", 9)
                    pdf.set_xy(15, 30)
                    pdf.cell(0, row_height, "TAX INVOICE", ln=True, align="C")

                    # GSTIN and HSN Code in the same row
                    pdf.set_xy(10, 26)
                    pdf.cell(90, 4, "GSTIN: 33ABRFA4846J1Z3", align="L")
                    pdf.set_xy(110, 26)
                    pdf.cell(90, 4, "HSN CODE: 36041000", align="R")
            
                elif self.selected_office == "A2":
                    pdf.cell(0, row_height, "ANGEL FIREWORKS INDUSTRIES                             ", ln=True, align="C")
                    # Address details
                    pdf.set_x(35)
                    pdf.cell(0, row_height, "FACTORY AT:O.KOVILPATTI,2/2204/W,DEVINAGAR                         ", ln=True, align="C")
                    pdf.set_x(40)
                    pdf.cell(0, row_height, "SIVAKASI-626123                                ", ln=True, align="C")

                    # Add Title.png to the top-right corner
                    try:
                        title_img_width = 50
                        title_img_height = 20
                        page_width = pdf.w
                        image_x = page_width - title_img_width - 6
                        image_y = 1
                        pdf.image("Title.png", x=image_x, y=image_y, w=title_img_width, h=title_img_height)
                    except:
                        pass

                    # Add the "Glory To God" slogan at the top-center
                    pdf.set_font("Arial", "I", 8)
                    slogan_text = "Glory To God                                                                           "
                    text_width = pdf.get_string_width(slogan_text)
                    center_x = (page_width - text_width) / 2
                    pdf.set_xy(center_x, 0)
                    pdf.cell(0, 10, slogan_text, ln=True, align='C')

                    # Move down and add "TAX INVOICE" centered
                    pdf.set_font("Arial", "B", 9)
                    pdf.set_xy(15, 30)
                    pdf.cell(0, row_height, "TAX INVOICE", ln=True, align="C")

                    # GSTIN and HSN Code in the same row
                    pdf.set_xy(10, 26)
                    pdf.cell(90, 4, "GSTIN: 33AARFA9673N2ZL", align="L")
                    pdf.set_xy(110, 26)
                    pdf.cell(90, 4, "HSN CODE: 36041000", align="R")

                elif self.selected_office == "A3":
                    pdf.cell(0, row_height, "ANGEL FIREWORKS FACTORY                             ", ln=True, align="C")
                    # Address details
                    pdf.set_x(35)
                    pdf.cell(0, row_height, "FACTORY AT:O.KOVILPATTI,2/2204/X,DEVINAGAR                         ", ln=True, align="C")
                    pdf.set_x(40)
                    pdf.cell(0, row_height, "VIRUTHUNAGAR-626123                                ", ln=True, align="C")

                    # Add Title.png to the top-right corner
                    try:
                        title_img_width = 50
                        title_img_height = 20
                        page_width = pdf.w
                        image_x = page_width - title_img_width - 6
                        image_y = 1
                        pdf.image("Title.png", x=image_x, y=image_y, w=title_img_width, h=title_img_height)
                    except:
                        pass

                    # Add the "Glory To God" slogan at the top-center
                    pdf.set_font("Arial", "I", 8)
                    slogan_text = "Glory To God                                                                           "
                    text_width = pdf.get_string_width(slogan_text)
                    center_x = (page_width - text_width) / 2
                    pdf.set_xy(center_x, 0)
                    pdf.cell(0, 10, slogan_text, ln=True, align='C')

                    # Move down and add "TAX INVOICE" centered
                    pdf.set_font("Arial", "B", 9)
                    pdf.set_xy(15, 30)
                    pdf.cell(0, row_height, "TAX INVOICE", ln=True, align="C")

                    # GSTIN and HSN Code in the same row
                    pdf.set_xy(10, 26)
                    pdf.cell(90, 4, "GSTIN: 33ABKFA4066F1ZN", align="L")
                    pdf.set_xy(110, 26)
                    pdf.cell(90, 4, "HSN CODE: 36041000", align="R")

                # Draw a line below GSTIN and HSN Code
                pdf.ln()
                pdf.line(10, 35, 200, 35)

            # Add header to the first page
            add_header()
            
            # Customer Information (Left) and Bill Details (Right)
            pdf.set_font("Arial", "B", 9)
            pdf.set_xy(10, 36)
            pdf.cell(90, row_height, "Customer Information", ln=False, align="L")
            pdf.set_xy(110, 36)
            pdf.cell(90, row_height, "Bill Details", ln=False, align="L")
            
            # Ensure customer details are updated before generating the PDF
            self.fill_customer_details()

            # Customer Details (Left Side)
            pdf.set_font("Arial", "", 9)
            pdf.set_xy(10, 42)
            pdf.cell(90, row_height, f"To           :      {self.to_name.get()}", ln=True, align="L")
        
            address = self.to_address.get()
            pdf.set_font("Arial", size=9)

            # First line: "Address : " (no line break)
            pdf.set_xy(10, 48)
            pdf.cell(20, 3, "Address  : ", ln=False, align="L")

            # Let FPDF handle wrapping (no manual splitting)
            pdf.multi_cell(
                w=80,  # Width after "Address : " (adjust as needed)
                h=3,   # Line height
                txt=address,
                align="L",
                border=0
            )

            # Bill Details (Right Side)
            pdf.set_xy(110, 44)
            pdf.cell(90, 0, f"Bill NO             :   {self.bill_no}", ln=True, align="L")
            pdf.set_xy(110, 48)
            pdf.cell(90, 0, f"Bill DATE         :  {self.bill_date.get()}", ln=True, align="L")
            pdf.set_xy(110, 52)
            pdf.cell(90, 0, f"L.R. NUMBER :   {self.lr_number.get()}", ln=True, align="L")
            
            pdf.set_xy(110, 56)
            pdf.cell(90, 0, f"GSTIN             :   {self.to_gstin.get()}", ln=True, align="L")
            

            # Draw a vertical line between Customer Information and Bill Details
            pdf.set_line_width(0.4)
            pdf.line(105, 35, 105, 60)

            # Final Line after Customer & Bill Details
            pdf.line(10, 60, 200, 60)

            # Product Table
            pdf.ln(3)
            pdf.set_x(10)
            pdf.set_font("Arial", "B", 10)
            pdf.cell(0, 10, "Product Details:", ln=True, align='L')

            pdf.set_font("Arial", "B", 10)

            # Table headers
            headers = ["S.No", "Product Name", "Case", "Per Case", "Quantity", "Rate", "Per", "Discount", "Amount"]
            col_widths = [10, 53, 10, 17, 15, 20, 19, 16, 30]

            # Set line width to make borders thinner
            pdf.set_line_width(0.4)

            # Draw header row
            pdf.set_x(10)
            for i, header in enumerate(headers):
                pdf.cell(col_widths[i], 6, header, border=1, align='C')
            
            pdf.set_font("Arial", "", 9)
                
            actual_rows = list(self.table.get_children())
            product_count = len(actual_rows)
            MAX_VISIBLE_ROWS = 22
            row_height = 6  # Fixed row height

            # ==== 1. Draw Table TOP border ====
            pdf.set_x(10)
            pdf.cell(sum(col_widths), row_height, "", border='T', ln=1)

            # ==== 2. Draw Product Rows (NO horizontal lines) ====
            for i, item in enumerate(actual_rows, 1):
                values = self.table.item(item, "values")
                
                pdf.set_x(10)
                # S.No column (left and right border)
                pdf.cell(col_widths[0], row_height, str(i), border='LR', align='C')
                
                # Middle columns (right border only)
                for j, value in enumerate(values[1:]):
                    if j == 5: continue  # Skip Unit Type column
                    adjusted_index = j - 1 if j > 5 else j
                    
                    if adjusted_index == 6:  # Discount column
                        match = re.search(r"(\d+)%", str(values[8]))
                        discount = match.group(1) if match else "0"
                        pdf.cell(col_widths[adjusted_index+1], row_height, f"{discount}%", border='R', align='C')
                    else:
                        pdf.cell(col_widths[adjusted_index+1], row_height, str(value), border='R', align='C')
                
                pdf.ln(row_height)

            # ==== 3. Fill remaining space with blank rows ====
            if product_count < MAX_VISIBLE_ROWS:
                for _ in range(MAX_VISIBLE_ROWS - product_count):
                    pdf.set_x(10)
                    pdf.cell(col_widths[0], row_height, "", border='LR')
                    for width in col_widths[1:-1]:
                        pdf.cell(width, row_height, "", border='R')
                    pdf.cell(col_widths[-1], row_height, "", border='R')
                    pdf.ln(row_height)

            # ==== 4. End Table (No Bottom Border + No Extra Space) ====
            pdf.set_x(10)
            pdf.cell(sum(col_widths), row_height, "", border=0)

            # Amount Section
            def check_page_break(pdf, content_height):
                # Check if adding content will exceed the bottom margin
                if pdf.get_y() + content_height > pdf.h - pdf.b_margin:
                    pdf.add_page()  
                    return pdf.get_y()  
                return pdf.get_y()  
            
            rect_y = check_page_break(pdf, 35)  

            pdf.set_line_width(0.4)
            rect_y = pdf.get_y()
            new_x = 10  
            new_height = 38  

            # Check if content will exceed the page and move to the next page
            rect_y = check_page_break(pdf, new_height)

            pdf.rect(new_x, rect_y, 190, new_height)
            pdf.line(new_x + 90, rect_y, new_x + 90, rect_y + new_height)

            # Calculate total number of cases (1st index of the table)
            total_no_of_cases = sum(int(self.table.item(item, "values")[2]) for item in self.table.get_children())

            # Check and move the content to the next page if needed
            rect_y = check_page_break(pdf, 10)

            # Left Side (From, To, Document Through)
            pdf.set_xy(20, rect_y + 2)
            pdf.set_x(20)
            pdf.set_font("Arial", "B", 9)
            pdf.cell(50, 3, f"No. of Cases          {total_no_of_cases}", ln=True, align="L")

            pdf.set_font("Arial", "", 9)
            if pdf.get_y() + 10 > pdf.h - pdf.b_margin:  
                pdf.add_page()  
                rect_y = pdf.get_y()  

            pdf.set_x(20)
            pdf.cell(50, 5, f"From        : {self.from_.get()}", ln=True, align="L")

            if pdf.get_y() + 10 > pdf.h - pdf.b_margin:  
                pdf.add_page()  
                rect_y = pdf.get_y()  

            pdf.set_x(20)
            pdf.cell(50, 3, f"To            : {self.to_.get()}", ln=True, align="L")

            if pdf.get_y() + 10 > pdf.h - pdf.b_margin:  
                pdf.add_page()  
                rect_y = pdf.get_y()  

            pdf.set_x(20)
            pdf.cell(50, 3, f"Through   : {self.document_through.get()}", ln=True, align="L")

            # Adding space before footer text
            pdf.ln(5)
            if pdf.get_y() + 10 > pdf.h - pdf.b_margin:  
                        pdf.add_page()  
                        rect_y = pdf.get_y()  
            # Footer Note
            pdf.set_font("Arial", "I", 9)
            footer_text = [
                "Note:",
                "1. Company not responsible for transit loss/damage",
                "2. subject to Sivakasi jurisdiction. E.& O.E"
            ]

            # Display footer note
            for line in footer_text:
                pdf.cell(0, 3, line, ln=True, align="L")

            pdf.set_font("Arial", "", 9)

            # Calculate GOODS VALUE (Sum of all amounts in the 9th index of the table)
            goods_value = sum(float(self.table.item(item, "values")[9]) for item in self.table.get_children())
            special_discount = sum(float(self.table.item(item, "values")[8].split('(')[0]) for item in self.table.get_children())
            sub_total = goods_value - special_discount
            packing_charge_percentage = float(self.packing_charge.get())
            packing_charges = (sub_total * packing_charge_percentage) / 100
            sub_total_with_packing = sub_total + packing_charges
            mahamai_percentage = 0.3  # Fixed 0.3%
            mahamai_charges = (sub_total_with_packing * mahamai_percentage) / 100
            taxable_value = sub_total_with_packing + mahamai_charges
            gst_percentage = float(self.gst_percentage.get())
            if self.region.get() == "South":
                cgst_amount = (taxable_value * (gst_percentage / 2)) / 100
                sgst_amount = (taxable_value * (gst_percentage / 2)) / 100
                igst_amount = 0
            else:
                cgst_amount = 0
                sgst_amount = 0
                igst_amount = (taxable_value * gst_percentage) / 100

            def custom_round(value):
                if value % 1 >= 0.5:
                    return int(value) + 1  
                else:
                    return int(value)  
                
            # Calculate the unrounded net amount
            unrounded_net_amount = taxable_value + cgst_amount + sgst_amount + igst_amount

            # Calculate the rounded net amount using custom_round
            rounded_net_amount = custom_round(unrounded_net_amount)

            # Calculate the round off amount
            round_off_amount = rounded_net_amount - unrounded_net_amount
            
            net_amount = custom_round(taxable_value + cgst_amount + sgst_amount + igst_amount)

            # Update the fields in the GUI
            self.before_discount_amount_field.set(round(goods_value, 2))  
            self.discount_amount_field.set(round(special_discount, 2))  
            self.After_discount_amount_field.set(round(sub_total, 2))  
            self.Packing_Amount.set(round(packing_charges, 2))  
            self.total_amount.set(round(net_amount, 2))  

            # Function to round the value to the nearest integer
            def format_value(value):
                return f"{value:.2f}"  
                
            
            pdf.set_xy(110, rect_y+2)
            pdf.cell(50, 3, "             GOODS VALUE", ln=False, align="L")
            pdf.cell(20, 3, f"{format_value(goods_value)}", ln=True, align="R")

            pdf.set_x(110)
            pdf.cell(50, 3, "    SPECIAL DISCOUNT", ln=False, align="L")
            pdf.cell(20, 3, f"-{format_value(special_discount)}", ln=True, align="R")

            pdf.set_x(110)
            pdf.cell(50, 3, "                  SUB TOTAL", ln=False, align="L")
            pdf.cell(20, 3, f"{format_value(sub_total)}", ln=True, align="R")

            pdf.set_x(100)
            pdf.cell(50, 3, f"PACKING CHARGES @ {packing_charge_percentage}%", ln=False, align="L")
            pdf.cell(30, 3, f"{format_value(packing_charges)}", ln=True, align="R")

            pdf.set_x(110)
            pdf.cell(50, 3, "                   SUB TOTAL", ln=False, align="L")
            pdf.cell(20, 3, f"{format_value(sub_total_with_packing)}", ln=True, align="R")

            pdf.set_x(100)
            pdf.cell(50, 3, f"                  MAHAMAI @ {mahamai_percentage}%", ln=False, align="L")
            pdf.cell(30, 3, f"{format_value(mahamai_charges)}", ln=True, align="R")

            pdf.set_x(110)
            pdf.cell(50, 3, "          TAXABLE VALUE", ln=False, align="L")
            pdf.cell(20, 3, f"{format_value(taxable_value)}", ln=True, align="R")

            if self.region.get() == "South":
                pdf.set_x(110)
                pdf.cell(50, 3, "                     CGST (9%)", ln=False, align="L")
                pdf.cell(20, 3, f"{format_value(cgst_amount)}", ln=True, align="R")

                pdf.set_x(110)
                pdf.cell(50, 3, "                     SGST (9%)", ln=False, align="L")
                pdf.cell(20, 3, f"{format_value(sgst_amount)}", ln=True, align="R")
            else:
                pdf.set_x(110)
                pdf.cell(50, 3, "                     IGST (18%)", ln=False, align="L")
                pdf.cell(20, 3, f"{format_value(igst_amount)}", ln=True, align="R")

            # Add the Round Off section to the PDF
            pdf.set_x(110)
            pdf.cell(50, 3, "                  ROUND OFF", ln=False, align="L")
            pdf.cell(20, 3, f"{round_off_amount:.2f}", ln=True, align="R")

            pdf.set_x(110)
            pdf.set_font("Arial", "B", 10)
            pdf.cell(50, 7, "           NET AMOUNT", ln=False, align="L")
            pdf.cell(20, 7, f"{format_value(net_amount)}", ln=True, align="R")

            if pdf.get_y() + 10 > pdf.h - pdf.b_margin:
                pdf.add_page()
                rect_y = pdf.get_y()

            # Amount in words
            pdf.ln(5)
            pdf.set_line_width(0.4)
            rect_y = pdf.get_y()
            new_x = 10
            width = 190
            height = 5
            pdf.rect(new_x, rect_y, width, height)
            pdf.set_xy(new_x + 5, rect_y - 2)
            pdf.set_font("Arial", "I", 9)
            amount_in_words = self.number_to_words(self.total_amount.get())
            pdf.cell(0, 10, f"Amount in Words: {amount_in_words}", ln=True, align="L")

            # Check for page break again after the NET AMOUNT
            rect_y = check_page_break(pdf, 10)
            
            # Set the y-coordinate to a specific value to move up
            pdf.set_y(pdf.get_y() - 3)

            # Determine the company name based on the selected office
            company_name = {
                "A1": "ANGEL PYROTECH",
                "A2": "Angel Fireworks Industries",
                "A3": "Angel Fireworks Factory"
            }.get(self.selected_office, "ANGEL PYROTECH")

            pdf.cell(0, 10, f"                                                                                                                For {company_name}", ln=True, align="C")
            pdf.cell(0, 10, "                                                                                                                Authorized Signature", ln=True, align="C")

            # Save the PDF file
            pdf.output(pdf_file_name)

            # Show success message with modern design
            self.show_status_message("✅ PDF invoice generated successfully!")
            messagebox.showinfo("✅ Success", 
                            f"Invoice saved successfully!\n\n"
                            f"📄 File: {os.path.basename(pdf_file_name)}\n"
                            f"📁 Location: {invoice_bill_dir}\n"
                            f"🧾 Bill No: {self.bill_no}")

            # Call display_pdf to show the saved PDF in the application
            self.display_pdf(pdf_file_name)

            # Save the bill details - Store RELATIVE path instead of absolute
            bill_details = {
                "bill_no": self.bill_no,
                "bill_date": self.bill_date.get(),
                "pdf_file_name": relative_pdf_path, 
                "customer_name": self.to_name.get(),
                "address": self.to_address.get(),
                "agent_name": self.agent_name.get(),
                "gstin": self.to_gstin.get(),
                "lr_number": self.lr_number.get(),
                "from_": self.from_.get(),
                "to_": self.to_.get(),
                "document_through": self.document_through.get(),
                "region": self.region.get(),
                "gst_percentage": float(self.gst_percentage.get()),  # ✅ Convert to float
                "packing_charge": float(self.packing_charge.get()),  # ✅ Convert to float
                "no_of_cases": sum(int(self.table.item(item, "values")[2]) for item in self.table.get_children()),
                "net_amount": float(self.total_amount.get()),  # ✅ Convert to float
                "payment_status": "Pending",
                "items": [self.table.item(item, "values") for item in self.table.get_children()],
                # Store RELATIVE path from Documents folder
                "pdf_file_name": os.path.relpath(pdf_file_name, os.path.join(os.path.expanduser("~"), "Documents")),
                "cgst_amount": float(self.cgst.get()),  # ✅ Convert to float
                "sgst_amount": float(self.sgst.get()),  # ✅ Convert to float
                "igst_amount": float(self.igst.get()),  # ✅ Convert to float
                "goods_value": float(goods_value),  # ✅ Convert to float
                "special_discount": float(special_discount),  # ✅ Convert to float
                "sub_total": float(sub_total),  # ✅ Convert to float
                "packing_charges": float(packing_charges),  # ✅ Convert to float
                "commission_rate": 0.0,
                "commission_amount": 0.0,
                "commission_calculated_on": "sub_total",
                
                # ✅ ADD THESE FOR COMPLETENESS:
                "office_type": self.selected_office,  # Track which office created this bill
                "created_timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # Audit trail
            }
            
            
            # SAVE ONLY THIS ONE BILL INTO FIREBASE UNDER ITS BILL NUMBER
            self.bills_ref.child(str(self.bill_no)).set(bill_details)

            # REFRESH LOCAL COPY FROM FIREBASE
            self.bills_data = self.load_data(self.bills_ref)

            print("🔥 Bill saved successfully to Firebase!")

            # Update the next bill number after saving
            self.bill_no = self.get_next_bill_number()
            self.bill_no_entry.delete(0, tk.END)
            self.bill_no_entry.insert(0, self.bill_no)

            # Reset the GUI
            self.reset_gui()
        
        except Exception as e:
            error_msg = f"Failed to generate PDF: {str(e)}"
            self.show_status_message(f"❌ {error_msg}", error=True)
            messagebox.showerror("❌ PDF Generation Error", error_msg)

    def display_pdf(self, pdf_filename):
        """Open the PDF in the user's default PDF viewer."""
        try:
            # Convert relative path to absolute path
            pdf_path = self.get_absolute_pdf_path(pdf_filename)

            if os.name == "nt":  # Windows
                os.startfile(pdf_path)
            elif os.name == "posix":  # macOS or Linux
                if os.uname().sysname == "Darwin":  # macOS
                    subprocess.run(["open", pdf_path], check=True)
                else:  # Linux
                    subprocess.run(["xdg-open", pdf_path], check=True)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to open PDF: {e}")

    def get_absolute_pdf_path(self, relative_path):
        documents_folder = os.path.join(os.path.expanduser("~"), "Documents")
        return os.path.join(documents_folder, relative_path)

    def number_to_words(self, num):
        """Convert number to words with enhanced formatting"""
        try:
            words = num2words(num, lang='en_IN').title() + " Rupees Only"
            return words
        except Exception as e:
            return f"Amount: ₹{num:.2f}"

    def get_pdf_path(self, relative_path):
        """
        Convert relative PDF path to absolute path based on current user's Documents folder
        Updated to handle year-based folder structure like Invoice_Bill_2025
        """
        try:
            documents_dir = os.path.join(os.path.expanduser("~"), "Documents")
            absolute_path = os.path.join(documents_dir, relative_path)
            
            # Check if file exists at the resolved path
            if os.path.exists(absolute_path):
                return absolute_path
            else:
                # Try to find the file by filename only (fallback)
                filename = os.path.basename(relative_path)
                
                # Search in all possible year folders
                invoice_app_dir = os.path.join(documents_dir, "InvoiceApp")
                
                if not os.path.exists(invoice_app_dir):
                    return absolute_path
                
                # Look for folders matching pattern "Invoice_Bill_20XX"
                for folder_name in os.listdir(invoice_app_dir):
                    if folder_name.startswith("Invoice_Bill_") and os.path.isdir(os.path.join(invoice_app_dir, folder_name)):
                        # Search in this year folder
                        year_folder_path = os.path.join(invoice_app_dir, folder_name)
                        for root, dirs, files in os.walk(year_folder_path):
                            if filename in files:
                                return os.path.join(root, filename)
                
                # If still not found, return the expected path
                return absolute_path
                
        except Exception as e:
            print(f"Error resolving PDF path: {e}")
            return relative_path

    def update_gst_fields(self):
        """Enhanced GST field management with modern UI updates"""
        try:
            if self.region.get() == "South":
                self.cgst_label.grid()
                self.cgst_entry.grid()
                self.sgst_label.grid()
                self.sgst_entry.grid()
                self.igst_label.grid_remove()
                self.igst_entry.grid_remove()
                
                # Update GST percentages (not amounts here - amounts calculated in calculate_total)
                # Just show that fields are for amounts
                self.cgst_label.config(text="CGST Amount:")
                self.sgst_label.config(text="SGST Amount:")
                
            else:
                self.cgst_label.grid_remove()
                self.cgst_entry.grid_remove()
                self.sgst_label.grid_remove()
                self.sgst_entry.grid_remove()
                self.igst_label.grid()
                self.igst_entry.grid()
                
                # Update IGST field for amount
                self.igst_label.config(text="IGST Amount:")
                
            self.show_status_message(f"📊 GST fields updated for {self.region.get()} region")
            
            # Recalculate GST amounts if there are items
            if hasattr(self, 'table') and self.table.get_children():
                self.calculate_total()
                
        except Exception as e:
            self.show_status_message(f"❌ Error updating GST fields: {e}", error=True)

    def calculate_total(self):
        """Calculate total amount based on items in the table"""
        try:
            # Calculate total amount from table items
            total_amount = 0.0
            total_discount = 0.0
            
            for item in self.table.get_children():
                values = self.table.item(item, "values")
                if len(values) >= 10:  # Ensure we have all columns
                    try:
                        amount = float(values[9])  # Amount is in column 9
                        total_amount += amount
                        
                        # Extract discount amount from discount column (format: "amount (percentage%)")
                        discount_str = values[8]
                        if '(' in discount_str and ')' in discount_str:
                            discount_amount_str = discount_str.split('(')[0].strip()
                            if discount_amount_str:
                                discount_amount = float(discount_amount_str)
                                total_discount += discount_amount
                    except (ValueError, IndexError):
                        continue
            
            # Update amount fields
            self.before_discount_amount_field.set(round(total_amount, 2))
            self.discount_amount_field.set(round(total_discount, 2))
            
            # Calculate after discount amount
            after_discount = total_amount - total_discount
            self.After_discount_amount_field.set(round(after_discount, 2))
            
            # Calculate packing charges
            try:
                packing_percentage = float(self.packing_charge.get() or 0)
                packing_amount = (after_discount * packing_percentage) / 100
                self.Packing_Amount.set(round(packing_amount, 2))
            except ValueError:
                packing_amount = 0.0
                self.Packing_Amount.set(0.0)
            
            # Calculate GST AMOUNTS (Rupees)
            subtotal_with_packing = after_discount + packing_amount
            
            try:
                gst_percentage = float(self.gst_percentage.get() or 0)
                
                if self.region.get() == "South":
                    # Calculate CGST and SGST amounts (Rupees)
                    cgst_amount = (subtotal_with_packing * (gst_percentage / 2)) / 100
                    sgst_amount = (subtotal_with_packing * (gst_percentage / 2)) / 100
                    igst_amount = 0.0
                    
                    # ✅ UPDATE GST ENTRY FIELDS WITH AMOUNTS (Rupees)
                    self.cgst.set(round(cgst_amount, 2))  # Set CGST AMOUNT in rupees
                    self.sgst.set(round(sgst_amount, 2))  # Set SGST AMOUNT in rupees
                    self.igst.set(0.0)  # Set IGST to 0
                else:
                    # Calculate IGST amount (Rupees)
                    cgst_amount = 0.0
                    sgst_amount = 0.0
                    igst_amount = (subtotal_with_packing * gst_percentage) / 100
                    
                    # ✅ UPDATE GST ENTRY FIELDS WITH AMOUNTS (Rupees)
                    self.cgst.set(0.0)  # Set CGST to 0
                    self.sgst.set(0.0)  # Set SGST to 0
                    self.igst.set(round(igst_amount, 2))  # Set IGST AMOUNT in rupees
                
                # Update GST amount fields if they exist
                if hasattr(self, 'cgst_amount_1'):
                    self.cgst_amount_1.set(round(cgst_amount, 2))
                if hasattr(self, 'sgst_amount_1'):
                    self.sgst_amount_1.set(round(sgst_amount, 2))
                if hasattr(self, 'igst_amount_1'):
                    self.igst_amount_1.set(round(igst_amount, 2))
                
                # Calculate final total
                final_total = subtotal_with_packing + cgst_amount + sgst_amount + igst_amount
                self.total_amount.set(round(final_total, 2))
                
            except (ValueError, tk.TclError):
                # If GST calculation fails, just use subtotal with packing
                self.total_amount.set(round(subtotal_with_packing, 2))
                # Reset GST amounts on error
                self.cgst.set(0.0)
                self.sgst.set(0.0)
                self.igst.set(0.0)
                
        except Exception as e:
            print(f"Error in calculate_total: {e}")
            # Set default values on error
            self.before_discount_amount_field.set(0.0)
            self.discount_amount_field.set(0.0)
            self.After_discount_amount_field.set(0.0)
            self.Packing_Amount.set(0.0)
            self.total_amount.set(0.0)
            # Also reset GST fields on error
            self.cgst.set(0.0)
            self.sgst.set(0.0)
            self.igst.set(0.0)

    def add_item(self):
        """Enhanced item addition with modern validation and feedback"""
        try:
            # Gather input data
            product_name = self.product_name_combobox.get().strip()
            if not product_name:
                messagebox.showwarning("⚠️ Input Required", "Please select or enter a product name!")
                self.product_name_combobox.focus_set()
                return

            # Validate numeric inputs
            try:
                no_of_case = int(self.no_of_case_entry.get())
                per_case = int(self.per_case_entry.get())
                rate = float(self.rate.get())
                per = int(self.per_entry.get())
                discount_percentage = float(self.discount.get())
            except ValueError as e:
                messagebox.showerror("❌ Invalid Input", "Please enter valid numeric values for all fields!")
                return

            unit_type = self.unit_type_combobox.get()
            if not unit_type:
                messagebox.showwarning("⚠️ Input Required", "Please select a unit type!")
                self.unit_type_combobox.focus_set()
                return
                
            # Calculate the quantity based on Per Case * No. of Case
            quantity = per_case * no_of_case
            self.quantity.set(quantity)

            # Calculate the amount based on the new formula
            amount = (rate / per) * quantity

            # Calculate the discount amount
            discount_amount = (discount_percentage / 100) * amount

            # Format the discount column
            discount_display = f"{int(discount_amount)} ({int(discount_percentage)}%)"

            # Final amount after applying discount
            final_amount = amount

            # Add row to the table
            current_row = len(self.table.get_children()) + 1
            self.table.insert("", "end", values=(
                current_row,
                product_name,
                no_of_case,
                per_case,
                f"{quantity} {unit_type}",
                rate,
                unit_type,
                f"{per} {unit_type}",
                discount_display,
                round(final_amount, 2)
            ))
            
            # Clear the product frame after adding the item
            self.clear_product_frame()

            # Call calculate_total() after adding the item to update totals
            self.calculate_total()
            
            self.show_status_message(f"📦 Added: {product_name}")
            
        except Exception as e:
            error_msg = f"Failed to add item: {str(e)}"
            self.show_status_message(f"❌ {error_msg}", error=True)
            messagebox.showerror("❌ Add Item Error", error_msg)

    def clear_product_frame(self):
        """Enhanced product frame clearing with modern feedback"""
        try:
            self.product_name_combobox.set('')
            self.no_of_case_entry.delete(0, tk.END)
            self.per_case_entry.delete(0, tk.END)
            self.unit_type_combobox.set('')
            self.rate.set('')
            self.quantity.set('')
            self.discount.set('')
            self.per_entry.delete(0, tk.END)
            
            # Set focus back to product name for quick entry
            self.product_name_combobox.focus_set()
            
            self.show_status_message("🔄 Product fields cleared")
            
        except Exception as e:
            self.show_status_message(f"❌ Error clearing fields: {e}", error=True)

    def remove_item(self):
        """Enhanced item removal with modern confirmation and feedback"""
        selected_item = self.table.selection()
        
        if not selected_item:
            messagebox.showwarning("⚠️ Selection Required", "Please select an item to remove!")
            return

        # Get item details for confirmation
        item_values = self.table.item(selected_item[0], "values")
        product_name = item_values[1] if len(item_values) > 1 else "Selected Item"
        
        # Confirm deletion
        confirm = messagebox.askyesno(
            "🗑️ Confirm Removal",
            f"Are you sure you want to remove this item?\n\n"
            f"Product: {product_name}\n"
            f"This action cannot be undone!",
            icon='warning'
        )
        
        if confirm:
            self.table.delete(selected_item)
            
            # Update the S.No of the remaining items
            for index, item in enumerate(self.table.get_children(), start=1):
                values = self.table.item(item, "values")
                updated_values = (index,) + values[1:]
                self.table.item(item, values=updated_values)

            self.calculate_total()
            self.show_status_message(f"🗑️ Removed: {product_name}")

    def reset_gui(self):
        # Reset the flag
        self.bill_no_edited = False

        # Reset the Bill No. entry
        self.bill_no = self.get_next_bill_number()
        self.bill_no_entry.delete(0, tk.END)
        self.bill_no_entry.insert(0, self.bill_no)

        # Reset other fields
        self.bill_date.set(datetime.now().strftime("%d/%m/%Y"))
        self.lr_number.set("")
        self.to_address.set("")
        self.to_gstin.set("")
        self.customer_combobox.set("") 
        self.to_name.set("")
        self.agent_name.set("")
        self.document_through.set("")
        self.from_.set("")  # Always reset the 'From' field
        self.to_.set("")
        self.product_code_2.set("")
        self.rate.set("")
        self.product_name.set("")
        self.type.set("")
        self.case_details.set("")
        self.quantity.set("")
        self.per.set("")  # Clear the 'per' field
        self.discount.set("")
        self.amount.set("")
        self.region.set("South")
        self.gst_percentage.set(18.0)
        self.cgst.set(0.0)
        self.sgst.set(0.0)
        self.igst.set(0.0)
        self.total_amount.set(0.0)
        self.cgst_amount_1.set(0.0)
        self.sgst_amount_1.set(0.0)
        self.igst_amount_1.set(0.0)
        self.discount_percentage.set(0.0)
        self.packing_charge.set(0.0)
        self.discount_amount_field.set(0.0)
        self.before_discount_amount_field.set(0.0)
        self.amount_before_discount.set(0.0)
        self.After_discount_amount_field.set(0.0)
        self.After_Discount_Total_Amount.set(0.0)
        self.Packing_Amount.set(0.0)

        # Clear the product name combobox and other related fields
        self.product_name_combobox.set('')  # Reset product name
        self.no_of_case_entry.delete(0, tk.END)  # Reset no. of case
        self.per_case_entry.delete(0, tk.END)  # Reset per case
        self.unit_type_combobox.set('')  # Reset unit type
        self.rate.set('')  # Reset rate
        self.quantity.set('')  # Reset quantity
        self.discount.set('')  # Reset discount percentage
        self.per_entry.delete(0, tk.END)  # Reset per entry

        # Clear the table
        for item in self.table.get_children():
            self.table.delete(item)

    def on_bill_no_edit(self, event):
        """Enhanced bill number edit handler with validation"""
        try:
            current_value = self.bill_no_entry.get().strip()
            
            # Only set edited flag if the value is valid
            if current_value and any(prefix in current_value for prefix in ['AP', 'AFI', 'AFF']):
                self.bill_no_edited = True
                print(f"Bill number manually edited to: {current_value}")
            else:
                print("Invalid bill number format, auto-generation will be used")
                self.bill_no_edited = False
                
        except Exception as e:
            print(f"Error in on_bill_no_edit: {e}")
            self.bill_no_edited = False

    def load_selected_item(self, event):
        # Get the selected item from the table
        selected_item = self.table.selection()
        if selected_item:
            # Get the values of the selected row
            values = self.table.item(selected_item, "values")
            
            # Load the values into the input fields
            self.product_name_combobox.set(values[1])  # Product Name
            self.no_of_case_entry.delete(0, tk.END)
            self.no_of_case_entry.insert(0, values[2])  # No. of Case
            
            self.per_case_entry.delete(0, tk.END)
            self.per_case_entry.insert(0, values[3])  # Per Case
            
            self.unit_type_combobox.set(values[6])  # Unit Type
            self.rate.set(float(values[5]))  # Rate
            
            # ✅ CORRECTION: Extract just the numeric part of quantity
            try:
                quantity_str = str(values[4])  # Quantity with unit
                # Extract numbers from string like "100 Box" -> 100
                quantity_num = int(''.join(filter(str.isdigit, quantity_str.split()[0])))
                self.quantity.set(quantity_num)
            except (ValueError, IndexError):
                self.quantity.set(0)
            
            # ✅ CORRECTION: Extract discount percentage properly
            try:
                discount_str = str(values[8])  # "450 (15%)"
                # Extract percentage from between parentheses
                if '(' in discount_str and '%)' in discount_str:
                    discount_percent = discount_str.split('(')[1].strip('%)')
                    self.discount.set(float(discount_percent))
                else:
                    self.discount.set(0.0)
            except (ValueError, IndexError):
                self.discount.set(0.0)
            
            # ✅ CORRECTION: Extract per value properly
            try:
                per_str = str(values[7])  # "1 Box"
                per_value = per_str.split()[0]  # Get the number part
                self.per_entry.delete(0, tk.END)
                self.per_entry.insert(0, per_value)
            except (ValueError, IndexError):
                self.per_entry.delete(0, tk.END)

    def update_item(self):
        """Update the selected item in the table with corrected calculations"""
        # Get the selected item from the table
        selected_item = self.table.selection()
        if selected_item:
            try:
                # Get the updated values from the input fields
                product_name = self.product_name_combobox.get()
                unit_type = self.unit_type_combobox.get()
                
                # Validate numeric inputs
                try:
                    no_of_case = int(self.no_of_case_entry.get())
                    per_case = int(self.per_case_entry.get())
                    rate = float(self.rate.get())
                    discount = float(self.discount.get())
                    per_value = int(self.per_entry.get())
                except ValueError:
                    messagebox.showerror("❌ Invalid Input", "Please enter valid numeric values!")
                    return
                
                # ✅ CORRECTION: Recalculate quantity = No. of Case × Per Case
                quantity = no_of_case * per_case
                
                # Calculate the updated amount
                amount = (rate / per_value) * quantity
                discount_amount = (discount / 100) * amount
                final_amount = amount

                # Update the table row
                self.table.item(selected_item, values=(
                    self.table.item(selected_item, "values")[0],  # S.No remains the same
                    product_name,
                    no_of_case,
                    per_case,
                    f"{quantity} {unit_type}",  # ✅ Updated quantity display
                    rate,
                    unit_type,
                    f"{per_value} {unit_type}",
                    f"{int(discount_amount)} ({int(discount)}%)",
                    round(final_amount, 2)
                ))
                
                # ✅ IMPORTANT: Update the quantity field in the GUI
                self.quantity.set(quantity)
                
                # ✅ Recalculate totals after updating item
                self.calculate_total()
                
                self.show_status_message(f"✅ Updated: {product_name}")
                
            except Exception as e:
                messagebox.showerror("❌ Update Error", f"Failed to update item:\n{str(e)}")
                return
            
        # Clear the product frame after updating the item
        self.clear_product_frame()

    def on_gst_percentage_change(self, *args):
        """Handle GST percentage change and update GST amounts accordingly"""
        try:
            # Don't show amounts in percentage field - just recalculate totals
            if hasattr(self, 'table') and self.table.get_children():
                self.calculate_total()
                
        except (ValueError, tk.TclError):
            # Handle invalid input gracefully
            pass

    def show_edit_bill(self):
        """Modern enhanced version of show_edit_bill for editing bills"""
        self.clear_screen()
        self.current_screen = "edit_bill"
        self.create_navigation_bar()
        self.create_status_bar()

        # Reset focusable widgets
        self.focusable_widgets.clear()

        # Main content area
        main_container = tk.Frame(self.root, bg=self.colors['light_bg'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Header frame
        header_frame = self.create_modern_frame(main_container, "✏️ EDIT EXISTING BILLS")
        header_frame.pack(fill=tk.X, pady=(0, 20))

        # Search and filter section
        search_frame = tk.Frame(header_frame, bg=self.colors['card_bg'])
        search_frame.pack(fill=tk.X, padx=20, pady=15)

        # Search entry
        tk.Label(search_frame, text="🔍 Search Bills:", font=("Segoe UI", 10, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).pack(side=tk.LEFT, padx=5)
        
        self.bill_search_entry = self.create_modern_entry(search_frame, width=30)
        self.bill_search_entry.pack(side=tk.LEFT, padx=10, pady=5)
        self.bill_search_entry.bind('<KeyRelease>', self.filter_bill_list)

        # Filter by date frame
        date_filter_frame = tk.Frame(search_frame, bg=self.colors['card_bg'])
        date_filter_frame.pack(side=tk.LEFT, padx=20)

        tk.Label(date_filter_frame, text="📅 Filter by Date:", font=("Segoe UI", 10, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).pack(side=tk.LEFT, padx=5)
        
        self.date_filter_var = tk.StringVar(value="All")
        date_filter_combo = self.create_modern_combobox(
            date_filter_frame, 
            values=["All", "Today", "This Week", "This Month", "Last Month"],
            textvariable=self.date_filter_var,
            width=12
        )
        date_filter_combo.pack(side=tk.LEFT, padx=10, pady=5)
        date_filter_combo.bind('<<ComboboxSelected>>', self.filter_bill_list)

        # ✅ ADDED: Office Filter
        office_filter_frame = tk.Frame(search_frame, bg=self.colors['card_bg'])
        office_filter_frame.pack(side=tk.LEFT, padx=20)

        tk.Label(office_filter_frame, text="🏢 Office:", font=("Segoe UI", 10, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).pack(side=tk.LEFT, padx=5)
        
        self.office_filter_var = tk.StringVar(value="Current")  # Default to current office only
        office_filter_combo = self.create_modern_combobox(
            office_filter_frame, 
            values=["Current", "All", "AP", "AFI", "AFF"],
            textvariable=self.office_filter_var,
            width=10
        )
        office_filter_combo.pack(side=tk.LEFT, padx=10, pady=5)
        office_filter_combo.bind('<<ComboboxSelected>>', self.filter_bill_list)


        # Action buttons
        action_frame = tk.Frame(search_frame, bg=self.colors['card_bg'])
        action_frame.pack(side=tk.RIGHT, padx=10)

        refresh_btn = self.create_modern_button(
            action_frame, "🔄 Refresh (F5)", self.refresh_bill_list,
            style="info", width=15, height=1
        )
        refresh_btn.pack(side=tk.LEFT, padx=5)

        # Table container
        table_container = self.create_modern_frame(main_container, "📋 ALL BILLS")
        table_container.pack(fill=tk.BOTH, expand=True, pady=10)

        # Create a frame for table and scrollbars
        table_main = tk.Frame(table_container, bg=self.colors['card_bg'])
        table_main.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        # ✅ Create Treeview with modern styling (fixed invisible header issue)
        style = ttk.Style()
        style.theme_use("default")

        # --- Fix invisible heading text layout ---
        style.layout("Modern.Treeview.Heading", [
            ("Treeheading.cell", {"sticky": "nswe"}),
            ("Treeheading.border", {"sticky": "nswe", "children": [
                ("Treeheading.padding", {"sticky": "nswe", "children": [
                    ("Treeheading.image", {"side": "right", "sticky": ""}),
                    ("Treeheading.text", {"sticky": "we"})
                ]})
            ]}),
        ])

        # --- Table row style ---
        style.configure("Modern.Treeview", 
            font=("Segoe UI", 10),
            rowheight=28,
            background=self.colors['card_bg'],
            fieldbackground=self.colors['card_bg'],
            foreground=self.colors['text_dark']
        )

        # --- Header style (always visible now) ---
        style.configure("Modern.Treeview.Heading", 
            font=("Segoe UI", 11, "bold"),
            background=self.colors['primary'],
            foreground=self.colors['text_light'],
            relief="flat"
        )

        # --- Hover / pressed effects for heading ---
        style.map("Modern.Treeview.Heading",
            background=[('active', self.colors['accent']), ('pressed', self.colors['secondary'])],
            relief=[('pressed', 'groove'), ('active', 'ridge')]
        )

        # ✅ Create Treeview with enhanced columns
        self.bill_table = ttk.Treeview(
            table_main, 
            columns=("Bill No", "Bill Date", "Customer Name", "Agent Name", 
                    "Net Amount", "Status", "Items Count", "Office"),  # ✅ ADDED Office column
            show="headings",
            style="Modern.Treeview",
            selectmode="extended",
            height=15
        )

        # Define column headings with appropriate widths
        columns = {
            "Bill No": {"width": 120, "anchor": "center"},
            "Bill Date": {"width": 100, "anchor": "center"},
            "Customer Name": {"width": 180, "anchor": "w"},
            "Agent Name": {"width": 120, "anchor": "w"},
            "Net Amount": {"width": 100, "anchor": "center"},
            "Status": {"width": 80, "anchor": "center"},
            "Items Count": {"width": 80, "anchor": "center"},
            "Office": {"width": 80, "anchor": "center"}  # ✅ ADDED Office column
        }

        for col, settings in columns.items():
            self.bill_table.heading(col, text=col)
            self.bill_table.column(col, width=settings["width"], anchor=settings["anchor"])

        # Add scrollbars
        v_scrollbar = ttk.Scrollbar(table_main, orient="vertical", command=self.bill_table.yview)
        h_scrollbar = ttk.Scrollbar(table_main, orient="horizontal", command=self.bill_table.xview)
        self.bill_table.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        # Grid layout for table and scrollbars
        self.bill_table.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")

        # Configure grid weights
        table_main.grid_rowconfigure(0, weight=1)
        table_main.grid_columnconfigure(0, weight=1)

        # Status bar for table info
        status_frame = tk.Frame(main_container, bg=self.colors['light_bg'])
        status_frame.pack(fill=tk.X, pady=10)

        self.bill_table_status = tk.Label(
            status_frame,
            text="📊 Total Bills: 0",
            font=("Segoe UI", 10),
            bg=self.colors['light_bg'],
            fg=self.colors['text_dark']
        )
        self.bill_table_status.pack(side=tk.LEFT, padx=10)

        # ✅ ADDED: Current Office Info
        current_office_info = tk.Label(
            status_frame,
            text=f"🏢 Current Office: {self.get_current_office_name()}",
            font=("Segoe UI", 10, "bold"),
            bg=self.colors['light_bg'],
            fg=self.colors['primary']
        )
        current_office_info.pack(side=tk.LEFT, padx=20)

        # Selected bill info
        self.selected_bill_info = tk.Label(
            status_frame,
            text="ℹ️  Select a bill to view details",
            font=("Segoe UI", 9),
            bg=self.colors['light_bg'],
            fg=self.colors['text_muted']
        )
        self.selected_bill_info.pack(side=tk.RIGHT, padx=10)

        # Action buttons frame
        action_buttons_frame = tk.Frame(main_container, bg=self.colors['light_bg'])
        action_buttons_frame.pack(fill=tk.X, pady=10)

        # Center the buttons
        button_container = tk.Frame(action_buttons_frame, bg=self.colors['light_bg'])
        button_container.pack(expand=True)

        # Edit selected button
        edit_btn = self.create_modern_button(
            button_container,
            "✏️ Edit Selected Bill",
            self.edit_selected_bill,
            style="primary",
            width=18,
            height=2
        )
        edit_btn.pack(side=tk.LEFT, padx=8)

        # Delete selected button
        delete_btn = self.create_modern_button(
            button_container,
            "🗑️ Delete Selected",
            self.delete_selected_bill,
            style="warning",
            width=18,
            height=2
        )
        delete_btn.pack(side=tk.LEFT, padx=8)

        # View PDF button
        view_pdf_btn = self.create_modern_button(
            button_container,
            "📄 View PDF",
            self.view_selected_bill_pdf,
            style="info",
            width=15,
            height=2
        )
        view_pdf_btn.pack(side=tk.LEFT, padx=8)

        # Back button
        back_btn = self.create_modern_button(
            button_container,
            "← Back to Billing (F4)",
            self.show_billing_dashboard,
            style="secondary",
            width=18,
            height=2
        )
        back_btn.pack(side=tk.LEFT, padx=8)

        # Populate the table
        self.populate_bill_table()

        # Bind keyboard shortcuts
        self.root.bind('<F5>', lambda e: self.refresh_bill_list())
        self.root.bind('<F4>', lambda e: self.show_billing_dashboard())
        self.root.bind('<Delete>', lambda e: self.delete_selected_bill())

        # Bind events to the table
        self.bill_table.bind('<Double-1>', self.on_double_click_edit)
        self.bill_table.bind('<Return>', self.on_double_click_edit)
        self.bill_table.bind('<<TreeviewSelect>>', self.on_bill_selection_change)
        self.bill_table.bind('<Button-3>', self.show_bill_context_menu)

        # Set focus to search field
        self.bill_search_entry.focus_set()

        self.show_status_message("✏️ Bill list loaded - Double-click or use context menu to edit bills")

    # ========== SUPPORTING METHODS FOR EDIT BILL ==========

    def get_current_office_name(self):
        """Get the display name of current office"""
        office_names = {
            "A1": "Angel Pyrotech (AP)",
            "A2": "Angel Fireworks Industries (AFI)", 
            "A3": "Angel Fireworks Factory (AFF)"
        }
        return office_names.get(self.selected_office, f"Office {self.selected_office}")

    def get_bill_office(self, bill_no):
        """Detect office from bill number"""
        if bill_no.startswith('AP'):
            return "A1"
        elif bill_no.startswith('AFI'):
            return "A2" 
        elif bill_no.startswith('AFF'):
            return "A3"
        else:
            return "A1"  # Default

    def get_office_prefix(self, office_code):
        """Get bill prefix for office code"""
        prefix_map = {
            "A1": "AP",
            "A2": "AFI",
            "A3": "AFF"
        }
        return prefix_map.get(office_code, "AP")

    def populate_bill_table(self, filtered_data=None):
        """Populate the bill table with data from bills.json"""
        try:
            # Clear existing data
            for item in self.bill_table.get_children():
                self.bill_table.delete(item)
            
            # Use provided data or all bills data
            bills_data = filtered_data if filtered_data is not None else self.bills_data
            
            # Calculate statistics
            total_bills = len(bills_data)
            total_amount = 0
            
            # Insert data with alternating row colors
            for index, (bill_no, bill_info) in enumerate(bills_data.items()):
                tags = ('evenrow',) if index % 2 == 0 else ('oddrow',)
                
                # Extract bill information
                customer_name = bill_info.get('customer_name', 'N/A')
                agent_name = bill_info.get('agent_name', 'N/A')
                bill_date = bill_info.get('bill_date', 'N/A')
                net_amount = float(bill_info.get('net_amount', 0))
                items_count = len(bill_info.get('items', []))
                status = bill_info.get('payment_status', 'Pending')
                
                # ✅ ADDED: Detect office from bill number
                bill_office = self.get_bill_office(bill_no)
                office_display = {"A1": "AP", "A2": "AFI", "A3": "AFF"}.get(bill_office, "AP")
                
                # Update statistics
                total_amount += net_amount
                
                self.bill_table.insert("", "end", values=(
                    bill_no,
                    bill_date,
                    customer_name,
                    agent_name,
                    f"₹{net_amount:,.2f}",
                    status,
                    items_count,
                    office_display  # ✅ ADDED Office column
                ), tags=tags)

            # Configure row colors
            self.bill_table.tag_configure('evenrow', background='#ffffff')
            self.bill_table.tag_configure('oddrow', background='#f0f8ff')
            
            # Update status
            if hasattr(self, 'bill_table_status'):
                self.bill_table_status.config(text=f"📊 Total Bills: {total_bills} | Total Amount: ₹ 0")
                
        except Exception as e:
            print(f"Error populating bill table: {e}")

    def filter_bill_list(self, event=None):
        """Filter bill list based on search criteria"""
        search_term = self.bill_search_entry.get().lower()
        date_filter = self.date_filter_var.get()
        office_filter = self.office_filter_var.get()
        
        filtered_data = {}
        
        for bill_no, bill_info in self.bills_data.items():
            # Search term filter
            matches_search = (
                search_term in bill_no.lower() or
                search_term in bill_info.get('customer_name', '').lower() or
                search_term in bill_info.get('agent_name', '').lower()
            )
            
            if not matches_search:
                continue
                
            # Date filter
            if date_filter != "All":
                bill_date_str = bill_info.get('bill_date', '')
                if bill_date_str:
                    try:
                        bill_date = parse_date_flexible(bill_date_str)
                        today = datetime.now()
                        
                        if date_filter == "Today":
                            if bill_date.date() != today.date():
                                continue
                        elif date_filter == "This Week":
                            start_of_week = today - timedelta(days=today.weekday())
                            if bill_date < start_of_week:
                                continue
                        elif date_filter == "This Month":
                            if bill_date.month != today.month or bill_date.year != today.year:
                                continue
                        elif date_filter == "Last Month":
                            first_day_this_month = today.replace(day=1)
                            last_month = first_day_this_month - timedelta(days=1)
                            if bill_date.month != last_month.month or bill_date.year != last_month.year:
                                continue
                    except Exception:
                        continue
            
            # ✅ ADDED: Office Filter
            if office_filter != "All":
                bill_office = self.get_bill_office(bill_no)
                office_prefix_map = {"AP": "AP", "AFI": "AFI", "AFF": "AFF"}
                
                if office_filter == "Current":
                    # Show only bills from current office
                    current_prefix = self.get_office_prefix(self.selected_office)
                    if not bill_no.startswith(current_prefix):
                        continue
                else:
                    # Show bills from specific office
                    if not bill_no.startswith(office_filter):
                        continue
                    
            filtered_data[bill_no] = bill_info
        
        self.populate_bill_table(filtered_data)

    def edit_selected_bill(self):
        """Edit the selected bill with office validation"""
        selected = self.bill_table.selection()
        if not selected:
            self.show_status_message("⚠️ Please select a bill to edit", "warning")
            return
        
        bill_data = self.bill_table.item(selected[0], 'values')
        if bill_data:
            bill_no = bill_data[0]
            
            # ✅ ADDED: Office Validation
            bill_office = self.get_bill_office(bill_no)
            
            if bill_office != self.selected_office:
                # Show warning message
                office_names = {
                    "A1": "Angel Pyrotech",
                    "A2": "Angel Fireworks Industries", 
                    "A3": "Angel Fireworks Factory"
                }
                
                messagebox.showwarning(
                    "Office Access Restricted", 
                    f"This bill ({bill_no}) belongs to {office_names.get(bill_office, 'another office')}.\n\n"
                    f"You are currently logged into {office_names.get(self.selected_office, 'this office')}.\n\n"
                    f"🔒 Please switch to {office_names.get(bill_office, 'the correct office')} to edit this bill."
                )
                return
            
            # If office matches, proceed with editing
            self.create_new_bill(bill_no=bill_no)

    def on_double_click_edit(self, event):
        """Handle double-click in edit view with office validation"""
        selected = self.bill_table.selection()
        if selected:
            bill_data = self.bill_table.item(selected[0], 'values')
            if bill_data:
                bill_no = bill_data[0]
                
                # ✅ ADDED: Office Validation for double-click
                bill_office = self.get_bill_office(bill_no)
                
                if bill_office != self.selected_office:
                    office_names = {
                        "A1": "Angel Pyrotech",
                        "A2": "Angel Fireworks Industries", 
                        "A3": "Angel Fireworks Factory"
                    }
                    
                    messagebox.showwarning(
                        "Office Access Restricted", 
                        f"This bill ({bill_no}) belongs to {office_names.get(bill_office, 'another office')}.\n\n"
                        f"You are currently logged into {office_names.get(self.selected_office, 'this office')}.\n\n"
                        f"🔒 Please switch to {office_names.get(bill_office, 'the correct office')} to edit this bill."
                    )
                    return
                
                # If office matches, proceed with editing
                self.create_new_bill(bill_no=bill_no)

    def refresh_bill_list(self):
        """Refresh the bill list"""
        self.bills_ref = db.reference('bills')
        self.populate_bill_table()
        self.show_status_message("🔄 Bill list refreshed successfully")

    def on_bill_selection_change(self, event):
        """Handle bill selection change"""
        selected = self.bill_table.selection()
        if selected:
            bill_data = self.bill_table.item(selected[0], 'values')
            if bill_data:
                bill_no = bill_data[0]
                customer_name = bill_data[2]
                net_amount = bill_data[4]
                office = bill_data[6]  # ✅ ADDED: Get office from table
                
                # ✅ ADDED: Color code based on office access
                bill_office = self.get_bill_office(bill_no)
                if bill_office == self.selected_office:
                    status_color = self.colors['success']
                    access_text = "✅ You can edit this bill"
                else:
                    status_color = self.colors['warning'] 
                    access_text = "⚠️ Switch office to edit"
                
                self.selected_bill_info.config(
                    text=f"ℹ️  Selected: {bill_no} - {customer_name} | Amount: {net_amount} | Office: {office} | {access_text}",
                    fg=status_color
                )

    def show_bill_context_menu(self, event):
        """Show context menu for bill table with office validation"""
        selected = self.bill_table.selection()
        if not selected:
            return
            
        bill_data = self.bill_table.item(selected[0], 'values')
        if not bill_data:
            return
            
        bill_no = bill_data[0]
        bill_office = self.get_bill_office(bill_no)
        
        menu = tk.Menu(self.root, tearoff=0, font=("Segoe UI", 9))
        
        if bill_office == self.selected_office:
            # User can edit this bill
            menu.add_command(label="✏️ Edit Bill", command=self.edit_selected_bill)
            menu.add_separator()
            menu.add_command(label="📄 View PDF", command=self.view_selected_bill_pdf)
            menu.add_separator()
            menu.add_command(label="🗑️ Delete Bill", command=self.delete_selected_bill)
        else:
            # User cannot edit this bill
            office_names = {
                "A1": "Angel Pyrotech",
                "A2": "Angel Fireworks Industries", 
                "A3": "Angel Fireworks Factory"
            }
            menu.add_command(
                label=f"🔒 Bill from {office_names.get(bill_office, 'Another Office')}", 
                state="disabled"
            )
            menu.add_separator()
            menu.add_command(label="📄 View PDF", command=self.view_selected_bill_pdf)
        
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    # Rest of the methods remain the same...
    def view_selected_bill_pdf(self):
        """View PDF for selected bill"""
        selected = self.bill_table.selection()
        if not selected:
            self.show_status_message("⚠️ Please select a bill to view", "warning")
            return
        
        bill_data = self.bill_table.item(selected[0], 'values')
        if bill_data:
            bill_no = bill_data[0]
            self.open_bill_pdf(bill_no)

    def delete_selected_bill(self):
        """Delete the selected bill with office validation"""
        selected = self.bill_table.selection()
        if not selected:
            self.show_status_message("⚠️ Please select a bill to delete", "warning")
            return
        
        bill_data = self.bill_table.item(selected[0], 'values')
        if bill_data:
            bill_no = bill_data[0]
            customer_name = bill_data[2]
            
            # ✅ ADDED: Office Validation for deletion
            bill_office = self.get_bill_office(bill_no)
            if bill_office != self.selected_office:
                office_names = {
                    "A1": "Angel Pyrotech",
                    "A2": "Angel Fireworks Industries", 
                    "A3": "Angel Fireworks Factory"
                }
                messagebox.showwarning(
                    "Office Access Restricted", 
                    f"You cannot delete this bill as it belongs to {office_names.get(bill_office, 'another office')}.\n\n"
                    f"Switch to {office_names.get(bill_office, 'the correct office')} to delete this bill."
                )
                return
            
            # Confirmation dialog
            if messagebox.askyesno(
                "Confirm Deletion",
                f"Are you sure you want to delete bill {bill_no} for {customer_name}?\n\nThis action cannot be undone!",
                icon='warning'
            ):
                if bill_no in self.bills_data:
                    del self.bills_data[bill_no]
                    self.save_data(self.bills_ref, self.bills_data)
                    self.refresh_bill_list()
                    self.show_status_message(f"🗑️ Bill {bill_no} deleted successfully", "success")

    def open_bill_pdf(self, bill_no):
        """Open the PDF for a specific bill"""
        try:
            if bill_no in self.bills_data:
                bill_info = self.bills_data[bill_no]
                pdf_path = bill_info.get("pdf_file_name", "")
                
                if pdf_path:
                    # Convert relative path to absolute path
                    absolute_pdf_path = self.get_pdf_path(pdf_path)
                    
                    if os.path.exists(absolute_pdf_path):
                        self.display_pdf(absolute_pdf_path)
                    else:
                        messagebox.showerror("Error", f"PDF file not found:\n{absolute_pdf_path}")
                else:
                    messagebox.showerror("Error", f"No PDF path found for bill {bill_no}")
            else:
                messagebox.showerror("Error", f"Bill {bill_no} not found in database")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open PDF: {str(e)}")

    def show_view_bill(self):
        """Modern enhanced version of show_view_bill for viewing bills"""
        self.clear_screen()
        self.current_screen = "view_bill"
        self.create_navigation_bar()
        self.create_status_bar()

        # Reset focusable widgets
        self.focusable_widgets.clear()

        # Reload the bill details from the JSON file to get latest data
        self.bills_ref = db.reference('bills')

        # Main content area
        main_container = tk.Frame(self.root, bg=self.colors['light_bg'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Header frame
        header_frame = self.create_modern_frame(main_container, "👀 VIEW ALL BILLS")
        header_frame.pack(fill=tk.X, pady=(0, 20))

        # Search and filter section
        search_frame = tk.Frame(header_frame, bg=self.colors['card_bg'])
        search_frame.pack(fill=tk.X, padx=20, pady=15)

        # Search entry
        tk.Label(search_frame, text="🔍 Search Bills:", font=("Segoe UI", 10, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).pack(side=tk.LEFT, padx=5)
        
        self.view_bill_search_entry = self.create_modern_entry(search_frame, width=30)
        self.view_bill_search_entry.pack(side=tk.LEFT, padx=10, pady=5)
        self.view_bill_search_entry.bind('<KeyRelease>', self.filter_view_bill_list)

        # Filter by date frame
        date_filter_frame = tk.Frame(search_frame, bg=self.colors['card_bg'])
        date_filter_frame.pack(side=tk.LEFT, padx=20)

        tk.Label(date_filter_frame, text="📅 Filter by Date:", font=("Segoe UI", 10, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).pack(side=tk.LEFT, padx=5)
        
        self.view_date_filter_var = tk.StringVar(value="All")
        date_filter_combo = self.create_modern_combobox(
            date_filter_frame, 
            values=["All", "Today", "This Week", "This Month", "Last Month"],
            textvariable=self.view_date_filter_var,
            width=12
        )
        date_filter_combo.pack(side=tk.LEFT, padx=10, pady=5)
        date_filter_combo.bind('<<ComboboxSelected>>', self.filter_view_bill_list)

        # ✅ ADDED: Office Filter
        office_filter_frame = tk.Frame(search_frame, bg=self.colors['card_bg'])
        office_filter_frame.pack(side=tk.LEFT, padx=20)

        tk.Label(office_filter_frame, text="🏢 Office:", font=("Segoe UI", 10, "bold"),
                bg=self.colors['card_bg'], fg=self.colors['text_dark']).pack(side=tk.LEFT, padx=5)
        
        self.view_office_filter_var = tk.StringVar(value="Current")  # Default to current office only
        office_filter_combo = self.create_modern_combobox(
            office_filter_frame, 
            values=["Current", "All", "AP", "AFI", "AFF"],
            textvariable=self.view_office_filter_var,
            width=10
        )
        office_filter_combo.pack(side=tk.LEFT, padx=10, pady=5)
        office_filter_combo.bind('<<ComboboxSelected>>', self.filter_view_bill_list)

        # Action buttons
        action_frame = tk.Frame(search_frame, bg=self.colors['card_bg'])
        action_frame.pack(side=tk.RIGHT, padx=10)

        refresh_btn = self.create_modern_button(
            action_frame, "🔄 Refresh (F5)", self.refresh_view_bill_list,
            style="info", width=15, height=1
        )
        refresh_btn.pack(side=tk.LEFT, padx=5)

        # Quick stats cards
        stats_frame = tk.Frame(main_container, bg=self.colors['light_bg'])
        stats_frame.pack(fill=tk.X, pady=(0, 15))

        # Create quick stat cards
        self.view_stats_cards = {}
        stats_data = [
            ("📋 Total Bills", "total_bills", self.colors['info']),
            ("💰 Total Revenue", "total_revenue", self.colors['success']),
            ("📦 Total Items", "total_items", self.colors['warning']),
            ("✅ Paid Bills", "paid_bills", self.colors['primary'])
        ]

        for title, key, color in stats_data:
            card = self.create_view_stat_card(stats_frame, title, "0", color)
            card.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=5, pady=5)
            self.view_stats_cards[key] = card

        # Table container
        table_container = self.create_modern_frame(main_container, "📋 BILLS OVERVIEW - Double-click to View PDF")
        table_container.pack(fill=tk.BOTH, expand=True, pady=10)

        # Create a frame for table and scrollbars
        table_main = tk.Frame(table_container, bg=self.colors['card_bg'])
        table_main.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        # ✅ Create Treeview with modern styling
        style = ttk.Style()
        style.theme_use("default")

        # Fix invisible heading text layout
        style.layout("Modern.Treeview.Heading", [
            ("Treeheading.cell", {"sticky": "nswe"}),
            ("Treeheading.border", {"sticky": "nswe", "children": [
                ("Treeheading.padding", {"sticky": "nswe", "children": [
                    ("Treeheading.image", {"side": "right", "sticky": ""}),
                    ("Treeheading.text", {"sticky": "we"})
                ]})
            ]}),
        ])

        # Table row style
        style.configure("Modern.Treeview", 
            font=("Segoe UI", 10),
            rowheight=28,
            background=self.colors['card_bg'],
            fieldbackground=self.colors['card_bg'],
            foreground=self.colors['text_dark']
        )

        # Header style
        style.configure("Modern.Treeview.Heading", 
            font=("Segoe UI", 11, "bold"),
            background=self.colors['primary'],
            foreground=self.colors['text_light'],
            relief="flat"
        )

        # Hover effects for heading
        style.map("Modern.Treeview.Heading",
            background=[('active', self.colors['accent']), ('pressed', self.colors['secondary'])],
            relief=[('pressed', 'groove'), ('active', 'ridge')]
        )

        # ✅ Create Treeview with enhanced columns (REMOVED Status column)
        self.view_bill_table = ttk.Treeview(
            table_main, 
            columns=("Bill No", "Bill Date", "Customer Name", "Agent Name", 
                    "Net Amount", "Items Count", "Office"),  # REMOVED Status column
            show="headings",
            style="Modern.Treeview",
            selectmode="extended",
            height=12
        )

        # Define column headings with appropriate widths (REMOVED Status column)
        columns = {
            "Bill No": {"width": 100, "anchor": "center"},
            "Bill Date": {"width": 100, "anchor": "center"},
            "Customer Name": {"width": 180, "anchor": "w"},
            "Agent Name": {"width": 120, "anchor": "w"},
            "Net Amount": {"width": 110, "anchor": "center"},
            "Items Count": {"width": 80, "anchor": "center"},
            "Office": {"width": 80, "anchor": "center"}  # ✅ ADDED Office column
        }

        for col, settings in columns.items():
            self.view_bill_table.heading(col, text=col)
            self.view_bill_table.column(col, width=settings["width"], anchor=settings["anchor"])

        

        # Add scrollbars
        v_scrollbar = ttk.Scrollbar(table_main, orient="vertical", command=self.view_bill_table.yview)
        h_scrollbar = ttk.Scrollbar(table_main, orient="horizontal", command=self.view_bill_table.xview)
        self.view_bill_table.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        # Grid layout for table and scrollbars
        self.view_bill_table.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")

        # Configure grid weights
        table_main.grid_rowconfigure(0, weight=1)
        table_main.grid_columnconfigure(0, weight=1)

        # Status bar for table info
        status_frame = tk.Frame(main_container, bg=self.colors['light_bg'])
        status_frame.pack(fill=tk.X, pady=10)

        self.view_bill_table_status = tk.Label(
            status_frame,
            text="📊 Total Bills: 0",
            font=("Segoe UI", 10),
            bg=self.colors['light_bg'],
            fg=self.colors['text_dark']
        )
        self.view_bill_table_status.pack(side=tk.LEFT, padx=10)

        # ✅ ADDED: Current Office Info
        current_office_info = tk.Label(
            status_frame,
            text=f"🏢 Current Office: {self.get_current_office_name()}",
            font=("Segoe UI", 10, "bold"),
            bg=self.colors['light_bg'],
            fg=self.colors['primary']
        )
        current_office_info.pack(side=tk.LEFT, padx=20)

        # Summary information
        self.bill_summary_info = tk.Label(
            status_frame,
            text="💰 Total Amount: ₹0.00 | 📦 Total Items: 0",
            font=("Segoe UI", 10),
            bg=self.colors['light_bg'],
            fg=self.colors['success']
        )
        self.bill_summary_info.pack(side=tk.LEFT, padx=20)

        # Selected bill info
        self.selected_view_bill_info = tk.Label(
            status_frame,
            text="ℹ️  Select a bill to view details and actions",
            font=("Segoe UI", 9),
            bg=self.colors['light_bg'],
            fg=self.colors['text_muted']
        )
        self.selected_view_bill_info.pack(side=tk.RIGHT, padx=10)

        # Action buttons frame
        action_buttons_frame = tk.Frame(main_container, bg=self.colors['light_bg'])
        action_buttons_frame.pack(fill=tk.X, pady=10)

        # Center the buttons
        button_container = tk.Frame(action_buttons_frame, bg=self.colors['light_bg'])
        button_container.pack(expand=True)

        # View PDF button
        view_pdf_btn = self.create_modern_button(
            button_container,
            "📄 View PDF (Double-click)",
            self.view_selected_bill_pdf_view,
            style="primary",
            width=20,
            height=2
        )
        view_pdf_btn.pack(side=tk.LEFT, padx=8)

        # Edit bill button
        edit_btn = self.create_modern_button(
            button_container,
            "✏️ Edit Bill",
            self.edit_selected_view_bill,
            style="warning",
            width=15,
            height=2
        )
        edit_btn.pack(side=tk.LEFT, padx=8)

        # Print PDF button
        print_pdf_btn = self.create_modern_button(
            button_container,
            "🖨️ Print PDF",
            self.print_selected_bill_pdf,
            style="info",
            width=15,
            height=2
        )
        print_pdf_btn.pack(side=tk.LEFT, padx=8)

        # Back button
        back_btn = self.create_modern_button(
            button_container,
            "← Back to Billing (F4)",
            self.show_billing_dashboard,
            style="secondary",
            width=18,
            height=2
        )
        back_btn.pack(side=tk.LEFT, padx=8)

        # Populate the table
        self.populate_view_bill_table()

        # Bind keyboard shortcuts
        self.root.bind('<F5>', lambda e: self.refresh_view_bill_list())
        self.root.bind('<F4>', lambda e: self.show_billing_dashboard())
        self.root.bind('<Control-p>', lambda e: self.print_selected_bill_pdf())

        # Bind events to the table
        self.view_bill_table.bind('<Double-1>', self.on_double_click_view)
        self.view_bill_table.bind('<Return>', self.on_double_click_view)
        self.view_bill_table.bind('<<TreeviewSelect>>', self.on_view_bill_selection_change)
        self.view_bill_table.bind('<Button-3>', self.show_view_bill_context_menu)

        # Set focus to search field
        self.view_bill_search_entry.focus_set()

        self.show_status_message("👀 Bill viewer loaded - Double-click any bill to view PDF or use context menu")

    # ========== SUPPORTING METHODS FOR VIEW BILL ==========

    def create_view_stat_card(self, parent, title, value, color):
        """Create a stat card for view statistics"""
        card = tk.Frame(
            parent,
            bg=self.colors['card_bg'],
            relief="raised",
            bd=1,
            width=150,
            height=70
        )
        card.pack_propagate(False)

        # Card content
        content_frame = tk.Frame(card, bg=self.colors['card_bg'])
        content_frame.pack(expand=True, fill=tk.BOTH, padx=8, pady=8)

        # Title
        title_label = tk.Label(
            content_frame,
            text=title,
            font=("Segoe UI", 8, "bold"),
            bg=self.colors['card_bg'],
            fg=self.colors['text_muted']
        )
        title_label.pack(anchor="w")

        # Value
        value_label = tk.Label(
            content_frame,
            text=str(value),
            font=("Segoe UI", 12, "bold"),
            bg=self.colors['card_bg'],
            fg=color
        )
        value_label.pack(expand=True)

        # Store reference to update value later
        card.value_label = value_label

        return card

    def populate_view_bill_table(self, filtered_data=None):
        """Populate the view bill table with data"""
        try:
            # Clear existing data
            for item in self.view_bill_table.get_children():
                self.view_bill_table.delete(item)
            
            # Use provided data or all bills data
            bills_data = filtered_data if filtered_data is not None else self.bills_data
            
            # Calculate statistics
            total_bills = len(bills_data)
            total_revenue = 0
            total_items = 0
            paid_bills = 0
            
            # Insert data with status-based coloring
            for index, (bill_no, bill_info) in enumerate(bills_data.items()):
                # Extract bill information
                customer_name = bill_info.get('customer_name', 'N/A')
                agent_name = bill_info.get('agent_name', 'N/A')
                bill_date = bill_info.get('bill_date', 'N/A')
                net_amount = float(bill_info.get('net_amount', 0))
                items_count = len(bill_info.get('items', []))
                
                # ✅ ADDED: Detect office from bill number
                bill_office = self.get_bill_office(bill_no)
                office_display = {"A1": "AP", "A2": "AFI", "A3": "AFF"}.get(bill_office, "AP")
                
                # Update statistics
                total_revenue += 0
                total_items += items_count

                self.view_bill_table.insert("", "end", values=(
                    bill_no,
                    bill_date,
                    customer_name,
                    agent_name,
                    f"₹{net_amount:,.2f}",
                    items_count,
                    office_display  # ✅ ADDED Office column
                ))

            # Update status and statistics
            if hasattr(self, 'view_bill_table_status'):
                self.view_bill_table_status.config(text=f"📊 Displaying {total_bills} bills")
                
            if hasattr(self, 'bill_summary_info'):
                self.bill_summary_info.config(text=f"💰 Total Revenue: ₹{total_revenue:,.2f} | 📦 Total Items: {total_items:,}")
            
            # Update summary cards
            if hasattr(self, 'view_stats_cards'):
                if 'total_bills' in self.view_stats_cards:
                    self.view_stats_cards['total_bills'].value_label.config(text=f"{total_bills:,}")
                if 'total_revenue' in self.view_stats_cards:
                    self.view_stats_cards['total_revenue'].value_label.config(text=f"₹{total_revenue:,.2f}")
                if 'total_items' in self.view_stats_cards:
                    self.view_stats_cards['total_items'].value_label.config(text=f"{total_items:,}")
                if 'paid_bills' in self.view_stats_cards:
                    self.view_stats_cards['paid_bills'].value_label.config(text=f"{paid_bills:,}")

        except Exception as e:
            print(f"Error populating view bill table: {e}")

    def filter_view_bill_list(self, event=None):
        """Filter view bill list based on search criteria"""
        search_term = self.view_bill_search_entry.get().lower()
        date_filter = self.view_date_filter_var.get()
        office_filter = self.view_office_filter_var.get()
        
        filtered_data = {}
        
        for bill_no, bill_info in self.bills_data.items():
            # Search term filter
            matches_search = (
                search_term in bill_no.lower() or
                search_term in bill_info.get('customer_name', '').lower() or
                search_term in bill_info.get('agent_name', '').lower()
            )
            
            if not matches_search:
                continue
                
            # Date filter
            if date_filter != "All":
                bill_date_str = bill_info.get('bill_date', '')
                if bill_date_str:
                    try:
                        bill_date = parse_date_flexible(bill_date_str)
                        today = datetime.now()
                        
                        if date_filter == "Today":
                            if bill_date.date() != today.date():
                                continue
                        elif date_filter == "This Week":
                            start_of_week = today - timedelta(days=today.weekday())
                            if bill_date < start_of_week:
                                continue
                        elif date_filter == "This Month":
                            if bill_date.month != today.month or bill_date.year != today.year:
                                continue
                        elif date_filter == "Last Month":
                            first_day_this_month = today.replace(day=1)
                            last_month = first_day_this_month - timedelta(days=1)
                            if bill_date.month != last_month.month or bill_date.year != last_month.year:
                                continue
                    except Exception:
                        continue
            
            # ✅ ADDED: Office Filter
            if office_filter != "All":
                bill_office = self.get_bill_office(bill_no)
                
                if office_filter == "Current":
                    # Show only bills from current office
                    current_prefix = self.get_office_prefix(self.selected_office)
                    if not bill_no.startswith(current_prefix):
                        continue
                else:
                    # Show bills from specific office
                    if not bill_no.startswith(office_filter):
                        continue
                    
            filtered_data[bill_no] = bill_info
        
        self.populate_view_bill_table(filtered_data)

    def refresh_view_bill_list(self):
        """Refresh the view bill list"""
        self.bills_ref = db.reference('bills')
        self.populate_view_bill_table()
        self.show_status_message("🔄 Bill view refreshed successfully")

    def on_view_bill_selection_change(self, event):
        """Handle bill selection change in view view"""
        selected = self.view_bill_table.selection()
        if selected:
            bill_data = self.view_bill_table.item(selected[0], 'values')
            if bill_data:
                bill_no = bill_data[0]
                customer_name = bill_data[2]
                net_amount = bill_data[4]
                items_count = bill_data[5]  # Items count at index 5
                office = bill_data[6]  # Office at index 6 (NOT 7)
                
                # ✅ ADDED: Color code based on office access
                bill_office = self.get_bill_office(bill_no)
                if bill_office == self.selected_office:
                    status_color = self.colors['success']
                    access_text = "✅ You can edit this bill"
                else:
                    status_color = self.colors['warning'] 
                    access_text = "⚠️ Switch office to edit"
                
                self.selected_view_bill_info.config(
                    text=f"ℹ️  Selected: {bill_no} - {customer_name} | Amount: {net_amount} | Items: {items_count} | Office: {office} | {access_text}",
                    fg=status_color
                )

    def show_view_bill_context_menu(self, event):
        """Show context menu for view bill table with office validation"""
        selected = self.view_bill_table.selection()
        if not selected:
            return
            
        bill_data = self.view_bill_table.item(selected[0], 'values')
        if not bill_data:
            return
            
        bill_no = bill_data[0]
        bill_office = self.get_bill_office(bill_no)
        
        menu = tk.Menu(self.root, tearoff=0, font=("Segoe UI", 9))
        
        menu.add_command(label="📄 View PDF", command=self.view_selected_bill_pdf_view)
        menu.add_command(label="🖨️ Print PDF", command=self.print_selected_bill_pdf)
        menu.add_separator()
        
        if bill_office == self.selected_office:
            # User can edit this bill
            menu.add_command(label="✏️ Edit Bill", command=self.edit_selected_view_bill)
        else:
            # User cannot edit this bill
            office_names = {
                "A1": "Angel Pyrotech",
                "A2": "Angel Fireworks Industries", 
                "A3": "Angel Fireworks Factory"
            }
            menu.add_command(
                label=f"🔒 Edit (Switch to {office_names.get(bill_office, 'Another Office')})", 
                state="disabled"
            )
        
        menu.add_separator()
        menu.add_command(label="📋 Copy Bill Number", command=self.copy_bill_number)
        
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def on_double_click_view(self, event):
        """Handle double-click in view view"""
        self.view_selected_bill_pdf_view()

    def view_selected_bill_pdf_view(self):
        """View PDF for selected bill in view view"""
        selected = self.view_bill_table.selection()
        if not selected:
            self.show_status_message("⚠️ Please select a bill to view", "warning")
            return
        
        bill_data = self.view_bill_table.item(selected[0], 'values')
        if bill_data:
            bill_no = bill_data[0]
            self.open_bill_pdf(bill_no)

    def edit_selected_view_bill(self):
        """Edit the selected bill from view screen with office validation"""
        selected = self.view_bill_table.selection()
        if not selected:
            self.show_status_message("⚠️ Please select a bill to edit", "warning")
            return
        
        bill_data = self.view_bill_table.item(selected[0], 'values')
        if bill_data:
            bill_no = bill_data[0]
            
            # ✅ ADDED: Office Validation
            bill_office = self.get_bill_office(bill_no)
            
            if bill_office != self.selected_office:
                # Show warning message
                office_names = {
                    "A1": "Angel Pyrotech",
                    "A2": "Angel Fireworks Industries", 
                    "A3": "Angel Fireworks Factory"
                }
                
                messagebox.showwarning(
                    "Office Access Restricted", 
                    f"This bill ({bill_no}) belongs to {office_names.get(bill_office, 'another office')}.\n\n"
                    f"You are currently logged into {office_names.get(self.selected_office, 'this office')}.\n\n"
                    f"🔒 Please switch to {office_names.get(bill_office, 'the correct office')} to edit this bill."
                )
                return
            
            # If office matches, proceed with editing
            self.create_new_bill(bill_no=bill_no)

    def print_selected_bill_pdf(self):
        """Print the selected bill PDF"""
        selected = self.view_bill_table.selection()
        if not selected:
            self.show_status_message("⚠️ Please select a bill to print", "warning")
            return
        
        bill_data = self.view_bill_table.item(selected[0], 'values')
        if bill_data:
            bill_no = bill_data[0]
            # Simple implementation - just open the PDF which user can then print
            self.open_bill_pdf(bill_no)
            self.show_status_message("🖨️ PDF opened - Use your PDF viewer's print function")

    def copy_bill_number(self):
        """Copy bill number to clipboard"""
        selected = self.view_bill_table.selection()
        if selected:
            bill_data = self.view_bill_table.item(selected[0], 'values')
            if bill_data:
                bill_no = bill_data[0]
                self.root.clipboard_clear()
                self.root.clipboard_append(bill_no)
                self.show_status_message(f"📋 Bill number {bill_no} copied to clipboard")

    def show_statement_options(self):
        """Show statement options with modern design"""
        self.clear_screen()
        self.current_screen = "statement_options"
        self.create_navigation_bar()
        self.create_status_bar()

        # Reset focusable widgets
        self.focusable_widgets.clear()

        # Main content area
        main_container = tk.Frame(self.root, bg=self.colors['light_bg'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Center container
        center_container = tk.Frame(main_container, bg=self.colors['light_bg'])
        center_container.place(relx=0.5, rely=0.5, anchor="center", width=600, height=400)

        # Header frame
        header_frame = self.create_modern_frame(center_container, "📊 STATEMENT GENERATION")
        header_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))

        # Options container
        options_frame = tk.Frame(header_frame, bg=self.colors['card_bg'])
        options_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)

        # Title
        title_label = tk.Label(
            options_frame,
            text="Select Statement Type",
            font=("Segoe UI", 16, "bold"),
            bg=self.colors['card_bg'],
            fg=self.colors['primary'],
            pady=20
        )
        title_label.pack()

        # Statement options
        statement_options = [
            {
                "text": "👥 Customer Statement",
                "description": "Generate customer-wise statement with all invoices",
                "command": self.show_customer_statement_page,
                "style": "primary"
            },
            {
                "text": "🤵 Agent Commission Statement", 
                "description": "Generate agent-wise commission statement",
                "command": self.show_agent_commission_page,
                "style": "info"
            },
            {
                "text": "📅 Date Range Statement",
                "description": "Generate statement for specific date range",
                "command": self.show_date_range_statement,
                "style": "success"
            }
        ]

        for option in statement_options:
            # Option frame
            option_frame = tk.Frame(options_frame, bg=self.colors['card_bg'], pady=10)
            option_frame.pack(fill=tk.X, pady=8)

            # Button
            btn = self.create_modern_button(
                option_frame,
                option["text"],
                option["command"],
                style=option["style"],
                width=25,
                height=2
            )
            btn.pack(side=tk.LEFT, padx=(0, 15))

            # Description
            desc_label = tk.Label(
                option_frame,
                text=option["description"],
                font=("Segoe UI", 10),
                bg=self.colors['card_bg'],
                fg=self.colors['text_muted'],
                justify=tk.LEFT
            )
            desc_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Back button
        back_btn = self.create_modern_button(
            center_container,
            "← Back to Billing (F4)",
            self.show_billing_dashboard,
            style="secondary",
            width=20,
            height=2
        )
        back_btn.pack(pady=20)

        self.show_status_message("📊 Select a statement type to generate reports")

    def show_customer_statement_page(self):
        """Modern customer statement generation page"""
        self.clear_screen()
        self.current_screen = "customer_statement"
        self.create_navigation_bar()
        self.create_status_bar()

        # Reset focusable widgets
        self.focusable_widgets.clear()

        # Main content area
        main_container = tk.Frame(self.root, bg=self.colors['light_bg'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Header frame
        header_frame = self.create_modern_frame(main_container, "👥 CUSTOMER STATEMENT GENERATION")
        header_frame.pack(fill=tk.X, pady=(0, 15))

        # Search and filter section
        search_frame = tk.Frame(header_frame, bg=self.colors['card_bg'])
        search_frame.pack(fill=tk.X, padx=20, pady=15)

        # --- Customer Selection Section ---
        customer_select_frame = tk.Frame(search_frame, bg=self.colors['card_bg'])
        customer_select_frame.pack(fill=tk.X, pady=10)

        tk.Label(
            customer_select_frame,
            text="👤 Select Customer:",
            font=("Segoe UI", 11, "bold"),
            bg=self.colors['card_bg'],
            fg=self.colors['text_dark']
        ).pack(side=tk.LEFT, padx=(0, 10))

        # Collect unique customer names from bills
        customer_names = set()
        for bill_no, bill in self.bills_data.items():
            if "customer_name" in bill and bill["customer_name"]:
                customer_names.add(bill["customer_name"].strip())

        # Store sorted list for global reuse
        self.all_customer_names = sorted(list(customer_names))

        # Create customer combobox
        self.customer_combobox = self.create_modern_combobox(
            customer_select_frame,
            values=self.all_customer_names,
            width=35,
            font_size=11
        )
        self.customer_combobox.pack(side=tk.LEFT, padx=5)

        # 🔥 Bind dynamic suggestion update
        self.customer_combobox.bind("<KeyRelease>", self.update_customer_suggestions)

        # Load bills button
        load_bills_btn = self.create_modern_button(
            customer_select_frame,
            "📥 Load Customer Bills",
            self.load_customer_bills_for_statement,
            style="info",
            width=18,
            height=1
        )
        load_bills_btn.pack(side=tk.LEFT, padx=10)


        # Date range filter
        date_frame = tk.Frame(search_frame, bg=self.colors['card_bg'])
        date_frame.pack(fill=tk.X, pady=10)

        tk.Label(
            date_frame,
            text="📅 Filter by Date Range:",
            font=("Segoe UI", 10, "bold"),
            bg=self.colors['card_bg'],
            fg=self.colors['text_dark']
        ).pack(side=tk.LEFT, padx=(0, 10))

        # From date
        tk.Label(date_frame, text="From:", font=("Segoe UI", 9),
                bg=self.colors['card_bg']).pack(side=tk.LEFT, padx=(0, 5))
        
        self.statement_from_date = self.create_modern_entry(date_frame, width=12, font_size=9)
        self.statement_from_date.pack(side=tk.LEFT, padx=5)
        self.statement_from_date.insert(0, "01/01/2024")

        # To date
        tk.Label(date_frame, text="To:", font=("Segoe UI", 9),
                bg=self.colors['card_bg']).pack(side=tk.LEFT, padx=(10, 5))
        
        self.statement_to_date = self.create_modern_entry(date_frame, width=12, font_size=9)
        self.statement_to_date.pack(side=tk.LEFT, padx=5)
        self.statement_to_date.insert(0, datetime.now().strftime("%d/%m/%Y"))

        # Apply date filter button
        apply_date_btn = self.create_modern_button(
            date_frame,
            "🔍 Apply Date Filter",
            self.apply_date_filter_customer,
            style="primary",
            width=16,
            height=1
        )
        apply_date_btn.pack(side=tk.LEFT, padx=15)

        # Clear filter button
        clear_filter_btn = self.create_modern_button(
            date_frame,
            "🔄 Clear Filter",
            self.clear_customer_filter,
            style="secondary",
            width=12,
            height=1
        )
        clear_filter_btn.pack(side=tk.LEFT, padx=5)

        # Bills table container
        table_container = self.create_modern_frame(main_container, "📋 CUSTOMER INVOICES - SELECT BILLS TO INCLUDE")
        table_container.pack(fill=tk.BOTH, expand=True, pady=10)

        # Create a frame for table and scrollbars
        table_main = tk.Frame(table_container, bg=self.colors['card_bg'])
        table_main.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        # Create Treeview with checkboxes
        style = ttk.Style()
        style.theme_use("default")

        # Configure modern style
        style.configure("Modern.Treeview", 
            font=("Segoe UI", 10),
            rowheight=28,
            background=self.colors['card_bg'],
            fieldbackground=self.colors['card_bg'],
            foreground=self.colors['text_dark']
        )

        style.configure("Modern.Treeview.Heading", 
            font=("Segoe UI", 11, "bold"),
            background=self.colors['primary'],
            foreground=self.colors['text_light'],
            relief="flat"
        )

        # Create Treeview with checkbox as first column
        self.customer_bills_table = ttk.Treeview(
            table_main, 
            columns=("Select", "Bill No", "Bill Date", "Customer Name", "Agent", "Amount", "Status", "PDF Available"), 
            show="headings",
            style="Modern.Treeview",
            selectmode="extended",
            height=12
        )

        # Define column headings
        columns = {
            "Select": {"width": 60, "anchor": "center"},
            "Bill No": {"width": 90, "anchor": "center"},
            "Bill Date": {"width": 90, "anchor": "center"},
            "Customer Name": {"width": 180, "anchor": "w"},
            "Agent": {"width": 120, "anchor": "w"},
            "Amount": {"width": 100, "anchor": "center"},
            "Status": {"width": 100, "anchor": "center"},
            "PDF Available": {"width": 100, "anchor": "center"}
        }

        for col, settings in columns.items():
            self.customer_bills_table.heading(col, text=col)
            self.customer_bills_table.column(col, width=settings["width"], anchor=settings["anchor"])

        # Add scrollbars
        v_scrollbar = ttk.Scrollbar(table_main, orient="vertical", command=self.customer_bills_table.yview)
        h_scrollbar = ttk.Scrollbar(table_main, orient="horizontal", command=self.customer_bills_table.xview)
        self.customer_bills_table.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        # Grid layout for table and scrollbars
        self.customer_bills_table.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")

        # Configure grid weights
        table_main.grid_rowconfigure(0, weight=1)
        table_main.grid_columnconfigure(0, weight=1)

        # Selection controls frame
        selection_frame = tk.Frame(main_container, bg=self.colors['light_bg'])
        selection_frame.pack(fill=tk.X, pady=10)

        # Left side - Selection controls
        left_selection_frame = tk.Frame(selection_frame, bg=self.colors['light_bg'])
        left_selection_frame.pack(side=tk.LEFT)

        # Select all checkbox
        self.select_all_customer = tk.BooleanVar()
        select_all_cb = tk.Checkbutton(
            left_selection_frame,
            text="Select All Bills",
            variable=self.select_all_customer,
            font=("Segoe UI", 10, "bold"),
            bg=self.colors['light_bg'],
            command=self.toggle_select_all_customer
        )
        select_all_cb.pack(anchor="w", pady=5)

        # Selection info
        self.customer_selection_info = tk.Label(
            left_selection_frame,
            text="Selected: 0 bills | Total Amount: ₹0.00",
            font=("Segoe UI", 10),
            bg=self.colors['light_bg'],
            fg=self.colors['success']
        )
        self.customer_selection_info.pack(anchor="w", pady=2)

        # Right side - Action buttons
        right_selection_frame = tk.Frame(selection_frame, bg=self.colors['light_bg'])
        right_selection_frame.pack(side=tk.RIGHT)

        # Generate statement button
        generate_btn = self.create_modern_button(
            right_selection_frame,
            "🚀 Generate Party Statement",
            self.generate_customer_statement,
            style="success",
            width=22,
            height=2
        )
        generate_btn.pack(side=tk.LEFT, padx=5)

        # Preview selected button
        preview_btn = self.create_modern_button(
            right_selection_frame,
            "👀 Preview Selected",
            self.preview_selected_customer_bills,
            style="info",
            width=16,
            height=2
        )
        preview_btn.pack(side=tk.LEFT, padx=5)

        # Clear selection button
        clear_btn = self.create_modern_button(
            right_selection_frame,
            "🗑️ Clear Selection",
            self.show_customer_statement_page,
            style="warning",
            width=14,
            height=2
        )
        clear_btn.pack(side=tk.LEFT, padx=5)

        # Back button
        back_btn = self.create_modern_button(
            main_container,
            "← Back to Statements",
            self.show_statement_options,
            style="secondary",
            width=16,
            height=2
        )
        back_btn.pack(pady=10)

        # Initialize customer bills data
        self.customer_bills_data = []
        self.selected_customer_bills = set()

        # Bind keyboard shortcuts
        self.root.bind('<Control-Return>', lambda e: self.generate_customer_statement())
        self.root.bind('<Control-a>', lambda e: self.toggle_select_all_customer())

        # Bind table events
        self.customer_bills_table.bind('<Button-1>', self.on_customer_bill_click)

        # Set focus to customer combobox
        self.customer_combobox.focus_set()

        self.show_status_message("👥 Select a customer and choose bills for statement generation")

    def update_customer_suggestions(self, event):
        """Update suggestions while typing but only show dropdown on Down arrow press."""
        try:
            # If Down arrow is pressed, show filtered dropdown
            if event.keysym == 'Down':
                typed = self.customer_combobox.get().strip().lower()
                
                # Generate filtered suggestions
                if not typed:
                    suggestions = self.customer_names
                else:
                    suggestions = [name for name in self.customer_names if typed.lower() in name.lower()]
                
                # Update combobox dropdown values and open it
                self.customer_combobox['values'] = suggestions
                if suggestions:
                    self.customer_combobox.event_generate('<Down>')
                return
            
            # For regular typing, just update the values without opening dropdown
            elif event.keysym not in ['Return', 'Up', 'Escape', 'Tab', 'Shift_L', 'Shift_R', 'Control_L', 'Control_R']:
                typed = self.customer_combobox.get().strip().lower()
                
                # Generate filtered suggestions
                if not typed:
                    suggestions = self.customer_names
                else:
                    suggestions = [name for name in self.customer_names if typed.lower() in name.lower()]
                
                # Only update the values, don't open dropdown
                self.customer_combobox['values'] = suggestions

        except Exception as e:
            print(f"Error updating customer suggestions: {e}")





    def generate_agent_statement(self):
        """Generate a simple Agent Statement (sales summary without commission)"""
        if not self.selected_agent_bills:
            messagebox.showwarning("Selection Required", "Please select at least one bill to generate statement.")
            return

        selected_bills = []
        total_sales = 0

        for bill in self.agent_bills_data:
            if bill['bill_no'] in self.selected_agent_bills:
                selected_bills.append(bill)
                total_sales += bill['amount']

        agent_name = self.agent_name_combobox.get().strip()

        from fpdf import FPDF
        pdf = FPDF()
        pdf.add_page()
        pdf.set_left_margin(15)
        pdf.set_right_margin(15)
        pdf.set_font("Arial", "", 11)

        # Header
        pdf.cell(0, 10, "AGENT STATEMENT (Sales Summary)", ln=True, align="C")
        pdf.ln(5)

        pdf.set_font("Arial", "B", 11)
        pdf.cell(0, 8, f"Agent: {agent_name}", ln=True)
        pdf.cell(0, 8, f"Statement Date: {datetime.now().strftime('%d/%m/%Y')}", ln=True)
        pdf.ln(5)

        # Summary
        pdf.set_font("Arial", "", 10)
        pdf.cell(0, 6, f"Total Bills: {len(selected_bills)}", ln=True)
        pdf.cell(0, 6, f"Total Sales: Rs.{total_sales:,.2f}", ln=True)
        pdf.ln(10)

        # Table Header
        headers = ["Bill No", "Date", "Customer", "Amount"]
        col_widths = [30, 25, 90, 35]

        pdf.set_font("Arial", "B", 9)
        for i, h in enumerate(headers):
            pdf.cell(col_widths[i], 8, h, border=1, align='C')
        pdf.ln()

        # Table Rows
        pdf.set_font("Arial", "", 9)
        for bill in selected_bills:
            pdf.cell(col_widths[0], 6, bill['bill_no'], border=1)
            pdf.cell(col_widths[1], 6, bill['bill_date'], border=1)
            pdf.cell(col_widths[2], 6, bill['customer_name'][:40], border=1)
            pdf.cell(col_widths[3], 6, f"Rs.{bill['amount']:,.2f}", border=1, align='R')
            pdf.ln()

        # Total
        pdf.set_font("Arial", "B", 9)
        pdf.cell(col_widths[0] + col_widths[1] + col_widths[2], 8, "TOTAL", border=1, align='C')
        pdf.cell(col_widths[3], 8, f"Rs.{total_sales:,.2f}", border=1, align='R')
        pdf.ln(10)

        pdf.set_font("Arial", "I", 8)
        pdf.cell(0, 6, "This is a computer-generated statement.", ln=True)

        # Save
        # 🆕 FIXED: Save in year-based folder
        current_year = str(datetime.now().year)
        documents_dir = os.path.join(os.path.expanduser("~"), "Documents")
        statements_dir = os.path.join(
            documents_dir, 
            "InvoiceApp", 
            f"Invoice_Bill_{current_year}",  # Year-based folder
            "STATEMENTS_FOLDER", 
            "Agent_Statement_Only"
        )
        os.makedirs(statements_dir, exist_ok=True)
        clean_agent = re.sub(r'[^\w\s-]', '', agent_name).strip().replace(' ', '_')
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(statements_dir, f"{clean_agent}_Statement_{timestamp}.pdf")

        pdf.output(output_file)

        # Show success
        messagebox.showinfo(
            "Success",
            f"Agent Statement Generated Successfully!\n\n"
            f"Agent: {agent_name}\n"
            f"Bills Included: {len(selected_bills)}\n"
            f"Total Sales: Rs.{total_sales:,.2f}\n\n"
            f"Saved to:\n{output_file}"
        )

        self.display_pdf(output_file)


    def load_customer_bills_for_statement(self):
        """Load bills for selected customer"""
        customer_name = self.customer_combobox.get().strip()

        if not customer_name:
            messagebox.showwarning("⚠️ Input Required", "Please select a customer name first.")
            return

        # Clear previous data
        for item in self.customer_bills_table.get_children():
            self.customer_bills_table.delete(item)
        
        self.customer_bills_data.clear()
        self.selected_customer_bills.clear()
        self.select_all_customer.set(False)

        # Show loading
        self.show_status_message(f"🔍 Loading bills for {customer_name}...")

        # Collect matching bills
        matching_bills = []
        total_amount = 0

        for bill_no, bill in self.bills_data.items():
            bill_customer_name = bill.get("customer_name", "").strip()
            
            # Check customer name match (case insensitive partial match)
            if customer_name.lower() not in bill_customer_name.lower():
                continue

            # Check if PDF exists
            pdf_available = "❌ No"
            pdf_path = bill.get("pdf_file_name", "")
            if pdf_path:
                # Resolve PDF path
                resolved_path = self.resolve_pdf_path_updated(pdf_path, bill_no)
                if resolved_path and os.path.exists(resolved_path):
                    pdf_available = "✅ Yes"

            amount = float(bill.get("net_amount", 0))
            total_amount += amount

            matching_bills.append({
                'bill_no': bill_no,
                'bill_date': bill.get("bill_date", ""),
                'customer_name': bill_customer_name,
                'agent_name': bill.get("agent_name", ""),
                'amount': amount,
                'status': bill.get("payment_status", "Pending"),
                'pdf_available': pdf_available,
                'pdf_path': pdf_path
            })

        if not matching_bills:
            messagebox.showinfo("ℹ️ No Bills Found", f"No bills found for customer: {customer_name}")
            self.show_status_message("❌ No bills found for selected customer")
            return

        # Sort by bill date (newest first)
        matching_bills.sort(key=lambda x: (parse_date_flexible(x.get('bill_date')) or datetime.min), reverse=True)

        # Populate table
        for bill in matching_bills:
            item_id = self.customer_bills_table.insert("", "end", values=(
                "☐",  # Checkbox placeholder
                bill['bill_no'],
                bill['bill_date'],
                bill['customer_name'],
                bill['agent_name'],
                f"₹{bill['amount']:,.2f}",
                bill['status'],
                bill['pdf_available']
            ))
            
            # Store bill data with treeview item ID
            bill['item_id'] = item_id
            self.customer_bills_data.append(bill)

        # Update selection info
        self.update_customer_selection_info()

        self.show_status_message(f"✅ Loaded {len(matching_bills)} bills for {customer_name} - Total: ₹{total_amount:,.2f}")

    def on_customer_bill_click(self, event):
        """Handle checkbox clicks in customer bills table"""
        item = self.customer_bills_table.identify_row(event.y)
        column = self.customer_bills_table.identify_column(event.x)

        if item and column == "#1":  # Checkbox column
            bill_no = self.customer_bills_table.item(item, "values")[1]
            
            if bill_no in self.selected_customer_bills:
                self.selected_customer_bills.remove(bill_no)
                self.customer_bills_table.set(item, "Select", "☐")
            else:
                self.selected_customer_bills.add(bill_no)
                self.customer_bills_table.set(item, "Select", "☑")
            
            self.update_customer_selection_info()

    def toggle_select_all_customer(self):
        """Toggle select all bills for customer"""
        if self.select_all_customer.get():
            # Select all
            self.selected_customer_bills.clear()
            for bill in self.customer_bills_data:
                self.selected_customer_bills.add(bill['bill_no'])
                self.customer_bills_table.set(bill['item_id'], "Select", "☑")
        else:
            # Deselect all
            self.selected_customer_bills.clear()
            for bill in self.customer_bills_data:
                self.customer_bills_table.set(bill['item_id'], "Select", "☐")
        
        self.update_customer_selection_info()

    def update_customer_selection_info(self):
        """Update selection information display"""
        selected_count = len(self.selected_customer_bills)
        total_amount = 0
        
        for bill in self.customer_bills_data:
            if bill['bill_no'] in self.selected_customer_bills:
                total_amount += bill['amount']
        
        self.customer_selection_info.config(
            text=f"Selected: {selected_count} bills | Total Amount: ₹{total_amount:,.2f}"
        )
        
        # Update select all checkbox state
        total_bills = len(self.customer_bills_data)
        if selected_count == total_bills and total_bills > 0:
            self.select_all_customer.set(True)
        else:
            self.select_all_customer.set(False)

    def apply_date_filter_customer(self):
        """Apply date filter to customer bills"""
        try:
            from_date = parse_date_flexible(self.statement_from_date.get())
            to_date   = parse_date_flexible(self.statement_to_date.get())

            if not from_date or not to_date:
                messagebox.showerror("❌ Error", "Please enter valid dates (e.g. DD/MM/YYYY or DD.MM.YYYY).")
                return

        except ValueError:
            messagebox.showerror("❌ Error", "Please enter valid dates in DD/MM/YYYY format.")
            return


        customer_name = self.customer_combobox.get().strip()
        if not customer_name:
            messagebox.showwarning("⚠️ Input Required", "Please select a customer first.")
            return

        # Reload bills with date filter
        self.load_customer_bills_with_date_filter(customer_name, from_date, to_date)

    def load_customer_bills_with_date_filter(self, customer_name, from_date, to_date):
        """Load customer bills with date filter"""
        # Clear previous data
        for item in self.customer_bills_table.get_children():
            self.customer_bills_table.delete(item)
        
        self.customer_bills_data.clear()
        self.selected_customer_bills.clear()
        self.select_all_customer.set(False)

        # Collect filtered bills
        filtered_bills = []
        total_amount = 0

        for bill_no, bill in self.bills_data.items():
            bill_customer_name = bill.get("customer_name", "").strip()
            
            # Check customer name match
            if customer_name.lower() not in bill_customer_name.lower():
                continue

            # Check date range
            try:
                bill_date = parse_date_flexible(bill.get("bill_date", ""))
                if not bill_date:
                    # skip bills with unparsable dates when filtering by date range
                    continue

                if not (from_date <= bill_date <= to_date):
                    continue
            except ValueError:
                continue

            # Check if PDF exists
            pdf_available = "❌ No"
            pdf_path = bill.get("pdf_file_name", "")
            if pdf_path:
                resolved_path = self.resolve_pdf_path_updated(pdf_path, bill_no)
                if resolved_path and os.path.exists(resolved_path):
                    pdf_available = "✅ Yes"

            amount = float(bill.get("net_amount", 0))
            total_amount += amount

            filtered_bills.append({
                'bill_no': bill_no,
                'bill_date': bill.get("bill_date", ""),
                'customer_name': bill_customer_name,
                'agent_name': bill.get("agent_name", ""),
                'amount': amount,
                'status': bill.get("payment_status", "Pending"),
                'pdf_available': pdf_available,
                'pdf_path': pdf_path
            })

        if not filtered_bills:
            messagebox.showinfo("ℹ️ No Bills Found", f"No bills found for {customer_name} in the selected date range.")
            self.show_status_message("❌ No bills found for selected criteria")
            return

        # Sort by bill date (newest first)
        filtered_bills.sort(key=lambda x: datetime.strptime(x['bill_date'], "%d/%m/%Y"), reverse=True)

        # Populate table
        for bill in filtered_bills:
            item_id = self.customer_bills_table.insert("", "end", values=(
                "☐",
                bill['bill_no'],
                bill['bill_date'],
                bill['customer_name'],
                bill['agent_name'],
                f"₹{bill['amount']:,.2f}",
                bill['status'],
                bill['pdf_available']
            ))
            
            bill['item_id'] = item_id
            self.customer_bills_data.append(bill)

        self.update_customer_selection_info()
        self.show_status_message(f"✅ Found {len(filtered_bills)} bills - Total: ₹{total_amount:,.2f}")

    def clear_customer_filter(self):
        """Clear date filter and reload all bills"""
        self.statement_from_date.delete(0, tk.END)
        self.statement_from_date.insert(0, "01/01/2024")
        self.statement_to_date.delete(0, tk.END)
        self.statement_to_date.insert(0, datetime.now().strftime("%d/%m/%Y"))
        
        customer_name = self.customer_combobox.get().strip()
        if customer_name:
            self.load_customer_bills_for_statement()

    def clear_customer_selection(self):
        """Clear all selections"""
        self.selected_customer_bills.clear()
        self.select_all_customer.set(False)
        
        for bill in self.customer_bills_data:
            self.customer_bills_table.set(bill['item_id'], "Select", "☐")
        
        self.update_customer_selection_info()

    def preview_selected_customer_bills(self):
        """Preview selected customer bills"""
        if not self.selected_customer_bills:
            messagebox.showwarning("⚠️ Selection Required", "Please select at least one bill to preview.")
            return

        selected_bills = []
        total_amount = 0

        for bill in self.customer_bills_data:
            if bill['bill_no'] in self.selected_customer_bills:
                selected_bills.append(bill)
                total_amount += bill['amount']

        # Create preview window
        preview_window = tk.Toplevel(self.root)
        preview_window.title(f"Preview - {len(selected_bills)} Selected Bills")
        preview_window.geometry("500x400")
        preview_window.configure(bg=self.colors['light_bg'])
        preview_window.transient(self.root)
        preview_window.grab_set()

        # Center the window
        preview_window.update_idletasks()
        x = (preview_window.winfo_screenwidth() // 2) - (500 // 2)
        y = (preview_window.winfo_screenheight() // 2) - (400 // 2)
        preview_window.geometry(f"500x400+{x}+{y}")

        # Header
        header_frame = tk.Frame(preview_window, bg=self.colors['primary'], height=50)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)

        title_label = tk.Label(
            header_frame,
            text=f"Preview - {len(selected_bills)} Selected Bills",
            font=("Segoe UI", 14, "bold"),
            bg=self.colors['primary'],
            fg=self.colors['text_light'],
            pady=15
        )
        title_label.pack()

        # Content
        content_frame = tk.Frame(preview_window, bg=self.colors['light_bg'], padx=20, pady=20)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # Summary
        summary_text = f"""
        📊 SELECTION SUMMARY:
        
        • Total Bills Selected: {len(selected_bills)}
        • Total Amount: ₹{total_amount:,.2f}
        • PDF Available: {sum(1 for b in selected_bills if b['pdf_available'] == '✅ Yes')} bills
        
        📋 Selected Bills:
        """

        # Add bill list
        for i, bill in enumerate(selected_bills[:10], 1):  # Show first 10
            summary_text += f"\n    {i}. {bill['bill_no']} - {bill['bill_date']} - ₹{bill['amount']:,.2f}"

        if len(selected_bills) > 10:
            summary_text += f"\n    ... and {len(selected_bills) - 10} more bills"

        summary_label = tk.Label(
            content_frame,
            text=summary_text.strip(),
            font=("Segoe UI", 10),
            bg=self.colors['light_bg'],
            fg=self.colors['text_dark'],
            justify=tk.LEFT
        )
        summary_label.pack(anchor="w", pady=10)

        # Action buttons
        button_frame = tk.Frame(content_frame, bg=self.colors['light_bg'])
        button_frame.pack(fill=tk.X, pady=20)

        generate_btn = self.create_modern_button(
            button_frame,
            "🚀 Generate Statement Now",
            lambda: [preview_window.destroy(), self.generate_customer_statement()],
            style="success",
            width=20,
            height=2
        )
        generate_btn.pack(side=tk.LEFT, padx=5)

        close_btn = self.create_modern_button(
            button_frame,
            "✖️ Close Preview",
            preview_window.destroy,
            style="secondary",
            width=15,
            height=2
        )
        close_btn.pack(side=tk.LEFT, padx=5)

    def generate_customer_statement(self, event=None):
        """Generate customer statement from selected bills - searches ALL office folders"""
        if not self.selected_customer_bills:
            messagebox.showwarning("⚠️ Selection Required", "Please select at least one bill to generate statement.")
            return

        # Get selected bills data
        selected_bills_data = []
        total_amount = 0
        pdfs_to_merge = []

        for bill in self.customer_bills_data:
            if bill['bill_no'] in self.selected_customer_bills:
                selected_bills_data.append(bill)
                total_amount += bill['amount']
                
                # Collect PDF paths for merging - searches ALL office folders
                if bill['pdf_available'] == '✅ Yes' and bill['pdf_path']:
                    resolved_path = self.resolve_pdf_path_updated(bill['pdf_path'], bill['bill_no'])
                    if resolved_path and os.path.exists(resolved_path):
                        pdfs_to_merge.append(resolved_path)
                    else:
                        print(f"DEBUG: ❌ Could not resolve PDF path for bill {bill['bill_no']}")

        customer_name = self.customer_combobox.get().strip()

        # Ask for confirmation
        confirm_msg = f"""
        Generate statement for {customer_name}?
        
        📊 Summary:
        • Bills to include: {len(selected_bills_data)}
        • PDFs available: {len(pdfs_to_merge)}
        • Total amount: ₹{total_amount:,.2f}
        
        Proceed with statement generation?
        """
        
        if not messagebox.askyesno("🚀 Confirm Statement Generation", confirm_msg):
            return

        try:
            from PyPDF2 import PdfMerger
            merger = PdfMerger()
            merged_pdfs_count = 0

            # Add available PDFs from ALL office folders
            for pdf_path in pdfs_to_merge:
                try:
                    merger.append(pdf_path)
                    merged_pdfs_count += 1
                    print(f"DEBUG: ✅ Merged PDF: {os.path.basename(pdf_path)}")
                except Exception as e:
                    print(f"DEBUG: ❌ Could not merge {pdf_path}: {e}")

            customer_name = self.customer_combobox.get().strip()

            # Get current year for statement folder
            current_year = str(datetime.now().year)

            # Create year-based statements directory
            documents_dir = os.path.join(os.path.expanduser("~"), "Documents")
            statements_dir = os.path.join(
                documents_dir, 
                "InvoiceApp", 
                f"Invoice_Bill_{current_year}",  # Year-based folder
                "STATEMENTS_FOLDER", 
                "Customer_All_Bills"
            )
            os.makedirs(statements_dir, exist_ok=True)
            
            clean_customer_name = re.sub(r'[^\w\s-]', '', customer_name).strip().replace(' ', '_')
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = os.path.join(statements_dir, f"{clean_customer_name}_Statement_{timestamp}.pdf")
            
            # Write merged PDF
            merger.write(output_filename)
            merger.close()

            # Show success message
            success_msg = f"""
            ✅ Customer Statement Generated Successfully!
            
            👤 Customer: {customer_name}
            📄 Bills Included: {len(selected_bills_data)}
            📁 PDFs Merged: {merged_pdfs_count} (searched across ALL office folders)
            💰 Total Amount: ₹{total_amount:,.2f}
            💾 Saved as: {os.path.basename(output_filename)}
            """
            
            messagebox.showinfo("✅ Success", success_msg.strip())

            # Open the generated statement
            self.display_pdf(output_filename)

            self.show_status_message(f"✅ Statement generated with {len(selected_bills_data)} bills - {merged_pdfs_count} PDFs merged from ALL offices")

        except Exception as e:
            messagebox.showerror("❌ Error", f"Failed to generate statement:\n{str(e)}")
            self.show_status_message("❌ Failed to generate statement", error=True)

    def show_agent_commission_page(self):
        """Modern enhanced agent commission statement generation with checkbox selection"""
        self.clear_screen()
        self.current_screen = "agent_commission"
        self.create_navigation_bar()
        self.create_status_bar()

        # Reset focusable widgets
        self.focusable_widgets.clear()

        # Main content area
        main_container = tk.Frame(self.root, bg=self.colors['light_bg'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Header frame
        header_frame = self.create_modern_frame(main_container, "🤵 AGENT COMMISSION STATEMENT")
        header_frame.pack(fill=tk.X, pady=(0, 15))

        # Search and filter section
        search_frame = tk.Frame(header_frame, bg=self.colors['card_bg'])
        search_frame.pack(fill=tk.X, padx=20, pady=15)

        # Agent selection
        agent_select_frame = tk.Frame(search_frame, bg=self.colors['card_bg'])
        agent_select_frame.pack(fill=tk.X, pady=10)

        tk.Label(
            agent_select_frame,
            text="🤵 Select Agent:",
            font=("Segoe UI", 11, "bold"),
            bg=self.colors['card_bg'],
            fg=self.colors['text_dark']
        ).pack(side=tk.LEFT, padx=(0, 10))

        # Get all unique agent names from bills data
        agent_names = set()
        for bill_no, bill in self.bills_data.items():
            if "agent_name" in bill and bill["agent_name"]:
                agent_names.add(bill["agent_name"].strip())
        
        agent_names = sorted(list(agent_names))
        self.all_agent_names = sorted(list(agent_names))

        
        # Agent combobox with search
        self.agent_name_combobox = self.create_modern_combobox(
            agent_select_frame,
            values=agent_names,
            width=35,
            font_size=11
        )
        self.agent_name_combobox.pack(side=tk.LEFT, padx=5)

        self.agent_name_combobox.bind("<KeyRelease>", self.update_agent_suggestions)
        
        # Load bills button
        load_bills_btn = self.create_modern_button(
            agent_select_frame,
            "📥 Load Agent Bills",
            self.load_agent_bills_for_commission,
            style="info",
            width=16,
            height=1
        )
        load_bills_btn.pack(side=tk.LEFT, padx=10)

        # Commission rate
        commission_frame = tk.Frame(search_frame, bg=self.colors['card_bg'])
        commission_frame.pack(fill=tk.X, pady=10)

        tk.Label(
            commission_frame,
            text="💰 Commission Rate:",
            font=("Segoe UI", 10, "bold"),
            bg=self.colors['card_bg'],
            fg=self.colors['text_dark']
        ).pack(side=tk.LEFT, padx=(0, 10))

        self.commission_rate = tk.DoubleVar(value=5.0)
        self.commission_entry = self.create_modern_entry(
            commission_frame, 
            textvariable=self.commission_rate,
            width=8,
            font_size=10
        )
        self.commission_entry.pack(side=tk.LEFT, padx=5)

        tk.Label(
            commission_frame,
            text="%",
            font=("Segoe UI", 10),
            bg=self.colors['card_bg'],
            fg=self.colors['text_dark']
        ).pack(side=tk.LEFT, padx=(0, 20))

        # Update commission button
        update_commission_btn = self.create_modern_button(
            commission_frame,
            "🔄 Update Commission",
            self.update_agent_commission_calculation,
            style="primary",
            width=16,
            height=1
        )
        update_commission_btn.pack(side=tk.LEFT, padx=10)

        # Date range
        date_frame = tk.Frame(search_frame, bg=self.colors['card_bg'])
        date_frame.pack(fill=tk.X, pady=10)

        tk.Label(
            date_frame,
            text="📅 Date Range:",
            font=("Segoe UI", 10, "bold"),
            bg=self.colors['card_bg'],
            fg=self.colors['text_dark']
        ).pack(side=tk.LEFT, padx=(0, 10))

        # From date
        tk.Label(date_frame, text="From:", font=("Segoe UI", 9),
                bg=self.colors['card_bg']).pack(side=tk.LEFT, padx=(0, 5))
        
        self.agent_from_date = self.create_modern_entry(date_frame, width=12, font_size=9)
        self.agent_from_date.pack(side=tk.LEFT, padx=5)
        self.agent_from_date.insert(0, "01/01/2024")

        # To date
        tk.Label(date_frame, text="To:", font=("Segoe UI", 9),
                bg=self.colors['card_bg']).pack(side=tk.LEFT, padx=(10, 5))
        
        self.agent_to_date = self.create_modern_entry(date_frame, width=12, font_size=9)
        self.agent_to_date.pack(side=tk.LEFT, padx=5)
        self.agent_to_date.insert(0, datetime.now().strftime("%d/%m/%Y"))

        # Apply filter button
        apply_btn = self.create_modern_button(
            date_frame,
            "🔍 Apply Filter",
            self.apply_agent_date_filter,
            style="primary",
            width=12,
            height=1
        )
        apply_btn.pack(side=tk.LEFT, padx=15)

        # Clear filter button
        clear_btn = self.create_modern_button(
            date_frame,
            "🔄 Clear Filter",
            self.clear_agent_filter,
            style="secondary",
            width=12,
            height=1
        )
        clear_btn.pack(side=tk.LEFT, padx=5)

        # Bills table container
        table_container = self.create_modern_frame(main_container, "📋 AGENT INVOICES - SELECT BILLS FOR COMMISSION")
        table_container.pack(fill=tk.BOTH, expand=True, pady=10)

        # Create table frame
        table_main = tk.Frame(table_container, bg=self.colors['card_bg'])
        table_main.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        # Create Treeview for agent bills
        self.agent_bills_table = ttk.Treeview(
            table_main, 
            columns=("Select", "Bill No", "Bill Date", "Customer", "Amount", "Commission", "Status", "PDF"), 
            show="headings",
            style="Modern.Treeview",
            selectmode="extended",
            height=12
        )

        # Define column headings
        columns = {
            "Select": {"width": 60, "anchor": "center"},
            "Bill No": {"width": 90, "anchor": "center"},
            "Bill Date": {"width": 90, "anchor": "center"},
            "Customer": {"width": 150, "anchor": "w"},
            "Amount": {"width": 100, "anchor": "center"},
            "Commission": {"width": 100, "anchor": "center"},
            "Status": {"width": 90, "anchor": "center"},
            "PDF": {"width": 80, "anchor": "center"}
        }

        for col, settings in columns.items():
            self.agent_bills_table.heading(col, text=col)
            self.agent_bills_table.column(col, width=settings["width"], anchor=settings["anchor"])

        # Add scrollbars
        v_scrollbar = ttk.Scrollbar(table_main, orient="vertical", command=self.agent_bills_table.yview)
        h_scrollbar = ttk.Scrollbar(table_main, orient="horizontal", command=self.agent_bills_table.xview)
        self.agent_bills_table.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        # Grid layout
        self.agent_bills_table.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")

        table_main.grid_rowconfigure(0, weight=1)
        table_main.grid_columnconfigure(0, weight=1)

        # Selection controls
        selection_frame = tk.Frame(main_container, bg=self.colors['light_bg'])
        selection_frame.pack(fill=tk.X, pady=10)

        # Left side - Selection info
        left_frame = tk.Frame(selection_frame, bg=self.colors['light_bg'])
        left_frame.pack(side=tk.LEFT)

        # Select all checkbox
        self.select_all_agent = tk.BooleanVar()
        select_all_cb = tk.Checkbutton(
            left_frame,
            text="Select All Bills",
            variable=self.select_all_agent,
            font=("Segoe UI", 10, "bold"),
            bg=self.colors['light_bg'],
            command=self.toggle_select_all_agent
        )
        select_all_cb.pack(anchor="w", pady=5)

        # Selection info
        self.agent_selection_info = tk.Label(
            left_frame,
            text="Selected: 0 bills | Sales: ₹0.00 | Commission: ₹0.00",
            font=("Segoe UI", 10),
            bg=self.colors['light_bg'],
            fg=self.colors['success']
        )
        self.agent_selection_info.pack(anchor="w", pady=2)

        # Right side - Action buttons
        right_selection_frame = tk.Frame(selection_frame, bg=self.colors['light_bg'])
        right_selection_frame.pack(side=tk.RIGHT)

        # Generate Agent Commission Button
        generate_commission_btn = self.create_modern_button(
            right_selection_frame,
            "💰Agent Commission & Bills",
            self.generate_agent_commission_statement,
            style="success",
            width=24,
            height=2
        )
        generate_commission_btn.pack(side=tk.LEFT, padx=5)

        # 🧾 NEW: Commission Only (no merge)
        commission_only_btn = self.create_modern_button(
            right_selection_frame,
            "🧾 Agent Commission Only",
            self.generate_agent_commission_only,  # new function
            style="primary",
            width=22,
            height=2
        )
        commission_only_btn.pack(side=tk.LEFT, padx=5)

        # NEW: Generate Agent Statement Bills (merged invoice PDFs only)
        generate_statement_btn = self.create_modern_button(
            right_selection_frame,
            "📄Agent All Bills Olny",
            self.generate_agent_statement_bills,  # <- new function
            style="primary",
            width=24,
            height=2
        )
        generate_statement_btn.pack(side=tk.LEFT, padx=5)

        # Generate Agent Statement Button
        generate_statement_btn = self.create_modern_button(
            right_selection_frame,
            "📄Agent Statement Only",
            self.generate_agent_statement,
            style="primary",
            width=22,
            height=2
        )
        generate_statement_btn.pack(side=tk.LEFT, padx=5)

        # Clear Selection Button
        clear_btn = self.create_modern_button(
            right_selection_frame,
            "🗑️ Clear Selection",
            self.show_agent_commission_page,
            style="warning",
            width=16,
            height=2
        )
        clear_btn.pack(side=tk.LEFT, padx=5)

        # Back button
        back_btn = self.create_modern_button(
            main_container,
            "← Back to Statements",
            self.show_statement_options,
            style="secondary",
            width=16,
            height=2
        )
        back_btn.pack(pady=10)

        # Initialize agent bills data
        self.agent_bills_data = []
        self.selected_agent_bills = set()

        # Bind events
        self.root.bind('<Control-a>', lambda e: self.toggle_select_all_agent())
        self.root.bind('<Control-Return>', lambda e: self.generate_agent_commission_statement())
        self.agent_bills_table.bind('<Button-1>', self.on_agent_bill_click)

        # Set focus to agent combobox
        self.agent_name_combobox.focus_set()

        self.show_status_message("🤵 Select an agent and choose bills for commission calculation")


    def generate_agent_commission_only(self):
        """Generate Agent Commission PDF only (no bill merging)."""
        if not self.selected_agent_bills:
            messagebox.showwarning("Selection Required", "Please select at least one bill for commission calculation.")
            return

        # Validate commission rate
        try:
            commission_rate = float(self.commission_rate.get())
            if commission_rate < 0 or commission_rate > 100:
                messagebox.showerror("Error", "Commission rate must be between 0 and 100%.")
                return
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid commission rate.")
            return

        selected_bills_data = []
        total_net_amount = 0
        total_sub_total = 0
        total_commission = 0
        bills_without_subtotal = 0

        for bill in self.agent_bills_data:
            if bill["bill_no"] in self.selected_agent_bills:
                sub_total = bill.get("sub_total", 0)
                net_amount = bill["amount"]

                if sub_total > 0:
                    bill_commission = (sub_total * commission_rate) / 100
                    total_sub_total += sub_total
                    total_commission += bill_commission
                else:
                    bill_commission = 0
                    bills_without_subtotal += 1

                total_net_amount += net_amount
                selected_bills_data.append({
                    "bill_no": bill["bill_no"],
                    "bill_date": bill["bill_date"],
                    "customer_name": bill["customer_name"],
                    "sub_total": sub_total,
                    "net_amount": net_amount,
                    "commission": bill_commission
                })

        agent_name = self.agent_name_combobox.get().strip()
        if not agent_name:
            messagebox.showwarning("Missing Agent", "Please select an agent.")
            return

        try:
            from fpdf import FPDF
            import re, os
            from datetime import datetime

            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", "", 11)
            pdf.cell(0, 10, "AGENT COMMISSION STATEMENT", ln=True, align="C")
            pdf.ln(5)

            pdf.set_font("Arial", "B", 11)
            pdf.cell(0, 8, f"Agent: {agent_name}", ln=True)
            pdf.cell(0, 8, f"Commission Rate: {commission_rate}%", ln=True)
            pdf.cell(0, 8, f"Statement Date: {datetime.now().strftime('%d/%m/%Y')}", ln=True)
            pdf.cell(0, 8, "Commission Calculated From: SUB TOTAL (After Discount)", ln=True)
            pdf.ln(8)

            pdf.set_font("Arial", "B", 10)
            pdf.cell(0, 8, "SUMMARY", ln=True)
            pdf.set_font("Arial", "", 10)
            pdf.cell(0, 6, f"Total Invoices: {len(selected_bills_data)}", ln=True)
            pdf.cell(0, 6, f"Total NET AMOUNT: Rs.{total_net_amount:,.2f}", ln=True)
            pdf.cell(0, 6, f"Commission Base (SUB TOTAL): Rs.{total_sub_total:,.2f}", ln=True)
            pdf.cell(0, 6, f"Total Commission: Rs.{total_commission:,.2f}", ln=True)
            pdf.ln(10)

            headers = ["Bill No", "Date", "Customer", "NET AMOUNT", "Commission"]
            col_widths = [30, 25, 60, 35, 30]

            pdf.set_font("Arial", "B", 9)
            for h, w in zip(headers, col_widths):
                pdf.cell(w, 8, h, border=1, align="C")
            pdf.ln()

            pdf.set_font("Arial", "", 9)
            for bill in selected_bills_data:
                customer_name = re.sub(r"[^\x00-\x7F]+", "", bill["customer_name"])[:20]
                pdf.cell(col_widths[0], 6, bill["bill_no"], border=1)
                pdf.cell(col_widths[1], 6, bill["bill_date"], border=1)
                pdf.cell(col_widths[2], 6, customer_name, border=1)
                pdf.cell(col_widths[3], 6, f"Rs.{bill['net_amount']:,.2f}", border=1, align="R")
                pdf.cell(col_widths[4], 6, f"Rs.{bill['commission']:,.2f}", border=1, align="R")
                pdf.ln()

            pdf.set_font("Arial", "B", 9)
            pdf.cell(sum(col_widths[:3]), 8, "TOTAL", border=1, align="C")
            pdf.cell(col_widths[3], 8, f"Rs.{total_net_amount:,.2f}", border=1, align="R")
            pdf.cell(col_widths[4], 8, f"Rs.{total_commission:,.2f}", border=1, align="R")
            pdf.ln(10)

            pdf.set_font("Arial", "I", 8)
            pdf.cell(0, 6, "This is a computer-generated commission statement.", ln=True)
            pdf.cell(0, 6, f"Commission calculated from SUB TOTAL (After Discount) at {commission_rate}% rate.", ln=True)

             # 🆕 FIXED: Create year-based statements directory
            current_year = str(datetime.now().year)
            documents_dir = os.path.join(os.path.expanduser("~"), "Documents")
            statements_dir = os.path.join(
                documents_dir, 
                "InvoiceApp", 
                f"Invoice_Bill_{current_year}",  # Year-based folder
                "STATEMENTS_FOLDER", 
                "Agent_Commission_Only"
            )
            os.makedirs(statements_dir, exist_ok=True)

            clean_agent = re.sub(r"[^\w\s-]", "", agent_name).strip().replace(" ", "_")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(statements_dir, f"{clean_agent}_Commission_Only_{timestamp}.pdf")

            pdf.output(output_file)
            messagebox.showinfo("Success", f"Commission PDF generated successfully!\n\nSaved to:\n{output_file}")
            self.display_pdf(output_file)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate Commission Only PDF:\n{e}")


    def generate_agent_statement_bills(self):
        """Merge all selected agent bill PDFs into one Agent Statement file - searches ALL office folders"""
        if not self.selected_agent_bills:
            messagebox.showwarning("Selection Required", "Please select at least one bill to merge.")
            return

        agent_name = self.agent_name_combobox.get().strip()
        if not agent_name:
            messagebox.showwarning("Missing Agent", "Please select an agent before merging bills.")
            return

        try:
            from PyPDF2 import PdfMerger
            import re, os
            from datetime import datetime

            # Get current year
            current_year = str(datetime.now().year)
            
            # Create year-based statements directory
            documents_dir = os.path.join(os.path.expanduser("~"), "Documents")
            statements_dir = os.path.join(
                documents_dir,
                "InvoiceApp",
                f"Invoice_Bill_{current_year}",  # Year-based folder
                "STATEMENTS_FOLDER", 
                "Agent_All_Bills_Olny"
            )
            os.makedirs(statements_dir, exist_ok=True)

            clean_agent = re.sub(r'[^\w\s-]', '', agent_name).strip().replace(' ', '_')
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(statements_dir, f"{clean_agent}_Bills_Merged_{timestamp}.pdf")

            merger = PdfMerger()
            missing = []

            # Merge selected bills from ALL office folders
            for bill in self.agent_bills_data:
                if bill["bill_no"] in self.selected_agent_bills:
                    pdf_path = bill.get("pdf_path", "")
                    if not pdf_path or not os.path.exists(pdf_path):
                        pdf_path = self.resolve_pdf_path_updated(pdf_path, bill["bill_no"])
                    if pdf_path and os.path.exists(pdf_path):
                        merger.append(pdf_path)
                        print(f"DEBUG: ✅ Merged PDF for bill {bill['bill_no']}: {pdf_path}")
                    else:
                        missing.append(bill["bill_no"])
                        print(f"DEBUG: ❌ PDF not found for bill {bill['bill_no']}")

            if merger.pages:  # Check if any PDFs were merged
                merger.write(output_file)
                merger.close()
                
                if missing:
                    messagebox.showwarning(
                        "Merged with Missing Bills",
                        f"Some bills were not found and skipped:\n{', '.join(missing)}\n\n"
                        f"Searched in ALL office folders (AP, AFI, AFF, Invoice_Bills)\n\n"
                        f"Merged file saved to:\n{output_file}"
                    )
                else:
                    messagebox.showinfo(
                        "Success", 
                        f"All selected bills merged successfully!\n\n"
                        f"Searched across ALL office folders\n\n"
                        f"Saved to:\n{output_file}"
                    )
                self.display_pdf(output_file)
            else:
                merger.close()
                messagebox.showerror(
                    "No PDFs Found", 
                    f"No PDFs found for the selected bills in ANY office folder.\n\n"
                    f"Missing bills: {', '.join(missing)}"
                )

        except Exception as e:
            messagebox.showerror("Error", f"Failed to merge Agent Statement Bills:\n{e}")

    def update_agent_suggestions(self, event=None):
        """Dynamically update agent name suggestions as the user types."""
        value = self.agent_name_combobox.get().strip().lower()
        if not hasattr(self, "all_agent_names"):
            return
        
        # Filter matching agents
        if value == "":
            data = self.all_agent_names
        else:
            data = [item for item in self.all_agent_names if value in item.lower()]
        
        # Update combobox dropdown values
        self.agent_name_combobox['values'] = data
        
        # If user typed something, open dropdown automatically
        if data:
            self.agent_name_combobox.event_generate('<Down>')


    def load_agent_bills_for_commission(self):
        """Load bills for selected agent with commission from SUB TOTAL and AUTO-SAVE to JSON"""
        agent_name = self.agent_name_combobox.get().strip()

        if not agent_name:
            messagebox.showwarning("⚠️ Input Required", "Please select an agent name first.")
            return

        # Clear previous data
        for item in self.agent_bills_table.get_children():
            self.agent_bills_table.delete(item)
        
        self.agent_bills_data.clear()
        self.selected_agent_bills.clear()
        self.select_all_agent.set(False)

        # Show loading
        self.show_status_message(f"🔍 Loading bills for {agent_name}...")

        # Collect matching bills
        matching_bills = []
        total_sales = 0
        commission_rate = float(self.commission_rate.get())
        
        # Track commission updates
        updated_commissions_count = 0
        bills_updated = []

        for bill_no, bill in self.bills_data.items():
            bill_agent_name = bill.get("agent_name", "").strip()
            
            # Check agent name match
            if agent_name.lower() != bill_agent_name.lower():
                continue

            # ✅ Calculate commission from SUB TOTAL instead of net_amount
            sub_total = bill.get("sub_total", 0)
            if sub_total > 0:
                commission_amount = (sub_total * commission_rate) / 100
            else:
                commission_amount = 0  # No SUB TOTAL = No commission
            
            amount = float(bill.get("net_amount", 0))
            total_sales += amount

            # ✅ AUTO-SAVE COMMISSION TO JSON
            old_commission = bill.get("commission_amount", 0)
            old_rate = bill.get("commission_rate", 0)
            
            # Check if commission needs updating
            if (commission_amount != old_commission or commission_rate != old_rate):
                # Update bill data with new commission
                bill.update({
                    "commission_rate": commission_rate,
                    "commission_amount": commission_amount,
                    "commission_calculated_on": "sub_total",
                    "commission_last_updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                })
                updated_commissions_count += 1
                bills_updated.append(bill_no)

            # Check if PDF exists
            pdf_available = "❌ No"
            pdf_path = bill.get("pdf_file_name", "")
            if pdf_path:
                resolved_path = self.resolve_pdf_path_updated(pdf_path, bill_no)
                if resolved_path and os.path.exists(resolved_path):
                    pdf_available = "✅ Yes"

            matching_bills.append({
                'bill_no': bill_no,
                'bill_date': bill.get("bill_date", ""),
                'customer_name': bill.get("customer_name", ""),
                'agent_name': bill_agent_name,
                'amount': amount,
                'commission': commission_amount,
                'sub_total': sub_total,  # Store for reference
                'status': bill.get("payment_status", "Pending"),
                'pdf_available': pdf_available,
                'pdf_path': pdf_path,
                'stored_commission': "commission_amount" in bill,
                'commission_rate': commission_rate  # Store current rate
            })

        # ✅ SAVE UPDATED COMMISSIONS TO JSON FILE
        if updated_commissions_count > 0:
            try:
                success = self.save_data(self.bills_ref, self.bills_data)
                if success:
                    self.show_status_message(f"✅ Auto-saved commissions for {updated_commissions_count} bills")
                    print(f"DEBUG: ✅ Auto-saved commissions for {updated_commissions_count} bills: {bills_updated}")
                else:
                    self.show_status_message("⚠️ Commissions calculated but save failed")
            except Exception as e:
                print(f"DEBUG: ❌ Error saving commissions: {e}")
                self.show_status_message("❌ Error saving commissions")

        if not matching_bills:
            messagebox.showinfo("ℹ️ No Bills Found", f"No bills found for agent: {agent_name}")
            self.show_status_message("❌ No bills found for selected agent")
            return

        # Sort by bill date (newest first)
        matching_bills.sort(key=lambda x: (parse_date_flexible(x.get('bill_date')) or datetime.min), reverse=True)

        # Populate table
        for bill in matching_bills:
            item_id = self.agent_bills_table.insert("", "end", values=(
                "☐",  # Checkbox placeholder
                bill['bill_no'],
                bill['bill_date'],
                bill['customer_name'],
                f"₹{bill['amount']:,.2f}",
                f"₹{bill['commission']:,.2f}",
                bill['status'],
                bill['pdf_available']
            ))
            
            # Store bill data with treeview item ID
            bill['item_id'] = item_id
            self.agent_bills_data.append(bill)

        # Update selection info
        self.update_agent_selection_info()

        # Count bills without SUB TOTAL
        bills_without_subtotal = sum(1 for bill in matching_bills if bill['sub_total'] == 0)
        
        # Final status message
        if bills_without_subtotal > 0:
            status_msg = f"📊 Loaded {len(matching_bills)} bills - {bills_without_subtotal} without SUB TOTAL"
            if updated_commissions_count > 0:
                status_msg += f" | ✅ Auto-saved {updated_commissions_count} commissions"
            self.show_status_message(status_msg)
        else:
            status_msg = f"📊 Loaded {len(matching_bills)} bills for {agent_name} | Sales: ₹{total_sales:,.2f}"
            if updated_commissions_count > 0:
                status_msg += f" | ✅ Auto-saved {updated_commissions_count} commissions"
            self.show_status_message(status_msg)

    def on_agent_bill_click(self, event):
        """Handle checkbox clicks in agent bills table"""
        item = self.agent_bills_table.identify_row(event.y)
        column = self.agent_bills_table.identify_column(event.x)

        if item and column == "#1":  # Checkbox column
            bill_no = self.agent_bills_table.item(item, "values")[1]
            
            if bill_no in self.selected_agent_bills:
                self.selected_agent_bills.remove(bill_no)
                self.agent_bills_table.set(item, "Select", "☐")
            else:
                self.selected_agent_bills.add(bill_no)
                self.agent_bills_table.set(item, "Select", "☑")
            
            self.update_agent_selection_info()

    def toggle_select_all_agent(self):
        """Toggle select all bills for agent"""
        if self.select_all_agent.get():
            # Select all
            self.selected_agent_bills.clear()
            for bill in self.agent_bills_data:
                self.selected_agent_bills.add(bill['bill_no'])
                self.agent_bills_table.set(bill['item_id'], "Select", "☑")
        else:
            # Deselect all
            self.selected_agent_bills.clear()
            for bill in self.agent_bills_data:
                self.agent_bills_table.set(bill['item_id'], "Select", "☐")
        
        self.update_agent_selection_info()

    def update_agent_selection_info(self):
        """Update agent selection information display"""
        selected_count = len(self.selected_agent_bills)
        total_sales = 0
        total_commission = 0
        
        for bill in self.agent_bills_data:
            if bill['bill_no'] in self.selected_agent_bills:
                total_sales += bill['amount']
                total_commission += bill['commission']
        
        commission_rate = float(self.commission_rate.get())
        
        self.agent_selection_info.config(
            text=f"Selected: {selected_count} bills | Sales: ₹{total_sales:,.2f} | Commission: ₹{total_commission:,.2f}"
        )
        
        # Update select all checkbox state
        total_bills = len(self.agent_bills_data)
        if selected_count == total_bills and total_bills > 0:
            self.select_all_agent.set(True)
        else:
            self.select_all_agent.set(False)

    def update_agent_commission_calculation(self):
        """Update commission calculations for all bills using SUB TOTAL"""
        agent_name = self.agent_name_combobox.get().strip()
        if not agent_name:
            messagebox.showwarning("⚠️ Input Required", "Please select an agent first.")
            return

        # Validate commission rate
        try:
            commission_rate = float(self.commission_rate.get())
            if commission_rate < 0 or commission_rate > 100:
                messagebox.showerror("❌ Error", "Commission rate must be between 0 and 100%.")
                return
        except ValueError:
            messagebox.showerror("❌ Error", "Please enter a valid commission rate.")
            return

        # Recalculate commission for all bills using SUB TOTAL
        for bill in self.agent_bills_data:
            # ✅ Direct calculation from SUB TOTAL
            sub_total = bill['sub_total']
            if sub_total > 0:
                new_commission = (sub_total * commission_rate) / 100
            else:
                new_commission = 0
            
            bill['commission'] = new_commission
            self.agent_bills_table.set(bill['item_id'], "Commission", f"₹{new_commission:,.2f}")

        # Update selection info
        self.update_agent_selection_info()

        self.show_status_message(f"✅ Commission rate updated to {commission_rate}% based on SUB TOTAL")

    def apply_agent_date_filter(self):
        """Apply date filter to agent bills"""
        try:
            from_date = parse_date_flexible(self.agent_from_date.get())
            to_date   = parse_date_flexible(self.agent_to_date.get())

            if not from_date or not to_date:
                messagebox.showerror("❌ Error", "Please enter valid dates (e.g. DD/MM/YYYY or DD.MM.YYYY).")
                return

            if from_date > to_date:
                messagebox.showerror("❌ Error", "From date cannot be after To date.")
                return

        except Exception as e:
            messagebox.showerror("❌ Error", f"Invalid date format. Details: {e}")
            return

        agent_name = self.agent_name_combobox.get().strip()
        if not agent_name:
            messagebox.showwarning("⚠️ Input Required", "Please select an agent first.")
            return

        # Reload bills with date filter
        self.load_agent_bills_with_date_filter(agent_name, from_date, to_date)

    def load_agent_bills_with_date_filter(self, agent_name, from_date, to_date):
        """Load agent bills with date filter"""
        # Clear previous data
        for item in self.agent_bills_table.get_children():
            self.agent_bills_table.delete(item)
        
        self.agent_bills_data.clear()
        self.selected_agent_bills.clear()
        self.select_all_agent.set(False)

        # Collect filtered bills
        filtered_bills = []
        total_sales = 0
        commission_rate = float(self.commission_rate.get())

        for bill_no, bill in self.bills_data.items():
            bill_agent_name = bill.get("agent_name", "").strip()
            
            # Check agent name match
            if agent_name.lower() != bill_agent_name.lower():
                continue

            # Check date range
            try:
                bill_date = datetime.strptime(bill.get("bill_date", ""), "%d/%m/%Y")
                if not (from_date <= bill_date <= to_date):
                    continue
            except ValueError:
                continue

            # Check if PDF exists
            pdf_available = " No"
            pdf_path = bill.get("pdf_file_name", "")
            if pdf_path:
                resolved_path = self.resolve_pdf_path_updated(pdf_path, bill_no)
                if resolved_path and os.path.exists(resolved_path):
                    pdf_available = " Yes"

            amount = float(bill.get("net_amount", 0))
            commission_amount = (amount * commission_rate) / 100
            total_sales += amount

            filtered_bills.append({
                'bill_no': bill_no,
                'bill_date': bill.get("bill_date", ""),
                'customer_name': bill.get("customer_name", ""),
                'agent_name': bill_agent_name,
                'amount': amount,
                'commission': commission_amount,
                'status': bill.get("payment_status", "Pending"),
                'pdf_available': pdf_available,
                'pdf_path': pdf_path
            })

        if not filtered_bills:
            messagebox.showinfo("ℹ️ No Bills Found", f"No bills found for {agent_name} in the selected date range.")
            self.show_status_message("❌ No bills found for selected criteria")
            return

        # Sort by bill date (newest first)
        filtered_bills.sort(key=lambda x: datetime.strptime(x['bill_date'], "%d/%m/%Y"), reverse=True)

        # Populate table
        for bill in filtered_bills:
            item_id = self.agent_bills_table.insert("", "end", values=(
                "☐",
                bill['bill_no'],
                bill['bill_date'],
                bill['customer_name'],
                f"₹{bill['amount']:,.2f}",
                f"₹{bill['commission']:,.2f}",
                bill['status'],
                bill['pdf_available']
            ))
            
            bill['item_id'] = item_id
            self.agent_bills_data.append(bill)

        self.update_agent_selection_info()
        self.show_status_message(f"✅ Found {len(filtered_bills)} bills - Total Sales: ₹{total_sales:,.2f}")

    def clear_agent_filter(self):
        """Clear date filter and reload all bills"""
        self.agent_from_date.delete(0, tk.END)
        self.agent_from_date.insert(0, "01/01/2024")
        self.agent_to_date.delete(0, tk.END)
        self.agent_to_date.insert(0, datetime.now().strftime("%d/%m/%Y"))
        
        agent_name = self.agent_name_combobox.get().strip()
        if agent_name:
            self.load_agent_bills_for_commission()

    def clear_agent_selection(self):
        """Completely clear all selected agent bills and refresh UI checkboxes."""
        try:
            # Clear tracking sets and checkbox variables
            self.selected_agent_bills.clear()
            self.select_all_agent.set(False)

            # Reset all rows' "Select" column to unchecked
            for bill in self.agent_bills_data:
                try:
                    self.agent_bills_table.set(bill['item_id'], "Select", "☐")
                except Exception:
                    pass  # In case an item was removed

            # Force table redraw to visually update checkboxes
            self.agent_bills_table.update_idletasks()

            # Update selection summary info
            self.update_agent_selection_info()

            # Status message
            self.show_status_message("🗑️ All selections cleared")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to clear selections:\n{e}")





    def generate_agent_commission_statement(self):
        """Generate agent commission statement - Commission from SUB TOTAL, Display NET AMOUNT"""
        if not self.selected_agent_bills:
            messagebox.showwarning("Selection Required", "Please select at least one bill for commission calculation.")
            return

        # Validate commission rate
        try:
            commission_rate = float(self.commission_rate.get())
            if commission_rate < 0 or commission_rate > 100:
                messagebox.showerror("Error", "Commission rate must be between 0 and 100%.")
                return
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid commission rate.")
            return

        selected_bills_data = []
        total_net_amount = 0  # For display
        total_sub_total = 0   # For commission calculation
        total_commission = 0
        bills_without_subtotal = 0

        for bill in self.agent_bills_data:
            if bill['bill_no'] in self.selected_agent_bills:
                # Get SUB TOTAL from bill data (for commission calculation)
                sub_total = bill.get('sub_total', 0)
                net_amount = bill['amount']  # NET AMOUNT for display
                
                if sub_total > 0:
                    # Calculate commission from SUB TOTAL
                    bill_commission = (sub_total * commission_rate) / 100
                    total_sub_total += sub_total
                    total_commission += bill_commission
                else:
                    # If no SUB TOTAL found, commission is 0
                    bill_commission = 0
                    bills_without_subtotal += 1
                
                total_net_amount += net_amount
                
                selected_bills_data.append({
                    'bill_no': bill['bill_no'],
                    'bill_date': bill['bill_date'],
                    'customer_name': bill['customer_name'],
                    'sub_total': sub_total,
                    'net_amount': net_amount,  # For display in table
                    'commission': bill_commission,
                    'pdf_path': bill.get('pdf_path', '')
                })

        agent_name = self.agent_name_combobox.get().strip()

        # Show warning if some bills don't have SUB TOTAL
        if bills_without_subtotal > 0:
            messagebox.showwarning(
                "Commission Calculation Warning", 
                f"{bills_without_subtotal} bills don't have SUB TOTAL stored. "
                f"Commission for these bills will be Rs.0.00. "
                f"Newer bills will have SUB TOTAL stored automatically."
            )

        confirm_msg = (
            f"Generate commission statement for {agent_name}?\n\n"
            f"Summary:\n"
            f"Bills to include: {len(selected_bills_data)}\n"
            f"Commission Rate: {commission_rate}%\n"
            f"Total NET AMOUNT: Rs.{total_net_amount:,.2f}\n"
            f"Commission Base (SUB TOTAL): Rs.{total_sub_total:,.2f}\n"
            f"Total Commission: Rs.{total_commission:,.2f}\n"
            f"Bills without SUB TOTAL: {bills_without_subtotal}\n\n"
            f"Proceed with commission statement generation?"
        )

        if not messagebox.askyesno("Confirm Commission Statement", confirm_msg):
            return

        try:
            from fpdf import FPDF

            pdf = FPDF()
            pdf.add_page()
            pdf.set_left_margin(15)
            pdf.set_right_margin(15)
            
            # Use only basic fonts and characters
            pdf.set_font("Arial", "", 11)

            # Header
            pdf.cell(0, 10, "AGENT COMMISSION STATEMENT", ln=True, align="C")
            pdf.ln(5)

            pdf.set_font("Arial", "B", 11)
            pdf.cell(0, 8, f"Agent: {agent_name}", ln=True)
            pdf.cell(0, 8, f"Commission Rate: {commission_rate}%", ln=True)
            pdf.cell(0, 8, f"Statement Date: {datetime.now().strftime('%d/%m/%Y')}", ln=True)
            pdf.cell(0, 8, "Commission Calculated From: SUB TOTAL (After Discount)", ln=True)
            pdf.ln(8)

            # Summary
            pdf.set_font("Arial", "B", 10)
            pdf.cell(0, 8, "SUMMARY", ln=True)
            pdf.set_font("Arial", "", 10)
            pdf.cell(0, 6, f"Total Invoices: {len(selected_bills_data)}", ln=True)
            pdf.cell(0, 6, f"Total NET AMOUNT: Rs.{total_net_amount:,.2f}", ln=True)  # Show NET AMOUNT
            pdf.cell(0, 6, f"Commission Base (SUB TOTAL): Rs.{total_sub_total:,.2f}", ln=True)  # Show commission base
            pdf.cell(0, 6, f"Total Commission: Rs.{total_commission:,.2f}", ln=True)
            if bills_without_subtotal > 0:
                pdf.cell(0, 6, f"Bills without SUB TOTAL: {bills_without_subtotal} (Commission: Rs.0.00)", ln=True)
            pdf.ln(10)

            # Table Header - Show NET AMOUNT instead of SUB TOTAL
            headers = ["Bill No", "Date", "Customer", "NET AMOUNT", "Commission"]
            col_widths = [30, 25, 60, 35, 30]

            pdf.set_font("Arial", "B", 9)
            for i, h in enumerate(headers):
                pdf.cell(col_widths[i], 8, h, border=1, align='C')
            pdf.ln()

            # Table Rows - Show NET AMOUNT instead of SUB TOTAL
            pdf.set_font("Arial", "", 9)
            for bill in selected_bills_data:
                # Clean customer name to remove special characters
                customer_name = re.sub(r'[^\x00-\x7F]+', '', bill['customer_name'])[:20]
                
                pdf.cell(col_widths[0], 6, bill['bill_no'], border=1)
                pdf.cell(col_widths[1], 6, bill['bill_date'], border=1)
                pdf.cell(col_widths[2], 6, customer_name, border=1)
                
                # Show NET AMOUNT in the table (always show net_amount, even if commission is 0)
                pdf.cell(col_widths[3], 6, f"Rs.{bill['net_amount']:,.2f}", border=1, align='R')
                
                if bill['sub_total'] == 0:
                    pdf.cell(col_widths[4], 6, "Rs.0.00", border=1, align='R')
                else:
                    pdf.cell(col_widths[4], 6, f"Rs.{bill['commission']:,.2f}", border=1, align='R')
                pdf.ln()

            # Total Row - Show NET AMOUNT total
            pdf.set_font("Arial", "B", 9)
            pdf.cell(col_widths[0] + col_widths[1] + col_widths[2], 8, "TOTAL", border=1, align='C')
            pdf.cell(col_widths[3], 8, f"Rs.{total_net_amount:,.2f}", border=1, align='R')  # NET AMOUNT total
            pdf.cell(col_widths[4], 8, f"Rs.{total_commission:,.2f}", border=1, align='R')  # Commission total
            pdf.ln(10)

            # Footer note - Explain commission calculation
            pdf.set_font("Arial", "I", 8)
            pdf.cell(0, 6, "This is a computer-generated commission statement.", ln=True)
            pdf.cell(0, 6, f"Commission calculated from SUB TOTAL (After Discount) at {commission_rate}% rate.", ln=True)
            pdf.cell(0, 6, "NET AMOUNT shown includes taxes and packing charges.", ln=True)
            if bills_without_subtotal > 0:
                pdf.cell(0, 6, f"Note: {bills_without_subtotal} bills have Rs.0.00 commission as SUB TOTAL was not available.", ln=True)

            # Get current year
            current_year = str(datetime.now().year)

            # Save commission PDF in year-based folder
            documents_dir = os.path.join(os.path.expanduser("~"), "Documents")
            statements_dir = os.path.join(
                documents_dir, 
                "InvoiceApp", 
                f"Invoice_Bill_{current_year}",  # Year-based folder
                "STATEMENTS_FOLDER", 
                "Agent_Commission_and_Bills"
            )
            os.makedirs(statements_dir, exist_ok=True)
            
            # Clean agent name for filename
            clean_agent = re.sub(r'[^\w\s-]', '', agent_name).strip().replace(' ', '_')
            clean_agent = re.sub(r'[^\x00-\x7F]+', '', clean_agent)  # Remove non-ASCII characters
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(statements_dir, f"{clean_agent}_Commission_{timestamp}.pdf")
            
            # Use proper encoding for PDF output
            try:
                pdf.output(output_file)
            except UnicodeEncodeError:
                # Fallback: Use the safe PDF creation method
                pdf = self.create_latin1_safe_pdf(selected_bills_data, agent_name, commission_rate, total_net_amount, total_commission, bills_without_subtotal)
                pdf.output(output_file)

            # Merge commission + bill PDFs
            try:
                from PyPDF2 import PdfMerger
                merger = PdfMerger()
                merger.append(output_file)

                missing = []
                for bill in selected_bills_data:
                    pdf_path = bill.get('pdf_path', '')
                    if not pdf_path or not os.path.exists(pdf_path):
                        pdf_path = self.resolve_pdf_path_updated(bill.get('pdf_path', ''), bill['bill_no'])
                    if pdf_path and os.path.exists(pdf_path):
                        merger.append(pdf_path)
                    else:
                        missing.append(bill['bill_no'])

                merged_output = os.path.join(statements_dir, f"{clean_agent}_Commission_Merged_{timestamp}.pdf")
                merger.write(merged_output)
                merger.close()

                if missing:
                    messagebox.showwarning("Missing PDFs", f"Some bills not merged:\n{', '.join(missing)}")

                messagebox.showinfo(
                    "Success",
                    f"Commission Statement Generated Successfully!\n\n"
                    f"Agent: {agent_name}\n"
                    f"Bills Included: {len(selected_bills_data)}\n"
                    f"Total NET AMOUNT: Rs.{total_net_amount:,.2f}\n"
                    f"Commission Base (SUB TOTAL): Rs.{total_sub_total:,.2f}\n"
                    f"Total Commission: Rs.{total_commission:,.2f}\n"
                    f"Bills without SUB TOTAL: {bills_without_subtotal}\n\n"
                    f"Saved to:\n{merged_output}"
                )

                self.display_pdf(merged_output)

            except Exception as e:
                messagebox.showerror("Error", f"Failed to merge PDFs:\n{e}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate commission statement:\n{e}")
            self.show_status_message("Failed to generate commission statement", error=True)

    def create_latin1_safe_pdf(self, bills_data, agent_name, commission_rate, total_net_amount, total_commission, bills_without_subtotal):
        """Create a Latin-1 safe PDF as fallback - Updated for NET AMOUNT display"""
        pdf = FPDF()
        pdf.add_page()
        pdf.set_left_margin(15)
        pdf.set_right_margin(15)
        pdf.set_font("Arial", "", 11)

        # Simple header without special characters
        pdf.cell(0, 10, "AGENT COMMISSION STATEMENT", ln=True, align="C")
        pdf.ln(5)

        pdf.set_font("Arial", "B", 11)
        pdf.cell(0, 8, f"Agent: {agent_name}", ln=True)
        pdf.cell(0, 8, f"Commission Rate: {commission_rate}%", ln=True)
        pdf.cell(0, 8, f"Statement Date: {datetime.now().strftime('%d/%m/%Y')}", ln=True)
        pdf.ln(8)

        # Simple table with NET AMOUNT
        headers = ["Bill No", "Date", "NET AMOUNT", "Commission"]
        col_widths = [40, 30, 40, 40]

        pdf.set_font("Arial", "B", 9)
        for i, h in enumerate(headers):
            pdf.cell(col_widths[i], 8, h, border=1, align='C')
        pdf.ln()

        pdf.set_font("Arial", "", 9)
        for bill in bills_data:
            pdf.cell(col_widths[0], 6, bill['bill_no'], border=1)
            pdf.cell(col_widths[1], 6, bill['bill_date'], border=1)
            pdf.cell(col_widths[2], 6, f"Rs.{bill['net_amount']:,.2f}", border=1, align='R')  # Show NET AMOUNT
            pdf.cell(col_widths[3], 6, f"Rs.{bill['commission']:,.2f}", border=1, align='R')
            pdf.ln()

        return pdf

    def save_bill_data(self):
        """Save bill data including SUB TOTAL for commission calculation"""
        # Your existing bill data collection code...
        
        # Calculate and store SUB TOTAL for commission
        goods_value = float(self.goods_value.get() or 0)
        special_discount = float(self.special_discount.get() or 0)
        sub_total = goods_value - special_discount
        
        bill_data = {
            # Your existing bill data...
            "sub_total": sub_total,  # Store SUB TOTAL for commission
            "goods_value": goods_value,
            "special_discount": special_discount,
            # ... other fields
        }
        
        self.save_data(self.bills_ref, self.bills_data)


    def calculate_agent_commission(self, bill_no, commission_rate=None):
        """Calculate and store agent commission from SUB TOTAL"""
        if bill_no not in self.bills_data:
            return 0
        
        bill = self.bills_data[bill_no]
        
        # Use provided rate or get from UI/default
        if commission_rate is None:
            commission_rate = float(getattr(self, 'commission_rate', 5.0))
        
        # Get SUB TOTAL - try different sources in order
        sub_total = 0
        
        # 1. First try: Get from stored bill data
        if "sub_total" in bill:
            sub_total = float(bill["sub_total"])
        # 2. Second try: Calculate from goods_value and special_discount
        elif "goods_value" in bill and "special_discount" in bill:
            goods_value = float(bill.get("goods_value", 0))
            special_discount = float(bill.get("special_discount", 0))
            sub_total = goods_value - special_discount
        # 3. Third try: Extract from PDF (for older bills)
        else:
            sub_total = self.extract_subtotal_from_pdf(bill_no)
            
            # If found from PDF, save it for future use
            if sub_total > 0:
                bill["sub_total"] = sub_total
                self.save_data(self.bills_ref, self.bills_data)
        
        # Calculate commission
        commission_amount = (sub_total * commission_rate) / 100
        
        # Store commission data in bill
        bill["commission_rate"] = commission_rate
        bill["commission_amount"] = commission_amount
        bill["commission_calculated_on"] = "sub_total"
        
        # Save updated bill data
        self.save_data(self.bills_ref, self.bills_data)
        
        return commission_amount

    def extract_subtotal_from_pdf(self, bill_no):
        """Extract SUB TOTAL from PDF file for older bills"""
        if bill_no not in self.bills_data:
            return 0
        
        bill = self.bills_data[bill_no]
        pdf_path = bill.get("pdf_file_name", "")
        
        if not pdf_path:
            return 0
        
        # Resolve PDF path
        resolved_path = self.resolve_pdf_path_updated(pdf_path, bill_no)
        if not resolved_path or not os.path.exists(resolved_path):
            return 0
        
        try:
            import PyPDF2
            with open(resolved_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                text = ""
                for page in pdf_reader.pages:
                    text += page.extract_text()
                
                # Look for SUB TOTAL in the text
                lines = text.split('\n')
                for i, line in enumerate(lines):
                    if "SUB TOTAL" in line.upper():
                        # Look for amount in current or next line
                        amount_line = line
                        if not any(char.isdigit() for char in line):
                            # Check next line if current line has no digits
                            if i + 1 < len(lines):
                                amount_line = lines[i + 1]
                        
                        # Extract numeric value
                        import re
                        numbers = re.findall(r"[\d,]+\.\d{2}", amount_line)
                        if numbers:
                            # Remove commas and convert to float
                            amount = float(numbers[0].replace(',', ''))
                            return amount
                
                # Alternative pattern: Look for amount after "SUB TOTAL"
                pattern = r"SUB TOTAL\s*[\d,]+\.\d{2}"
                matches = re.findall(pattern, text.upper())
                if matches:
                    numbers = re.findall(r"[\d,]+\.\d{2}", matches[0])
                    if numbers:
                        return float(numbers[0].replace(',', ''))
                        
        except Exception as e:
            print(f"DEBUG: Error extracting SUB TOTAL from PDF: {e}")
        
        return 0

    def batch_update_commissions(self):
        """Update commissions for all bills in database"""
        confirm = messagebox.askyesno(
            "🔄 Batch Update Commissions", 
            "This will calculate commissions from SUB TOTAL for ALL bills.\n\n"
            "For bills without SUB TOTAL stored, it will try to extract from PDFs.\n\n"
            "This may take a while. Continue?"
        )
        
        if not confirm:
            return
        
        updated_count = 0
        commission_rate = float(self.commission_rate.get())
        
        for bill_no in self.bills_data:
            old_commission = self.bills_data[bill_no].get("commission_amount", 0)
            new_commission = self.calculate_agent_commission(bill_no, commission_rate)
            
            if new_commission != old_commission:
                updated_count += 1
        
        messagebox.showinfo(
            "✅ Batch Update Complete", 
            f"Updated commissions for {updated_count} bills.\n\n"
            f"Commission rate: {commission_rate}%\n"
            f"Calculated from: SUB TOTAL (After Discount)"
        )

    def calculate_and_save_commission(self, bill_no, commission_rate=5.0):
        """Calculate and save commission to JSON for a specific bill"""
        if bill_no not in self.bills_data:
            return
        
        bill = self.bills_data[bill_no]
        
        # Get SUB TOTAL for commission calculation
        sub_total = bill.get("sub_total", 0)
        
        if sub_total > 0:
            # Calculate commission from SUB TOTAL
            commission_amount = (sub_total * commission_rate) / 100
        else:
            # If no SUB TOTAL, commission is 0
            commission_amount = 0
        
        # Update bill data with commission information
        bill.update({
            "commission_rate": commission_rate,
            "commission_amount": commission_amount,
            "commission_calculated_on": "sub_total",
            "commission_last_updated": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        })
        
        # Save to JSON
        self.save_data(self.bills_ref, self.bills_data)
        
        print(f"DEBUG: Commission saved for bill {bill_no} - Rate: {commission_rate}%, Amount: ₹{commission_amount:,.2f}")


    def update_commissions_for_all_bills(self):
        """Update commissions for all existing bills in the database"""
        confirm = messagebox.askyesno(
            "Update All Commissions",
            "This will calculate and save commissions for ALL bills in the database.\n\n"
            "Commission will be calculated from SUB TOTAL at the current rate.\n\n"
            "Continue?"
        )
        
        if not confirm:
            return
        
        commission_rate = float(self.commission_rate.get())
        updated_count = 0
        
        for bill_no in self.bills_data:
            old_commission = self.bills_data[bill_no].get("commission_amount", 0)
            self.calculate_and_save_commission(bill_no, commission_rate)
            new_commission = self.bills_data[bill_no].get("commission_amount", 0)
            
            if new_commission != old_commission:
                updated_count += 1
        
        messagebox.showinfo(
            "Commissions Updated",
            f"Commission calculation completed!\n\n"
            f"Updated: {updated_count} bills\n"
            f"Commission Rate: {commission_rate}%\n"
            f"Calculation Base: SUB TOTAL"
        )
        
    def resolve_pdf_path_updated(self, pdf_file_path, bill_no):
        """
        RESOLVE PDF PATH - Checks in ALL office folders within year-based folders
        Structure: Documents/InvoiceApp/Invoice_Bill_2025/[Office]/[Agent]/
        """
        # If it's already an absolute path and exists, return it
        if os.path.isabs(pdf_file_path) and os.path.exists(pdf_file_path):
            return pdf_file_path
        
        # Get the user's documents folder
        documents_dir = os.path.join(os.path.expanduser("~"), "Documents")
        invoice_app_dir = os.path.join(documents_dir, "InvoiceApp")
        
        # Define office folders
        office_folders = ["AP", "AFI", "AFF"]
        
        # Try to extract year from bill data
        bill_year = None
        if bill_no in self.bills_data:
            bill_date = self.bills_data[bill_no].get("bill_date", "")
            try:
                bill_date_obj = parse_date_flexible(bill_date)
                if bill_date_obj:
                    bill_year = str(bill_date_obj.year)
            except:
                pass
        
        # Extract filename from the path
        filename = os.path.basename(pdf_file_path)
        
        # Search strategy:
        # 1. First try specific year folder (if we know the year)
        # 2. Then try all year folders
        
        years_to_search = []
        if bill_year:
            years_to_search.append(f"Invoice_Bill_{bill_year}")
        
        # Add all existing year folders
        if os.path.exists(invoice_app_dir):
            for folder in os.listdir(invoice_app_dir):
                if folder.startswith("Invoice_Bill_") and os.path.isdir(os.path.join(invoice_app_dir, folder)):
                    if folder not in years_to_search:
                        years_to_search.append(folder)
        
        # Search in year folders
        for year_folder in years_to_search:
            year_folder_path = os.path.join(invoice_app_dir, year_folder)
            
            if not os.path.exists(year_folder_path):
                continue
                
            # Search in each office folder within the year folder
            for office_folder in office_folders:
                office_path = os.path.join(year_folder_path, office_folder)
                
                if not os.path.exists(office_path):
                    continue
                
                # Try direct path in office folder
                direct_path = os.path.join(office_path, filename)
                if os.path.exists(direct_path):
                    print(f"DEBUG: ✅ PDF found in {year_folder}/{office_folder}: {direct_path}")
                    return direct_path
                
                # Try with agent subfolder
                agent_name = self.bills_data.get(bill_no, {}).get("agent_name", "Unknown_Agent")
                if agent_name:
                    clean_agent_name = re.sub(r'[^\w\s-]', '', agent_name).strip().replace(' ', '_')
                    agent_folder_path = os.path.join(office_path, clean_agent_name)
                    agent_file_path = os.path.join(agent_folder_path, filename)
                    
                    if os.path.exists(agent_file_path):
                        print(f"DEBUG: ✅ PDF found in {year_folder}/{office_folder}/{clean_agent_name}: {agent_file_path}")
                        return agent_file_path
        
        # Fallback: Search recursively in all year folders
        if os.path.exists(invoice_app_dir):
            for year_folder in os.listdir(invoice_app_dir):
                if year_folder.startswith("Invoice_Bill_"):
                    year_path = os.path.join(invoice_app_dir, year_folder)
                    try:
                        for root, dirs, files in os.walk(year_path):
                            if filename in files:
                                found_path = os.path.join(root, filename)
                                print(f"DEBUG: ✅ PDF found via recursive search in {year_folder}: {found_path}")
                                return found_path
                    except Exception:
                        continue
        
        print(f"DEBUG: ❌ PDF NOT FOUND for bill {bill_no} in any year folder")
        return None

    def locate_missing_pdf_updated(self, bill_no):
        """Locate missing PDF file - searches ALL office folders"""
        if bill_no not in self.bills_data:
            return
        
        customer_name = self.bills_data[bill_no].get("customer_name", "")
        agent_name = self.bills_data[bill_no].get("agent_name", "")
        
        # Get the base folder
        documents_dir = os.path.join(os.path.expanduser("~"), "Documents")
        base_folder = os.path.join(documents_dir, "InvoiceApp", "Invoice_Bill")
        
        # Define office folders for the message
        office_folders = ["AP", "AFI", "AFF", "Invoice_Bills"]
        
        # Ask user to manually locate the file
        manual_choice = messagebox.askyesno(
            "📄 PDF Not Found", 
            f"PDF for bill {bill_no} not found in any office folder:\n"
            f"{', '.join(office_folders)}\n\n"
            f"Customer: {customer_name}\n"
            f"Agent: {agent_name}\n\n"
            f"Would you like to manually locate the PDF file?"
        )
        
        if manual_choice:
            # Start in the base invoice folder
            initial_dir = base_folder if os.path.exists(base_folder) else documents_dir
            
            file_path = filedialog.askopenfilename(
                title=f"Locate PDF for Bill {bill_no} - {customer_name}",
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
                initialdir=initial_dir
            )
            
            if file_path and os.path.exists(file_path):
                # Update the database with the correct path
                self.bills_data[bill_no]["pdf_file_name"] = file_path
                self.save_data(self.bills_ref, self.bills_data)
                
                # Open the PDF
                self.display_pdf(file_path)
                self.show_status_message("✅ PDF located and database updated!")
                return True
        
        return False


    def show_date_range_statement(self):
        """Show date range statement generation (placeholder)"""
        messagebox.showinfo("📅 Date Range Statement", "Date range statement feature will be implemented soon!")
        self.show_status_message("📅 Date range statement - Coming soon!")

    def show_settings(self):
        """Enhanced System Settings with Theme Selection & Keyboard Shortcuts Guide"""
        self.clear_screen()
        self.current_screen = "settings"
        self.create_navigation_bar()
        self.create_status_bar()

        # Main container
        main_container = tk.Frame(self.root, bg=self.colors['light_bg'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Header
        header_frame = self.create_modern_frame(main_container, "⚙️ SYSTEM SETTINGS")
        header_frame.pack(fill=tk.X, pady=(0, 20))

        tk.Label(
            header_frame,
            text="Configure application appearance and view keyboard shortcuts",
            font=("Segoe UI", 11),
            bg=self.colors['card_bg'],
            fg=self.colors['text_muted'],
            pady=10
        ).pack()

        # Create notebook for tabs
        notebook = ttk.Notebook(header_frame)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Tab 1: Appearance Settings
        appearance_tab = tk.Frame(notebook, bg=self.colors['card_bg'])
        notebook.add(appearance_tab, text="🎨 Appearance")

        appearance_container = tk.Frame(appearance_tab, bg=self.colors['card_bg'])
        appearance_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Theme Selection Section
        tk.Label(
            appearance_container,
            text="Select Theme",
            font=("Segoe UI", 14, "bold"),
            bg=self.colors['card_bg'],
            fg=self.colors['primary']
        ).pack(anchor="w", pady=(0, 20))

        # Theme description
        tk.Label(
            appearance_container,
            text="Choose your preferred color theme for the application",
            font=("Segoe UI", 10),
            bg=self.colors['card_bg'],
            fg=self.colors['text_muted']
        ).pack(anchor="w", pady=(0, 30))

        # Theme Selection Grid
        themes_frame = tk.Frame(appearance_container, bg=self.colors['card_bg'])
        themes_frame.pack(fill=tk.BOTH, expand=True)

        themes = [
            {
                "name": "Light",
                "icon": "☀️",
                "description": "Clean white theme for daytime use",
                "key": "light",
                "preview_colors": ["#C62828", "#283593", "#2e7d32"]  # Primary, Secondary, Success
            },
            {
                "name": "Dark",
                "icon": "🌙",
                "description": "Dark theme for reduced eye strain",
                "key": "dark",
                "preview_colors": ["#BB86FC", "#03DAC6", "#CF6679"]  # Primary, Secondary, Warning
            },
            {
                "name": "Blue",
                "icon": "🔵",
                "description": "Professional blue theme",
                "key": "blue",
                "preview_colors": ["#1976D2", "#2196F3", "#4CAF50"]  # Primary, Accent, Success
            },
            {
                "name": "Green",
                "icon": "🍃",
                "description": "Calming green theme",
                "key": "green",
                "preview_colors": ["#2E7D32", "#4CAF50", "#FFB300"]  # Primary, Secondary, Warning
            }
        ]

        # Create theme cards in a 2x2 grid
        for i, theme in enumerate(themes):
            row = i // 2
            col = i % 2
            
            # Theme card frame
            card_frame = tk.Frame(
                themes_frame,
                bg=self.colors['card_bg'],
                relief="solid",
                bd=1,
                width=250,
                height=180
            )
            card_frame.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")
            card_frame.grid_propagate(False)
            
            # Configure grid weights
            themes_frame.grid_rowconfigure(row, weight=1)
            themes_frame.grid_columnconfigure(col, weight=1)
            
            # Card content
            content_frame = tk.Frame(card_frame, bg=self.colors['card_bg'])
            content_frame.pack(expand=True, fill=tk.BOTH, padx=15, pady=15)
            
            # Theme header
            header_frame_card = tk.Frame(content_frame, bg=self.colors['card_bg'])
            header_frame_card.pack(fill=tk.X, pady=(0, 10))
            
            tk.Label(
                header_frame_card,
                text=theme["icon"],
                font=("Segoe UI", 16),
                bg=self.colors['card_bg']
            ).pack(side=tk.LEFT)
            
            tk.Label(
                header_frame_card,
                text=theme["name"],
                font=("Segoe UI", 12, "bold"),
                bg=self.colors['card_bg'],
                fg=self.colors['text_dark']
            ).pack(side=tk.LEFT, padx=5)
            
            # Current theme indicator
            if theme["key"] == self.current_theme:
                current_indicator = tk.Label(
                    header_frame_card,
                    text="✓ Current",
                    font=("Segoe UI", 8, "bold"),
                    bg=self.colors['card_bg'],
                    fg=self.colors['success']
                )
                current_indicator.pack(side=tk.RIGHT)
            
            # Theme description
            tk.Label(
                content_frame,
                text=theme["description"],
                font=("Segoe UI", 9),
                bg=self.colors['card_bg'],
                fg=self.colors['text_muted'],
                wraplength=200,
                justify=tk.LEFT
            ).pack(anchor="w", pady=(0, 10))
            
            # Color preview
            colors_frame = tk.Frame(content_frame, bg=self.colors['card_bg'])
            colors_frame.pack(fill=tk.X, pady=(0, 10))
            
            for color in theme["preview_colors"]:
                color_box = tk.Frame(
                    colors_frame,
                    bg=color,
                    width=25,
                    height=25,
                    relief="solid",
                    bd=1
                )
                color_box.pack(side=tk.LEFT, padx=2)
                color_box.pack_propagate(False)
            
            # Apply button
            apply_btn = tk.Button(
                content_frame,
                text="Apply Theme" if theme["key"] != self.current_theme else "✓ Applied",
                font=("Segoe UI", 9, "bold" if theme["key"] == self.current_theme else "normal"),
                bg=self.colors['primary'] if theme["key"] != self.current_theme else self.colors['success'],
                fg="white",
                relief="raised" if theme["key"] != self.current_theme else "sunken",
                bd=2,
                cursor="hand2",
                command=lambda t=theme["key"]: self.apply_theme_with_feedback(t)
            )
            apply_btn.pack(fill=tk.X, pady=(5, 0))
            
            # Add hover effect for non-current themes
            if theme["key"] != self.current_theme:
                def on_enter(e, btn=apply_btn):
                    btn.config(bg=self.darken_color(self.colors['primary'], 15))
                
                def on_leave(e, btn=apply_btn):
                    btn.config(bg=self.colors['primary'])
                
                apply_btn.bind("<Enter>", on_enter)
                apply_btn.bind("<Leave>", on_leave)

        # Tab 2: Keyboard Shortcuts Guide
        shortcuts_tab = tk.Frame(notebook, bg=self.colors['card_bg'])
        notebook.add(shortcuts_tab, text="⌨️ Keyboard Shortcuts")

        shortcuts_container = tk.Frame(shortcuts_tab, bg=self.colors['card_bg'])
        shortcuts_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Title
        tk.Label(
            shortcuts_container,
            text="Complete Keyboard Shortcuts Guide",
            font=("Segoe UI", 14, "bold"),
            bg=self.colors['card_bg'],
            fg=self.colors['primary']
        ).pack(anchor="w", pady=(0, 20))

        # Introduction
        intro_text = """Master these keyboard shortcuts to work faster and more efficiently in Angel Invoice Pro.
    All shortcuts work globally throughout the application."""
        
        tk.Label(
            shortcuts_container,
            text=intro_text,
            font=("Segoe UI", 10),
            bg=self.colors['card_bg'],
            fg=self.colors['text_dark'],
            wraplength=700,
            justify=tk.LEFT
        ).pack(anchor="w", pady=(0, 30))

        # Create scrollable frame for shortcuts
        shortcuts_canvas = tk.Canvas(shortcuts_container, bg=self.colors['card_bg'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(shortcuts_container, orient="vertical", command=shortcuts_canvas.yview)
        scrollable_frame = tk.Frame(shortcuts_canvas, bg=self.colors['card_bg'])

        scrollable_frame.bind(
            "<Configure>",
            lambda e: shortcuts_canvas.configure(scrollregion=shortcuts_canvas.bbox("all"))
        )

        shortcuts_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        shortcuts_canvas.configure(yscrollcommand=scrollbar.set)

        # Pack canvas and scrollbar
        shortcuts_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Bind mouse wheel for scrolling
        def _on_mousewheel(event):
            self.shortcuts_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        # Keep references for cleanup/local use
        self.shortcuts_canvas = shortcuts_canvas 
        self.mousewheel_handler = _on_mousewheel 

        # Bind locally to the canvas and its scrollable frame
        shortcuts_canvas.bind("<MouseWheel>", self.mousewheel_handler) 
        scrollable_frame.bind("<MouseWheel>", self.mousewheel_handler)
        
        

        # Keyboard Shortcuts Categories
        categories = [
            {
                "title": "🎯 NAVIGATION & BASIC CONTROLS",
                "shortcuts": [
                    ("TAB", "Move to next field/widget"),
                    ("SHIFT + TAB", "Move to previous field/widget"),
                    ("ENTER / RETURN", "Confirm selection / Move next"),
                    ("ESCAPE", "Clear focus / Close dialog"),
                    ("UP / DOWN ARROWS", "Navigate lists and dropdowns"),
                    ("LEFT / RIGHT ARROWS", "Navigate within fields"),
                    ("F1", "Open this keyboard shortcuts guide"),
                    ("F12", "Go to Dashboard from anywhere")
                ]
            },
            {
                "title": "⚡ QUICK ACTIONS",
                "shortcuts": [
                    ("CTRL + S", "Quick Save current data"),
                    ("CTRL + Z", "Undo last action"),
                    ("CTRL + Y", "Redo last undone action"),
                    ("CTRL + N", "Create New Bill/Invoice"),
                    ("CTRL + Q", "Quit application"),
                    ("CTRL + P", "Print current document/PDF"),
                    ("CTRL + F", "Find/Search in lists"),
                    ("DELETE", "Delete selected item(s)")
                ]
            },
            {
                "title": "🚀 FUNCTION KEY NAVIGATION",
                "shortcuts": [
                    ("F2", "Party Management - Add/Edit Customers"),
                    ("F3", "Product Management - Manage Products"),
                    ("F4", "Billing Center - Create/Edit Invoices"),
                    ("F5", "Refresh Data - Reload all data"),
                    ("F6", "Reports Dashboard - View analytics"),
                    ("F7", "System Settings - This screen"),
                    ("F8", "Toggle Fullscreen mode"),
                    ("F9", "Quick Calculator overlay")
                ]
            },
            {
                "title": "📊 BILLING & INVOICING",
                "shortcuts": [
                    ("ALT + A", "Add item to invoice table"),
                    ("ALT + R", "Remove selected item from invoice"),
                    ("ALT + U", "Update selected item in invoice"),
                    ("ALT + C", "Calculate totals"),
                    ("ALT + P", "Preview PDF before saving"),
                    ("CTRL + B", "Open bill browser"),
                    ("CTRL + I", "Insert new row in table"),
                    ("CTRL + D", "Duplicate selected item")
                ]
            },
            {
                "title": "🎨 THEME & APPEARANCE",
                "shortcuts": [
                    ("CTRL + ALT + L", "Switch to Light theme"),
                    ("CTRL + ALT + D", "Switch to Dark theme"),
                    ("CTRL + ALT + B", "Switch to Blue theme"),
                    ("CTRL + ALT + G", "Switch to Green theme"),
                    ("CTRL + ALT + T", "Show theme preview"),
                    ("CTRL + +", "Zoom in / Increase font size"),
                    ("CTRL + -", "Zoom out / Decrease font size"),
                    ("CTRL + 0", "Reset zoom to default")
                ]
            },
            {
                "title": "🔄 DATA MANAGEMENT",
                "shortcuts": [
                    ("CTRL + R", "Refresh current view"),
                    ("CTRL + E", "Export data to CSV/Excel"),
                    ("CTRL + O", "Open/Import data"),
                    ("CTRL + SHIFT + S", "Save As / Export PDF"),
                    ("CTRL + SHIFT + B", "Backup database"),
                    ("CTRL + SHIFT + R", "Restore from backup"),
                    ("ALT + F", "Filter/Sort data"),
                    ("ALT + E", "Edit selected record")
                ]
            },
            {
                "title": "🖥️ WINDOW & DISPLAY",
                "shortcuts": [
                    ("ALT + ENTER", "Toggle fullscreen mode"),
                    ("CTRL + W", "Close current tab/window"),
                    ("CTRL + TAB", "Switch between tabs"),
                    ("CTRL + SHIFT + TAB", "Switch tabs in reverse"),
                    ("ALT + LEFT", "Go back to previous screen"),
                    ("ALT + RIGHT", "Go forward to next screen"),
                    ("CTRL + M", "Minimize to system tray"),
                    ("CTRL + SHIFT + M", "Maximize window")
                ]
            },
            {
                "title": "🔧 ADVANCED & UTILITY",
                "shortcuts": [
                    ("CTRL + SHIFT + I", "Developer tools/info"),
                    ("CTRL + SHIFT + D", "Toggle debug mode"),
                    ("CTRL + SHIFT + L", "Clear application logs"),
                    ("CTRL + SHIFT + C", "Clear cache"),
                    ("ALT + SHIFT + R", "Reset application"),
                    ("CTRL + SHIFT + P", "Performance monitor"),
                    ("CTRL + SHIFT + H", "Toggle hidden features"),
                    ("CTRL + ALT + SHIFT + S", "System diagnostics")
                ]
            }
        ]

        # Display all categories
        current_row = 0
        for category in categories:
            # Category header
            cat_frame = tk.Frame(scrollable_frame, bg=self.colors['card_bg'])
            cat_frame.grid(row=current_row, column=0, sticky="ew", pady=(0, 15))
            scrollable_frame.grid_columnconfigure(0, weight=1)
            
            tk.Label(
                cat_frame,
                text=category["title"],
                font=("Segoe UI", 12, "bold"),
                bg=self.colors['card_bg'],
                fg=self.colors['primary'],
                anchor="w"
            ).pack(fill=tk.X, pady=(0, 10))
            
            # Shortcuts in this category
            for shortcut, description in category["shortcuts"]:
                shortcut_frame = tk.Frame(scrollable_frame, bg=self.colors['card_bg'])
                shortcut_frame.grid(row=current_row + 1, column=0, sticky="ew", pady=(0, 8))
                
                # Shortcut key styling
                shortcut_label = tk.Label(
                    shortcut_frame,
                    text=shortcut,
                    font=("Consolas", 10, "bold"),
                    bg=self.colors['light_bg'],
                    fg=self.colors['text_dark'],
                    relief="solid",
                    bd=1,
                    padx=10,
                    pady=2
                )
                shortcut_label.pack(side=tk.LEFT, padx=(0, 15))
                
                # Description
                tk.Label(
                    shortcut_frame,
                    text=description,
                    font=("Segoe UI", 10),
                    bg=self.colors['card_bg'],
                    fg=self.colors['text_dark'],
                    anchor="w"
                ).pack(side=tk.LEFT, fill=tk.X, expand=True)
                
                current_row += 1
            
            current_row += 2  # Extra space between categories

        # Tips and Tricks Section
        tips_frame = tk.Frame(scrollable_frame, bg=self.colors['light_bg'], relief="solid", bd=1)
        tips_frame.grid(row=current_row, column=0, sticky="ew", pady=(30, 0))
        
        tk.Label(
            tips_frame,
            text="💡 PRO TIPS",
            font=("Segoe UI", 11, "bold"),
            bg=self.colors['light_bg'],
            fg=self.colors['primary'],
            anchor="w"
        ).pack(fill=tk.X, padx=15, pady=(10, 5))
        
        tips = [
            "• Press SPACEBAR to click the currently focused button",
            "• Double-click any table row to edit that item",
            "• Right-click on tables for context menu options",
            "• Use arrow keys in dropdowns to navigate without mouse",
            "• CTRL+Z works for text fields, dropdowns, and checkboxes",
            "• Press ESC to cancel dialogs and clear focus",
            "• F1 shows context-sensitive help based on where you are",
            "• Most shortcuts can be used even when menus are open"
        ]
        
        for tip in tips:
            tk.Label(
                tips_frame,
                text=tip,
                font=("Segoe UI", 9),
                bg=self.colors['light_bg'],
                fg=self.colors['text_dark'],
                anchor="w"
            ).pack(fill=tk.X, padx=15, pady=2)

        # Bottom Action Buttons
        button_frame = tk.Frame(main_container, bg=self.colors['light_bg'])
        button_frame.pack(fill=tk.X, pady=20)

        # Print Shortcuts Button
        print_btn = self.create_modern_button(
            button_frame,
            "🖨️ Print Shortcuts Guide",
            self.print_shortcuts_guide,
            style="info",
            width=20,
            height=2
        )
        print_btn.pack(side=tk.LEFT, padx=5)

        # Practice Mode Button
        practice_btn = self.create_modern_button(
            button_frame,
            "🎮 Practice Shortcuts",
            self.practice_shortcuts_mode,
            style="success",
            width=18,
            height=2
        )
        practice_btn.pack(side=tk.LEFT, padx=5)

        # Close Button
        close_btn = self.create_modern_button(
            button_frame,
            "✖️ Close Settings",
            self.show_modern_dashboard,
            style="secondary",
            width=15,
            height=2
        )
        close_btn.pack(side=tk.RIGHT, padx=5)

        # Default Settings Button
        default_btn = self.create_modern_button(
            button_frame,
            "🔄 Reset to Defaults",
            self.reset_theme_to_default,
            style="warning",
            width=18,
            height=2
        )
        default_btn.pack(side=tk.RIGHT, padx=5)

        # Status indicator
        status_frame = tk.Frame(main_container, bg=self.colors['light_bg'])
        status_frame.pack(fill=tk.X, pady=10)

        self.settings_status = tk.Label(
            status_frame,
            text=f"Current Theme: {self.current_theme.title()} • Press F1 anytime for help",
            font=("Segoe UI", 9),
            bg=self.colors['light_bg'],
            fg=self.colors['success']
        )
        self.settings_status.pack()

        self.show_status_message("⚙️ System Settings loaded - Configure theme and learn keyboard shortcuts")

    def apply_theme_with_feedback(self, theme_name):
        """Apply theme with visual feedback"""
        old_theme = self.current_theme
        
        if theme_name == old_theme:
            messagebox.showinfo("Theme", f"{theme_name.title()} theme is already active!")
            return
        
        # Show applying message
        self.settings_status.config(
            text=f"🔄 Applying {theme_name.title()} theme...",
            fg=self.colors['warning']
        )
        self.root.update()
        
        # Change theme
        self.change_theme(theme_name)
        
        # Update feedback
        self.settings_status.config(
            text=f"✅ {theme_name.title()} theme applied successfully! • Press F7 to reopen settings",
            fg=self.colors['success']
        )
        
        # Play success sound if enabled
        try:
            import winsound
            winsound.MessageBeep()
        except:
            pass

    def reset_theme_to_default(self):
        """Reset theme to default (Light)"""
        confirm = messagebox.askyesno(
            "Reset Theme",
            "Reset theme to default Light mode?\n\nAll customizations will be lost.",
            icon='question'
        )
        
        if confirm:
            self.apply_theme_with_feedback('light')

    def print_shortcuts_guide(self):
        """Print or export keyboard shortcuts guide"""
        try:
            from fpdf import FPDF
            import os
            
            # Create PDF
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", "B", 16)
            
            # Title
            pdf.cell(0, 10, "Angel Invoice Pro - Keyboard Shortcuts Guide", ln=True, align="C")
            pdf.ln(10)
            
            pdf.set_font("Arial", "", 12)
            pdf.cell(0, 10, f"Generated on: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", ln=True)
            pdf.cell(0, 10, f"Current Theme: {self.current_theme.title()}", ln=True)
            pdf.ln(10)
            
            # Create directory for guides
            documents_dir = os.path.join(os.path.expanduser("~"), "Documents")
            guides_dir = os.path.join(documents_dir, "InvoiceApp", "User_Guides")
            os.makedirs(guides_dir, exist_ok=True)
            
            # Save PDF
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            pdf_path = os.path.join(guides_dir, f"Shortcuts_Guide_{timestamp}.pdf")
            pdf.output(pdf_path)
            
            # Show success
            messagebox.showinfo(
                "Guide Printed",
                f"Keyboard shortcuts guide saved as PDF:\n{pdf_path}\n\nYou can print this file for reference."
            )
            
            # Open the PDF
            self.display_pdf(pdf_path)
            
        except Exception as e:
            messagebox.showerror("Print Error", f"Failed to print guide:\n{str(e)}")

    def practice_shortcuts_mode(self):
        """Open practice mode for learning shortcuts"""
        practice_window = tk.Toplevel(self.root)
        practice_window.title("🎮 Practice Keyboard Shortcuts")
        practice_window.geometry("600x400")
        practice_window.configure(bg=self.colors['light_bg'])
        practice_window.transient(self.root)
        practice_window.grab_set()
        
        # Center window
        practice_window.update_idletasks()
        x = (practice_window.winfo_screenwidth() // 2) - (600 // 2)
        y = (practice_window.winfo_screenheight() // 2) - (400 // 2)
        practice_window.geometry(f"600x400+{x}+{y}")
        
        # Header
        header = tk.Label(
            practice_window,
            text="🎮 Practice Keyboard Shortcuts",
            font=("Segoe UI", 16, "bold"),
            bg=self.colors['primary'],
            fg=self.colors['text_light'],
            pady=15
        )
        header.pack(fill=tk.X)
        
        # Instructions
        content = tk.Frame(practice_window, bg=self.colors['light_bg'], padx=20, pady=20)
        content.pack(fill=tk.BOTH, expand=True)
        
        instructions = """PRACTICE MODE
        
    Try these essential shortcuts:
        
    1. Navigation:
    • TAB - Move between fields
    • SHIFT+TAB - Move backward
    • ENTER - Confirm/Select
    • ESCAPE - Cancel/Close
        
    2. Actions:
    • CTRL+S - Save
    • CTRL+Z - Undo
    • CTRL+N - New Invoice
        
    3. Function Keys:
    • F1 - Help
    • F2 - Parties
    • F3 - Products
    • F4 - Billing
        
    Press the shortcut keys and see the action!
    """
        
        tk.Label(
            content,
            text=instructions,
            font=("Segoe UI", 11),
            bg=self.colors['light_bg'],
            fg=self.colors['text_dark'],
            justify=tk.LEFT
        ).pack(anchor="w", pady=10)
        
        # Practice area
        practice_frame = tk.Frame(content, bg=self.colors['card_bg'], relief="solid", bd=1)
        practice_frame.pack(fill=tk.BOTH, expand=True, pady=20)
        
        self.practice_output = tk.Label(
            practice_frame,
            text="Press any shortcut key combination...",
            font=("Consolas", 12),
            bg=self.colors['card_bg'],
            fg=self.colors['text_muted'],
            pady=20
        )
        self.practice_output.pack(expand=True)
        
        # Key press handler
        def on_key_press(event):
            keys = []
            if event.state & 0x4:  # Control
                keys.append("CTRL")
            if event.state & 0x8:  # Alt
                keys.append("ALT")
            if event.state & 0x1:  # Shift
                keys.append("SHIFT")
            
            if event.keysym not in ["Control_L", "Control_R", "Alt_L", "Alt_R", "Shift_L", "Shift_R"]:
                keys.append(event.keysym.upper())
            
            key_combo = " + ".join(keys)
            
            # Known shortcuts
            known_shortcuts = {
                "CTRL + S": "✅ Save Action",
                "CTRL + Z": "✅ Undo Action", 
                "CTRL + N": "✅ New Invoice",
                "F1": "✅ Help Screen",
                "F2": "✅ Party Management",
                "F3": "✅ Product Management",
                "F4": "✅ Billing Center",
                "TAB": "✅ Next Field",
                "SHIFT + TAB": "✅ Previous Field",
                "ESCAPE": "✅ Clear Focus"
            }
            
            if key_combo in known_shortcuts:
                self.practice_output.config(
                    text=f"🎯 {key_combo}\n\n{known_shortcuts[key_combo]}",
                    fg=self.colors['success']
                )
            else:
                self.practice_output.config(
                    text=f"❓ {key_combo}\n\nNot a registered shortcut",
                    fg=self.colors['warning']
                )
        
        practice_window.bind("<KeyPress>", on_key_press)
        
        # Close button
        close_btn = self.create_modern_button(
            content,
            "✖️ Close Practice",
            practice_window.destroy,
            style="secondary",
            width=15,
            height=2
        )
        close_btn.pack(pady=10)


    def change_theme(self, theme_name):
        """Change the application theme"""
        # Define color schemes for different themes
        themes = {
            'light': {
                'primary': "#C62828",
                'secondary': '#283593',
                'accent': '#3949ab',
                'success': '#2e7d32',
                'warning': '#d32f2f',
                'info': '#0288d1',
                'light_bg': '#f5f5f5',
                'dark_bg': '#121212',
                'card_bg': '#ffffff',
                'text_light': '#ffffff',
                'text_dark': '#212121',
                'text_muted': '#757575',
                'focus_border': '#ff4081',
                'hover_effect': '#e3f2fd'
            },
            'dark': {
                'primary': "#BB86FC",
                'secondary': '#03DAC6',
                'accent': '#3700B3',
                'success': '#03DAC6',
                'warning': '#CF6679',
                'info': '#018786',
                'light_bg': '#121212',
                'dark_bg': '#000000',
                'card_bg': '#1E1E1E',
                'text_light': '#FFFFFF',
                'text_dark': '#E0E0E0',
                'text_muted': '#9E9E9E',
                'focus_border': '#BB86FC',
                'hover_effect': '#2C2C2C'
            },
            'blue': {
                'primary': "#1976D2",
                'secondary': '#2196F3',
                'accent': '#2196F3',
                'success': '#4CAF50',
                'warning': '#FF9800',
                'info': '#00BCD4',
                'light_bg': '#E3F2FD',
                'dark_bg': '#1565C0',
                'card_bg': '#FFFFFF',
                'text_light': '#FFFFFF',
                'text_dark': '#212121',
                'text_muted': '#757575',
                'focus_border': '#1976D2',
                'hover_effect': '#E8F5E9'
            },
            'green': {
                'primary': "#2E7D32",
                'secondary': '#4CAF50',
                'accent': '#388E3C',
                'success': '#4CAF50',
                'warning': '#FFB300',
                'info': '#0097A7',
                'light_bg': '#E8F5E9',
                'dark_bg': '#1B5E20',
                'card_bg': '#FFFFFF',
                'text_light': '#FFFFFF',
                'text_dark': '#212121',
                'text_muted': '#757575',
                'focus_border': '#2E7D32',
                'hover_effect': '#F1F8E9'
            }
        }
        
        # Update current theme
        self.current_theme = theme_name
        
        # Update colors dictionary
        if theme_name in themes:
            self.colors = themes[theme_name]
        
        # Store the theme preference (optional - save to file)
        try:
            # Save theme preference to a settings file
            settings = {}
            try:
                with open("settings.json", "r") as f:
                    settings = json.load(f)
            except FileNotFoundError:
                pass
            
            settings["theme"] = theme_name
            with open("settings.json", "w") as f:
                json.dump(settings, f, indent=4)
        except Exception as e:
            print(f"Error saving theme preference: {e}")

    def initialize_theme_colors(self):
        """Initialize colors based on current theme"""
        themes = {
            'light': {
                'primary': "#C62828",
                'secondary': '#283593',
                'accent': '#3949ab',
                'success': '#2e7d32',
                'warning': '#d32f2f',
                'info': '#0288d1',
                'light_bg': '#f5f5f5',
                'dark_bg': '#121212',
                'card_bg': '#ffffff',
                'text_light': '#ffffff',
                'text_dark': '#212121',
                'text_muted': '#757575',
                'focus_border': '#ff4081',
                'hover_effect': '#e3f2fd'
            },
            'dark': {
                'primary': "#BB86FC",
                'secondary': '#03DAC6',
                'accent': '#3700B3',
                'success': '#03DAC6',
                'warning': '#CF6679',
                'info': '#018786',
                'light_bg': '#121212',
                'dark_bg': '#000000',
                'card_bg': '#1E1E1E',
                'text_light': '#FFFFFF',
                'text_dark': '#E0E0E0',
                'text_muted': '#9E9E9E',
                'focus_border': '#BB86FC',
                'hover_effect': '#2C2C2C'
            },
            'blue': {
                'primary': "#1976D2",
                'secondary': '#2196F3',
                'accent': '#2196F3',
                'success': '#4CAF50',
                'warning': '#FF9800',
                'info': '#00BCD4',
                'light_bg': '#E3F2FD',
                'dark_bg': '#1565C0',
                'card_bg': '#FFFFFF',
                'text_light': '#FFFFFF',
                'text_dark': '#212121',
                'text_muted': '#757575',
                'focus_border': '#1976D2',
                'hover_effect': '#E8F5E9'
            },
            'green': {
                'primary': "#2E7D32",
                'secondary': '#4CAF50',
                'accent': '#388E3C',
                'success': '#4CAF50',
                'warning': '#FFB300',
                'info': '#0097A7',
                'light_bg': '#E8F5E9',
                'dark_bg': '#1B5E20',
                'card_bg': '#FFFFFF',
                'text_light': '#FFFFFF',
                'text_dark': '#212121',
                'text_muted': '#757575',
                'focus_border': '#2E7D32',
                'hover_effect': '#F1F8E9'
            }
        }
        
        # Set colors based on current theme
        self.colors = themes.get(self.current_theme, themes['light'])
    
        



# Main application entry point
if __name__ == "__main__":
    root = tk.Tk()
    app = ModernInvoiceApp(root)
    root.mainloop()