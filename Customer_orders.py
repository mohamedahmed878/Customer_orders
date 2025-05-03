import tkinter as tk
from tkinter import messagebox, ttk
import openpyxl
import os
from datetime import datetime
import sys # Needed for platform check

# Get the directory where the script is located
script_dir = os.path.dirname(os.path.abspath(__file__))
# Excel file name (relative to the script's directory)
file_name = os.path.join(script_dir, 'customer_orders.xlsx')


# Create new Excel file with headers if it doesn't exist
if not os.path.exists(file_name):
    try:
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = 'Orders'
        # Set sheet direction to Right-to-Left
        sheet.sheet_view.rightToLeft = True
        sheet.append(['الاسم بالكامل', 'رقم المحمول', 'الرقم البديل', 'المحافظة',
                    'العنوان بالتفصيل', 'الكمية', 'سعر بيع القطعة',
                    'ملاحظات', 'الوقت'])
        wb.save(file_name)
    except PermissionError:
         print(f"خطأ: لا يمكن إنشاء ملف '{os.path.basename(file_name)}'. قد يكون مفتوحاً أو لا تملك الصلاحيات.")
         # Optionally show a messagebox if GUI is already running, but usually this runs before mainloop
         # messagebox.showerror("خطأ تهيئة", f"لا يمكن إنشاء ملف '{os.path.basename(file_name)}'. قد يكون مفتوحاً أو لا تملك الصلاحيات.")
         sys.exit(1) # Exit if we can't create the file
    except Exception as e:
        print(f"خطأ غير متوقع عند إنشاء الملف: {e}")
        # messagebox.showerror("خطأ تهيئة", f"خطأ غير متوقع عند إنشاء الملف: {e}")
        sys.exit(1)

# Available governorates
governorates = ['القاهرة', 'الجيزة', 'الإسكندرية', 'الدقهلية', 'الشرقية', 'الغربية', 'البحيرة', 'المنوفية', 'القليوبية', 'كفر الشيخ', 'دمياط', 'بورسعيد', 'الإسماعيلية', 'السويس', 'شمال سيناء', 'جنوب سيناء', 'البحر الأحمر', 'الفيوم', 'بني سويف', 'المنيا', 'أسيوط', 'سوهاج', 'قنا', 'الأقصر', 'أسوان', 'الوادي الجديد', 'مطروح']
# Available quantities
quantities = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
# Background colors
entry_normal_bg = "white"
entry_error_bg = "#FFDDDD"  # Light red

# Save data to Excel
def save_data():
    full_name = full_name_entry.get()
    mobile_number = mobile_number_entry.get()
    alternative_number = alternative_number_entry.get()
    governorate = governorate_var.get()
    detailed_address = detailed_address_entry.get()
    quantity_str = quantity_var.get() # Get as string first for validation
    price_per_piece_str = price_per_piece_entry.get() # Get as string first
    notes = notes_entry.get("1.0", tk.END).strip()
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    required_fields_widgets = {
        "الاسم بالكامل": full_name_entry,
        "رقم المحمول": mobile_number_entry,
        "المحافظة": governorate_combo,
        "العنوان بالتفصيل": detailed_address_entry,
        "الكمية": quantity_combo,
        "سعر بيع القطعة": price_per_piece_entry
    }

    missing_fields_labels = []
    has_error = False # Flag to track if any error occurred

    # Reset backgrounds and check for missing required fields
    for label, widget in required_fields_widgets.items():
        # Use 'winfo_class' to check if it's a Combobox or Entry
        if widget.winfo_class() == 'TCombobox':
            widget.configure(style="TCombobox") # Reset style (assuming default is TCombobox)
            if not widget.get():
                missing_fields_labels.append(label)
                # Combobox doesn't have a 'bg' option, style might need custom error state if desired
                # For simplicity, we'll just list it as missing.
                has_error = True
        elif widget.winfo_class() == 'TEntry':
            widget.config(bg=entry_normal_bg)
            if not widget.get():
                missing_fields_labels.append(label)
                widget.config(bg=entry_error_bg)
                has_error = True
        # Add handling for Text widget if it becomes required
        # elif widget.winfo_class() == 'Text':
        #     widget.config(bg=entry_normal_bg)
        #     if not widget.get("1.0", tk.END).strip(): # Check if Text is empty
        #          missing_fields_labels.append(label)
        #          widget.config(bg=entry_error_bg)
        #          has_error = True


    if missing_fields_labels:
        messagebox.showerror("خطأ", "الرجاء ملء الحقول المطلوبة: \n" + "\n".join(missing_fields_labels))
        # No return here yet, check type errors next

    # Validate numeric fields only if they are not empty
    quantity = None
    price_per_piece = None
    type_error = False

    if quantity_str: # If quantity field is not empty
        try:
            quantity = int(quantity_str)
            # quantity_combo.configure(style="TCombobox") # Reset style if validation passes
        except ValueError:
            messagebox.showerror("خطأ", "الكمية يجب أن تكون رقماً صحيحاً.")
            # Highlight - again, styling combobox background is tricky, list error instead
            # quantity_combo.config(bg=entry_error_bg) # This won't work directly on ttk.Combobox
            has_error = True
            type_error = True # Mark that a type error occurred

    if price_per_piece_str: # If price field is not empty
        try:
            price_per_piece = float(price_per_piece_str)
            price_per_piece_entry.config(bg=entry_normal_bg) # Reset background
        except ValueError:
            messagebox.showerror("خطأ", "سعر بيع القطعة يجب أن يكون رقماً.")
            price_per_piece_entry.config(bg=entry_error_bg) # Highlight Entry field
            has_error = True
            type_error = True # Mark that a type error occurred

    # If any validation error occurred (missing or type error), stop.
    if has_error:
        return

    # --- Try saving to Excel ---
    try:
        wb = openpyxl.load_workbook(file_name)
        sheet = wb.active
        # Ensure sheet direction is RTL (might be needed if file was created differently)
        if not sheet.sheet_view.rightToLeft:
             sheet.sheet_view.rightToLeft = True

        sheet.append([full_name, mobile_number, alternative_number, governorate,
                    detailed_address, quantity, price_per_piece,
                    notes, timestamp])
        wb.save(file_name)

    # ***** MODIFICATION: Catch PermissionError specifically *****
    except PermissionError:
        messagebox.showerror("خطأ في الحفظ", f"لا يمكن حفظ البيانات.\nيبدو أن ملف الإكسل '{os.path.basename(file_name)}' مفتوح حالياً في برنامج آخر.\n\nالرجاء إغلاق الملف ثم حاول الإرسال مرة أخرى.")
        return # Stop processing
    # ***** END MODIFICATION *****

    except Exception as e:
        messagebox.showerror("خطأ في الملف", f"حدث خطأ غير متوقع أثناء محاولة حفظ البيانات في ملف Excel.\nالخطأ: {e}")
        return # Stop processing

    # If saving was successful
    messagebox.showinfo("نجاح", "تم إرسال الطلب وحفظه في ملف Excel بنجاح!")
    clear_fields()


def clear_fields():
    # Clear entry fields
    full_name_entry.delete(0, tk.END)
    mobile_number_entry.delete(0, tk.END)
    alternative_number_entry.delete(0, tk.END)
    detailed_address_entry.delete(0, tk.END)
    price_per_piece_entry.delete(0, tk.END)
    # Reset comboboxes
    governorate_var.set(governorates[0])
    quantity_var.set(quantities[0])
    # Clear text widget
    notes_entry.delete("1.0", tk.END)

    # Reset background/styles for all input widgets
    full_name_entry.config(bg=entry_normal_bg)
    mobile_number_entry.config(bg=entry_normal_bg)
    alternative_number_entry.config(bg=entry_normal_bg)
    detailed_address_entry.config(bg=entry_normal_bg)
    price_per_piece_entry.config(bg=entry_normal_bg)
    notes_entry.config(bg=entry_normal_bg)
    governorate_combo.configure(style="TCombobox") # Reset style
    quantity_combo.configure(style="TCombobox")   # Reset style


# GUI setup
root = tk.Tk()
root.title("متجر الملابس الذكية - تسجيل طلب جديد")

# Adjust window size (reduced height as buttons were removed)
root.geometry("600x700") # Adjusted height

# Styling with ttk
style = ttk.Style()
# Configure styles for Right-to-Left alignment if possible
# Note: ttk respects system locale often, but explicit alignment helps
style.configure("TLabel", font=("Helvetica", 12), anchor="e") # Anchor East (right)
style.configure("TButton", font=("Helvetica", 12, "bold"))
style.configure("TEntry", font=("Helvetica", 12)) # Default entry style
style.configure("TCombobox", font=("Helvetica", 12)) # Default combobox style
# You might need more specific styling for RTL languages depending on OS/Theme
# style.map("TEntry", fieldbackground=[("rtl", "white")]) # Example concept

# --- Main Frame for Padding ---
main_frame = ttk.Frame(root, padding="10 10 10 10")
main_frame.pack(fill=tk.BOTH, expand=True)

# Configure grid weights for responsiveness
main_frame.columnconfigure(1, weight=1)


# --- Personal Information ---
personal_label = ttk.Label(main_frame, text="البيانات الشخصية", font=("Helvetica", 14, "bold"), anchor="center")
personal_label.grid(row=0, column=0, columnspan=2, pady=(0, 10), sticky="ew")

ttk.Label(main_frame, text=":الاسم بالكامل").grid(row=1, column=0, padx=5, pady=2, sticky=tk.E)
full_name_entry = ttk.Entry(main_frame, style="TEntry", justify='right')
full_name_entry.grid(row=1, column=1, padx=5, pady=2, sticky=tk.EW)

ttk.Label(main_frame, text=":رقم المحمول").grid(row=2, column=0, padx=5, pady=2, sticky=tk.E)
mobile_number_entry = ttk.Entry(main_frame, style="TEntry", justify='right')
mobile_number_entry.grid(row=2, column=1, padx=5, pady=2, sticky=tk.EW)

ttk.Label(main_frame, text=":(اختياري) الرقم البديل").grid(row=3, column=0, padx=5, pady=2, sticky=tk.E)
alternative_number_entry = ttk.Entry(main_frame, style="TEntry", justify='right')
alternative_number_entry.grid(row=3, column=1, padx=5, pady=2, sticky=tk.EW)

# Separator
ttk.Separator(main_frame, orient='horizontal').grid(row=4, column=0, columnspan=2, pady=10, sticky='ew')

# --- Shipping Information ---
shipping_label = ttk.Label(main_frame, text="بيانات الشحن", font=("Helvetica", 14, "bold"), anchor="center")
shipping_label.grid(row=5, column=0, columnspan=2, pady=(0, 10), sticky="ew")

ttk.Label(main_frame, text=":المحافظة").grid(row=6, column=0, padx=5, pady=2, sticky=tk.E)
governorate_var = tk.StringVar(value=governorates[0])
governorate_combo = ttk.Combobox(main_frame, textvariable=governorate_var, values=governorates, style="TCombobox", justify='right', state='readonly') # Make readonly
governorate_combo.grid(row=6, column=1, padx=5, pady=2, sticky=tk.EW)

ttk.Label(main_frame, text=":العنوان بالتفصيل").grid(row=7, column=0, padx=5, pady=2, sticky=tk.E)
detailed_address_entry = ttk.Entry(main_frame, style="TEntry", justify='right')
detailed_address_entry.grid(row=7, column=1, padx=5, pady=2, sticky=tk.EW)

# Separator
ttk.Separator(main_frame, orient='horizontal').grid(row=8, column=0, columnspan=2, pady=10, sticky='ew')

# --- Order Details ---
order_label = ttk.Label(main_frame, text="تفاصيل الطلب", font=("Helvetica", 14, "bold"), anchor="center")
order_label.grid(row=9, column=0, columnspan=2, pady=(0, 10), sticky="ew")

ttk.Label(main_frame, text=":الكمية").grid(row=10, column=0, padx=5, pady=2, sticky=tk.E)
quantity_var = tk.IntVar(value=quantities[0])
quantity_combo = ttk.Combobox(main_frame, textvariable=quantity_var, values=quantities, style="TCombobox", justify='right', state='readonly') # Make readonly
quantity_combo.grid(row=10, column=1, padx=5, pady=2, sticky=tk.EW)

ttk.Label(main_frame, text=":سعر بيع القطعة").grid(row=11, column=0, padx=5, pady=2, sticky=tk.E)
price_per_piece_entry = ttk.Entry(main_frame, style="TEntry", justify='right')
price_per_piece_entry.grid(row=11, column=1, padx=5, pady=2, sticky=tk.EW)

# Separator
ttk.Separator(main_frame, orient='horizontal').grid(row=12, column=0, columnspan=2, pady=10, sticky='ew')

# --- Optional Information ---
optional_label = ttk.Label(main_frame, text="ملاحظات (اختياري)", font=("Helvetica", 14, "bold"), anchor="center")
optional_label.grid(row=13, column=0, columnspan=2, pady=(0, 10), sticky="ew")

ttk.Label(main_frame, text=":ملاحظات").grid(row=14, column=0, padx=5, pady=2, sticky=tk.NE) # Align top-right
notes_entry = tk.Text(main_frame, font=("Helvetica", 12), height=4, width=30, wrap=tk.WORD, relief=tk.SOLID, borderwidth=1) # Added border
notes_entry.grid(row=14, column=1, padx=5, pady=2, sticky=tk.EW)
# Simple border for Text widget
# For RTL text input in Text widget, it often depends on the underlying Tk version and system settings.


# Separator
ttk.Separator(main_frame, orient='horizontal').grid(row=15, column=0, columnspan=2, pady=15, sticky='ew')


# --- Buttons ---
# Frame for buttons at the bottom
button_frame = ttk.Frame(main_frame)
button_frame.grid(row=16, column=0, columnspan=2, pady=(5, 0), sticky="ew")
button_frame.columnconfigure(0, weight=1) # Make buttons expand
button_frame.columnconfigure(1, weight=1)


# Original Save Button
save_button = ttk.Button(button_frame, text="إرسال الطلب", command=save_data, style="TButton")
# Function to animate the button
def animate_button(event):
    event.widget.config(relief=tk.SUNKEN)
    root.after(100, lambda: event.widget.config(relief=tk.RAISED))
save_button.bind("<Button-1>", animate_button)
save_button.grid(row=0, column=0, padx=(0, 5), pady=5, sticky=tk.EW) # Place in grid

# Original Add Another Button
add_another_button = ttk.Button(button_frame, text="+ إضافة طلب آخر", command=clear_fields, style="TButton")
add_another_button.grid(row=0, column=1, padx=(5, 0), pady=5, sticky=tk.EW) # Place in grid


# Make the main content area resizeable if needed
main_frame.rowconfigure(14, weight=1) # Allow notes section to expand vertically if desired


root.mainloop()