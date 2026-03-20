import customtkinter as ctk
from tkinter import ttk, Entry, filedialog, Menu
import openpyxl
import csv
import sys
import os

# Settings 
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.title("Lee's Sales Tax Calculator")
app.geometry("900x750")
app.iconbitmap(r"C:\Users\lucas.roff.ITCMI\OneDrive - Inter-Tribal Council of Michigan, Inc\Desktop\Coding\MyApps\Python\Lee's Sales Tax App\CodingLogo.ico")

edit_box = None

# Container for pages 
container = ctk.CTkFrame(app)
container.pack(fill="both", expand=True)
container.pack_propagate(False)
pages = {}

def show_page(page_name):
    pages[page_name].tkraise()



# Homepage Page 
page_home = ctk.CTkFrame(container)
page_home.place(relwidth=1, relheight=1)
pages["home"] = page_home

lbl_home = ctk.CTkLabel(page_home, text="Welcome to Lee's Sales Tax Calculator!", font=("Zefani", 35, "bold"), text_color="#0bf565")
lbl_home.pack(pady=(150, 20))

lbl_intro = ctk.CTkLabel(
    page_home,
    text=(
        "Upload an Excel file only containing a 'Product' and 'Price' column\n"
        "to calculate taxes automatically. You can also manually add, edit,\n"
        "copy, or paste entries in the table."
        " Enjoy!\n"
    ),
    font=("Arial", 14),
    justify="center"
)
lbl_intro.pack(pady=(0, 80))

btn_start = ctk.CTkButton(
    page_home,
    text="Open Calculator",
    width=220,
    height=50,
    fg_color="#0fa84a",
    hover_color="#0a7a7e",
    command=lambda: show_page("calculator")
)
btn_start.pack()


# 
# Calculator Page 

page_calc = ctk.CTkFrame(container)
page_calc.place(relwidth=1, relheight=1)
pages["calculator"] = page_calc

#  Functions 
def recalc_totals():
    total_tax = 0
    total_sum = 0
    try:
        tax_rate = float(entry_tax.get()) / 100
    except:
        tax_rate = 0.06
    for item in tree.get_children():
        values = list(tree.item(item, "values"))
        try:
            price = float(values[1])
        except:
            price = 0.0
        tax = price * tax_rate
        total = price + tax  
        values[2] = f"{tax:.2f}"
        values[3] = f"{total:.2f}"
        tree.item(item, values=values)
        total_tax += tax
        total_sum += total
    lbl_tax_total.configure(text=f"Sales Tax Total: ${total_tax:.2f}",)
    lbl_total.configure(text=f"Grand Total: ${total_sum:.2f}")

def add_row_popup():
    popup = ctk.CTkToplevel(app)
    popup.title("Add New Entry - Sales Tax Calculator")
    popup.geometry("450x350")
    
    ctk.CTkLabel(popup, text="Product Name:", text_color="#00f566").pack(pady=5)
    entry_product = ctk.CTkEntry(popup, width=220)
    entry_product.pack(pady=5)

    ctk.CTkLabel(popup, text="Price:", text_color="#f51d00").pack(pady=5)
    entry_price = ctk.CTkEntry(popup, width=220)
    entry_price.pack(pady=5)

    def submit_row():
        product = entry_product.get().strip()
        price_text = entry_price.get().strip()

        if product == "":
            product = "New Product"

        try:
            price = float(price_text)
        except:
            price = 0.0

        try:
            tax_rate = float(entry_tax.get()) / 100
        except:
            tax_rate = 0.06

        tax = price * tax_rate
        total = price + tax  

        tree.insert("", "end", values=(
            product,
            f"{price:.2f}",
            f"{tax:.2f}",
            f"{total:.2f}"
        ))

        recalc_totals()
        popup.destroy()

    ctk.CTkButton(popup, text="Add New Entry", fg_color="#f5550b", hover_color="#970a0a", command=submit_row).pack(pady=15) 
frame_buttons = ctk.CTkFrame(page_calc)
frame_buttons.pack(pady=(0, 20), padx=20, fill="x")
ctk.CTkButton(frame_buttons, text="Add New Entry",
              command=add_row_popup,
              fg_color="#f5550b", hover_color="#d90606").pack(padx=10, pady= 10)

def add_product():
    tree.insert("", "end", values=("New Product", "0.00", "0.00", "0.00"))

def delete_product():
    for item in tree.selection():
        tree.delete(item)
    recalc_totals()

def on_double_click(event):
    global edit_box
    if edit_box:
        return
    region = tree.identify("region", event.x, event.y)
    if region != "cell":
        return
    row_id = tree.identify_row(event.y)
    column = tree.identify_column(event.x)
    col_index = int(column.replace("#","")) - 1
    x, y, width, height = tree.bbox(row_id, column)
    value = tree.item(row_id, "values")[col_index]
    edit_box = Entry(tree)
    edit_box.place(x=x, y=y, width=width, height=height)
    edit_box.insert(0, value)
    edit_box.focus_set()
    def save_edit(event=None):
        global edit_box
        new_value = edit_box.get()
        values = list(tree.item(row_id, "values"))
        if col_index >= 1:
            try:
                values[col_index] = f"{float(new_value):.2f}"
            except:
                values[col_index] = "0.00"
        else:
            values[col_index] = new_value
        tree.item(row_id, values=values)
        edit_box.destroy()
        edit_box = None
        recalc_totals()
    edit_box.bind("<Return>", save_edit)
    edit_box.bind("<FocusOut>", save_edit)

def save_to_excel():
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files","*.xlsx *.xls")])
    if not file_path:
        return
    data = []
    for item in tree.get_children():
        values = tree.item(item,"values")
        data.append({"Product": values[0], "Price": float(values[1]), "Tax": float(values[2]), "Total": float(values[3])})  
    pd.DataFrame(data).to_excel(file_path, index=False)

def load_excel_dynamic():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel/CSV Files", "*.xlsx *.csv")]
    )
    if not file_path:
        return

    try:
        # Clear existing rows
        for row in tree.get_children():
            tree.delete(row)

        tax_rate = float(entry_tax.get()) / 100
         
        # Handle XLSX (Excel)
        
        if file_path.endswith(".xlsx"):
            wb = openpyxl.load_workbook(file_path)
            sheet = wb.active

            headers = [cell.value for cell in sheet[1]]

            if "Product" not in headers or "Price" not in headers:
                raise ValueError("File must have 'Product' and 'Price' columns.")

            product_idx = headers.index("Product")
            price_idx = headers.index("Price")

            for row in sheet.iter_rows(min_row=2, values_only=True):
                product = str(row[product_idx])
                try:
                    price = float(row[price_idx])
                except:
                    price = 0.0

                tax = price * tax_rate
                total = price + tax

                tree.insert("", "end", values=(
                    product,
                    f"{price:.2f}",
                    f"{tax:.2f}",
                    f"{total:.2f}"
                ))

         
        # Handle CSV
        
        elif file_path.endswith(".csv"):
            with open(file_path, newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f)

                if "Product" not in reader.fieldnames or "Price" not in reader.fieldnames:
                    raise ValueError("CSV must have 'Product' and 'Price' columns.")

                for row in reader:
                    product = row["Product"]
                    try:
                        price = float(row["Price"])
                    except:
                        price = 0.0

                    tax = price * tax_rate
                    total = price + tax

                    tree.insert("", "end", values=(
                        product,
                        f"{price:.2f}",
                        f"{tax:.2f}",
                        f"{total:.2f}"
                    ))

        recalc_totals()

    except Exception as e:
        lbl_total.configure(text=f"Error: {e}")
# Copy / Paste
def copy_selected(event=None):
    rows = ["\t".join(tree.item(i,"values")) for i in tree.selection()]
    if not rows: return
    app.clipboard_clear()
    app.clipboard_append("\n".join(rows))
    app.update()

def paste_to_tree(event=None):
    try:
        lines = app.clipboard_get().strip().split("\n")
        for line in lines:
            values = line.split("\t")
            while len(values) < 4:
                values.append("0.00")
            tree.insert("", "end", values=values[:4])
        recalc_totals()
    except:
        pass

# Right-click Menu 
right_click_menu = Menu(app, tearoff=0)
right_click_menu.add_command(label="Add Product", command=add_product)
right_click_menu.add_command(label="Delete Selected", command=delete_product)
right_click_menu.add_separator()
right_click_menu.add_command(label="Copy", command=copy_selected)
right_click_menu.add_command(label="Paste", command=paste_to_tree)
right_click_menu.add_command(label="Add Row...", command=add_row_popup)  


def show_context_menu(event):
    row_id = tree.identify_row(event.y)
    if row_id:
        tree.selection_set(row_id)
    right_click_menu.tk_popup(event.x_root, event.y_root)

# Layout 
lbl_title = ctk.CTkLabel(page_calc, text="Dynamic Sales Tax Calculator", font=("Arial", 22, "bold"))
lbl_title.pack(pady=(20, 10))

# Tax controls
frame_tax = ctk.CTkFrame(page_calc)
frame_tax.pack(pady=(10, 20))
ctk.CTkLabel(frame_tax, text="Tax Rate (%):", font=("Arial",12)).pack(side="left", padx=5)
entry_tax = ctk.CTkEntry(frame_tax, width=80)
entry_tax.insert(0, "6")
entry_tax.pack(side="left", padx=5)
ctk.CTkButton(frame_tax, text="Recalculate", fg_color="#f50baf", hover_color="#d97706", command=recalc_totals).pack(side="left", padx=10) 

# File buttons
frame_buttons = ctk.CTkFrame(page_calc)
frame_buttons.pack(pady=(0, 20), padx=20, fill="x")
ctk.CTkButton(frame_buttons, text="Load Excel", command=load_excel_dynamic, fg_color="#22c55e", hover_color="#16a34a").pack(side="left", padx=10)
ctk.CTkButton(frame_buttons, text="Save Excel", fg_color="#2bcf34", hover_color="#16a34a", command=save_to_excel).pack(side="right", padx=10)  

# Table
frame_table = ctk.CTkFrame(page_calc)
frame_table.pack(padx=20, pady=10, fill="both", expand=True)
columns = ("Product","Price","Tax","Total") 
tree = ttk.Treeview(frame_table, columns=columns, show="headings", selectmode="extended")
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, anchor="center", width=150)

# Style
style = ttk.Style()
style.theme_use("clam")
style.configure("Treeview",
                background="#2b2b2b",
                foreground="white",
                rowheight=26,
                fieldbackground="#2b2b2b",
                bordercolor="#444",
                borderwidth=1)
style.configure("Treeview.Heading", font=("Arial",12,"bold"))
style.map("Treeview", background=[('selected', '#3b82f6')], foreground=[('selected', 'white')])

vsb = ttk.Scrollbar(frame_table, orient="vertical", command=tree.yview)
tree.configure(yscroll=vsb.set)
vsb.pack(side="right", fill="y")
tree.pack(fill="both", expand=True)

# Totals
lbl_tax_total = ctk.CTkLabel(page_calc, text="Tax Total: $0.00", font=("Arial",20), text_color="#08fc82")
lbl_tax_total.pack(pady=(10,0))
lbl_total = ctk.CTkLabel(page_calc, text="Grand Total: $0.00", font=("Arial",14))
lbl_total.pack(pady=(0,10))

# Bindings 
tree.bind("<Double-1>", on_double_click)
tree.bind("<Button-3>", show_context_menu)
tree.bind("<Control-c>", copy_selected)
tree.bind("<Control-v>", paste_to_tree)

# Drag selection 
drag_selecting = False
start_item = None
def on_button_press(event):
    global drag_selecting, start_item
    region = tree.identify("region", event.x, event.y)
    if region != "cell":
        return
    drag_selecting = True
    start_item = tree.identify_row(event.y)
    if start_item:
        tree.selection_set(start_item)
def on_mouse_drag(event):
    global drag_selecting, start_item
    if not drag_selecting or not start_item:
        return
    current_item = tree.identify_row(event.y)
    if not current_item:
        return
    items = tree.get_children()
    if start_item in items and current_item in items:
        start, end = sorted([items.index(start_item), items.index(current_item)])
        tree.selection_set(items[start:end+1])
def on_button_release(event):
    global drag_selecting, start_item
    drag_selecting = False
    start_item = None
tree.bind("<Button-1>", on_button_press)
tree.bind("<B1-Motion>", on_mouse_drag)
tree.bind("<ButtonRelease-1>", on_button_release)


#Show homepage first
show_page("home")
app.mainloop()
