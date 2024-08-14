import tkinter as tk
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime

window = tk.Tk()
window.title("Invoice builder")
window.geometry("1000x500")

products = []

def generateInvoice():
    v = name_var.get()
    a = add_var.get()
    p = phone_var.get()
    total = sum(item[3] for item in products)
    context = {
        "NAME": v,
        "ADDRESS": a,
        "PHONE": p,
        "itemList": products,
        'total': total
    }
    doc = DocxTemplate("tt1.docx")
    try:
        doc.render(context)
        filename=v+datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S")+".docx"
        doc.save(filename)
        print("Invoice generated successfully.")
    except Exception as e:
        print(f"Error generating invoice: {e}")
        reset()
def clearProductfield():
    product_var.set('')
    quantity_var.set(1)
    price_per_unit_var.set(0.0)

def reset():
    clearProductfield()
    name_var.set('')
    add_var.set('')
    phone_var.set('')
    table.delete(*table.get_children())
    products.clear()


table = ttk.Treeview(window,
                     columns=('product_name',
                                      'quantity',
                                      'price_per_unit',
                                      'amount'),
                     show='headings')
table.grid(row=6, column=0, columnspan=3, sticky='ewns', padx=30)

table.heading('product_name', text='Product Name', anchor='w')
table.heading('quantity', text='Quantity', anchor='w')
table.heading('price_per_unit', text='Price per Unit', anchor='w')
table.heading('amount', text='Amount', anchor='w')

def addProduct():
    p = product_var.get()
    q = quantity_var.get()
    pr = price_per_unit_var.get()
    t = q * pr
    table.insert('', 0, values=(p, q, pr, t))
    products.append((p, q, pr, t))
    clearProductfield()

name_label = ttk.Label(window, text="Full Name")
name_label.grid(row=0, column=0)
name_var = tk.StringVar()
name_entry = ttk.Entry(window, textvariable=name_var)
name_entry.grid(row=1, column=0)

add_label = ttk.Label(window, text="Address")
add_label.grid(row=0, column=1)
add_var = tk.StringVar()
add_entry = ttk.Entry(window, textvariable=add_var)
add_entry.grid(row=1, column=1)

phone_label = ttk.Label(window, text="Phone")
phone_label.grid(row=0, column=2)
phone_var = tk.StringVar()
phone_entry = ttk.Entry(window, textvariable=phone_var)
phone_entry.grid(row=1, column=2)

product_label = ttk.Label(window, text="Product Name")
product_label.grid(row=3, column=0)
product_var = tk.StringVar()
product_entry = ttk.Entry(window, textvariable=product_var)
product_entry.grid(row=4, column=0)

quantity_label = ttk.Label(window, text="Quantity")
quantity_label.grid(row=3, column=1)
quantity_var = tk.IntVar(value=1)
quantity_entry = ttk.Entry(window, textvariable=quantity_var)
quantity_entry.grid(row=4, column=1)

price_per_unit_label = ttk.Label(window, text="Price per Unit")
price_per_unit_label.grid(row=3, column=2)
price_per_unit_var = tk.DoubleVar()
price_per_unit_entry = ttk.Entry(window, textvariable=price_per_unit_var)
price_per_unit_entry.grid(row=4, column=2)

add_button = ttk.Button(window, text="Add Product", command=addProduct)
add_button.grid(row=5, column=2, columnspan=3)

new_button = ttk.Button(window, text="New Invoice", command=reset)
new_button.grid(row=7, column=1)

generateinvoice_button = ttk.Button(window, text="Generate Invoice", command=generateInvoice)
generateinvoice_button.grid(row=7, column=2)

for i in range(9):
    weight = 1
    if i == 6:
        weight = 6
    window.rowconfigure(i, weight=weight)

for u in range(3):
    window.columnconfigure(u, weight=1)

window.mainloop()
