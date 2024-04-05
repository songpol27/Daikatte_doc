import tkinter as tk
from tkinter import ttk
from google_api_manage import manager_google_sheet_data
from operate_os import Opreate_os
import pandas as pd
import locale, random, string
from tkcalendar import Calendar
from tkinter import messagebox
import tkinter.filedialog
from bahttext import bahttext

locale.setlocale(locale.LC_ALL, 'th_TH')
class GoogleSheetGUI:

    def __init__(self, master):
        self.master = master
        master.title("Document Manager by jel")

        # Set up styles
        self.style = ttk.Style()
        self.style.theme_use("clam")  # Choose a theme (e.g., "clam", "alt", "default")
        
        # Create a notebook
        self.notebook = ttk.Notebook(master)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        # Create tabs
        self.tab1 = ttk.Frame(self.notebook)
        self.tab2 = ttk.Frame(self.notebook)

        self.notebook.add(self.tab1, text="Document manager")
        self.notebook.add(self.tab2, text="Invoice manager")

        # Call the method to populate Tab 1
        self.populate_daily_note()

        # Call the method to populate Tab 2
        self.populate_invoice()

    def populate_daily_note(self):
        # Populate Tab 1 with widgets
        self.label = ttk.Label(self.tab1, text="Document Manager by jel", font=("Helvetica", 16, "bold"))
        self.label.grid(row=0, column=0, columnspan=2, pady=10)

        self.tree = ttk.Treeview(self.tab1, columns=('Year', 'Month', 'total_amount'), show='headings', style="Treeview")
        self.tree.heading('Year', text='Year')
        self.tree.heading('Month', text='Month')
        self.tree.heading('total_amount', text='Total Amount')
        self.tree.grid(row=1, column=0, columnspan=2, padx=20, pady=10, sticky="nsew")

        self.display_df_daily_note()

        self.tree.bind("<Double-1>", self.on_double_click_to_doc)

        self.button_frame = ttk.Frame(self.tab1)
        self.button_frame.grid(row=2, column=0, columnspan=2, pady=10)

        #self.Add_daily_note_button = ttk.Button(self.button_frame, text="ADD", command=self.add_window)
        #self.Add_daily_note_button.grid(row=0, column=0, padx=5)

        self.refresh_button = ttk.Button(self.button_frame, text="Refresh", command=self.refresh_data)
        self.refresh_button.grid(row=0, column=1, padx=5)

    def populate_invoice(self):
        
        label_tab2 = ttk.Label(self.tab2, text="Invoice Manager by jel")
        label_tab2.grid(row=0, column=0, columnspan=2, pady=10)

        self.tabel_invoice = ttk.Treeview(self.tab2, columns=('id_inv','no_inv', 'Date', 'name', 'total_amount'), show='headings', style="Treeview")
        self.tabel_invoice.heading('id_inv', text='ID')
        self.tabel_invoice.heading('no_inv', text='Number Invoice')
        self.tabel_invoice.heading('Date', text='Date')
        self.tabel_invoice.heading('name', text='Name')
        self.tabel_invoice.heading('total_amount', text='Total Amount')
        self.tabel_invoice.grid(row=1, column=0, columnspan=1, padx=10, pady=10, sticky="nsew")
        
        self.display_df_invoice()
        self.tabel_invoice.bind("<Double-1>", self.on_double_click_invoice)

        self.refresh_button = ttk.Button(self.tab2, text="Refresh", command=self.refresh_data)
        self.refresh_button.grid(row=2, column=0, padx=5)

    def on_double_click_invoice(self, event):

        def open_sum_invoice(filename):
            opreate_os.open_sum_invoice_file(filename)

        def delete_sum_invoice(filename):
            opreate_os.delete_sum_invoice_file(filename)

        def create_sum_invoice(filename, main_data, sub_data):
            opreate_os.create_sum_invoice(filename, main_data, sub_data)


        item = self.tabel_invoice.selection()[0]
        id_inv = self.tabel_invoice.item(item, "values")[0]

        matching_rows = self.df_invoice.loc[self.df_invoice['id_inv'] == id_inv]
        no_inv = matching_rows["no_inv"].values[0]
        date_inv = pd.Timestamp(matching_rows["Date"].values[0]).strftime("%d/%m/%Y")
        total_price = matching_rows["total_price"].values[0]
        vat_price = matching_rows["vat_price"].values[0]
        total_amount = matching_rows["total_amount"].values[0]

        id_customer = matching_rows["id_company"].values[0]
        customer_rows = self.df_company.loc[self.df_company['id_company'] == id_customer]
        company = customer_rows["name"].values[0]
        company_address = customer_rows["address"].values[0]
        main_office = customer_rows["branch_office"].values[0]
        main_address = customer_rows["delivery_address"].values[0]
        tax = customer_rows["tax_number"].values[0]

        text_show = "Date: " + date_inv + "\n"
        text_show += "Office: " + company + "\n"
        text_show += "Total Price: " + total_price + "\n"
        text_show += "Vat Price: " + vat_price + "\n"
        text_show += "Total Amount: " + total_amount + "\n"

        filtered_data = self.invoice_items[(self.invoice_items['id_inv'] == id_inv)]
        invoice_merged_df = pd.merge(filtered_data, self.df_daily_note, on='id_dn', how='inner')
        invoice_merged_df = invoice_merged_df[['id_item_inv', 'Date', 'no_dn', 'id_inv_y','id_po', 'qty', 'total_price_y']]

        main_data = [company, company_address, main_office, main_address, tax, no_inv, date_inv, total_price, vat_price, total_amount]
        sub_data = invoice_merged_df
        
        new_window_invoice = tk.Toplevel(self.tab2)
        new_window_invoice.title(f"Details for {id_inv}")

        df_daily_label = ttk.Label(new_window_invoice, text=str(no_inv))
        df_daily_label.grid(row=0, column=0, pady=10)

        df_daily_text_label = ttk.Label(new_window_invoice, text=text_show)
        df_daily_text_label.grid(row=1, column=0, pady=10)

        #Create and pack widgets
        self.group_tree_invoice = ttk.Treeview(new_window_invoice, columns=('id_item_inv', 'Date', 'no_dn', 'id_inv_y','id_po', 'qty', 'total_price_y'), show='headings', style="Treeview")
        self.group_tree_invoice.heading('id_item_inv', text='ID')
        self.group_tree_invoice.heading('Date', text='Date')
        self.group_tree_invoice.heading('no_dn', text='DN number')
        self.group_tree_invoice.heading('id_inv_y', text='INV number')
        self.group_tree_invoice.heading('id_po', text='PO number')
        self.group_tree_invoice.heading('qty', text='quantity')
        self.group_tree_invoice.heading('total_price_y', text='Total Price')
        self.group_tree_invoice.grid(row=3, column=0, padx=20, pady=10, sticky='nsew')
        for index, row in invoice_merged_df.iterrows():
            self.group_tree_invoice.insert('', 'end', values=(row['id_item_inv'], row['Date'], row['no_dn'], row['id_inv_y'], row['id_po'], row['qty'], row['total_price_y']))

        upload_frame = ttk.Frame(new_window_invoice)
        upload_frame.grid(row=2, column=0, pady=5)

        open_button = ttk.Button(upload_frame, text="Open Invoice", command=lambda: open_sum_invoice(str(no_inv))) #
        open_button.grid(row=0, column=0, padx=5)

        create_button = ttk.Button(upload_frame, text="Create Invoice", command=lambda: create_sum_invoice(no_inv, main_data, sub_data)) 
        create_button.grid(row=0, column=1, padx=5)

        delete_button = ttk.Button(upload_frame, text="Delete Invoice", command=lambda: delete_sum_invoice(str(no_inv))) #command=lambda: delete_sum_invoice(entry_po_file_path)
        delete_button.grid(row=0, column=2, padx=5)

    def display_df_invoice(self):
        #self.df_daily_note, self.df_aily_note_item = google_sheet_manager.read_dynamic_data()
        
        self.df_invoice, self.invoice_items = google_sheet_manager.read_invoice()
        self.df_invoice['Date'] = pd.to_datetime(self.df_invoice['Date'])
        

        self.invoice_merged_df = pd.merge(self.df_invoice, self.df_company, on='id_company', how='left')

        for index, row in self.invoice_merged_df.iterrows():
            self.tabel_invoice.insert('', 'end', values=(row['id_inv'],row['no_inv'], row['Date'], row['name'], row['total_amount']))

    def display_df_daily_note(self):
        self.df_daily_note, self.df_daily_note_item = google_sheet_manager.read_dynamic_data()
        self.df_company, self.df_products = google_sheet_manager.read_static_data()
        self.df_daily_note['Date'] = pd.to_datetime(self.df_daily_note['Date'])
        
        self.df_daily_note['total_amount'] = self.df_daily_note['total_amount'].str.replace('฿', '').str.replace(',', '').astype(float)
        self.df_daily_note['Month'] = self.df_daily_note['Date'].dt.month.astype(str)
        self.df_daily_note['Year'] = self.df_daily_note['Date'].dt.year.astype(str)
        
        grouped_df = self.df_daily_note.groupby(['Year', 'Month']).agg({'total_amount': 'sum'}).reset_index()

        for index, row in grouped_df.iterrows():
            total_amount_thai_baht = f"฿{row['total_amount']:,.0f}"
            self.tree.insert('', 'end', values=(row['Year'], row['Month'], total_amount_thai_baht))

    def on_double_click_to_doc(self, event):
        # Handle double click event
        item = self.tree.selection()[0]
        year_click = self.tree.item(item, "values")[0]
        month_click = self.tree.item(item, "values")[1]

        filtered_data = self.df_daily_note[(self.df_daily_note['Year'] == year_click) & (self.df_daily_note['Month'] == month_click)]
        new_window = tk.Toplevel(self.master)
        new_window.title(f"Details for {month_click}/{year_click}")

        # Create and pack widgets
        self.group_tree = ttk.Treeview(new_window, columns=('no_dn', 'Date', 'id_company', 'total_amount'), show='headings', style="Treeview")
        self.group_tree.heading('no_dn', text='DN_No')
        self.group_tree.heading('Date', text='Date')
        self.group_tree.heading('id_company', text='Customer')
        self.group_tree.heading('total_amount', text='Total Amount')
        self.group_tree.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)


        # Display filtered data
        for index, row in filtered_data.iterrows():
            total_amount_thai_baht = f"฿{row['total_amount']:,.0f}"  # Format total amount with Thai baht symbol
            customer_rows = self.df_company.loc[self.df_company['id_company'] == row['id_company']]
            branch_office = customer_rows["branch_office"].values[0]
            date_con= pd.Timestamp(row['Date']).strftime("%d/%m/%Y") 
            self.group_tree.insert('', 'end', values=(row['no_dn'], date_con, branch_office, total_amount_thai_baht))

        # Add a button to show product details
        show_detail_button = ttk.Button(new_window, text="Show Detail", command=self.open_products)
        show_detail_button.pack(pady=10)

    def open_products(self):
        item = self.group_tree.selection()[0]
        DN_No = self.group_tree.item(item, "values")[0]

        matching_rows = self.df_daily_note.loc[self.df_daily_note['no_dn'] == DN_No]
        id_customer = matching_rows["id_company"].values[0]
        id_products = matching_rows["id_dn"].values[0]

        daily_note_drop_list = ['id_dn', 'id_company', 'vat', 'Month', 'Year']
        df_daily_noted_cleaned = matching_rows.drop(columns=daily_note_drop_list)

        customer_rows = self.df_company.loc[self.df_company['id_company'] == id_customer]
        products_rows = self.df_daily_note_item.loc[self.df_daily_note_item['id_dn'] == id_products]

        branch_office = customer_rows["branch_office"].values[0]
        df_daily_noted_cleaned["branch_office"] = branch_office
        df_daily_noted_cleaned = df_daily_noted_cleaned[['no_dn', 'Date', 'branch_office', 'id_inv',
                                                        'id_po', 'total_price', 'vat_price',
                                                        'total_amount']]

        df_merged_product = pd.merge(products_rows, self.df_products, on='id_product', how='inner')
        product_drop_list = ['id_item', 'id_dn', 'id_product', 'unit', 'kind', 'Detail']
        df_merged_product_cleaned = df_merged_product.drop(columns=product_drop_list)
        df_merged_product_cleaned = df_merged_product_cleaned.reset_index(drop=True)
        df_merged_product_cleaned = df_merged_product_cleaned[['item_code', 'description', 'qty', "p/u_y", 'total_item']]

        text_show = "Date: " + pd.Timestamp(df_daily_noted_cleaned["Date"].iloc[0]).strftime("%d/%m/%Y") + "\n"
        text_show += "Branch Office: " + df_daily_noted_cleaned["branch_office"].iloc[0] + "\n"
        text_show += "PO name: " + df_daily_noted_cleaned["id_po"].iloc[0] + "\n"
        text_show += "INV name: " + df_daily_noted_cleaned["id_inv"].iloc[0] + "\n"
        text_show += "Total price : " + str(df_daily_noted_cleaned["total_price"].iloc[0]) + "\n"
        text_show += "Vat price : " + str(df_daily_noted_cleaned["vat_price"].iloc[0]) + "\n"
        text_show += "Total Amount : " + locale.currency(df_daily_noted_cleaned["total_amount"].iloc[0], grouping=True, symbol=True) + "\n"

        def add_newlines_to_address(address):
            # Add newline characters to the address string
            modified_address = address.replace(",", ",\n")
            return modified_address

        po_name = df_daily_noted_cleaned["id_po"].iloc[0]
        inv_name = df_daily_noted_cleaned["id_inv"].iloc[0]
        dn_name = str(df_daily_noted_cleaned["no_dn"].values[0])
        all_name = "ALL" + (str(df_daily_noted_cleaned["no_dn"].values[0]))[2:]

        
        Date = pd.Timestamp(df_daily_noted_cleaned["Date"].iloc[0]).strftime("%d/%m/%Y")
        

        text_company = customer_rows["name"].iloc[0]
        text_date = f"No. : {dn_name} \nDate : {Date}"
        tax_text = customer_rows["tax_number"].iloc[0]
        text_company_main_address = add_newlines_to_address(customer_rows["address"].iloc[0])
        text_company_delivery_address = add_newlines_to_address(customer_rows["delivery_address"].iloc[0])
        text_total = str(df_daily_noted_cleaned["total_price"].iloc[0])
        text_vat = str(df_daily_noted_cleaned["vat_price"].iloc[0])
        text_total_amount = locale.currency(df_daily_noted_cleaned["total_amount"].iloc[0], grouping=True, symbol=True)
        text_thai_amount = f'ตัวอักษร({bahttext(float(df_daily_noted_cleaned["total_amount"].iloc[0]))})'
        

        id_code = []
        description = []
        qty = []
        unit = []
        unit_price = []
        amount = []

        for index, row in df_merged_product.iterrows():
            id_code.append(row.get("item_code", ""))
            description.append(row.get("description", ""))
            qty.append(row.get("qty", ""))
            unit.append(row.get("unit", ""))
            unit_price.append(row.get("p/u_y", ""))
            amount.append(row.get("total_item", ""))

        max_number = len(id_code)
        if not max_number == 10 :
            for _ in range(10 - max_number):
                #id_code.append("")
                description.append("")
                qty.append("")
                unit.append("")
                unit_price.append("")
                amount.append("")

        if len(id_code) <= 5 :

            page_1 = [
                ["Company :", text_company, "", text_date, "", ""],
                ["Address :", text_company_main_address, "10", "Payment terms\n\nCredit 30 Day", "12", "13"],
                ["Delivery Address :", text_company_delivery_address, "5", "14", "15", "16"],
                ["Tax :", tax_text, "17", "18", "19", "29"],
                ["Item code", "Description", "Quantity", "Unit", "Unit Price", "Amount"],
                ["", description[0], qty[0], unit[0], unit_price[0], amount[0]],
                ["", description[1], qty[1], unit[1], unit_price[1], amount[1]],
                ["", description[2], qty[2], unit[2], unit_price[2], amount[2]],
                ["", description[3], qty[3], unit[3], unit_price[3], amount[3]],
                ["", description[4], qty[4], unit[4], unit_price[4], amount[4]],

                [text_thai_amount, "2", "3", "4", "Sub Total", text_total],
                ["5", "6", "7", "8", "Vat 7%", text_vat],
                ["9", "10", "11", "12", "Total Amount", text_total_amount]
            ]
            list_page = [page_1]
        
        else:

            page_1 = [
                ["Company :", text_company, "", text_date, "", ""],
                ["Address :", text_company_main_address, "10", "Payment terms\n\nCredit 30 Day", "12", "13"],
                ["Delivery Address :", text_company_delivery_address, "5", "14", "15", "16"],
                ["Tax :", tax_text, "17", "18", "19", "29"],
                ["Item code", "Description", "Quantity", "Unit", "Unit Price", "Amount"],
                ["", description[0], qty[0], unit[0], unit_price[0], amount[0]],
                ["", description[1], qty[1], unit[1], unit_price[1], amount[1]],
                ["", description[2], qty[2], unit[2], unit_price[2], amount[2]],
                ["", description[3], qty[3], unit[3], unit_price[3], amount[3]],
                ["", description[4], qty[4], unit[4], unit_price[4], amount[4]],

            ]

            page_2 = [
                ["Company :", text_company, "", text_date, "", ""],
                ["Address :", text_company_main_address, "10", "Payment terms\n\nCredit 30 Day", "12", "13"],
                ["Delivery Address :", text_company_delivery_address, "5", "14", "15", "16"],
                ["Tax :", tax_text, "17", "18", "19", "29"],
                ["Item code", "Description", "Quantity", "Unit", "Unit Price", "Amount"],
                ["", description[5], qty[5], unit[5], unit_price[5], amount[5]],
                ["", description[6], qty[6], unit[6], unit_price[6], amount[6]],
                ["", description[7], qty[7], unit[7], unit_price[7], amount[7]],
                ["", description[8], qty[8], unit[8], unit_price[8], amount[8]],
                ["", description[9], qty[9], unit[9], unit_price[9], amount[9]],

                [text_thai_amount, "2", "3", "4", "Sub Total", text_total],
                ["5", "6", "7", "8", "Vat 7%", text_vat],
                ["9", "10", "11", "12", "Total Amount", text_total_amount]
            ]

            list_page = [page_1, page_2]
         

        def browse_file(entry):
            file_path = tkinter.filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
            if file_path:
                entry.delete(0, tk.END)  # Clear any existing text in the entry field
                entry.insert(0, file_path)  # Insert the selected file path into the entry field

        def upload_file_po_inv(po_path, inv_path, po_name, inv_name):
            if po_path == "":
                messagebox.showerror("Error", "Please enter a PO file.")
                return
            
            if inv_path =="":
                messagebox.showerror("Error", "Please enter a INV file.")
                return
            
            if po_name =="":
                messagebox.showerror("Error", "No PO data in database.")
                return
            
            if inv_name =="":
                messagebox.showerror("Error", "No INV data in database.")
                return
            try:
                opreate_os.copy_addnew_inv(inv_path, inv_name)
                opreate_os.copy_addnew_po(po_path, po_name)

                entry_po_file_path.delete(0, tk.END)  
                entry_po_file_path.insert(0, "")

                entry_inv_file_path.delete(0, tk.END)
                entry_inv_file_path.insert(0, "")

                open_po_button.config(state=tk.NORMAL)
                delete_po_button.config(state=tk.NORMAL)

                open_inv_button.config(state=tk.NORMAL)
                delete_inv_button.config(state=tk.NORMAL)

                messagebox.showinfo("Successful", "Successful to upload file.")
                
            except:
                messagebox.showerror("Error", "Fail to upload file .")
                return

        def chk_file(po_filename, inv_filename, dn_filename):
            po_have , inv_have, dn_have = opreate_os.chk_file(po_filename, inv_filename, dn_filename)

            if po_have:
                open_po_button.config(state=tk.NORMAL)
                delete_po_button.config(state=tk.NORMAL)
            else:
                open_po_button.config(state=tk.DISABLED)
                delete_po_button.config(state=tk.DISABLED)

            if inv_have:
                open_inv_button.config(state=tk.NORMAL)
                delete_inv_button.config(state=tk.NORMAL)
            else:
                open_inv_button.config(state=tk.DISABLED)
                delete_inv_button.config(state=tk.DISABLED)

            if dn_have:
                open_daily_button.config(state=tk.NORMAL)
                delete_daily_button.config(state=tk.NORMAL)
            else:
                open_daily_button.config(state=tk.DISABLED)
                delete_daily_button.config(state=tk.DISABLED)

            if po_have == True and inv_have == True and dn_have == True:
                
                gen_sum_button.config(state=tk.NORMAL)
                open_sum_button.config(state=tk.NORMAL)
                delete_sum_button.config(state=tk.NORMAL)
            else:
                gen_sum_button.config(state=tk.DISABLED)
                open_sum_button.config(state=tk.DISABLED)
                delete_sum_button.config(state=tk.DISABLED)

        def open_po(po_filename):
            opreate_os.open_po_file(po_filename)
            
        def open_inv(inv_filename):
            opreate_os.open_inv_file(inv_filename)

        def open_dn(dn_filename):
            opreate_os.open_dn_file(dn_filename)

        def open_all(all_filename):
            opreate_os.open_all_file(all_filename)

        def delete_po(po_filename):
            opreate_os.delete_po_file(po_filename)
            open_po_button.config(state=tk.DISABLED)
            delete_po_button.config(state=tk.DISABLED)
            open_sum_button.config(state=tk.DISABLED)

        def delete_inv(inv_filename):
            opreate_os.delete_inv_file(inv_filename)
            open_inv_button.config(state=tk.DISABLED)
            delete_inv_button.config(state=tk.DISABLED)
            open_sum_button.config(state=tk.DISABLED)

        def delete_dn(dn_filename):
            opreate_os.delete_dn_file(dn_filename)
            open_daily_button.config(state=tk.DISABLED)
            delete_daily_button.config(state=tk.DISABLED)
            open_sum_button.config(state=tk.DISABLED)
            
        def delete_all(all_filename):
            opreate_os.delete_all_file(all_filename)
            open_sum_button.config(state=tk.DISABLED)
            delete_sum_button.config(state=tk.DISABLED)

        def create_dn(file_name,data, date_text, image_list):
            opreate_os.create_delivery_note_pdf(file_name,data, date_text, image_list)
            open_daily_button.config(state=tk.NORMAL)
            delete_daily_button.config(state=tk.NORMAL)

        def merge_pdf(po_filename, inv_filename, dn_filename, output_filename):
            opreate_os.merge_3_file_order_by(po_filename, inv_filename, dn_filename, output_filename)
            delete_sum_button.config(state=tk.NORMAL)
            open_sum_button.config(state=tk.NORMAL)

        new_window2 = tk.Toplevel(self.master)
        new_window2.title("Details")

        df_daily_label = ttk.Label(new_window2, text=str(df_daily_noted_cleaned["no_dn"].values[0]))
        df_daily_label.grid(row=0, column=0, pady=10)

        df_daily_text_label = ttk.Label(new_window2, text=text_show)
        df_daily_text_label.grid(row=1, column=0, pady=10)

        tree = ttk.Treeview(new_window2, columns=('item_code', 'description', 'qty', 'p/u_y', 'total_item'), show='headings', style="Treeview")
        tree.heading('item_code', text='Item code')
        tree.heading('description', text='Description')
        tree.heading('qty', text='Quantity')
        tree.heading('p/u_y', text='Price/unit')
        tree.heading('total_item', text='Total Item Cost')
        tree.grid(row=3, column=0, padx=20, pady=10, sticky='nsew')

        for index, row in df_merged_product_cleaned.iterrows():
            tree.insert('', 'end', values=(row['item_code'], row['description'], row['qty'], row["p/u_y"], row['total_item']))

        upload_frame = ttk.Frame(new_window2)
        upload_frame.grid(row=4, column=0, pady=5)

        po_file_label = ttk.Label(upload_frame, text="PO file: ")
        po_file_label.grid(row=0, column=0, pady=10)

        entry_po_file_path = ttk.Entry(upload_frame)
        entry_po_file_path.grid(row=0, column=1, padx=0, pady=5, ipadx = 40)

        po_browsw_button = ttk.Button(upload_frame, text="Browse", command=lambda: browse_file(entry_po_file_path))
        po_browsw_button.grid(row=0, column=2, padx=5)

        inv_file_label = ttk.Label(upload_frame, text="INV file: ")
        inv_file_label.grid(row=1, column=0, pady=10)

        entry_inv_file_path = ttk.Entry(upload_frame)
        entry_inv_file_path.grid(row=1, column=1, padx=0, pady=5, ipadx = 40)

        inv_browsw_button = ttk.Button(upload_frame, text="Browse", command=lambda: browse_file(entry_inv_file_path))
        inv_browsw_button.grid(row=1, column=2, padx=5)

        upload_button = ttk.Button(upload_frame, text="Upload", command=lambda:upload_file_po_inv(entry_po_file_path.get(), entry_inv_file_path.get(), po_name, inv_name))
        upload_button.grid(row=1, column=3)

        

        style = ttk.Style()
        style.configure("Green.TButton", background="#4CAF50")
        style.configure("Red.TButton", background="#FF0000")
        style.configure("Yellow.TButton", background="#FFD966")


        gen_daily_button = ttk.Button(upload_frame, text="Create Daily Note", style="Yellow.TButton", command=lambda:create_dn(dn_name, list_page, Date, id_code))
        gen_daily_button.grid(row=4, column=1, padx=5, pady=5)

        open_daily_button = ttk.Button(upload_frame, text="Open Daily Note", style="Green.TButton", command=lambda:open_dn(dn_name))
        open_daily_button.grid(row=4, column=0, padx=5, pady=5)

        delete_daily_button = ttk.Button(upload_frame, text="Delete Daily Note", style="Red.TButton", command=lambda:delete_dn(dn_name))
        delete_daily_button.grid(row=4, column=2, padx=5, pady=5)

        open_po_button = ttk.Button(upload_frame, text="Open PO", style="Green.TButton", command=lambda:open_po(po_name))
        open_po_button.grid(row=2, column=0, padx=5, pady=5)

        delete_po_button = ttk.Button(upload_frame, text="Delete PO", style="Red.TButton", command=lambda:delete_po(po_name))
        delete_po_button.grid(row=2, column=2, padx=5, pady=5)

        open_inv_button = ttk.Button(upload_frame, text="Open INV", style="Green.TButton", command=lambda:open_inv(inv_name))
        open_inv_button.grid(row=3, column=0, padx=5, pady=5)

        delete_inv_button = ttk.Button(upload_frame, text="Delete INV", style="Red.TButton", command=lambda:delete_inv(inv_name))
        delete_inv_button.grid(row=3, column=2, padx=5, pady=5)

        gen_sum_button = ttk.Button(upload_frame, text="Create sum file", style="Yellow.TButton", command=lambda:merge_pdf(po_name, inv_name, dn_name, all_name))
        gen_sum_button.grid(row=5, column=1, padx=5, pady=5)

        open_sum_button = ttk.Button(upload_frame, text="Open sum file", style="Green.TButton", command=lambda:open_all(all_name))
        open_sum_button.grid(row=5, column=0, padx=5, pady=5)

        delete_sum_button = ttk.Button(upload_frame, text="Delete sum file", style="Red.TButton", command=lambda:delete_all(all_name))
        delete_sum_button.grid(row=5, column=2, padx=5, pady=5)
        
        chk_file(po_name, inv_name, dn_name)

    def refresh_data(self):
        # Clear the current data displayed in the treeview
        for item in self.tree.get_children():
            self.tree.delete(item)

        for item in self.tabel_invoice.get_children():
            self.tabel_invoice.delete(item)

        # Call the method to display the latest data
        self.display_df_daily_note()
        self.display_df_invoice()

if __name__ == "__main__":
    root = tk.Tk()
    google_sheet_manager = manager_google_sheet_data()
    opreate_os = Opreate_os()
    gui = GoogleSheetGUI(root)
    root.iconbitmap('contract.ico')

    root.mainloop()
