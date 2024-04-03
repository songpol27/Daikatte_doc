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
        
        # Create and pack main widgets
        self.label = ttk.Label(master, text="Document Manager by jel", font=("Helvetica", 16, "bold"))
        self.label.grid(row=0, column=0, columnspan=2, pady=10)

        self.tree = ttk.Treeview(master, columns=('Year', 'Month', 'total_amount'), show='headings', style="Treeview")
        self.tree.heading('Year', text='Year')
        self.tree.heading('Month', text='Month')
        self.tree.heading('total_amount', text='Total Amount')
        self.tree.grid(row=1, column=0, columnspan=2, padx=20, pady=10, sticky="nsew")

        self.display_df_daily_note()

        self.tree.bind("<Double-1>", self.on_double_click)

        self.button_frame = ttk.Frame(master)
        self.button_frame.grid(row=2, column=0, columnspan=2, pady=10)

        #self.Add_daily_note_button = ttk.Button(self.button_frame, text="ADD", command=self.add_window)
        #self.Add_daily_note_button.grid(row=0, column=0, padx=5)

        self.refresh_button = ttk.Button(self.button_frame, text="Refresh", command=self.refresh_data)
        self.refresh_button.grid(row=0, column=1, padx=5)

    def add_window(self):
        self.add_window = tk.Toplevel(self.master)
        self.add_window.title(f"Add Daily Note")

        def generate_random_string():
            return ''.join(random.choices(string.ascii_letters + string.digits, k=8))

        def get_customer_names():
            return self.df_company['Branch office'].unique().tolist()

        DN_ID_input = generate_random_string()
        
        label_date = ttk.Label(self.add_window, text="Date:")
        label_date.grid(row=0, column=0, padx=5, pady=5)

        # Use Calendar widget for date selection
        cal = Calendar(self.add_window, selectmode="day", date_pattern="yyyy-mm-dd")
        cal.grid(row=0, column=1, padx=5, pady=5)

        label_customer = ttk.Label(self.add_window, text="Customer:")
        label_customer.grid(row=1, column=0, padx=5, pady=5)

        # Dropdown menu for customer selection
        customer_names = get_customer_names()
        selected_customer = tk.StringVar()
        customer_combobox = ttk.Combobox(self.add_window, textvariable=selected_customer, values=customer_names)
        customer_combobox.grid(row=1, column=1, padx=0, pady=5, ipadx = 40)
        customer_combobox.current(0)

        label_vat = ttk.Label(self.add_window, text="VAT Statute:")
        label_vat.grid(row=2, column=0, padx=5, pady=5)
        vat_combobox = ttk.Combobox(self.add_window, values=["True", "False"])
        vat_combobox.grid(row=2, column=1, padx=0, pady=5)
        vat_combobox.current(0)

        label_po_file_path = ttk.Label(self.add_window, text="PO file:")
        label_po_file_path.grid(row=3, column=0, padx=5, pady=5)

        def browse_po_file():
            file_path = tkinter.filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
            if file_path:
                entry_po_file_path.delete(0, tk.END)  # Clear any existing text in the entry field
                entry_po_file_path.insert(0, file_path)  # Insert the selected file path into the entry field


        entry_po_file_path = ttk.Entry(self.add_window)
        entry_po_file_path.grid(row=3, column=1, padx=0, pady=5, ipadx = 40)

        button_po_browse = ttk.Button(self.add_window, text="Browse", command=browse_po_file)
        button_po_browse.grid(row=3, column=2, padx=5, pady=5)

        label_invioce_file_path = ttk.Label(self.add_window, text="Invoice file:")
        label_invioce_file_path.grid(row=4, column=0, padx=5, pady=5)

        def browse_invoice_file():
            file_path = tkinter.filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
            entry_invoice_file_path.delete(0, tk.END)  # Clear any existing text in the entry field
            entry_invoice_file_path.insert(0, file_path)  # Insert the selected file path into the entry field

        entry_invoice_file_path = ttk.Entry(self.add_window)
        entry_invoice_file_path.grid(row=4, column=1, padx=0, pady=5, ipadx = 40)

        button_invoice_browse = ttk.Button(self.add_window, text="Browse", command=browse_invoice_file)
        button_invoice_browse.grid(row=4, column=2, padx=5, pady=5)

        def call_product():
            filtered_data = self.df_products[(self.df_products['Detail'] == "Available")]
            products_data = []
            for index, row in filtered_data.iterrows():
                Product_ID = row["Product_ID"]
                Description = row["Description"]
                PricePerunit = row["Price/unit"]
                products_data.append((Product_ID, Description, PricePerunit))
            return products_data

        Product_data = call_product()
        running_num = 1

        product_frame = ttk.Frame(self.add_window)
        product_frame.grid(row=5, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

        label_product_name = ttk.Label(product_frame, text="Product name")
        label_product_name.grid(row=0, column=0, padx=0, pady=5)

        label_price_product = ttk.Label(product_frame, text="Product price")
        label_price_product.grid(row=0, column=1, padx=5, pady=5)

        label_input_qty = ttk.Label(product_frame, text="Input qty")
        label_input_qty.grid(row=0, column=2, padx=5, pady=5)

        for product in Product_data:
            Product_ID, Description, PricePerunit = product

            label_desc = ttk.Label(product_frame, text=Description)
            label_desc.grid(row=running_num, column=0, padx=0, pady=5)

            label_price = ttk.Label(product_frame, text=PricePerunit)
            label_price.grid(row=running_num, column=1, padx=5, pady=5)

            entry_qty = ttk.Entry(product_frame)
            entry_qty.grid(row=running_num, column=2, padx=0, pady=5, ipadx = 10)

            running_num += 1

        def submit():
            selected_customer = customer_combobox.get()
            selected_date = cal.get_date()
            selected_vat = vat_combobox.get()
            selected_file_po_path = entry_po_file_path.get()
            selected_file_invoice_path = entry_invoice_file_path.get()
            

            if selected_date == "":
                messagebox.showerror("Error", "Please select a date.")
                return
            if selected_customer == "":
                messagebox.showerror("Error", "Please select a customer.")
                return
            if selected_vat == "":
                messagebox.showerror("Error", "Please select a VAT.")
                return
            if selected_file_po_path == "":
                messagebox.showerror("Error", "Please enter a PO file.")
                return
            if selected_file_invoice_path == "":
                messagebox.showerror("Error", "Please enter an Invoice file.")
                return
            
            def DN_No_cal(selected_customer,selected_date):

                if  selected_customer == "DON DON DONKI Thonglor":
                    dn_no = "DN" + "01"
                elif selected_customer == "DON DON DONKI MBK Center":
                    dn_no = "DN" + "02"
                elif selected_customer == "DON DON DONKI Seacon Square":
                    dn_no = "DN" + "03"
                elif selected_customer == "DON DON DONKI Seacon Bangkae":
                    dn_no = "DN" + "04"
                elif selected_customer == "DON DON DONKI J-park Sriracha":
                    dn_no = "DN" + "05"
                elif selected_customer == "DON DON DONKI Thaniya Plaza":
                    dn_no = "DN" + "06"
                elif selected_customer == "DON DON DONKI Fashion lsland":
                    dn_no = "DN" + "07"
                elif selected_customer == "DON DON DONKI The Mall Bangkapi":
                    dn_no = "DN" + "08"
                else:
                    messagebox.showerror("Error", "Please Contact to Jel.")
                    return
                dn_no = dn_no + "-" +selected_date
                return dn_no

            quantities = []
            for product_frame in self.add_window.winfo_children():
                if isinstance(product_frame, ttk.Frame):
                    for widget in product_frame.winfo_children():
                        if isinstance(widget, ttk.Entry):
                            quantity_str = widget.get()
                            try:
                                if not quantity_str == "":
                                    quantity = int(quantity_str)
                                    quantities.append(quantity)

                                else:
                                    quantities.append(0)
                            except ValueError:
                                messagebox.showerror("Error", "Invalid quantity. Please enter a valid integer.")
                                return


            
            DN_No = DN_No_cal(selected_customer, selected_date)
            
            
            self.add_window.withdraw()  # Hide the self.add_window
            self.comfrim_add(DN_ID_input, DN_No, selected_date,selected_customer, selected_vat, selected_file_po_path, selected_file_invoice_path, Product_data, quantities)



        submit_button = ttk.Button(self.add_window, text="Submit", command=submit)
        submit_button.grid(row=6, columnspan=2, padx=5, pady=10)

    def comfrim_add(self, DN_ID, DN_No, Date, Customer, Vat_staute, File_PO_path, Invoice_number, Product_data, quantities):
        self.comfrim_add_window = tk.Toplevel(self.master)
        self.comfrim_add_window.title(f"Confirm")

        # Frame to contain the information
        info_frame = ttk.Frame(self.comfrim_add_window)
        info_frame.pack(padx=20, pady=20)

        # Display DN ID
        label_dn_id = ttk.Label(info_frame, text=f"DN ID: {DN_ID}")
        label_dn_id.grid(row=0, column=0, sticky="w")

        # Display DN Number
        label_dn_no = ttk.Label(info_frame, text=f"DN Number: {DN_No}")
        label_dn_no.grid(row=1, column=0, sticky="w")

        # Display Date
        label_date = ttk.Label(info_frame, text=f"Date: {Date}")
        label_date.grid(row=2, column=0, sticky="w")

        # Display Customer
        label_customer = ttk.Label(info_frame, text=f"Customer: {Customer}")
        label_customer.grid(row=3, column=0, sticky="w")

        # Display VAT Statute
        label_vat = ttk.Label(info_frame, text=f"VAT Statute: {Vat_staute}")
        label_vat.grid(row=4, column=0, sticky="w")

        # Display PO File Path
        label_po_path = ttk.Label(info_frame, text=f"PO File Path: {File_PO_path}")
        label_po_path.grid(row=5, column=0, sticky="w")

        # Display Invoice Number
        label_invoice_no = ttk.Label(info_frame, text=f"Invoice Number: {Invoice_number}")
        label_invoice_no.grid(row=6, column=0, sticky="w")

        df = pd.DataFrame(Product_data, columns=['Product_ID', 'Description', 'Price_per_unit'])
        df['Quantity'] = quantities
        df['Total Item Cost'] = df['Quantity'].astype(int) * df['Price_per_unit'].astype(int)
        df['DN_ID'] = DN_ID

        filtered_data = df[(df['Total Item Cost'] != 0 )]
        
        DN_Item_ID_list = []
        for _ in range(len(filtered_data)):
            DN_Item_ID = ''.join(random.choices(string.ascii_letters + string.digits, k=8))
            DN_Item_ID_list.append(DN_Item_ID)
        filtered_data['DN_Item_ID'] = DN_Item_ID_list
        df_product_cleaned = filtered_data[['DN_Item_ID', 'DN_ID', 'Product_ID', 'Description',
                                                        'Price_per_unit', 'Quantity', 'Total Item Cost']]

        drop_list = ['DN_Item_ID', 'DN_ID']
        df_product_cleaned_for_show = df_product_cleaned.drop(columns=drop_list)

        table_frame = ttk.Frame(self.comfrim_add_window)
        table_frame.pack(padx=20, pady=10)
        
        tree = ttk.Treeview(table_frame)
        tree["columns"] = list(df_product_cleaned_for_show.columns)
        tree["show"] = "headings"

        for column in df_product_cleaned_for_show.columns:
            tree.heading(column, text=column)
        
        for index, row in df_product_cleaned_for_show.iterrows():
            tree.insert("", "end", values=list(row))
        
        tree.pack(expand=True, fill=tk.BOTH)


        # Calculate the total cost
        total_cost = df_product_cleaned['Total Item Cost'].sum()

        total_cost = df_product_cleaned['Total Item Cost'].sum()
        vat_total_cost = total_cost * 0.07 if Vat_staute == "True" else 0
        total_amount = total_cost + vat_total_cost

        # Format the values with two decimal places
        total_cost_str = f"฿{total_cost:,.2f}"
        vat_total_cost_str = f"฿{vat_total_cost:,.2f}"
        total_amount_str = f"฿{total_amount:,.2f}"

        # Display total cost, VAT cost, and total amount
        label_total_cost = ttk.Label(info_frame, text=f"Total Cost: {total_cost_str}")
        label_total_cost.grid(row=7, column=0, sticky="w")

        label_vat_total_cost = ttk.Label(info_frame, text=f"Vat Cost: {vat_total_cost_str}")
        label_vat_total_cost.grid(row=8, column=0, sticky="w")

        label_total_amount = ttk.Label(info_frame, text=f"Total Amount: {total_amount_str}")
        label_total_amount.grid(row=9, column=0, sticky="w")
        
        def customer_to_id(selected_date):
                filtered_data = self.df_company[(self.df_company['Branch office'] == selected_date)]
                id_customer = filtered_data["ID_Company"].values[0]
                return id_customer
        id_company = customer_to_id(Customer)

        # Button Frame
        self.button_frame = ttk.Frame(self.comfrim_add_window)
        self.button_frame.pack(pady=20)
        

        # Confirm Button
        confirm_button = ttk.Button(self.button_frame, text="Confirm", command=lambda: self.comfrim_add_confirm(DN_ID, DN_No, Date, id_company, Vat_staute, File_PO_path, Invoice_number, df_product_cleaned))
        confirm_button.grid(row=0, column=0, padx=10)

        # Back Button
        back_button = ttk.Button(self.button_frame, text="Back", command=self.comfrim_add_back)
        back_button.grid(row=0, column=1, padx=10)

    def comfrim_add_confirm(self, id_dn, dn_no, date, id_company, vat, File_PO_path, Invoice_number, df_product):
        
        try:
            PO_name_file = dn_no.replace("DN", "PO")
            IV_name_file = dn_no.replace("DN", "IV")

            opreate_os.copy_addnew_to_po(File_PO_path, PO_name_file)
            inv_name= opreate_os.copy_addnew_to_inv(Invoice_number)

            google_sheet_manager.add_data_daily_note(id_dn, dn_no, date, str(id_company), vat, inv_name, PO_name_file)

            
            for index, row in df_product.iterrows():
                google_sheet_manager.add_data_daily_note_items(row["DN_ID"], row["Product_ID"], row["Quantity"], row["Price_per_unit"])


            self.add_window.destroy()
            self.comfrim_add_window.destroy()

            messagebox.showinfo("Successful", "Data added successfully to Google Sheets.")
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {e}")
            
    def comfrim_add_back(self):
        self.add_window.deiconify()
        self.comfrim_add_window.destroy()

    def display_df_daily_note(self):
        self.df_daily_note, self.df_aily_note_item = google_sheet_manager.read_dynamic_data()
        self.df_company, self.df_products = google_sheet_manager.read_static_data()
        self.df_daily_note['Date'] = pd.to_datetime(self.df_daily_note['Date'])

        self.df_daily_note['total_amount'] = self.df_daily_note['total_amount'].str.replace('฿', '').str.replace(',', '').astype(float)
        self.df_daily_note['Month'] = self.df_daily_note['Date'].dt.month.astype(str)
        self.df_daily_note['Year'] = self.df_daily_note['Date'].dt.year.astype(str)
        
        grouped_df = self.df_daily_note.groupby(['Year', 'Month']).agg({'total_amount': 'sum'}).reset_index()

        for index, row in grouped_df.iterrows():
            total_amount_thai_baht = f"฿{row['total_amount']:,.0f}"
            self.tree.insert('', 'end', values=(row['Year'], row['Month'], total_amount_thai_baht))

    def on_double_click(self, event):
        # Handle double click event
        item = self.tree.selection()[0]
        year = self.tree.item(item, "values")[0]
        month = self.tree.item(item, "values")[1]
        filtered_data = self.df_daily_note[(self.df_daily_note['Year'] == year) & (self.df_daily_note['Month'] == month)]
        new_window = tk.Toplevel(self.master)
        new_window.title(f"Details for {month}/{year}")

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
        products_rows = self.df_aily_note_item.loc[self.df_aily_note_item['id_dn'] == id_products]

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

        # Call the method to display the latest data
        self.display_df_daily_note()

if __name__ == "__main__":
    root = tk.Tk()
    google_sheet_manager = manager_google_sheet_data()
    opreate_os = Opreate_os()
    gui = GoogleSheetGUI(root)
    root.iconbitmap('contract.ico')

    root.mainloop()
