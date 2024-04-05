import shutil, os
import webbrowser
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import PyPDF2
from bahttext import bahttext
import re

class Opreate_os:
    def __init__(self):
        self.path_all_folder = "source/all_folder"
        self.path_dn_folder = "source/dn_folder"
        self.path_po_folder = "source/po_folder"
        self.path_inv_folder = "source/inv_folder"

        self.path_sum_invoice = "source/sum_invoice_folder"

    def copy_addnew_po(self, source_file, file_name):
        try:
            # Construct the destination file path
            destination_file = os.path.join(self.path_po_folder, f"{file_name}.pdf")

            # Check if a file with the same name already exists in the destination folder
            if os.path.exists(destination_file):
                # If it exists, delete the file
                os.remove(destination_file)
                print(f"Deleted existing file: {destination_file}")

            # Copy the new file to the destination folder
            shutil.copyfile(source_file, destination_file)
            print("PDF file copied successfully.")
        except FileNotFoundError:
            print("Error: Source file not found.")
        except PermissionError:
            print("Error: Permission denied.")
        except Exception as e:
            print(f"An error occurred: {e}")

    def copy_addnew_inv(self, source_file, file_name):
        try:
            # Construct the destination file path
            destination_file = os.path.join(self.path_inv_folder, f"{file_name}.pdf")

            # Check if a file with the same name already exists in the destination folder
            if os.path.exists(destination_file):
                # If it exists, delete the file
                os.remove(destination_file)
                print(f"Deleted existing file: {destination_file}")

            # Copy the new file to the destination folder
            shutil.copyfile(source_file, destination_file)
            print("PDF file copied successfully.")
        except FileNotFoundError:
            print("Error: Source file not found.")
        except PermissionError:
            print("Error: Permission denied.")
        except Exception as e:
            print(f"An error occurred: {e}")

    def chk_file(self,po_filename, inv_filename, dn_filename):
        
        
        po_file_path = os.path.join(self.path_po_folder, f"{po_filename}.pdf")
        inv_file_path = os.path.join(self.path_inv_folder, f"{inv_filename}.pdf")
        dn_file_path = os.path.join(self.path_dn_folder, f"{dn_filename}.pdf")

        if os.path.exists(po_file_path):
            po_exists = True
        else:
            po_exists = False

        if os.path.exists(inv_file_path):
            inv_exists = True
        else:
            inv_exists = False

        if os.path.exists(dn_file_path):
            dn_exists = True
        else:
            dn_exists = False

        return po_exists, inv_exists, dn_exists
    
    def open_po_file(self, file_name):
        file_path = os.path.join(self.path_po_folder, f"{file_name}.pdf")
        webbrowser.open(file_path)

    def open_inv_file(self, file_name):
        file_path = os.path.join(self.path_inv_folder, f"{file_name}.pdf")
        webbrowser.open(file_path)

    def open_dn_file(self, file_name):
        file_path = os.path.join(self.path_dn_folder, f"{file_name}.pdf")
        webbrowser.open(file_path)

    def open_all_file(self, file_name):
        file_path = os.path.join(self.path_all_folder, f"{file_name}.pdf")
        webbrowser.open(file_path)

    def open_sum_invoice_file(self, file_name):
        file_path = os.path.join(self.path_sum_invoice, f"{file_name}.pdf")
        webbrowser.open(file_path)

    def delete_po_file(self, filename):
        file_path = os.path.join(self.path_po_folder, f"{filename}.pdf")
        os.remove(file_path)

    def delete_inv_file(self, filename):
        file_path = os.path.join(self.path_inv_folder, f"{filename}.pdf")
        os.remove(file_path)

    def delete_dn_file(self, filename):
        file_path = os.path.join(self.path_dn_folder, f"{filename}.pdf")
        os.remove(file_path)

    def delete_all_file(self, filename):
        file_path = os.path.join(self.path_all_folder, f"{filename}.pdf")
        os.remove(file_path)

    def delete_sum_invoice_file(self, filename):
        file_path = os.path.join(self.path_sum_invoice, f"{filename}.pdf")
        os.remove(file_path)
    
    def create_delivery_note_pdf(self, filename, data_list, date_text, image_list):

        def one_page(file_path, data_List, image_list):
            n = 0.3
            pdfmetrics.registerFont(TTFont('ThaiFont', 'image/THSarabunNew Bold.ttf'))

            c = canvas.Canvas(file_path)
            logo_path = "image/logo_daikatte.png"
            c.drawImage(logo_path, inch, A4[1] - 1.67*inch- n*inch, width=1.43*inch, height=1.25*inch)

            sign_path = "image\sign_praew.png"
            c.drawImage(sign_path, 1.3*inch, 1.1*inch, width=1.12*inch, height=0.80*inch)

            number = 0
            for barcode in image_list:
                barcode_path = os.path.join("image/", f"{barcode}.gif")
                if number < 5  :

                    c.drawImage(barcode_path, 0.9*inch, A4[1] - 5.2*inch - number*0.5*inch, width=1.26*inch, height=0.35*inch)
                    number = number + 1

            c.drawString(2.74*inch, A4[1] - 0.67*inch - n*inch, "New Line Production Co., Ltd. (Head Office)")
            c.setFont("Helvetica", 10)
            c.drawString(2.74*inch, A4[1] - 0.98*inch - n*inch, "425 Moo 2, Soi Sukhumvit Rd. 13, Pak Nam, Mueang, Samut Prakan, 10270.")
            c.drawString(2.74*inch, A4[1] - 1.26*inch - n*inch, "TEL : 063-1684278 (Praew) , 081-8350860 (Pair)    E-mail : daikatte@gmail.com")
            c.drawString(2.74*inch, A4[1] - 1.61*inch - n*inch, "TAX : 0115560000981")
            c.setFont("ThaiFont", 14)
            c.drawString(2.5*inch,  3.0*inch , "Note : ใบวางบิลชำระแบบเครดิต 30 วัน นับตั้งแต่วันที่จังส่งใบวางบิล")
            c.drawString(1.8*inch,  2.7*inch , "Term Of Payment : Transfer to account name : New Line Production Co., Ltd.( บจก.นิวไลน์โปรดักชั่น )")
            c.drawString(2.90*inch,  2.5*inch , "Swift code: KASITHBK , NO. 066-3-51404-0, KasikornBank")
            c.drawString(1.61*inch,  2.0*inch , "LSSUED BY")
            c.drawString(1.3*inch,  1.0*inch , f"Date: {date_text}")
            c.drawString(3.83*inch,  2.0*inch , "DELIVERED BY")
            c.drawString(3.44*inch,  1.0*inch , "Date:  ............................... ")
            c.drawString(6.19*inch,  2.0*inch , "RECEIVED BY")
            c.drawString(5.82*inch,  1.0*inch , "Date:  ............................... ")
            
            c.setFont("Helvetica", 16)
            c.drawString(3.56*inch, A4[1] - 2.10*inch - n*inch, "Delivery note")

            thai_style = ParagraphStyle(
                name='ThaiFont',
                fontName='ThaiFont',
                fontSize=10,
            )

            # Set row heights for all rows in the table
            row_heights = [0.47*inch, 0.62*inch, 0.56*inch, 0.24*inch, 0.3*inch, 0.5*inch, 0.5*inch, 0.5*inch, 0.5*inch, 0.5*inch, 0.3*inch, 0.3*inch, 0.3*inch]

            # Create table and set column widths and row heights
            table = Table(data_List, colWidths=[1.62*inch, 2.17*inch, 0.63*inch, 0.55*inch, 0.96*inch, 0.87*inch], rowHeights=row_heights)

            table_style = [
                ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Add border to the table
                #('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Bold font for headers
                #('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),  # Background color for headers
                ('SPAN', (1, 0), (2, 0)),#DONKI (THAILAND) CO., LTD.
                ('SPAN', (1, 1), (2, 1)),#Main Live
                ('SPAN', (1, 2), (2, 2)),#Delivary live
                ('SPAN', (1, 3), (2, 3)),#Tax 
                ('SPAN', (3, 0), (5, 0)),#Number 
                ('SPAN', (3, 1), (5, 3)),#Payment
                ('SPAN', (0, 10), (3, 12)),#thai baht

                ('VALIGN', (0, 0), (0, 0), 'TOP'), #Company :
                ('VALIGN', (0, 1), (0, 1), 'TOP'), #Address :
                ('VALIGN', (0, 2), (0, 2), 'TOP'), #Delivery Address :
                ('FONTSIZE', (1, 1), (1, 1), 8),
                ('FONTSIZE', (1, 2), (1, 2), 8),

                ('VALIGN', (1, 0), (2, 0), 'TOP'), #DONKI (THAILAND) CO., LTD.
                ('VALIGN', (1, 1), (2, 1), 'TOP'), #Main Live
                ('VALIGN', (1, 2), (2, 2), 'TOP'), #Delivary live
                ('VALIGN', (3, 0), (5, 0), 'TOP'),#Number 


                ('ALIGN', (1, 0), (2, 3), 'LEFT'),
                ('VALIGN', (3, 1), (5, 3), 'TOP'),
                ('VALIGN', (1, 1), (2, 1), 'TOP'),
                ('VALIGN', (1, 3), (2, 3), 'TOP'),
                ('ALIGN', (0, 4), (0, 4), 'CENTER'),#Item code
                ('ALIGN', (1, 4), (1, 4), 'CENTER'),#Description
                ('ALIGN', (2, 4), (2, 4), 'CENTER'),#Quantity
                ('ALIGN', (3, 4), (3, 4), 'CENTER'),#Unit
                ('ALIGN', (4, 4), (4, 4), 'CENTER'),#Unit Price
                ('ALIGN', (5, 4), (5, 4), 'CENTER'),#Amount

                ('VALIGN', (0, 5), (0, 5), 'MIDDLE'),#list item code1
                ('ALIGN', (0, 5), (0, 5), 'CENTER'),#list item code1
                ('VALIGN', (0, 6), (0, 6), 'MIDDLE'),#list item code2
                ('ALIGN', (0, 6), (0, 6), 'CENTER'),#list item code2
                ('VALIGN', (0, 7), (0, 7), 'MIDDLE'),#list item code3
                ('ALIGN', (0, 7), (0, 7), 'CENTER'),#list item code3
                ('VALIGN', (0, 8), (0, 8), 'MIDDLE'),#list item code4
                ('ALIGN', (0, 8), (0, 8), 'CENTER'),#list item code4
                ('VALIGN', (0, 9), (0, 9), 'MIDDLE'),#list item code5
                ('ALIGN', (0, 9), (0, 9), 'CENTER'),#list item code5

                ('VALIGN', (1, 5), (1, 5), 'MIDDLE'),#Description1
                ('ALIGN', (1, 5), (1, 5), 'CENTER'),#Description1
                ('VALIGN', (1, 6), (1, 6), 'MIDDLE'),#Description2
                ('ALIGN', (1, 6), (1, 6), 'CENTER'),#Descriptione2
                ('VALIGN', (1, 7), (1, 7), 'MIDDLE'),#Description3
                ('ALIGN', (1, 7), (1, 7), 'CENTER'),#Description3
                ('VALIGN', (1, 8), (1, 8), 'MIDDLE'),#Description4
                ('ALIGN', (1, 8), (1, 8), 'CENTER'),#Description4
                ('VALIGN', (1, 9), (1, 9), 'MIDDLE'),#Description5
                ('ALIGN', (1, 9), (1, 9), 'CENTER'),#Description5

                ('FONTSIZE', (1, 5), (1, 5), 8),#Description1
                ('FONTSIZE', (1, 6), (1, 6), 8),#Description1
                ('FONTSIZE', (1, 7), (1, 7), 8),#Description1
                ('FONTSIZE', (1, 8), (1, 8), 8),#Description1
                ('FONTSIZE', (1, 9), (1, 9), 8),#Description1

                ('VALIGN', (2, 5), (2, 5), 'MIDDLE'),#Quantity1
                ('ALIGN', (2, 5), (2, 5), 'CENTER'),#Quantity1
                ('VALIGN', (2, 6), (2, 6), 'MIDDLE'),#Quantity2
                ('ALIGN', (2, 6), (2, 6), 'CENTER'),#Quantity2
                ('VALIGN', (2, 7), (2, 7), 'MIDDLE'),#Quantity3
                ('ALIGN', (2, 7), (2, 7), 'CENTER'),#Quantity3
                ('VALIGN', (2, 8), (2, 8), 'MIDDLE'),#Quantity4
                ('ALIGN', (2, 8), (2, 8), 'CENTER'),#Quantity4
                ('VALIGN', (2, 9), (2, 9), 'MIDDLE'),#Quantity5
                ('ALIGN', (2, 9), (2, 9), 'CENTER'),#Quantity5

                ('VALIGN', (3, 5), (3, 5), 'MIDDLE'),#Unit1
                ('ALIGN', (3, 5), (3, 5), 'CENTER'),#Unit1
                ('VALIGN', (3, 6), (3, 6), 'MIDDLE'),#Unit2
                ('ALIGN', (3, 6), (3, 6), 'CENTER'),#Unit2
                ('VALIGN', (3, 7), (3, 7), 'MIDDLE'),#Unit3
                ('ALIGN', (3, 7), (3, 7), 'CENTER'),#Unit3
                ('VALIGN', (3, 8), (3, 8), 'MIDDLE'),#Unit4
                ('ALIGN', (3, 8), (3, 8), 'CENTER'),#Unit4
                ('VALIGN', (3, 9), (3, 9), 'MIDDLE'),#Unit5
                ('ALIGN', (3, 9), (3, 9), 'CENTER'),#Unit5

                ('VALIGN', (4, 5), (4, 5), 'MIDDLE'),#Unit Price1
                ('ALIGN', (4, 5), (4, 5), 'CENTER'),#Unit Price1
                ('VALIGN', (4, 6), (4, 6), 'MIDDLE'),#Unit Price2
                ('ALIGN', (4, 6), (4, 6), 'CENTER'),#Unit Price2
                ('VALIGN', (4, 7), (4, 7), 'MIDDLE'),#Unit Price3
                ('ALIGN', (4, 7), (4, 7), 'CENTER'),#Unit Price3
                ('VALIGN', (4, 8), (4, 8), 'MIDDLE'),#Unit Price4
                ('ALIGN', (4, 8), (4, 8), 'CENTER'),#Unit Price4
                ('VALIGN', (4, 9), (4, 9), 'MIDDLE'),#Unit Price5
                ('ALIGN', (4, 9), (4, 9), 'CENTER'),#Unit Price5

                

                ('FONTNAME', (5, 5), (5, 5), 'ThaiFont'),#Amount1
                ('FONTNAME', (5, 6), (5, 6), 'ThaiFont'),#Amount2
                ('FONTNAME', (5, 7), (5, 7), 'ThaiFont'),#Amount3
                ('FONTNAME', (5, 8), (5, 8), 'ThaiFont'),#Amount4
                ('FONTNAME', (5, 9), (5, 9), 'ThaiFont'),#Amount5

                ('FONTSIZE', (5, 5), (5, 5), 12),#Amount1
                ('FONTSIZE', (5, 6), (5, 6), 12),#Amount2
                ('FONTSIZE', (5, 7), (5, 7), 12),#Amount3
                ('FONTSIZE', (5, 8), (5, 8), 12),#Amount4
                ('FONTSIZE', (5, 9), (5, 9), 12),#Amount5

                ('VALIGN', (5, 10), (5, 10), 'MIDDLE'),#
                ('ALIGN', (5, 10), (5, 10), 'CENTER'),#
                ('VALIGN', (5, 11), (5, 11), 'MIDDLE'),#vat
                ('ALIGN', (5, 11), (5, 11), 'CENTER'),#vat
                ('VALIGN', (5, 12), (5, 12), 'MIDDLE'),#total anmount
                ('ALIGN', (5, 12), (5, 12), 'CENTER'),#total anmount

                ('FONTNAME', (5, 10), (5, 10), 'ThaiFont'),#Amount3
                ('FONTNAME', (5, 11), (5, 11), 'ThaiFont'),#Amount4
                ('FONTNAME', (5, 12), (5, 12), 'ThaiFont'),
                ('FONTSIZE', (5, 10), (5, 10), 12),#Amount5
                ('FONTSIZE', (5, 11), (5, 11), 12),#Amount5
                ('FONTSIZE', (5, 12), (5, 12), 12),#Amount5

                ('VALIGN', (0, 10), (5, 12), 'MIDDLE'),# Thai font
                ('ALIGN', (0, 10), (5, 12), 'CENTER'),# Thai font
                ('FONTNAME', (0, 10), (3, 12), 'ThaiFont'),  # Thai font
                ('FONTSIZE', (0, 10), (3, 12), 14),
                
                ('VALIGN', (5, 5), (5, 5), 'MIDDLE'),#Amount1
                ('ALIGN', (5, 5), (5, 5), 'CENTER'),#Amount1
                ('VALIGN', (5, 6), (5, 6), 'MIDDLE'),#Amount2
                ('ALIGN', (5, 6), (5, 6), 'CENTER'),#Amount2
                ('VALIGN', (5, 7), (5, 7), 'MIDDLE'),#Amount3
                ('ALIGN', (5, 7), (5, 7), 'CENTER'),#Amount3
                ('VALIGN', (5, 8), (5, 8), 'MIDDLE'),#Amount4
                ('ALIGN', (5, 8), (5, 8), 'CENTER'),#Amount4
                ('VALIGN', (5, 9), (5, 9), 'MIDDLE'),#Amount5
                ('ALIGN', (5, 9), (5, 9), 'CENTER'),#Amount5




            ]

            table.setStyle(TableStyle(table_style))

            table.wrapOn(c, 6.81*inch, 5*inch)
            table.drawOn(c, 0.75*inch, 3.5*inch)

            c.save()
        
        def two_page_1(file_path, data_List, image_list):
            n = 0.3
            pdfmetrics.registerFont(TTFont('ThaiFont', 'image/THSarabunNew Bold.ttf'))

            c = canvas.Canvas(file_path)
            logo_path = "image/logo_daikatte.png"
            c.drawImage(logo_path, inch, A4[1] - 1.8*inch- n*inch, width=1.43*inch, height=1.25*inch)

            sign_path = "image\sign_praew.png"
            c.drawImage(sign_path, 1.3*inch, 1.1*inch, width=1.12*inch, height=0.80*inch)

            number = 0
            for barcode in image_list:
                barcode_path = os.path.join("image/", f"{barcode}.gif")
                if number < 5  :

                    c.drawImage(barcode_path, 0.9*inch, A4[1] - 5.3*inch - number*0.5*inch, width=1.26*inch, height=0.35*inch)
                    number = number + 1

            c.drawString(2.74*inch, A4[1] - 0.67*inch - n*inch, "New Line Production Co., Ltd. (Head Office)")
            c.setFont("Helvetica", 10)
            c.drawString(2.74*inch, A4[1] - 0.98*inch - n*inch, "425 Moo 2, Soi Sukhumvit Rd. 13, Pak Nam, Mueang, Samut Prakan, 10270.")
            c.drawString(2.74*inch, A4[1] - 1.26*inch - n*inch, "TEL : 063-1684278 (Praew) , 081-8350860 (Pair)    E-mail : daikatte@gmail.com")
            c.drawString(2.74*inch, A4[1] - 1.61*inch - n*inch, "TAX : 0115560000981")
            c.setFont("ThaiFont", 14)
            c.drawString(2.5*inch,  3.0*inch , "Note : ใบวางบิลชำระแบบเครดิต 30 วัน นับตั้งแต่วันที่จังส่งใบวางบิล")
            c.drawString(1.8*inch,  2.7*inch , "Term Of Payment : Transfer to account name : New Line Production Co., Ltd.( บจก.นิวไลน์โปรดักชั่น )")
            c.drawString(2.90*inch,  2.5*inch , "Swift code: KASITHBK , NO. 066-3-51404-0, KasikornBank")
            c.drawString(1.61*inch,  2.0*inch , "LSSUED BY")
            c.drawString(1.3*inch,  1.0*inch , f"Date: {date_text}")
            c.drawString(3.83*inch,  2.0*inch , "DELIVERED BY")
            c.drawString(3.44*inch,  1.0*inch , "Date:  ............................... ")
            c.drawString(6.19*inch,  2.0*inch , "RECEIVED BY")
            c.drawString(5.82*inch,  1.0*inch , "Date:  ............................... ")
            c.drawString(7.5*inch,  0.2*inch , "Page 1/2")


            c.setFont("Helvetica", 16)
            c.drawString(3.56*inch, A4[1] - 2.10*inch - n*inch, "Delivery note")

            # Set row heights for all rows in the table
            row_heights = [0.47*inch, 0.62*inch, 0.56*inch, 0.24*inch, 0.3*inch, 0.5*inch, 0.5*inch, 0.5*inch, 0.5*inch, 0.5*inch]

            # Create table and set column widths and row heights
            table = Table(data_List, colWidths=[1.62*inch, 2.17*inch, 0.63*inch, 0.55*inch, 0.96*inch, 0.87*inch], rowHeights=row_heights)

            table_style = [
                ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Add border to the table
                #('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Bold font for headers
                #('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),  # Background color for headers
                ('SPAN', (1, 0), (2, 0)),#DONKI (THAILAND) CO., LTD.
                ('SPAN', (1, 1), (2, 1)),#Main Live
                ('SPAN', (1, 2), (2, 2)),#Delivary live
                ('SPAN', (1, 3), (2, 3)),#Tax 
                ('SPAN', (3, 0), (5, 0)),#Number 
                ('SPAN', (3, 1), (5, 3)),#Payment
                ('SPAN', (0, 10), (3, 12)),#thai baht

                ('VALIGN', (0, 0), (0, 0), 'TOP'), #Company :
                ('VALIGN', (0, 1), (0, 1), 'TOP'), #Address :
                ('VALIGN', (0, 2), (0, 2), 'TOP'), #Delivery Address :
                ('FONTSIZE', (1, 1), (1, 1), 8),
                ('FONTSIZE', (1, 2), (1, 2), 8),

                ('VALIGN', (1, 0), (2, 0), 'TOP'), #DONKI (THAILAND) CO., LTD.
                ('VALIGN', (1, 1), (2, 1), 'TOP'), #Main Live
                ('VALIGN', (1, 2), (2, 2), 'TOP'), #Delivary live
                ('VALIGN', (3, 0), (5, 0), 'TOP'),#Number 


                ('ALIGN', (1, 0), (2, 3), 'LEFT'),
                ('VALIGN', (3, 1), (5, 3), 'TOP'),
                ('VALIGN', (1, 1), (2, 1), 'TOP'),
                ('VALIGN', (1, 3), (2, 3), 'TOP'),
                ('ALIGN', (0, 4), (0, 4), 'CENTER'),#Item code
                ('ALIGN', (1, 4), (1, 4), 'CENTER'),#Description
                ('ALIGN', (2, 4), (2, 4), 'CENTER'),#Quantity
                ('ALIGN', (3, 4), (3, 4), 'CENTER'),#Unit
                ('ALIGN', (4, 4), (4, 4), 'CENTER'),#Unit Price
                ('ALIGN', (5, 4), (5, 4), 'CENTER'),#Amount

                ('VALIGN', (0, 5), (0, 5), 'MIDDLE'),#list item code1
                ('ALIGN', (0, 5), (0, 5), 'CENTER'),#list item code1
                ('VALIGN', (0, 6), (0, 6), 'MIDDLE'),#list item code2
                ('ALIGN', (0, 6), (0, 6), 'CENTER'),#list item code2
                ('VALIGN', (0, 7), (0, 7), 'MIDDLE'),#list item code3
                ('ALIGN', (0, 7), (0, 7), 'CENTER'),#list item code3
                ('VALIGN', (0, 8), (0, 8), 'MIDDLE'),#list item code4
                ('ALIGN', (0, 8), (0, 8), 'CENTER'),#list item code4
                ('VALIGN', (0, 9), (0, 9), 'MIDDLE'),#list item code5
                ('ALIGN', (0, 9), (0, 9), 'CENTER'),#list item code5

                ('VALIGN', (1, 5), (1, 5), 'MIDDLE'),#Description1
                ('ALIGN', (1, 5), (1, 5), 'CENTER'),#Description1
                ('VALIGN', (1, 6), (1, 6), 'MIDDLE'),#Description2
                ('ALIGN', (1, 6), (1, 6), 'CENTER'),#Descriptione2
                ('VALIGN', (1, 7), (1, 7), 'MIDDLE'),#Description3
                ('ALIGN', (1, 7), (1, 7), 'CENTER'),#Description3
                ('VALIGN', (1, 8), (1, 8), 'MIDDLE'),#Description4
                ('ALIGN', (1, 8), (1, 8), 'CENTER'),#Description4
                ('VALIGN', (1, 9), (1, 9), 'MIDDLE'),#Description5
                ('ALIGN', (1, 9), (1, 9), 'CENTER'),#Description5

                ('FONTSIZE', (1, 5), (1, 5), 8),#Description1
                ('FONTSIZE', (1, 6), (1, 6), 8),#Description1
                ('FONTSIZE', (1, 7), (1, 7), 8),#Description1
                ('FONTSIZE', (1, 8), (1, 8), 8),#Description1
                ('FONTSIZE', (1, 9), (1, 9), 8),#Description1

                ('VALIGN', (2, 5), (2, 5), 'MIDDLE'),#Quantity1
                ('ALIGN', (2, 5), (2, 5), 'CENTER'),#Quantity1
                ('VALIGN', (2, 6), (2, 6), 'MIDDLE'),#Quantity2
                ('ALIGN', (2, 6), (2, 6), 'CENTER'),#Quantity2
                ('VALIGN', (2, 7), (2, 7), 'MIDDLE'),#Quantity3
                ('ALIGN', (2, 7), (2, 7), 'CENTER'),#Quantity3
                ('VALIGN', (2, 8), (2, 8), 'MIDDLE'),#Quantity4
                ('ALIGN', (2, 8), (2, 8), 'CENTER'),#Quantity4
                ('VALIGN', (2, 9), (2, 9), 'MIDDLE'),#Quantity5
                ('ALIGN', (2, 9), (2, 9), 'CENTER'),#Quantity5

                ('VALIGN', (3, 5), (3, 5), 'MIDDLE'),#Unit1
                ('ALIGN', (3, 5), (3, 5), 'CENTER'),#Unit1
                ('VALIGN', (3, 6), (3, 6), 'MIDDLE'),#Unit2
                ('ALIGN', (3, 6), (3, 6), 'CENTER'),#Unit2
                ('VALIGN', (3, 7), (3, 7), 'MIDDLE'),#Unit3
                ('ALIGN', (3, 7), (3, 7), 'CENTER'),#Unit3
                ('VALIGN', (3, 8), (3, 8), 'MIDDLE'),#Unit4
                ('ALIGN', (3, 8), (3, 8), 'CENTER'),#Unit4
                ('VALIGN', (3, 9), (3, 9), 'MIDDLE'),#Unit5
                ('ALIGN', (3, 9), (3, 9), 'CENTER'),#Unit5

                ('VALIGN', (4, 5), (4, 5), 'MIDDLE'),#Unit Price1
                ('ALIGN', (4, 5), (4, 5), 'CENTER'),#Unit Price1
                ('VALIGN', (4, 6), (4, 6), 'MIDDLE'),#Unit Price2
                ('ALIGN', (4, 6), (4, 6), 'CENTER'),#Unit Price2
                ('VALIGN', (4, 7), (4, 7), 'MIDDLE'),#Unit Price3
                ('ALIGN', (4, 7), (4, 7), 'CENTER'),#Unit Price3
                ('VALIGN', (4, 8), (4, 8), 'MIDDLE'),#Unit Price4
                ('ALIGN', (4, 8), (4, 8), 'CENTER'),#Unit Price4
                ('VALIGN', (4, 9), (4, 9), 'MIDDLE'),#Unit Price5
                ('ALIGN', (4, 9), (4, 9), 'CENTER'),#Unit Price5

                ('FONTNAME', (5, 5), (5, 5), 'ThaiFont'),#Amount1
                ('FONTNAME', (5, 6), (5, 6), 'ThaiFont'),#Amount2
                ('FONTNAME', (5, 7), (5, 7), 'ThaiFont'),#Amount3
                ('FONTNAME', (5, 8), (5, 8), 'ThaiFont'),#Amount4
                ('FONTNAME', (5, 9), (5, 9), 'ThaiFont'),#Amount5

                ('FONTSIZE', (5, 5), (5, 5), 12),#Amount1
                ('FONTSIZE', (5, 6), (5, 6), 12),#Amount2
                ('FONTSIZE', (5, 7), (5, 7), 12),#Amount3
                ('FONTSIZE', (5, 8), (5, 8), 12),#Amount4
                ('FONTSIZE', (5, 9), (5, 9), 12),#Amount5

                
                ('VALIGN', (5, 5), (5, 5), 'MIDDLE'),#Amount1
                ('ALIGN', (5, 5), (5, 5), 'CENTER'),#Amount1
                ('VALIGN', (5, 6), (5, 6), 'MIDDLE'),#Amount2
                ('ALIGN', (5, 6), (5, 6), 'CENTER'),#Amount2
                ('VALIGN', (5, 7), (5, 7), 'MIDDLE'),#Amount3
                ('ALIGN', (5, 7), (5, 7), 'CENTER'),#Amount3
                ('VALIGN', (5, 8), (5, 8), 'MIDDLE'),#Amount4
                ('ALIGN', (5, 8), (5, 8), 'CENTER'),#Amount4
                ('VALIGN', (5, 9), (5, 9), 'MIDDLE'),#Amount5
                ('ALIGN', (5, 9), (5, 9), 'CENTER'),#Amount5




            ]

            table.setStyle(TableStyle(table_style))

            table.wrapOn(c, 6.81*inch, 5*inch)
            table.drawOn(c, 0.75*inch, 4.3*inch)


            c.save()

        def two_page_2(file_path, data_List, image_list):
            n = 0.3
            pdfmetrics.registerFont(TTFont('ThaiFont', 'image/THSarabunNew Bold.ttf'))

            c = canvas.Canvas(file_path)
            logo_path = "image/logo_daikatte.png"
            c.drawImage(logo_path, inch, A4[1] - 1.8*inch- n*inch, width=1.43*inch, height=1.25*inch)

            sign_path = "image\sign_praew.png"
            c.drawImage(sign_path, 1.3*inch, 1.1*inch, width=1.12*inch, height=0.80*inch)

            number = 0
            del image_list[0:5]
            for barcode in image_list:
                if number < 5 :
                    barcode_path = os.path.join("image/", f"{barcode}.gif")
                    c.drawImage(barcode_path, 0.9*inch, A4[1] - 5.3*inch - (number)*0.5*inch, width=1.26*inch, height=0.35*inch)
                    number = number + 1

            c.drawString(2.74*inch, A4[1] - 0.67*inch - n*inch, "New Line Production Co., Ltd. (Head Office)")
            c.setFont("Helvetica", 10)
            c.drawString(2.74*inch, A4[1] - 0.98*inch - n*inch, "425 Moo 2, Soi Sukhumvit Rd. 13, Pak Nam, Mueang, Samut Prakan, 10270.")
            c.drawString(2.74*inch, A4[1] - 1.26*inch - n*inch, "TEL : 063-1684278 (Praew) , 081-8350860 (Pair)    E-mail : daikatte@gmail.com")
            c.drawString(2.74*inch, A4[1] - 1.61*inch - n*inch, "TAX : 0115560000981")
            c.setFont("ThaiFont", 14)
            c.drawString(2.5*inch,  3.0*inch , "Note : ใบวางบิลชำระแบบเครดิต 30 วัน นับตั้งแต่วันที่จังส่งใบวางบิล")
            c.drawString(1.8*inch,  2.7*inch , "Term Of Payment : Transfer to account name : New Line Production Co., Ltd.( บจก.นิวไลน์โปรดักชั่น )")
            c.drawString(2.90*inch,  2.5*inch , "Swift code: KASITHBK , NO. 066-3-51404-0, KasikornBank")
            c.drawString(1.61*inch,  2.0*inch , "LSSUED BY")
            c.drawString(1.3*inch,  1.0*inch , f"Date: {date_text}")
            c.drawString(3.83*inch,  2.0*inch , "DELIVERED BY")
            c.drawString(3.44*inch,  1.0*inch , "Date:  ............................... ")
            c.drawString(6.19*inch,  2.0*inch , "RECEIVED BY")
            c.drawString(5.82*inch,  1.0*inch , "Date:  ............................... ")
            c.drawString(7.5*inch,  0.2*inch , "Page 2/2")
            
            c.setFont("Helvetica", 16)
            c.drawString(3.56*inch, A4[1] - 2.10*inch - n*inch, "Delivery note")

            thai_style = ParagraphStyle(
                name='ThaiFont',
                fontName='ThaiFont',
                fontSize=10,
            )

            # Set row heights for all rows in the table
            row_heights = [0.47*inch, 0.62*inch, 0.56*inch, 0.24*inch, 0.3*inch, 0.5*inch, 0.5*inch, 0.5*inch, 0.5*inch, 0.5*inch, 0.3*inch, 0.3*inch, 0.3*inch]

            # Create table and set column widths and row heights
            table = Table(data_List, colWidths=[1.62*inch, 2.17*inch, 0.63*inch, 0.55*inch, 0.96*inch, 0.87*inch], rowHeights=row_heights)

            table_style = [
                ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Add border to the table
                #('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Bold font for headers
                #('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),  # Background color for headers
                ('SPAN', (1, 0), (2, 0)),#DONKI (THAILAND) CO., LTD.
                ('SPAN', (1, 1), (2, 1)),#Main Live
                ('SPAN', (1, 2), (2, 2)),#Delivary live
                ('SPAN', (1, 3), (2, 3)),#Tax 
                ('SPAN', (3, 0), (5, 0)),#Number 
                ('SPAN', (3, 1), (5, 3)),#Payment
                ('SPAN', (0, 10), (3, 12)),#thai baht

                ('VALIGN', (0, 0), (0, 0), 'TOP'), #Company :
                ('VALIGN', (0, 1), (0, 1), 'TOP'), #Address :
                ('VALIGN', (0, 2), (0, 2), 'TOP'), #Delivery Address :
                ('FONTSIZE', (1, 1), (1, 1), 8),
                ('FONTSIZE', (1, 2), (1, 2), 8),

                ('VALIGN', (1, 0), (2, 0), 'TOP'), #DONKI (THAILAND) CO., LTD.
                ('VALIGN', (1, 1), (2, 1), 'TOP'), #Main Live
                ('VALIGN', (1, 2), (2, 2), 'TOP'), #Delivary live
                ('VALIGN', (3, 0), (5, 0), 'TOP'),#Number 


                ('ALIGN', (1, 0), (2, 3), 'LEFT'),
                ('VALIGN', (3, 1), (5, 3), 'TOP'),
                ('VALIGN', (1, 1), (2, 1), 'TOP'),
                ('VALIGN', (1, 3), (2, 3), 'TOP'),
                ('ALIGN', (0, 4), (0, 4), 'CENTER'),#Item code
                ('ALIGN', (1, 4), (1, 4), 'CENTER'),#Description
                ('ALIGN', (2, 4), (2, 4), 'CENTER'),#Quantity
                ('ALIGN', (3, 4), (3, 4), 'CENTER'),#Unit
                ('ALIGN', (4, 4), (4, 4), 'CENTER'),#Unit Price
                ('ALIGN', (5, 4), (5, 4), 'CENTER'),#Amount

                ('VALIGN', (0, 5), (0, 5), 'MIDDLE'),#list item code1
                ('ALIGN', (0, 5), (0, 5), 'CENTER'),#list item code1
                ('VALIGN', (0, 6), (0, 6), 'MIDDLE'),#list item code2
                ('ALIGN', (0, 6), (0, 6), 'CENTER'),#list item code2
                ('VALIGN', (0, 7), (0, 7), 'MIDDLE'),#list item code3
                ('ALIGN', (0, 7), (0, 7), 'CENTER'),#list item code3
                ('VALIGN', (0, 8), (0, 8), 'MIDDLE'),#list item code4
                ('ALIGN', (0, 8), (0, 8), 'CENTER'),#list item code4
                ('VALIGN', (0, 9), (0, 9), 'MIDDLE'),#list item code5
                ('ALIGN', (0, 9), (0, 9), 'CENTER'),#list item code5

                ('VALIGN', (1, 5), (1, 5), 'MIDDLE'),#Description1
                ('ALIGN', (1, 5), (1, 5), 'CENTER'),#Description1
                ('VALIGN', (1, 6), (1, 6), 'MIDDLE'),#Description2
                ('ALIGN', (1, 6), (1, 6), 'CENTER'),#Descriptione2
                ('VALIGN', (1, 7), (1, 7), 'MIDDLE'),#Description3
                ('ALIGN', (1, 7), (1, 7), 'CENTER'),#Description3
                ('VALIGN', (1, 8), (1, 8), 'MIDDLE'),#Description4
                ('ALIGN', (1, 8), (1, 8), 'CENTER'),#Description4
                ('VALIGN', (1, 9), (1, 9), 'MIDDLE'),#Description5
                ('ALIGN', (1, 9), (1, 9), 'CENTER'),#Description5

                ('FONTSIZE', (1, 5), (1, 5), 8),#Description1
                ('FONTSIZE', (1, 6), (1, 6), 8),#Description1
                ('FONTSIZE', (1, 7), (1, 7), 8),#Description1
                ('FONTSIZE', (1, 8), (1, 8), 8),#Description1
                ('FONTSIZE', (1, 9), (1, 9), 8),#Description1

                ('VALIGN', (2, 5), (2, 5), 'MIDDLE'),#Quantity1
                ('ALIGN', (2, 5), (2, 5), 'CENTER'),#Quantity1
                ('VALIGN', (2, 6), (2, 6), 'MIDDLE'),#Quantity2
                ('ALIGN', (2, 6), (2, 6), 'CENTER'),#Quantity2
                ('VALIGN', (2, 7), (2, 7), 'MIDDLE'),#Quantity3
                ('ALIGN', (2, 7), (2, 7), 'CENTER'),#Quantity3
                ('VALIGN', (2, 8), (2, 8), 'MIDDLE'),#Quantity4
                ('ALIGN', (2, 8), (2, 8), 'CENTER'),#Quantity4
                ('VALIGN', (2, 9), (2, 9), 'MIDDLE'),#Quantity5
                ('ALIGN', (2, 9), (2, 9), 'CENTER'),#Quantity5

                ('VALIGN', (3, 5), (3, 5), 'MIDDLE'),#Unit1
                ('ALIGN', (3, 5), (3, 5), 'CENTER'),#Unit1
                ('VALIGN', (3, 6), (3, 6), 'MIDDLE'),#Unit2
                ('ALIGN', (3, 6), (3, 6), 'CENTER'),#Unit2
                ('VALIGN', (3, 7), (3, 7), 'MIDDLE'),#Unit3
                ('ALIGN', (3, 7), (3, 7), 'CENTER'),#Unit3
                ('VALIGN', (3, 8), (3, 8), 'MIDDLE'),#Unit4
                ('ALIGN', (3, 8), (3, 8), 'CENTER'),#Unit4
                ('VALIGN', (3, 9), (3, 9), 'MIDDLE'),#Unit5
                ('ALIGN', (3, 9), (3, 9), 'CENTER'),#Unit5

                ('VALIGN', (4, 5), (4, 5), 'MIDDLE'),#Unit Price1
                ('ALIGN', (4, 5), (4, 5), 'CENTER'),#Unit Price1
                ('VALIGN', (4, 6), (4, 6), 'MIDDLE'),#Unit Price2
                ('ALIGN', (4, 6), (4, 6), 'CENTER'),#Unit Price2
                ('VALIGN', (4, 7), (4, 7), 'MIDDLE'),#Unit Price3
                ('ALIGN', (4, 7), (4, 7), 'CENTER'),#Unit Price3
                ('VALIGN', (4, 8), (4, 8), 'MIDDLE'),#Unit Price4
                ('ALIGN', (4, 8), (4, 8), 'CENTER'),#Unit Price4
                ('VALIGN', (4, 9), (4, 9), 'MIDDLE'),#Unit Price5
                ('ALIGN', (4, 9), (4, 9), 'CENTER'),#Unit Price5

                

                ('FONTNAME', (5, 5), (5, 5), 'ThaiFont'),#Amount1
                ('FONTNAME', (5, 6), (5, 6), 'ThaiFont'),#Amount2
                ('FONTNAME', (5, 7), (5, 7), 'ThaiFont'),#Amount3
                ('FONTNAME', (5, 8), (5, 8), 'ThaiFont'),#Amount4
                ('FONTNAME', (5, 9), (5, 9), 'ThaiFont'),#Amount5

                ('FONTSIZE', (5, 5), (5, 5), 12),#Amount1
                ('FONTSIZE', (5, 6), (5, 6), 12),#Amount2
                ('FONTSIZE', (5, 7), (5, 7), 12),#Amount3
                ('FONTSIZE', (5, 8), (5, 8), 12),#Amount4
                ('FONTSIZE', (5, 9), (5, 9), 12),#Amount5

                ('VALIGN', (5, 10), (5, 10), 'MIDDLE'),#
                ('ALIGN', (5, 10), (5, 10), 'CENTER'),#
                ('VALIGN', (5, 11), (5, 11), 'MIDDLE'),#vat
                ('ALIGN', (5, 11), (5, 11), 'CENTER'),#vat
                ('VALIGN', (5, 12), (5, 12), 'MIDDLE'),#total anmount
                ('ALIGN', (5, 12), (5, 12), 'CENTER'),#total anmount

                ('FONTNAME', (5, 10), (5, 10), 'ThaiFont'),#Amount3
                ('FONTNAME', (5, 11), (5, 11), 'ThaiFont'),#Amount4
                ('FONTNAME', (5, 12), (5, 12), 'ThaiFont'),
                ('FONTSIZE', (5, 10), (5, 10), 12),#Amount5
                ('FONTSIZE', (5, 11), (5, 11), 12),#Amount5
                ('FONTSIZE', (5, 12), (5, 12), 12),#Amount5

                ('VALIGN', (0, 10), (5, 12), 'MIDDLE'),# Thai font
                ('ALIGN', (0, 10), (5, 12), 'CENTER'),# Thai font
                ('FONTNAME', (0, 10), (3, 12), 'ThaiFont'),  # Thai font
                ('FONTSIZE', (0, 10), (3, 12), 14),
                
                ('VALIGN', (5, 5), (5, 5), 'MIDDLE'),#Amount1
                ('ALIGN', (5, 5), (5, 5), 'CENTER'),#Amount1
                ('VALIGN', (5, 6), (5, 6), 'MIDDLE'),#Amount2
                ('ALIGN', (5, 6), (5, 6), 'CENTER'),#Amount2
                ('VALIGN', (5, 7), (5, 7), 'MIDDLE'),#Amount3
                ('ALIGN', (5, 7), (5, 7), 'CENTER'),#Amount3
                ('VALIGN', (5, 8), (5, 8), 'MIDDLE'),#Amount4
                ('ALIGN', (5, 8), (5, 8), 'CENTER'),#Amount4
                ('VALIGN', (5, 9), (5, 9), 'MIDDLE'),#Amount5
                ('ALIGN', (5, 9), (5, 9), 'CENTER'),#Amount5




            ]

            table.setStyle(TableStyle(table_style))

            table.wrapOn(c, 6.81*inch, 5*inch)
            table.drawOn(c, 0.75*inch, 3.5*inch)


            c.save()

        if len(data_list) == 1:

            file_path = os.path.join(self.path_dn_folder, f"{filename}.pdf")
            doc = SimpleDocTemplate(file_path, pagesize=A4)

            styles = getSampleStyleSheet()
            normal_style = styles["Normal"]

            content = []

            text = "This is a sample text for the delivery note."
            paragraph = Paragraph(text, normal_style)
            content.append(paragraph)
            doc.build(content)
            one_page(file_path, data_list[0], image_list)

            print(f"Delivery note PDF created successfully at: {file_path}")
        

        elif len(data_list) == 2:

            file_path_1 = os.path.join(self.path_dn_folder, f"{filename}_1.pdf")
            doc_page_1 = SimpleDocTemplate(file_path_1, pagesize=A4)
            styles_page_1 = getSampleStyleSheet()
            normal_style_page_1 = styles_page_1["Normal"]

            content_page_1 = []

            text_page_1 = "This is a sample text for the delivery note."
            paragraph_page_1 = Paragraph(text_page_1, normal_style_page_1)
            content_page_1.append(paragraph_page_1)
            doc_page_1.build(content_page_1)
            two_page_1(file_path_1, data_list[0], image_list)

            # Page 2 
            file_path_2 = os.path.join(self.path_dn_folder, f"{filename}_2.pdf")
            doc_page_2 = SimpleDocTemplate(file_path_2, pagesize=A4)
            stylesdoc_page_2 = getSampleStyleSheet()
            normal_style_page_2  = stylesdoc_page_2["Normal"]

            content_page_2 = []

            text_page_2  = "This is a sample text for the delivery note."
            paragraph_page_2  = Paragraph(text_page_2 , normal_style_page_2 )
            content_page_2 .append(paragraph_page_2 )
            doc_page_2 .build(content_page_2)
            two_page_2(file_path_2, data_list[1], image_list)

            def merge_pdfs(pdf1_path, pdf2_path, output_path):
                output_path = os.path.join(self.path_dn_folder, f"{output_path}.pdf")
                # Open the first PDF
                with open(pdf1_path, 'rb') as pdf1_file:
                    pdf1_reader = PyPDF2.PdfReader(pdf1_file)
                    pdf1_pages = pdf1_reader.pages

                    # Open the second PDF
                    with open(pdf2_path, 'rb') as pdf2_file:
                        pdf2_reader = PyPDF2.PdfReader(pdf2_file)

                        # Create a new PDF writer
                        pdf_writer = PyPDF2.PdfWriter()

                        # Add pages from the first PDF
                        for page in pdf1_pages:
                            pdf_writer.add_page(page)

                        # Add pages from the second PDF
                        for page in pdf2_reader.pages:
                            pdf_writer.add_page(page)

                        # Write the merged PDF to the output file
                        with open(output_path, 'wb') as output_file:
                            pdf_writer.write(output_file)
                            print(f"Delivery note PDF created successfully at: {output_path}")
                        

            merge_pdfs(file_path_1, file_path_2, filename)
            os.remove(file_path_1)
            os.remove(file_path_2)

        else:
            pass

    def merge_3_file_order_by(self, po_filename, inv_filename, dn_filename, output_filename):
        po_file_path = os.path.join(self.path_po_folder, f"{po_filename}.pdf")
        inv_file_path = os.path.join(self.path_inv_folder, f"{inv_filename}.pdf")
        dn_file_path = os.path.join(self.path_dn_folder, f"{dn_filename}.pdf")
        all_file_path = os.path.join(self.path_all_folder, f"{output_filename}.pdf")

        copy_page = []

        with open(inv_file_path, 'rb') as pdf1_file:
            pdf1_reader = PyPDF2.PdfReader(pdf1_file)
            pdf1_pages = pdf1_reader.pages

            # Open the second PDF
            with open(dn_file_path, 'rb') as pdf2_file:
                pdf2_reader = PyPDF2.PdfReader(pdf2_file)
                pdf2_pages = pdf2_reader.pages

                # Open the third PDF
                with open(po_file_path, 'rb') as pdf3_file:
                    pdf3_reader = PyPDF2.PdfReader(pdf3_file)
                    pdf3_pages = pdf3_reader.pages

                    # Create a new PDF writer
                    pdf_writer = PyPDF2.PdfWriter()

                    if not len(pdf1_pages) <= 2:
                        real_inv_page = len(pdf1_pages) // 2

                        running_inv = 1
                        for page in pdf1_pages:
                            if running_inv <= real_inv_page:
                                pdf_writer.add_page(page)
                            else:
                                copy_page.append(page)
                            running_inv += 1
                    else:
                        pdf_writer.add_page(pdf1_pages[0])
                        copy_page.append(pdf1_pages[1])

                    for page in pdf2_pages:
                        pdf_writer.add_page(page)
                        copy_page.append(page)

                    for page in pdf3_pages:
                        pdf_writer.add_page(page)
                        copy_page.append(page)

                    for page_copy in copy_page:
                        pdf_writer.add_page(page_copy)



                    # Write the merged PDF to the output file
                    with open(all_file_path, 'wb') as output_file:
                        pdf_writer.write(output_file)

    def create_sum_invoice(self, file_name, main_data, sub_data):

        def one_page_invoice(filename, main, sub):
            file_path = os.path.join(self.path_sum_invoice, f"{filename}_inv.pdf")
            doc_page = SimpleDocTemplate(file_path, pagesize=A4)
            styles_page = getSampleStyleSheet()
            normal_style_page = styles_page["Normal"]
            content_page = []

            text_page = "This is a sample text for the delivery note."
            paragraph_page = Paragraph(text_page, normal_style_page)
            content_page.append(paragraph_page)
            doc_page.build(content_page)

            #main_data = [company, company_address, main_office, main_address, tax, no_inv, date_inv, total_price, vat_price, total_amount]
            date_text = main[6]

            numeric_value = float(re.sub(r'[^\d.]', '', main[9]))
            baht_text = bahttext(numeric_value)


            def add_newlines_to_address(address):
                # Add newline characters to the address string
                modified_address = address.replace(",", ",\n")
                return modified_address
            
            n = 0.1
            pdfmetrics.registerFont(TTFont('ThaiFont', 'image/THSarabunNew Bold.ttf'))

            c = canvas.Canvas(file_path)
            logo_path = "image/logo_daikatte.png"
            c.drawImage(logo_path, inch, A4[1] - 1.8*inch- n*inch, width=1.43*inch, height=1.25*inch)

            sign_path = "image\sign_praew.png"
            c.drawImage(sign_path, 1.3*inch, 0.6*inch, width=1.12*inch, height=0.80*inch)

            c.drawString(2.74*inch, A4[1] - 0.67*inch - n*inch, "New Line Production Co., Ltd. (Head Office)")
            c.setFont("Helvetica", 10)
            c.drawString(2.74*inch, A4[1] - 0.98*inch - n*inch, "425 Moo 2, Soi Sukhumvit Rd. 13, Pak Nam, Mueang, Samut Prakan, 10270.")
            c.drawString(2.74*inch, A4[1] - 1.26*inch - n*inch, "TEL : 063-1684278 (Praew) , 081-8350860 (Pair)    E-mail : daikatte@gmail.com")
            c.drawString(2.74*inch, A4[1] - 1.61*inch - n*inch, "TAX : 0115560000981")
            c.setFont("ThaiFont", 14)
            c.drawString(2.5*inch,  2.2*inch , "Note : ใบวางบิลชำระแบบเครดิต 30 วัน นับตั้งแต่วันที่จังส่งใบวางบิล")
            c.drawString(1.8*inch,  2.0*inch , "Term Of Payment : Transfer to account name : New Line Production Co., Ltd.( บจก.นิวไลน์โปรดักชั่น )")
            c.drawString(2.90*inch,  1.8*inch , "Swift code: KASITHBK , NO. 066-3-51404-0, KasikornBank")
            c.drawString(1.61*inch,  1.5*inch , "LSSUED BY")
            c.drawString(1.3*inch,  0.5*inch , f"Date: {date_text}")
            c.drawString(3.83*inch,  1.5*inch , "DELIVERED BY")
            c.drawString(3.44*inch,  0.5*inch , "Date:  ............................... ")
            c.drawString(6.19*inch,  1.5*inch , "RECEIVED BY")
            c.drawString(5.82*inch,  0.5*inch , "Date:  ............................... ")
            c.drawString(7.5*inch,  0.2*inch , "Page 1/1")

            c.setFont("Helvetica", 16)
            c.drawString(3.56*inch, A4[1] - 2.10*inch - n*inch, "Invoice")

            c.setFont("Helvetica", 8)

            Number = []
            delivery_note = []
            description = []
            qty = []
            unit = []
            amount = []
            number = 1
            for index, row in sub.iterrows():
                Number.append(number)
                delivery_note.append(row.get("no_dn", ""))
                description.append("Nama Pudding & Cheese cake can")
                qty.append(row.get("qty", ""))
                unit.append("Bottle")
                amount.append(row.get("total_price_y", ""))
                number = number + 1

            max_number = len(sub)
            if not max_number == 30 :
                for _ in range(30):
                    Number.append("")
                    delivery_note.append("")
                    description.append("")
                    qty.append("")
                    unit.append("")
                    amount.append("")

            IV_name = "No. : " + main[5]
            date_text_ = "Date : " + main[6]
            data_list_page_1 = [
                ['Company:', '', main[0],'','Original','' ],
                ['Address:', '', add_newlines_to_address(main[1]),'',IV_name,'' ],
                ['', '', '','',date_text_,'' ],
                ['Delivery Address :', '', main[2],'','Payment terms\n\nCredit : 30 Day','' ],
                ['', '', add_newlines_to_address(main[3]),'','','' ],
                ['Tax :', '', main[4],'','','' ],
                ['No.', 'Delivery Note', 'Description','quantity','Unit','Amount' ],
                [Number[0], delivery_note[0], description[0],qty[0],unit[0],amount[0] ],
                [Number[1], delivery_note[1], description[1],qty[1],unit[1],amount[1] ],
                [Number[2], delivery_note[2], description[2],qty[2],unit[2],amount[2] ],
                [Number[3], delivery_note[3], description[3],qty[3],unit[3],amount[3] ],
                [Number[4], delivery_note[4], description[4],qty[4],unit[4],amount[4] ],
                [Number[5], delivery_note[5], description[5],qty[5],unit[5],amount[5] ],
                [Number[6], delivery_note[6], description[6],qty[6],unit[6],amount[6] ],
                [Number[7], delivery_note[7], description[7],qty[7],unit[7],amount[7] ],
                [Number[8], delivery_note[8], description[8],qty[8],unit[8],amount[8] ],
                [Number[9], delivery_note[9], description[9],qty[9],unit[9],amount[9] ],
                [Number[10], delivery_note[10], description[10],qty[10],unit[10],amount[10] ],
                [Number[11], delivery_note[11], description[11],qty[11],unit[11],amount[11] ],
                [Number[12], delivery_note[12], description[12],qty[12],unit[12],amount[12] ],
                [Number[13], delivery_note[13], description[13],qty[13],unit[13],amount[13] ],
                [Number[14], delivery_note[14], description[14],qty[14],unit[14],amount[14] ],
                [baht_text, '', '','',"Sub Total", main[7] ],
                ['', '', '','',"Vat 7%", main[8] ],
                ['', '', '','', "Total Amount", main[9] ],


            ]
            item_heigh = 0.25
            row_heights = [0.3*inch, 0.35*inch, 0.35*inch, 0.24*inch, 0.5*inch
                           , 0.3*inch, 0.35*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch]

            # Create table and set column widths and row heights
            table = Table(data_list_page_1, colWidths=[0.5*inch, 1.35*inch, 2.5*inch, 0.7*inch, 0.7*inch, 1*inch], rowHeights=row_heights)
            table_style = [
                ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Add border to the table
                #('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Bold font for headers
                #('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),  # Background color for headers
                ('SPAN', (0, 0), (1, 0)),#Company
                ('SPAN', (2, 0), (3, 0)),#Company ans
                ('SPAN', (4, 0), (5, 0)),#original
                ('SPAN', (0, 1), (1, 2)),#addres
                ('SPAN', (2, 1), (3, 2)),#addres ans
                ('SPAN', (4, 1), (5, 1)),#No
                ('SPAN', (4, 2), (5, 2)),#Date
                ('SPAN', (0, 3), (1, 4)),#Delivery ad
                ('SPAN', (2, 3), (3, 3)),#main office
                ('SPAN', (4, 3), (5, 5)),#Payment
                ('SPAN', (2, 4), (3, 4)),#Delivery ad ans
                ('SPAN', (0, 5), (1, 5)), #Tax
                ('SPAN', (2, 5), (3, 5)),#Tax ans
                ('SPAN', (0, 22), (3, 24)), #thai baht

                ('ALIGN', (0, 0), (3, 5), 'LEFT'),
                ('VALIGN', (0, 0), (3, 5), 'TOP'),
                
                ('VALIGN', (4, 0), (5, 5), 'MIDDLE'),#list item code5
                ('ALIGN', (4, 0), (5, 2), 'CENTER'),
                ('ALIGN', (2, 5), (3, 5), 'LEFT'),

                ('FONTSIZE', (2, 4), (3, 4), 8),
                ('FONTNAME', (5, 7), (5, 29), 'ThaiFont'),  # Thai font
                ('FONTSIZE', (5, 7), (5, 29), 12),
                ('FONTSIZE', (0, 6), (4, 29), 8),
                ('VALIGN', (0, 6), (5, 29), 'MIDDLE'),#list item code5
                ('ALIGN', (0, 6), (5, 29), 'CENTER'),
                ('FONTNAME', (0, 22), (3, 24), 'ThaiFont'),
                ('VALIGN', (0, 22), (3, 24), 'MIDDLE'),#list item code5
                ('ALIGN', (0, 22), (3, 24), 'CENTER'),
                ('FONTSIZE', (0, 22), (3, 24), 14),
            ]

            table.setStyle(TableStyle(table_style))

            table.wrapOn(c, 6.81*inch, 5*inch)
            table.drawOn(c, 0.75*inch, 2.5*inch)

            c.save()

            return file_path

        def one_page_receipt(filename, main, sub):
            file_path = os.path.join(self.path_sum_invoice, f"{filename}_receipt.pdf")
            doc_page = SimpleDocTemplate(file_path, pagesize=A4)
            styles_page = getSampleStyleSheet()
            normal_style_page = styles_page["Normal"]
            content_page = []

            text_page = "This is a sample text for the delivery note."
            paragraph_page = Paragraph(text_page, normal_style_page)
            content_page.append(paragraph_page)
            doc_page.build(content_page)

            #main_data = [company, company_address, main_office, main_address, tax, no_inv, date_inv, total_price, vat_price, total_amount]
            date_text = main[6]

            numeric_value = float(re.sub(r'[^\d.]', '', main[9]))
            baht_text = bahttext(numeric_value)


            def add_newlines_to_address(address):
                # Add newline characters to the address string
                modified_address = address.replace(",", ",\n")
                return modified_address
            
            n = 0.1
            pdfmetrics.registerFont(TTFont('ThaiFont', 'image/THSarabunNew Bold.ttf'))

            c = canvas.Canvas(file_path)
            logo_path = "image/logo_daikatte.png"
            c.drawImage(logo_path, inch, A4[1] - 1.8*inch- n*inch, width=1.43*inch, height=1.25*inch)

            sign_path = "image\sign_praew.png"
            c.drawImage(sign_path, 1.3*inch, 0.6*inch, width=1.12*inch, height=0.80*inch)

            c.drawString(2.74*inch, A4[1] - 0.67*inch - n*inch, "New Line Production Co., Ltd. (Head Office)")
            c.setFont("Helvetica", 10)
            c.drawString(2.74*inch, A4[1] - 0.98*inch - n*inch, "425 Moo 2, Soi Sukhumvit Rd. 13, Pak Nam, Mueang, Samut Prakan, 10270.")
            c.drawString(2.74*inch, A4[1] - 1.26*inch - n*inch, "TEL : 063-1684278 (Praew) , 081-8350860 (Pair)    E-mail : daikatte@gmail.com")
            c.drawString(2.74*inch, A4[1] - 1.61*inch - n*inch, "TAX : 0115560000981")
            c.setFont("ThaiFont", 14)
            c.drawString(2.5*inch,  2.2*inch , "Note : ใบวางบิลชำระแบบเครดิต 30 วัน นับตั้งแต่วันที่จังส่งใบวางบิล")
            c.drawString(1.8*inch,  2.0*inch , "Term Of Payment : Transfer to account name : New Line Production Co., Ltd.( บจก.นิวไลน์โปรดักชั่น )")
            c.drawString(2.90*inch,  1.8*inch , "Swift code: KASITHBK , NO. 066-3-51404-0, KasikornBank")
            c.drawString(1.61*inch,  1.5*inch , "LSSUED BY")
            c.drawString(1.3*inch,  0.5*inch , f"Date: {date_text}")
            c.drawString(3.83*inch,  1.5*inch , "DELIVERED BY")
            c.drawString(3.44*inch,  0.5*inch , "Date:  ............................... ")
            c.drawString(6.19*inch,  1.5*inch , "RECEIVED BY")
            c.drawString(5.82*inch,  0.5*inch , "Date:  ............................... ")
            c.drawString(7.5*inch,  0.2*inch , "Page 1/1")

            c.setFont("Helvetica", 16)
            c.drawString(3.56*inch, A4[1] - 2.10*inch - n*inch, "Receipt")

            c.setFont("Helvetica", 8)

            Number = []
            delivery_note = []
            description = []
            qty = []
            unit = []
            amount = []
            number = 1
            for index, row in sub.iterrows():
                Number.append(number)
                delivery_note.append(row.get("no_dn", ""))
                description.append("Nama Pudding & Cheese cake can")
                qty.append(row.get("qty", ""))
                unit.append("Bottle")
                amount.append(row.get("total_price_y", ""))
                number = number + 1

            max_number = len(sub)
            if not max_number == 30 :
                for _ in range(30):
                    Number.append("")
                    delivery_note.append("")
                    description.append("")
                    qty.append("")
                    unit.append("")
                    amount.append("")

            IV_name = "No. : " + main[5]
            date_text_ = "Date : " + main[6]
            data_list_page_1 = [
                ['Company:', '', main[0],'','Original','' ],
                ['Address:', '', add_newlines_to_address(main[1]),'',IV_name,'' ],
                ['', '', '','',date_text_,'' ],
                ['Delivery Address :', '', main[2],'','Payment terms\n\nCredit : 30 Day','' ],
                ['', '', add_newlines_to_address(main[3]),'','','' ],
                ['Tax :', '', main[4],'','','' ],
                ['No.', 'Delivery Note', 'Description','quantity','Unit','Amount' ],
                [Number[0], delivery_note[0], description[0],qty[0],unit[0],amount[0] ],
                [Number[1], delivery_note[1], description[1],qty[1],unit[1],amount[1] ],
                [Number[2], delivery_note[2], description[2],qty[2],unit[2],amount[2] ],
                [Number[3], delivery_note[3], description[3],qty[3],unit[3],amount[3] ],
                [Number[4], delivery_note[4], description[4],qty[4],unit[4],amount[4] ],
                [Number[5], delivery_note[5], description[5],qty[5],unit[5],amount[5] ],
                [Number[6], delivery_note[6], description[6],qty[6],unit[6],amount[6] ],
                [Number[7], delivery_note[7], description[7],qty[7],unit[7],amount[7] ],
                [Number[8], delivery_note[8], description[8],qty[8],unit[8],amount[8] ],
                [Number[9], delivery_note[9], description[9],qty[9],unit[9],amount[9] ],
                [Number[10], delivery_note[10], description[10],qty[10],unit[10],amount[10] ],
                [Number[11], delivery_note[11], description[11],qty[11],unit[11],amount[11] ],
                [Number[12], delivery_note[12], description[12],qty[12],unit[12],amount[12] ],
                [Number[13], delivery_note[13], description[13],qty[13],unit[13],amount[13] ],
                [Number[14], delivery_note[14], description[14],qty[14],unit[14],amount[14] ],
                [baht_text, '', '','',"Sub Total", main[7] ],
                ['', '', '','',"Vat 7%", main[8] ],
                ['', '', '','', "Total Amount", main[9] ],


            ]
            item_heigh = 0.25
            row_heights = [0.3*inch, 0.35*inch, 0.35*inch, 0.24*inch, 0.5*inch
                           , 0.3*inch, 0.35*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch]

            # Create table and set column widths and row heights
            table = Table(data_list_page_1, colWidths=[0.5*inch, 1.35*inch, 2.5*inch, 0.7*inch, 0.7*inch, 1*inch], rowHeights=row_heights)
            table_style = [
                ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Add border to the table
                #('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Bold font for headers
                #('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),  # Background color for headers
                ('SPAN', (0, 0), (1, 0)),#Company
                ('SPAN', (2, 0), (3, 0)),#Company ans
                ('SPAN', (4, 0), (5, 0)),#original
                ('SPAN', (0, 1), (1, 2)),#addres
                ('SPAN', (2, 1), (3, 2)),#addres ans
                ('SPAN', (4, 1), (5, 1)),#No
                ('SPAN', (4, 2), (5, 2)),#Date
                ('SPAN', (0, 3), (1, 4)),#Delivery ad
                ('SPAN', (2, 3), (3, 3)),#main office
                ('SPAN', (4, 3), (5, 5)),#Payment
                ('SPAN', (2, 4), (3, 4)),#Delivery ad ans
                ('SPAN', (0, 5), (1, 5)), #Tax
                ('SPAN', (2, 5), (3, 5)),#Tax ans
                ('SPAN', (0, 22), (3, 24)), #thai baht

                ('ALIGN', (0, 0), (3, 5), 'LEFT'),
                ('VALIGN', (0, 0), (3, 5), 'TOP'),
                
                ('VALIGN', (4, 0), (5, 5), 'MIDDLE'),#list item code5
                ('ALIGN', (4, 0), (5, 2), 'CENTER'),
                ('ALIGN', (2, 5), (3, 5), 'LEFT'),

                ('FONTSIZE', (2, 4), (3, 4), 8),
                ('FONTNAME', (5, 7), (5, 29), 'ThaiFont'),  # Thai font
                ('FONTSIZE', (5, 7), (5, 29), 12),
                ('FONTSIZE', (0, 6), (4, 29), 8),
                ('VALIGN', (0, 6), (5, 29), 'MIDDLE'),#list item code5
                ('ALIGN', (0, 6), (5, 29), 'CENTER'),
                ('FONTNAME', (0, 22), (3, 24), 'ThaiFont'),
                ('VALIGN', (0, 22), (3, 24), 'MIDDLE'),#list item code5
                ('ALIGN', (0, 22), (3, 24), 'CENTER'),
                ('FONTSIZE', (0, 22), (3, 24), 14),
            ]

            table.setStyle(TableStyle(table_style))

            table.wrapOn(c, 6.81*inch, 5*inch)
            table.drawOn(c, 0.75*inch, 2.5*inch)

            c.save()
            return file_path

        def two_page_1_invoice(filename, main, sub):

            file_path_1 = os.path.join(self.path_sum_invoice, f"{filename}_1_invoice.pdf")
            doc_page_1 = SimpleDocTemplate(file_path_1, pagesize=A4)
            styles_page_1 = getSampleStyleSheet()
            normal_style_page_1 = styles_page_1["Normal"]
            content_page_1 = []

            text_page_1 = "This is a sample text for the delivery note."
            paragraph_page_1 = Paragraph(text_page_1, normal_style_page_1)
            content_page_1.append(paragraph_page_1)
            doc_page_1.build(content_page_1)

            #main_data = [company, company_address, main_office, main_address, tax, no_inv, date_inv, total_price, vat_price, total_amount]
            date_text = main[6]

            def add_newlines_to_address(address):
                # Add newline characters to the address string
                modified_address = address.replace(",", ",\n")
                return modified_address
            
            n = 0.1
            pdfmetrics.registerFont(TTFont('ThaiFont', 'image/THSarabunNew Bold.ttf'))

            c = canvas.Canvas(file_path_1)
            logo_path = "image/logo_daikatte.png"
            c.drawImage(logo_path, inch, A4[1] - 1.8*inch- n*inch, width=1.43*inch, height=1.25*inch)

            sign_path = "image\sign_praew.png"
            c.drawImage(sign_path, 1.3*inch, 0.9*inch, width=1.12*inch, height=0.80*inch)

            c.drawString(2.74*inch, A4[1] - 0.67*inch - n*inch, "New Line Production Co., Ltd. (Head Office)")
            c.setFont("Helvetica", 10)
            c.drawString(2.74*inch, A4[1] - 0.98*inch - n*inch, "425 Moo 2, Soi Sukhumvit Rd. 13, Pak Nam, Mueang, Samut Prakan, 10270.")
            c.drawString(2.74*inch, A4[1] - 1.26*inch - n*inch, "TEL : 063-1684278 (Praew) , 081-8350860 (Pair)    E-mail : daikatte@gmail.com")
            c.drawString(2.74*inch, A4[1] - 1.61*inch - n*inch, "TAX : 0115560000981")
            c.setFont("ThaiFont", 14)
            c.drawString(2.5*inch,  2.8*inch , "Note : ใบวางบิลชำระแบบเครดิต 30 วัน นับตั้งแต่วันที่จังส่งใบวางบิล")
            c.drawString(1.8*inch,  2.5*inch , "Term Of Payment : Transfer to account name : New Line Production Co., Ltd.( บจก.นิวไลน์โปรดักชั่น )")
            c.drawString(2.90*inch,  2.3*inch , "Swift code: KASITHBK , NO. 066-3-51404-0, KasikornBank")
            c.drawString(1.61*inch,  1.8*inch , "LSSUED BY")
            c.drawString(1.3*inch,  0.8*inch , f"Date: {date_text}")
            c.drawString(3.83*inch,  1.8*inch , "DELIVERED BY")
            c.drawString(3.44*inch,  0.8*inch , "Date:  ............................... ")
            c.drawString(6.19*inch,  1.8*inch , "RECEIVED BY")
            c.drawString(5.82*inch,  0.8*inch , "Date:  ............................... ")
            c.drawString(7.5*inch,  0.2*inch , "Page 1/2")

            c.setFont("Helvetica", 16)
            c.drawString(3.56*inch, A4[1] - 2.10*inch - n*inch, "Invoice")

            c.setFont("Helvetica", 8)

            Number = []
            delivery_note = []
            description = []
            qty = []
            unit = []
            amount = []
            number = 1
            for index, row in sub.iterrows():
                Number.append(number)
                delivery_note.append(row.get("no_dn", ""))
                description.append("Nama Pudding & Cheese cake can")
                qty.append(row.get("qty", ""))
                unit.append("Bottle")
                amount.append(row.get("total_price_y", ""))
                number = number + 1

            max_number = len(sub)
            if not max_number == 15 :
                for _ in range(10 - max_number):
                    Number.append("")
                    delivery_note.append("")
                    description.append("")
                    qty.append("")
                    unit.append("")
                    amount.append("")

            IV_name = "No. : " + main[5]
            date_text_ = "Date : " + main[6]
            data_list_page_1 = [
                ['Company:', '', main[0],'','Original','' ],
                ['Address:', '', add_newlines_to_address(main[1]),'',IV_name,'' ],
                ['', '', '','',date_text_,'' ],
                ['Delivery Address :', '', main[2],'','Payment terms\n\nCredit : 30 Day','' ],
                ['', '', add_newlines_to_address(main[3]),'','','' ],
                ['Tax :', '', main[4],'','','' ],
                ['No.', 'Delivery Note', 'Description','quantity','Unit','Amount' ],
                [Number[0], delivery_note[0], description[0],qty[0],unit[0],amount[0] ],
                [Number[1], delivery_note[1], description[1],qty[1],unit[1],amount[1] ],
                [Number[2], delivery_note[2], description[2],qty[2],unit[2],amount[2] ],
                [Number[3], delivery_note[3], description[3],qty[3],unit[3],amount[3] ],
                [Number[4], delivery_note[4], description[4],qty[4],unit[4],amount[4] ],
                [Number[5], delivery_note[5], description[5],qty[5],unit[5],amount[5] ],
                [Number[6], delivery_note[6], description[6],qty[6],unit[6],amount[6] ],
                [Number[7], delivery_note[7], description[7],qty[7],unit[7],amount[7] ],
                [Number[8], delivery_note[8], description[8],qty[8],unit[8],amount[8] ],
                [Number[9], delivery_note[9], description[9],qty[9],unit[9],amount[9] ],
                [Number[10], delivery_note[10], description[10],qty[10],unit[10],amount[10] ],
                [Number[11], delivery_note[11], description[11],qty[11],unit[11],amount[11] ],
                [Number[12], delivery_note[12], description[12],qty[12],unit[12],amount[12] ],
                [Number[13], delivery_note[13], description[13],qty[13],unit[13],amount[13] ],
                [Number[14], delivery_note[14], description[14],qty[14],unit[14],amount[14] ]


            ]
            item_heigh = 0.25
            row_heights = [0.3*inch, 0.35*inch, 0.35*inch, 0.24*inch, 0.5*inch
                           , 0.3*inch, 0.35*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch]

            # Create table and set column widths and row heights
            table = Table(data_list_page_1, colWidths=[0.5*inch, 1.35*inch, 2.5*inch, 0.7*inch, 0.7*inch, 1*inch], rowHeights=row_heights)
            table_style = [
                ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Add border to the table
                #('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Bold font for headers
                #('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),  # Background color for headers
                ('SPAN', (0, 0), (1, 0)),#Company
                ('SPAN', (2, 0), (3, 0)),#Company ans
                ('SPAN', (4, 0), (5, 0)),#original
                ('SPAN', (0, 1), (1, 2)),#addres
                ('SPAN', (2, 1), (3, 2)),#addres ans
                ('SPAN', (4, 1), (5, 1)),#No
                ('SPAN', (4, 2), (5, 2)),#Date
                ('SPAN', (0, 3), (1, 4)),#Delivery ad
                ('SPAN', (2, 3), (3, 3)),#main office
                ('SPAN', (4, 3), (5, 5)),#Payment
                ('SPAN', (2, 4), (3, 4)),#Delivery ad ans
                ('SPAN', (0, 5), (1, 5)), #Tax
                ('SPAN', (2, 5), (3, 5)),#Tax ans

                ('ALIGN', (0, 0), (3, 5), 'LEFT'),
                ('VALIGN', (0, 0), (3, 5), 'TOP'),
                
                ('VALIGN', (4, 0), (5, 5), 'MIDDLE'),#list item code5
                ('ALIGN', (4, 0), (5, 2), 'CENTER'),
                ('ALIGN', (2, 5), (3, 5), 'LEFT'),

                ('FONTSIZE', (2, 4), (3, 4), 8),
                ('FONTNAME', (5, 7), (5, 29), 'ThaiFont'),  # Thai font
                ('FONTSIZE', (5, 7), (5, 29), 12),
                ('FONTSIZE', (0, 6), (4, 29), 8),
                ('VALIGN', (0, 6), (5, 29), 'MIDDLE'),#list item code5
                ('ALIGN', (0, 6), (5, 29), 'CENTER'),


            ]

            table.setStyle(TableStyle(table_style))

            table.wrapOn(c, 6.81*inch, 5*inch)
            table.drawOn(c, 0.75*inch, 3.2*inch)

            c.save()
            return file_path_1

        def two_page_2_invoice(filename, main, sub):
            file_path_2 = os.path.join(self.path_sum_invoice, f"{filename}_2_invoice.pdf")
            doc_page_2 = SimpleDocTemplate(file_path_2, pagesize=A4)
            styles_page_2 = getSampleStyleSheet()
            normal_style_page_2 = styles_page_2["Normal"]
            content_page_2 = []

            text_page_2 = "This is a sample text for the delivery note."
            paragraph_page_2 = Paragraph(text_page_2, normal_style_page_2)
            content_page_2.append(paragraph_page_2)
            doc_page_2.build(content_page_2)

            #main_data = [company, company_address, main_office, main_address, tax, no_inv, date_inv, total_price, vat_price, total_amount]
            date_text = main[6]

            numeric_value = float(re.sub(r'[^\d.]', '', main[9]))
            baht_text = bahttext(numeric_value)


            def add_newlines_to_address(address):
                # Add newline characters to the address string
                modified_address = address.replace(",", ",\n")
                return modified_address
            
            n = 0.1
            pdfmetrics.registerFont(TTFont('ThaiFont', 'image/THSarabunNew Bold.ttf'))

            c = canvas.Canvas(file_path_2)
            logo_path = "image/logo_daikatte.png"
            c.drawImage(logo_path, inch, A4[1] - 1.8*inch- n*inch, width=1.43*inch, height=1.25*inch)

            sign_path = "image\sign_praew.png"
            c.drawImage(sign_path, 1.3*inch, 0.6*inch, width=1.12*inch, height=0.80*inch)

            c.drawString(2.74*inch, A4[1] - 0.67*inch - n*inch, "New Line Production Co., Ltd. (Head Office)")
            c.setFont("Helvetica", 10)
            c.drawString(2.74*inch, A4[1] - 0.98*inch - n*inch, "425 Moo 2, Soi Sukhumvit Rd. 13, Pak Nam, Mueang, Samut Prakan, 10270.")
            c.drawString(2.74*inch, A4[1] - 1.26*inch - n*inch, "TEL : 063-1684278 (Praew) , 081-8350860 (Pair)    E-mail : daikatte@gmail.com")
            c.drawString(2.74*inch, A4[1] - 1.61*inch - n*inch, "TAX : 0115560000981")
            c.setFont("ThaiFont", 14)
            c.drawString(2.5*inch,  2.2*inch , "Note : ใบวางบิลชำระแบบเครดิต 30 วัน นับตั้งแต่วันที่จังส่งใบวางบิล")
            c.drawString(1.8*inch,  2.0*inch , "Term Of Payment : Transfer to account name : New Line Production Co., Ltd.( บจก.นิวไลน์โปรดักชั่น )")
            c.drawString(2.90*inch,  1.8*inch , "Swift code: KASITHBK , NO. 066-3-51404-0, KasikornBank")
            c.drawString(1.61*inch,  1.5*inch , "LSSUED BY")
            c.drawString(1.3*inch,  0.5*inch , f"Date: {date_text}")
            c.drawString(3.83*inch,  1.5*inch , "DELIVERED BY")
            c.drawString(3.44*inch,  0.5*inch , "Date:  ............................... ")
            c.drawString(6.19*inch,  1.5*inch , "RECEIVED BY")
            c.drawString(5.82*inch,  0.5*inch , "Date:  ............................... ")
            c.drawString(7.5*inch,  0.2*inch , "Page 2/2")

            c.setFont("Helvetica", 16)
            c.drawString(3.56*inch, A4[1] - 2.10*inch - n*inch, "Invoice")

            c.setFont("Helvetica", 8)

            Number = []
            delivery_note = []
            description = []
            qty = []
            unit = []
            amount = []
            number = 1
            for index, row in sub.iterrows():
                Number.append(number)
                delivery_note.append(row.get("no_dn", ""))
                description.append("Nama Pudding & Cheese cake can")
                qty.append(row.get("qty", ""))
                unit.append("Bottle")
                amount.append(row.get("total_price_y", ""))
                number = number + 1

            max_number = len(sub)
            if not max_number == 30 :
                for _ in range(30):
                    Number.append("")
                    delivery_note.append("")
                    description.append("")
                    qty.append("")
                    unit.append("")
                    amount.append("")

            IV_name = "No. : " + main[5]
            date_text_ = "Date : " + main[6]
            data_list_page_1 = [
                ['Company:', '', main[0],'','Original','' ],
                ['Address:', '', add_newlines_to_address(main[1]),'',IV_name,'' ],
                ['', '', '','',date_text_,'' ],
                ['Delivery Address :', '', main[2],'','Payment terms\n\nCredit : 30 Day','' ],
                ['', '', add_newlines_to_address(main[3]),'','','' ],
                ['Tax :', '', main[4],'','','' ],
                ['No.', 'Delivery Note', 'Description','quantity','Unit','Amount' ],
                [Number[15], delivery_note[15], description[15],qty[15],unit[15],amount[15] ],
                [Number[16], delivery_note[16], description[16],qty[16],unit[16],amount[16] ],
                [Number[17], delivery_note[17], description[17],qty[17],unit[17],amount[17] ],
                [Number[18], delivery_note[18], description[18],qty[18],unit[18],amount[18] ],
                [Number[19], delivery_note[19], description[19],qty[19],unit[19],amount[19] ],
                [Number[20], delivery_note[20], description[20],qty[20],unit[20],amount[20] ],
                [Number[21], delivery_note[21], description[21],qty[21],unit[21],amount[21] ],
                [Number[22], delivery_note[22], description[22],qty[22],unit[22],amount[22] ],
                [Number[23], delivery_note[23], description[23],qty[23],unit[23],amount[23] ],
                [Number[24], delivery_note[24], description[24],qty[24],unit[24],amount[24] ],
                [Number[25], delivery_note[25], description[25],qty[25],unit[25],amount[25] ],
                [Number[26], delivery_note[26], description[26],qty[26],unit[26],amount[26] ],
                [Number[27], delivery_note[27], description[27],qty[27],unit[27],amount[27] ],
                [Number[28], delivery_note[28], description[28],qty[28],unit[28],amount[28] ],
                [Number[29], delivery_note[29], description[29],qty[29],unit[29],amount[29] ],
                [baht_text, '', '','',"Sub Total", main[7] ],
                ['', '', '','',"Vat 7%", main[8] ],
                ['', '', '','', "Total Amount", main[9] ],


            ]
            item_heigh = 0.25
            row_heights = [0.3*inch, 0.35*inch, 0.35*inch, 0.24*inch, 0.5*inch
                           , 0.3*inch, 0.35*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch]

            # Create table and set column widths and row heights
            table = Table(data_list_page_1, colWidths=[0.5*inch, 1.35*inch, 2.5*inch, 0.7*inch, 0.7*inch, 1*inch], rowHeights=row_heights)
            table_style = [
                ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Add border to the table
                #('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Bold font for headers
                #('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),  # Background color for headers
                ('SPAN', (0, 0), (1, 0)),#Company
                ('SPAN', (2, 0), (3, 0)),#Company ans
                ('SPAN', (4, 0), (5, 0)),#original
                ('SPAN', (0, 1), (1, 2)),#addres
                ('SPAN', (2, 1), (3, 2)),#addres ans
                ('SPAN', (4, 1), (5, 1)),#No
                ('SPAN', (4, 2), (5, 2)),#Date
                ('SPAN', (0, 3), (1, 4)),#Delivery ad
                ('SPAN', (2, 3), (3, 3)),#main office
                ('SPAN', (4, 3), (5, 5)),#Payment
                ('SPAN', (2, 4), (3, 4)),#Delivery ad ans
                ('SPAN', (0, 5), (1, 5)), #Tax
                ('SPAN', (2, 5), (3, 5)),#Tax ans
                ('SPAN', (0, 22), (3, 24)), #thai baht

                ('ALIGN', (0, 0), (3, 5), 'LEFT'),
                ('VALIGN', (0, 0), (3, 5), 'TOP'),
                
                ('VALIGN', (4, 0), (5, 5), 'MIDDLE'),#list item code5
                ('ALIGN', (4, 0), (5, 2), 'CENTER'),
                ('ALIGN', (2, 5), (3, 5), 'LEFT'),

                ('FONTSIZE', (2, 4), (3, 4), 8),
                ('FONTNAME', (5, 7), (5, 29), 'ThaiFont'),  # Thai font
                ('FONTSIZE', (5, 7), (5, 29), 12),
                ('FONTSIZE', (0, 6), (4, 29), 8),
                ('VALIGN', (0, 6), (5, 29), 'MIDDLE'),#list item code5
                ('ALIGN', (0, 6), (5, 29), 'CENTER'),
                ('FONTNAME', (0, 22), (3, 24), 'ThaiFont'),
                ('VALIGN', (0, 22), (3, 24), 'MIDDLE'),#list item code5
                ('ALIGN', (0, 22), (3, 24), 'CENTER'),
                ('FONTSIZE', (0, 22), (3, 24), 14),
            ]

            table.setStyle(TableStyle(table_style))

            table.wrapOn(c, 6.81*inch, 5*inch)
            table.drawOn(c, 0.75*inch, 2.5*inch)

            c.save()
            return file_path_2
        
        def two_page_1_receipt(filename, main, sub):

            file_path_1 = os.path.join(self.path_sum_invoice, f"{filename}_1_receipt.pdf")
            doc_page_1 = SimpleDocTemplate(file_path_1, pagesize=A4)
            styles_page_1 = getSampleStyleSheet()
            normal_style_page_1 = styles_page_1["Normal"]
            content_page_1 = []

            text_page_1 = "This is a sample text for the delivery note."
            paragraph_page_1 = Paragraph(text_page_1, normal_style_page_1)
            content_page_1.append(paragraph_page_1)
            doc_page_1.build(content_page_1)

            #main_data = [company, company_address, main_office, main_address, tax, no_inv, date_inv, total_price, vat_price, total_amount]
            date_text = main[6]

            def add_newlines_to_address(address):
                # Add newline characters to the address string
                modified_address = address.replace(",", ",\n")
                return modified_address
            
            n = 0.1
            pdfmetrics.registerFont(TTFont('ThaiFont', 'image/THSarabunNew Bold.ttf'))

            c = canvas.Canvas(file_path_1)
            logo_path = "image/logo_daikatte.png"
            c.drawImage(logo_path, inch, A4[1] - 1.8*inch- n*inch, width=1.43*inch, height=1.25*inch)

            sign_path = "image\sign_praew.png"
            c.drawImage(sign_path, 1.3*inch, 0.9*inch, width=1.12*inch, height=0.80*inch)

            c.drawString(2.74*inch, A4[1] - 0.67*inch - n*inch, "New Line Production Co., Ltd. (Head Office)")
            c.setFont("Helvetica", 10)
            c.drawString(2.74*inch, A4[1] - 0.98*inch - n*inch, "425 Moo 2, Soi Sukhumvit Rd. 13, Pak Nam, Mueang, Samut Prakan, 10270.")
            c.drawString(2.74*inch, A4[1] - 1.26*inch - n*inch, "TEL : 063-1684278 (Praew) , 081-8350860 (Pair)    E-mail : daikatte@gmail.com")
            c.drawString(2.74*inch, A4[1] - 1.61*inch - n*inch, "TAX : 0115560000981")
            c.setFont("ThaiFont", 14)
            c.drawString(2.5*inch,  2.8*inch , "Note : ใบวางบิลชำระแบบเครดิต 30 วัน นับตั้งแต่วันที่จังส่งใบวางบิล")
            c.drawString(1.8*inch,  2.5*inch , "Term Of Payment : Transfer to account name : New Line Production Co., Ltd.( บจก.นิวไลน์โปรดักชั่น )")
            c.drawString(2.90*inch,  2.3*inch , "Swift code: KASITHBK , NO. 066-3-51404-0, KasikornBank")
            c.drawString(1.61*inch,  1.8*inch , "LSSUED BY")
            c.drawString(1.3*inch,  0.8*inch , f"Date: {date_text}")
            c.drawString(3.83*inch,  1.8*inch , "DELIVERED BY")
            c.drawString(3.44*inch,  0.8*inch , "Date:  ............................... ")
            c.drawString(6.19*inch,  1.8*inch , "RECEIVED BY")
            c.drawString(5.82*inch,  0.8*inch , "Date:  ............................... ")
            c.drawString(7.5*inch,  0.2*inch , "Page 1/2")

            c.setFont("Helvetica", 16)
            c.drawString(3.56*inch, A4[1] - 2.10*inch - n*inch, "Receipt")

            c.setFont("Helvetica", 8)

            Number = []
            delivery_note = []
            description = []
            qty = []
            unit = []
            amount = []
            number = 1
            for index, row in sub.iterrows():
                Number.append(number)
                delivery_note.append(row.get("no_dn", ""))
                description.append("Nama Pudding & Cheese cake can")
                qty.append(row.get("qty", ""))
                unit.append("Bottle")
                amount.append(row.get("total_price_y", ""))
                number = number + 1

            max_number = len(sub)
            if not max_number == 15 :
                for _ in range(10 - max_number):
                    Number.append("")
                    delivery_note.append("")
                    description.append("")
                    qty.append("")
                    unit.append("")
                    amount.append("")

            IV_name = "No. : " + main[5]
            date_text_ = "Date : " + main[6]
            data_list_page_1 = [
                ['Company:', '', main[0],'','Original','' ],
                ['Address:', '', add_newlines_to_address(main[1]),'',IV_name,'' ],
                ['', '', '','',date_text_,'' ],
                ['Delivery Address :', '', main[2],'','Payment terms\n\nCredit : 30 Day','' ],
                ['', '', add_newlines_to_address(main[3]),'','','' ],
                ['Tax :', '', main[4],'','','' ],
                ['No.', 'Delivery Note', 'Description','quantity','Unit','Amount' ],
                [Number[0], delivery_note[0], description[0],qty[0],unit[0],amount[0] ],
                [Number[1], delivery_note[1], description[1],qty[1],unit[1],amount[1] ],
                [Number[2], delivery_note[2], description[2],qty[2],unit[2],amount[2] ],
                [Number[3], delivery_note[3], description[3],qty[3],unit[3],amount[3] ],
                [Number[4], delivery_note[4], description[4],qty[4],unit[4],amount[4] ],
                [Number[5], delivery_note[5], description[5],qty[5],unit[5],amount[5] ],
                [Number[6], delivery_note[6], description[6],qty[6],unit[6],amount[6] ],
                [Number[7], delivery_note[7], description[7],qty[7],unit[7],amount[7] ],
                [Number[8], delivery_note[8], description[8],qty[8],unit[8],amount[8] ],
                [Number[9], delivery_note[9], description[9],qty[9],unit[9],amount[9] ],
                [Number[10], delivery_note[10], description[10],qty[10],unit[10],amount[10] ],
                [Number[11], delivery_note[11], description[11],qty[11],unit[11],amount[11] ],
                [Number[12], delivery_note[12], description[12],qty[12],unit[12],amount[12] ],
                [Number[13], delivery_note[13], description[13],qty[13],unit[13],amount[13] ],
                [Number[14], delivery_note[14], description[14],qty[14],unit[14],amount[14] ]


            ]
            item_heigh = 0.25
            row_heights = [0.3*inch, 0.35*inch, 0.35*inch, 0.24*inch, 0.5*inch
                           , 0.3*inch, 0.35*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch]

            # Create table and set column widths and row heights
            table = Table(data_list_page_1, colWidths=[0.5*inch, 1.35*inch, 2.5*inch, 0.7*inch, 0.7*inch, 1*inch], rowHeights=row_heights)
            table_style = [
                ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Add border to the table
                #('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Bold font for headers
                #('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),  # Background color for headers
                ('SPAN', (0, 0), (1, 0)),#Company
                ('SPAN', (2, 0), (3, 0)),#Company ans
                ('SPAN', (4, 0), (5, 0)),#original
                ('SPAN', (0, 1), (1, 2)),#addres
                ('SPAN', (2, 1), (3, 2)),#addres ans
                ('SPAN', (4, 1), (5, 1)),#No
                ('SPAN', (4, 2), (5, 2)),#Date
                ('SPAN', (0, 3), (1, 4)),#Delivery ad
                ('SPAN', (2, 3), (3, 3)),#main office
                ('SPAN', (4, 3), (5, 5)),#Payment
                ('SPAN', (2, 4), (3, 4)),#Delivery ad ans
                ('SPAN', (0, 5), (1, 5)), #Tax
                ('SPAN', (2, 5), (3, 5)),#Tax ans

                ('ALIGN', (0, 0), (3, 5), 'LEFT'),
                ('VALIGN', (0, 0), (3, 5), 'TOP'),
                
                ('VALIGN', (4, 0), (5, 5), 'MIDDLE'),#list item code5
                ('ALIGN', (4, 0), (5, 2), 'CENTER'),
                ('ALIGN', (2, 5), (3, 5), 'LEFT'),

                ('FONTSIZE', (2, 4), (3, 4), 8),
                ('FONTNAME', (5, 7), (5, 29), 'ThaiFont'),  # Thai font
                ('FONTSIZE', (5, 7), (5, 29), 12),
                ('FONTSIZE', (0, 6), (4, 29), 8),
                ('VALIGN', (0, 6), (5, 29), 'MIDDLE'),#list item code5
                ('ALIGN', (0, 6), (5, 29), 'CENTER'),


            ]

            table.setStyle(TableStyle(table_style))

            table.wrapOn(c, 6.81*inch, 5*inch)
            table.drawOn(c, 0.75*inch, 3.2*inch)

            c.save()
            return file_path_1

        def two_page_2_receipt(filename, main, sub):
            file_path_2 = os.path.join(self.path_sum_invoice, f"{filename}_2_receipt.pdf")
            doc_page_2 = SimpleDocTemplate(file_path_2, pagesize=A4)
            styles_page_2 = getSampleStyleSheet()
            normal_style_page_2 = styles_page_2["Normal"]
            content_page_2 = []

            text_page_2 = "This is a sample text for the delivery note."
            paragraph_page_2 = Paragraph(text_page_2, normal_style_page_2)
            content_page_2.append(paragraph_page_2)
            doc_page_2.build(content_page_2)

            #main_data = [company, company_address, main_office, main_address, tax, no_inv, date_inv, total_price, vat_price, total_amount]
            date_text = main[6]

            numeric_value = float(re.sub(r'[^\d.]', '', main[9]))
            baht_text = bahttext(numeric_value)


            def add_newlines_to_address(address):
                # Add newline characters to the address string
                modified_address = address.replace(",", ",\n")
                return modified_address
            
            n = 0.1
            pdfmetrics.registerFont(TTFont('ThaiFont', 'image/THSarabunNew Bold.ttf'))

            c = canvas.Canvas(file_path_2)
            logo_path = "image/logo_daikatte.png"
            c.drawImage(logo_path, inch, A4[1] - 1.8*inch- n*inch, width=1.43*inch, height=1.25*inch)

            sign_path = "image\sign_praew.png"
            c.drawImage(sign_path, 1.3*inch, 0.6*inch, width=1.12*inch, height=0.80*inch)

            c.drawString(2.74*inch, A4[1] - 0.67*inch - n*inch, "New Line Production Co., Ltd. (Head Office)")
            c.setFont("Helvetica", 10)
            c.drawString(2.74*inch, A4[1] - 0.98*inch - n*inch, "425 Moo 2, Soi Sukhumvit Rd. 13, Pak Nam, Mueang, Samut Prakan, 10270.")
            c.drawString(2.74*inch, A4[1] - 1.26*inch - n*inch, "TEL : 063-1684278 (Praew) , 081-8350860 (Pair)    E-mail : daikatte@gmail.com")
            c.drawString(2.74*inch, A4[1] - 1.61*inch - n*inch, "TAX : 0115560000981")
            c.setFont("ThaiFont", 14)
            c.drawString(2.5*inch,  2.2*inch , "Note : ใบวางบิลชำระแบบเครดิต 30 วัน นับตั้งแต่วันที่จังส่งใบวางบิล")
            c.drawString(1.8*inch,  2.0*inch , "Term Of Payment : Transfer to account name : New Line Production Co., Ltd.( บจก.นิวไลน์โปรดักชั่น )")
            c.drawString(2.90*inch,  1.8*inch , "Swift code: KASITHBK , NO. 066-3-51404-0, KasikornBank")
            c.drawString(1.61*inch,  1.5*inch , "LSSUED BY")
            c.drawString(1.3*inch,  0.5*inch , f"Date: {date_text}")
            c.drawString(3.83*inch,  1.5*inch , "DELIVERED BY")
            c.drawString(3.44*inch,  0.5*inch , "Date:  ............................... ")
            c.drawString(6.19*inch,  1.5*inch , "RECEIVED BY")
            c.drawString(5.82*inch,  0.5*inch , "Date:  ............................... ")
            c.drawString(7.5*inch,  0.2*inch , "Page 2/2")

            c.setFont("Helvetica", 16)
            c.drawString(3.56*inch, A4[1] - 2.10*inch - n*inch, "Receipt")

            c.setFont("Helvetica", 8)

            Number = []
            delivery_note = []
            description = []
            qty = []
            unit = []
            amount = []
            number = 1
            for index, row in sub.iterrows():
                Number.append(number)
                delivery_note.append(row.get("no_dn", ""))
                description.append("Nama Pudding & Cheese cake can")
                qty.append(row.get("qty", ""))
                unit.append("Bottle")
                amount.append(row.get("total_price_y", ""))
                number = number + 1

            max_number = len(sub)
            if not max_number == 30 :
                for _ in range(30):
                    Number.append("")
                    delivery_note.append("")
                    description.append("")
                    qty.append("")
                    unit.append("")
                    amount.append("")

            IV_name = "No. : " + main[5]
            date_text_ = "Date : " + main[6]
            data_list_page_1 = [
                ['Company:', '', main[0],'','Original','' ],
                ['Address:', '', add_newlines_to_address(main[1]),'',IV_name,'' ],
                ['', '', '','',date_text_,'' ],
                ['Delivery Address :', '', main[2],'','Payment terms\n\nCredit : 30 Day','' ],
                ['', '', add_newlines_to_address(main[3]),'','','' ],
                ['Tax :', '', main[4],'','','' ],
                ['No.', 'Delivery Note', 'Description','quantity','Unit','Amount' ],
                [Number[15], delivery_note[15], description[15],qty[15],unit[15],amount[15] ],
                [Number[16], delivery_note[16], description[16],qty[16],unit[16],amount[16] ],
                [Number[17], delivery_note[17], description[17],qty[17],unit[17],amount[17] ],
                [Number[18], delivery_note[18], description[18],qty[18],unit[18],amount[18] ],
                [Number[19], delivery_note[19], description[19],qty[19],unit[19],amount[19] ],
                [Number[20], delivery_note[20], description[20],qty[20],unit[20],amount[20] ],
                [Number[21], delivery_note[21], description[21],qty[21],unit[21],amount[21] ],
                [Number[22], delivery_note[22], description[22],qty[22],unit[22],amount[22] ],
                [Number[23], delivery_note[23], description[23],qty[23],unit[23],amount[23] ],
                [Number[24], delivery_note[24], description[24],qty[24],unit[24],amount[24] ],
                [Number[25], delivery_note[25], description[25],qty[25],unit[25],amount[25] ],
                [Number[26], delivery_note[26], description[26],qty[26],unit[26],amount[26] ],
                [Number[27], delivery_note[27], description[27],qty[27],unit[27],amount[27] ],
                [Number[28], delivery_note[28], description[28],qty[28],unit[28],amount[28] ],
                [Number[29], delivery_note[29], description[29],qty[29],unit[29],amount[29] ],
                [baht_text, '', '','',"Sub Total", main[7] ],
                ['', '', '','',"Vat 7%", main[8] ],
                ['', '', '','', "Total Amount", main[9] ],


            ]
            item_heigh = 0.25
            row_heights = [0.3*inch, 0.35*inch, 0.35*inch, 0.24*inch, 0.5*inch
                           , 0.3*inch, 0.35*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch]

            # Create table and set column widths and row heights
            table = Table(data_list_page_1, colWidths=[0.5*inch, 1.35*inch, 2.5*inch, 0.7*inch, 0.7*inch, 1*inch], rowHeights=row_heights)
            table_style = [
                ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Add border to the table
                #('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Bold font for headers
                #('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),  # Background color for headers
                ('SPAN', (0, 0), (1, 0)),#Company
                ('SPAN', (2, 0), (3, 0)),#Company ans
                ('SPAN', (4, 0), (5, 0)),#original
                ('SPAN', (0, 1), (1, 2)),#addres
                ('SPAN', (2, 1), (3, 2)),#addres ans
                ('SPAN', (4, 1), (5, 1)),#No
                ('SPAN', (4, 2), (5, 2)),#Date
                ('SPAN', (0, 3), (1, 4)),#Delivery ad
                ('SPAN', (2, 3), (3, 3)),#main office
                ('SPAN', (4, 3), (5, 5)),#Payment
                ('SPAN', (2, 4), (3, 4)),#Delivery ad ans
                ('SPAN', (0, 5), (1, 5)), #Tax
                ('SPAN', (2, 5), (3, 5)),#Tax ans
                ('SPAN', (0, 22), (3, 24)), #thai baht

                ('ALIGN', (0, 0), (3, 5), 'LEFT'),
                ('VALIGN', (0, 0), (3, 5), 'TOP'),
                
                ('VALIGN', (4, 0), (5, 5), 'MIDDLE'),#list item code5
                ('ALIGN', (4, 0), (5, 2), 'CENTER'),
                ('ALIGN', (2, 5), (3, 5), 'LEFT'),

                ('FONTSIZE', (2, 4), (3, 4), 8),
                ('FONTNAME', (5, 7), (5, 29), 'ThaiFont'),  # Thai font
                ('FONTSIZE', (5, 7), (5, 29), 12),
                ('FONTSIZE', (0, 6), (4, 29), 8),
                ('VALIGN', (0, 6), (5, 29), 'MIDDLE'),#list item code5
                ('ALIGN', (0, 6), (5, 29), 'CENTER'),
                ('FONTNAME', (0, 22), (3, 24), 'ThaiFont'),
                ('VALIGN', (0, 22), (3, 24), 'MIDDLE'),#list item code5
                ('ALIGN', (0, 22), (3, 24), 'CENTER'),
                ('FONTSIZE', (0, 22), (3, 24), 14),
            ]

            table.setStyle(TableStyle(table_style))

            table.wrapOn(c, 6.81*inch, 5*inch)
            table.drawOn(c, 0.75*inch, 2.5*inch)

            c.save()
            return file_path_2
        
        def three_page_1_invoice(filename, main, sub):

            file_path_1 = os.path.join(self.path_sum_invoice, f"{filename}_invoice_1.pdf")
            doc_page_1 = SimpleDocTemplate(file_path_1, pagesize=A4)
            styles_page_1 = getSampleStyleSheet()
            normal_style_page_1 = styles_page_1["Normal"]
            content_page_1 = []

            text_page_1 = "This is a sample text for the delivery note."
            paragraph_page_1 = Paragraph(text_page_1, normal_style_page_1)
            content_page_1.append(paragraph_page_1)
            doc_page_1.build(content_page_1)

            #main_data = [company, company_address, main_office, main_address, tax, no_inv, date_inv, total_price, vat_price, total_amount]
            date_text = main[6]

            def add_newlines_to_address(address):
                # Add newline characters to the address string
                modified_address = address.replace(",", ",\n")
                return modified_address
            
            n = 0.1
            pdfmetrics.registerFont(TTFont('ThaiFont', 'image/THSarabunNew Bold.ttf'))

            c = canvas.Canvas(file_path_1)
            logo_path = "image/logo_daikatte.png"
            c.drawImage(logo_path, inch, A4[1] - 1.8*inch- n*inch, width=1.43*inch, height=1.25*inch)

            sign_path = "image\sign_praew.png"
            c.drawImage(sign_path, 1.3*inch, 0.9*inch, width=1.12*inch, height=0.80*inch)

            c.drawString(2.74*inch, A4[1] - 0.67*inch - n*inch, "New Line Production Co., Ltd. (Head Office)")
            c.setFont("Helvetica", 10)
            c.drawString(2.74*inch, A4[1] - 0.98*inch - n*inch, "425 Moo 2, Soi Sukhumvit Rd. 13, Pak Nam, Mueang, Samut Prakan, 10270.")
            c.drawString(2.74*inch, A4[1] - 1.26*inch - n*inch, "TEL : 063-1684278 (Praew) , 081-8350860 (Pair)    E-mail : daikatte@gmail.com")
            c.drawString(2.74*inch, A4[1] - 1.61*inch - n*inch, "TAX : 0115560000981")
            c.setFont("ThaiFont", 14)
            c.drawString(2.5*inch,  2.8*inch , "Note : ใบวางบิลชำระแบบเครดิต 30 วัน นับตั้งแต่วันที่จังส่งใบวางบิล")
            c.drawString(1.8*inch,  2.5*inch , "Term Of Payment : Transfer to account name : New Line Production Co., Ltd.( บจก.นิวไลน์โปรดักชั่น )")
            c.drawString(2.90*inch,  2.3*inch , "Swift code: KASITHBK , NO. 066-3-51404-0, KasikornBank")
            c.drawString(1.61*inch,  1.8*inch , "LSSUED BY")
            c.drawString(1.3*inch,  0.8*inch , f"Date: {date_text}")
            c.drawString(3.83*inch,  1.8*inch , "DELIVERED BY")
            c.drawString(3.44*inch,  0.8*inch , "Date:  ............................... ")
            c.drawString(6.19*inch,  1.8*inch , "RECEIVED BY")
            c.drawString(5.82*inch,  0.8*inch , "Date:  ............................... ")
            c.drawString(7.5*inch,  0.2*inch , "Page 1/3")

            c.setFont("Helvetica", 16)
            c.drawString(3.56*inch, A4[1] - 2.10*inch - n*inch, "Invoice")

            c.setFont("Helvetica", 8)

            Number = []
            delivery_note = []
            description = []
            qty = []
            unit = []
            amount = []
            number = 1
            for index, row in sub.iterrows():
                Number.append(number)
                delivery_note.append(row.get("no_dn", ""))
                description.append("Nama Pudding & Cheese cake can")
                qty.append(row.get("qty", ""))
                unit.append("Bottle")
                amount.append(row.get("total_price_y", ""))
                number = number + 1

            max_number = len(sub)
            if not max_number == 15 :
                for _ in range(10 - max_number):
                    Number.append("")
                    delivery_note.append("")
                    description.append("")
                    qty.append("")
                    unit.append("")
                    amount.append("")

            IV_name = "No. : " + main[5]
            date_text_ = "Date : " + main[6]
            data_list_page_1 = [
                ['Company:', '', main[0],'','Original','' ],
                ['Address:', '', add_newlines_to_address(main[1]),'',IV_name,'' ],
                ['', '', '','',date_text_,'' ],
                ['Delivery Address :', '', main[2],'','Payment terms\n\nCredit : 30 Day','' ],
                ['', '', add_newlines_to_address(main[3]),'','','' ],
                ['Tax :', '', main[4],'','','' ],
                ['No.', 'Delivery Note', 'Description','quantity','Unit','Amount' ],
                [Number[0], delivery_note[0], description[0],qty[0],unit[0],amount[0] ],
                [Number[1], delivery_note[1], description[1],qty[1],unit[1],amount[1] ],
                [Number[2], delivery_note[2], description[2],qty[2],unit[2],amount[2] ],
                [Number[3], delivery_note[3], description[3],qty[3],unit[3],amount[3] ],
                [Number[4], delivery_note[4], description[4],qty[4],unit[4],amount[4] ],
                [Number[5], delivery_note[5], description[5],qty[5],unit[5],amount[5] ],
                [Number[6], delivery_note[6], description[6],qty[6],unit[6],amount[6] ],
                [Number[7], delivery_note[7], description[7],qty[7],unit[7],amount[7] ],
                [Number[8], delivery_note[8], description[8],qty[8],unit[8],amount[8] ],
                [Number[9], delivery_note[9], description[9],qty[9],unit[9],amount[9] ],
                [Number[10], delivery_note[10], description[10],qty[10],unit[10],amount[10] ],
                [Number[11], delivery_note[11], description[11],qty[11],unit[11],amount[11] ],
                [Number[12], delivery_note[12], description[12],qty[12],unit[12],amount[12] ],
                [Number[13], delivery_note[13], description[13],qty[13],unit[13],amount[13] ],
                [Number[14], delivery_note[14], description[14],qty[14],unit[14],amount[14] ]


            ]
            item_heigh = 0.25
            row_heights = [0.3*inch, 0.35*inch, 0.35*inch, 0.24*inch, 0.5*inch
                           , 0.3*inch, 0.35*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch]

            # Create table and set column widths and row heights
            table = Table(data_list_page_1, colWidths=[0.5*inch, 1.35*inch, 2.5*inch, 0.7*inch, 0.7*inch, 1*inch], rowHeights=row_heights)
            table_style = [
                ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Add border to the table
                #('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Bold font for headers
                #('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),  # Background color for headers
                ('SPAN', (0, 0), (1, 0)),#Company
                ('SPAN', (2, 0), (3, 0)),#Company ans
                ('SPAN', (4, 0), (5, 0)),#original
                ('SPAN', (0, 1), (1, 2)),#addres
                ('SPAN', (2, 1), (3, 2)),#addres ans
                ('SPAN', (4, 1), (5, 1)),#No
                ('SPAN', (4, 2), (5, 2)),#Date
                ('SPAN', (0, 3), (1, 4)),#Delivery ad
                ('SPAN', (2, 3), (3, 3)),#main office
                ('SPAN', (4, 3), (5, 5)),#Payment
                ('SPAN', (2, 4), (3, 4)),#Delivery ad ans
                ('SPAN', (0, 5), (1, 5)), #Tax
                ('SPAN', (2, 5), (3, 5)),#Tax ans

                ('ALIGN', (0, 0), (3, 5), 'LEFT'),
                ('VALIGN', (0, 0), (3, 5), 'TOP'),
                
                ('VALIGN', (4, 0), (5, 5), 'MIDDLE'),#list item code5
                ('ALIGN', (4, 0), (5, 2), 'CENTER'),
                ('ALIGN', (2, 5), (3, 5), 'LEFT'),

                ('FONTSIZE', (2, 4), (3, 4), 8),
                ('FONTNAME', (5, 7), (5, 29), 'ThaiFont'),  # Thai font
                ('FONTSIZE', (5, 7), (5, 29), 12),
                ('FONTSIZE', (0, 6), (4, 29), 8),
                ('VALIGN', (0, 6), (5, 29), 'MIDDLE'),#list item code5
                ('ALIGN', (0, 6), (5, 29), 'CENTER'),


            ]

            table.setStyle(TableStyle(table_style))

            table.wrapOn(c, 6.81*inch, 5*inch)
            table.drawOn(c, 0.75*inch, 3.2*inch)

            c.save()
            return file_path_1

        def three_page_2_invoice(filename, main, sub):
            file_path_2 = os.path.join(self.path_sum_invoice, f"{filename}_invoice_2.pdf")
            doc_page_2 = SimpleDocTemplate(file_path_2, pagesize=A4)
            styles_page_2 = getSampleStyleSheet()
            normal_style_page_2 = styles_page_2["Normal"]
            content_page_2 = []

            text_page_2 = "This is a sample text for the delivery note."
            paragraph_page_2 = Paragraph(text_page_2, normal_style_page_2)
            content_page_2.append(paragraph_page_2)
            doc_page_2.build(content_page_2)

            #main_data = [company, company_address, main_office, main_address, tax, no_inv, date_inv, total_price, vat_price, total_amount]
            date_text = main[6]

            def add_newlines_to_address(address):
                # Add newline characters to the address string
                modified_address = address.replace(",", ",\n")
                return modified_address
            
            n = 0.1
            pdfmetrics.registerFont(TTFont('ThaiFont', 'image/THSarabunNew Bold.ttf'))

            c = canvas.Canvas(file_path_2)
            logo_path = "image/logo_daikatte.png"
            c.drawImage(logo_path, inch, A4[1] - 1.8*inch- n*inch, width=1.43*inch, height=1.25*inch)

            sign_path = "image\sign_praew.png"
            c.drawImage(sign_path, 1.3*inch, 0.9*inch, width=1.12*inch, height=0.80*inch)

            c.drawString(2.74*inch, A4[1] - 0.67*inch - n*inch, "New Line Production Co., Ltd. (Head Office)")
            c.setFont("Helvetica", 10)
            c.drawString(2.74*inch, A4[1] - 0.98*inch - n*inch, "425 Moo 2, Soi Sukhumvit Rd. 13, Pak Nam, Mueang, Samut Prakan, 10270.")
            c.drawString(2.74*inch, A4[1] - 1.26*inch - n*inch, "TEL : 063-1684278 (Praew) , 081-8350860 (Pair)    E-mail : daikatte@gmail.com")
            c.drawString(2.74*inch, A4[1] - 1.61*inch - n*inch, "TAX : 0115560000981")
            c.setFont("ThaiFont", 14)
            c.drawString(2.5*inch,  2.8*inch , "Note : ใบวางบิลชำระแบบเครดิต 30 วัน นับตั้งแต่วันที่จังส่งใบวางบิล")
            c.drawString(1.8*inch,  2.5*inch , "Term Of Payment : Transfer to account name : New Line Production Co., Ltd.( บจก.นิวไลน์โปรดักชั่น )")
            c.drawString(2.90*inch,  2.3*inch , "Swift code: KASITHBK , NO. 066-3-51404-0, KasikornBank")
            c.drawString(1.61*inch,  1.8*inch , "LSSUED BY")
            c.drawString(1.3*inch,  0.8*inch , f"Date: {date_text}")
            c.drawString(3.83*inch,  1.8*inch , "DELIVERED BY")
            c.drawString(3.44*inch,  0.8*inch , "Date:  ............................... ")
            c.drawString(6.19*inch,  1.8*inch , "RECEIVED BY")
            c.drawString(5.82*inch,  0.8*inch , "Date:  ............................... ")
            c.drawString(7.5*inch,  0.2*inch , "Page 2/3")

            c.setFont("Helvetica", 16)
            c.drawString(3.56*inch, A4[1] - 2.10*inch - n*inch, "Invoice")

            c.setFont("Helvetica", 8)

            Number = []
            delivery_note = []
            description = []
            qty = []
            unit = []
            amount = []
            number = 1
            for index, row in sub.iterrows():
                Number.append(number)
                delivery_note.append(row.get("no_dn", ""))
                description.append("Nama Pudding & Cheese cake can")
                qty.append(row.get("qty", ""))
                unit.append("Bottle")
                amount.append(row.get("total_price_y", ""))
                number = number + 1

            max_number = len(sub)
            if not max_number == 30 :
                for _ in range(30):
                    Number.append("")
                    delivery_note.append("")
                    description.append("")
                    qty.append("")
                    unit.append("")
                    amount.append("")

            IV_name = "No. : " + main[5]
            date_text_ = "Date : " + main[6]
            data_list_page_1 = [
                ['Company:', '', main[0],'','Original','' ],
                ['Address:', '', add_newlines_to_address(main[1]),'',IV_name,'' ],
                ['', '', '','',date_text_,'' ],
                ['Delivery Address :', '', main[2],'','Payment terms\n\nCredit : 30 Day','' ],
                ['', '', add_newlines_to_address(main[3]),'','','' ],
                ['Tax :', '', main[4],'','','' ],
                ['No.', 'Delivery Note', 'Description','quantity','Unit','Amount' ],
                [Number[15], delivery_note[15], description[15],qty[15],unit[15],amount[15] ],
                [Number[16], delivery_note[16], description[16],qty[16],unit[16],amount[16] ],
                [Number[17], delivery_note[17], description[17],qty[17],unit[17],amount[17] ],
                [Number[18], delivery_note[18], description[18],qty[18],unit[18],amount[18] ],
                [Number[19], delivery_note[19], description[19],qty[19],unit[19],amount[19] ],
                [Number[20], delivery_note[20], description[20],qty[20],unit[20],amount[20] ],
                [Number[21], delivery_note[21], description[21],qty[21],unit[21],amount[21] ],
                [Number[22], delivery_note[22], description[22],qty[22],unit[22],amount[22] ],
                [Number[23], delivery_note[23], description[23],qty[23],unit[23],amount[23] ],
                [Number[24], delivery_note[24], description[24],qty[24],unit[24],amount[24] ],
                [Number[25], delivery_note[25], description[25],qty[25],unit[25],amount[25] ],
                [Number[26], delivery_note[26], description[26],qty[26],unit[26],amount[26] ],
                [Number[27], delivery_note[27], description[27],qty[27],unit[27],amount[27] ],
                [Number[28], delivery_note[28], description[28],qty[28],unit[28],amount[28] ],
                [Number[29], delivery_note[29], description[29],qty[29],unit[29],amount[29] ]


            ]
            item_heigh = 0.25
            row_heights = [0.3*inch, 0.35*inch, 0.35*inch, 0.24*inch, 0.5*inch
                           , 0.3*inch, 0.35*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch]

            # Create table and set column widths and row heights
            table = Table(data_list_page_1, colWidths=[0.5*inch, 1.35*inch, 2.5*inch, 0.7*inch, 0.7*inch, 1*inch], rowHeights=row_heights)
            table_style = [
                ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Add border to the table
                #('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Bold font for headers
                #('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),  # Background color for headers
                ('SPAN', (0, 0), (1, 0)),#Company
                ('SPAN', (2, 0), (3, 0)),#Company ans
                ('SPAN', (4, 0), (5, 0)),#original
                ('SPAN', (0, 1), (1, 2)),#addres
                ('SPAN', (2, 1), (3, 2)),#addres ans
                ('SPAN', (4, 1), (5, 1)),#No
                ('SPAN', (4, 2), (5, 2)),#Date
                ('SPAN', (0, 3), (1, 4)),#Delivery ad
                ('SPAN', (2, 3), (3, 3)),#main office
                ('SPAN', (4, 3), (5, 5)),#Payment
                ('SPAN', (2, 4), (3, 4)),#Delivery ad ans
                ('SPAN', (0, 5), (1, 5)), #Tax
                ('SPAN', (2, 5), (3, 5)),#Tax ans

                ('ALIGN', (0, 0), (3, 5), 'LEFT'),
                ('VALIGN', (0, 0), (3, 5), 'TOP'),
                
                ('VALIGN', (4, 0), (5, 5), 'MIDDLE'),#list item code5
                ('ALIGN', (4, 0), (5, 2), 'CENTER'),
                ('ALIGN', (2, 5), (3, 5), 'LEFT'),

                ('FONTSIZE', (2, 4), (3, 4), 8),
                ('FONTNAME', (5, 7), (5, 29), 'ThaiFont'),  # Thai font
                ('FONTSIZE', (5, 7), (5, 29), 12),
                ('FONTSIZE', (0, 6), (4, 29), 8),
                ('VALIGN', (0, 6), (5, 29), 'MIDDLE'),#list item code5
                ('ALIGN', (0, 6), (5, 29), 'CENTER'),


            ]

            table.setStyle(TableStyle(table_style))

            table.wrapOn(c, 6.81*inch, 5*inch)
            table.drawOn(c, 0.75*inch, 3.2*inch)

            c.save()
            return file_path_2
        
        def three_page_3_invoice(filename, main, sub):
            file_path_3 = os.path.join(self.path_sum_invoice, f"{filename}_invoice_3.pdf")
            doc_page_3 = SimpleDocTemplate(file_path_3, pagesize=A4)
            styles_page_3 = getSampleStyleSheet()
            normal_style_page_3 = styles_page_3["Normal"]
            content_page_3 = []

            text_page_3 = "This is a sample text for the delivery note."
            paragraph_page_3 = Paragraph(text_page_3, normal_style_page_3)
            content_page_3.append(paragraph_page_3)
            doc_page_3.build(content_page_3)

            #main_data = [company, company_address, main_office, main_address, tax, no_inv, date_inv, total_price, vat_price, total_amount]
            date_text = main[6]

            numeric_value = float(re.sub(r'[^\d.]', '', main[9]))
            baht_text = bahttext(numeric_value)


            def add_newlines_to_address(address):
                # Add newline characters to the address string
                modified_address = address.replace(",", ",\n")
                return modified_address
            
            n = 0.1
            pdfmetrics.registerFont(TTFont('ThaiFont', 'image/THSarabunNew Bold.ttf'))

            c = canvas.Canvas(file_path_3)
            logo_path = "image/logo_daikatte.png"
            c.drawImage(logo_path, inch, A4[1] - 1.8*inch- n*inch, width=1.43*inch, height=1.25*inch)

            sign_path = "image\sign_praew.png"
            c.drawImage(sign_path, 1.3*inch, 0.6*inch, width=1.12*inch, height=0.80*inch)

            c.drawString(2.74*inch, A4[1] - 0.67*inch - n*inch, "New Line Production Co., Ltd. (Head Office)")
            c.setFont("Helvetica", 10)
            c.drawString(2.74*inch, A4[1] - 0.98*inch - n*inch, "425 Moo 2, Soi Sukhumvit Rd. 13, Pak Nam, Mueang, Samut Prakan, 10270.")
            c.drawString(2.74*inch, A4[1] - 1.26*inch - n*inch, "TEL : 063-1684278 (Praew) , 081-8350860 (Pair)    E-mail : daikatte@gmail.com")
            c.drawString(2.74*inch, A4[1] - 1.61*inch - n*inch, "TAX : 0115560000981")
            c.setFont("ThaiFont", 14)
            c.drawString(2.5*inch,  2.2*inch , "Note : ใบวางบิลชำระแบบเครดิต 30 วัน นับตั้งแต่วันที่จังส่งใบวางบิล")
            c.drawString(1.8*inch,  2.0*inch , "Term Of Payment : Transfer to account name : New Line Production Co., Ltd.( บจก.นิวไลน์โปรดักชั่น )")
            c.drawString(2.90*inch,  1.8*inch , "Swift code: KASITHBK , NO. 066-3-51404-0, KasikornBank")
            c.drawString(1.61*inch,  1.5*inch , "LSSUED BY")
            c.drawString(1.3*inch,  0.5*inch , f"Date: {date_text}")
            c.drawString(3.83*inch,  1.5*inch , "DELIVERED BY")
            c.drawString(3.44*inch,  0.5*inch , "Date:  ............................... ")
            c.drawString(6.19*inch,  1.5*inch , "RECEIVED BY")
            c.drawString(5.82*inch,  0.5*inch , "Date:  ............................... ")
            c.drawString(7.5*inch,  0.2*inch , "Page 3/3")

            c.setFont("Helvetica", 16)
            c.drawString(3.56*inch, A4[1] - 2.10*inch - n*inch, "Invoice")

            c.setFont("Helvetica", 8)

            Number = []
            delivery_note = []
            description = []
            qty = []
            unit = []
            amount = []
            number = 1
            for index, row in sub.iterrows():
                Number.append(number)
                delivery_note.append(row.get("no_dn", ""))
                description.append("Nama Pudding & Cheese cake can")
                qty.append(row.get("qty", ""))
                unit.append("Bottle")
                amount.append(row.get("total_price_y", ""))
                number = number + 1

            max_number = len(sub)
            if not max_number == 45 :
                for _ in range(45):
                    Number.append("")
                    delivery_note.append("")
                    description.append("")
                    qty.append("")
                    unit.append("")
                    amount.append("")

            IV_name = "No. : " + main[5]
            date_text_ = "Date : " + main[6]
            data_list_page_1 = [
                ['Company:', '', main[0],'','Original','' ],
                ['Address:', '', add_newlines_to_address(main[1]),'',IV_name,'' ],
                ['', '', '','',date_text_,'' ],
                ['Delivery Address :', '', main[2],'','Payment terms\n\nCredit : 30 Day','' ],
                ['', '', add_newlines_to_address(main[3]),'','','' ],
                ['Tax :', '', main[4],'','','' ],
                ['No.', 'Delivery Note', 'Description','quantity','Unit','Amount' ],
                [Number[30], delivery_note[30], description[30],qty[30],unit[30],amount[30] ],
                [Number[31], delivery_note[31], description[31],qty[31],unit[31],amount[31] ],
                [Number[32], delivery_note[32], description[32],qty[32],unit[32],amount[32] ],
                [Number[33], delivery_note[33], description[33],qty[33],unit[33],amount[33] ],
                [Number[34], delivery_note[34], description[34],qty[34],unit[34],amount[34] ],
                [Number[35], delivery_note[35], description[35],qty[35],unit[35],amount[35] ],
                [Number[36], delivery_note[36], description[36],qty[36],unit[36],amount[36] ],
                [Number[37], delivery_note[37], description[37],qty[37],unit[37],amount[37] ],
                [Number[38], delivery_note[38], description[38],qty[38],unit[38],amount[38] ],
                [Number[39], delivery_note[39], description[39],qty[39],unit[39],amount[39] ],
                [Number[40], delivery_note[40], description[40],qty[40],unit[40],amount[40] ],
                [Number[41], delivery_note[41], description[41],qty[41],unit[41],amount[41] ],
                [Number[42], delivery_note[42], description[42],qty[42],unit[42],amount[42] ],
                [Number[43], delivery_note[43], description[43],qty[43],unit[43],amount[43] ],
                [Number[44], delivery_note[44], description[44],qty[44],unit[44],amount[44] ],
                [baht_text, '', '','',"Sub Total", main[7] ],
                ['', '', '','',"Vat 7%", main[8] ],
                ['', '', '','', "Total Amount", main[9] ],


            ]
            item_heigh = 0.25
            row_heights = [0.3*inch, 0.35*inch, 0.35*inch, 0.24*inch, 0.5*inch
                           , 0.3*inch, 0.35*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch]

            # Create table and set column widths and row heights
            table = Table(data_list_page_1, colWidths=[0.5*inch, 1.35*inch, 2.5*inch, 0.7*inch, 0.7*inch, 1*inch], rowHeights=row_heights)
            table_style = [
                ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Add border to the table
                #('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Bold font for headers
                #('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),  # Background color for headers
                ('SPAN', (0, 0), (1, 0)),#Company
                ('SPAN', (2, 0), (3, 0)),#Company ans
                ('SPAN', (4, 0), (5, 0)),#original
                ('SPAN', (0, 1), (1, 2)),#addres
                ('SPAN', (2, 1), (3, 2)),#addres ans
                ('SPAN', (4, 1), (5, 1)),#No
                ('SPAN', (4, 2), (5, 2)),#Date
                ('SPAN', (0, 3), (1, 4)),#Delivery ad
                ('SPAN', (2, 3), (3, 3)),#main office
                ('SPAN', (4, 3), (5, 5)),#Payment
                ('SPAN', (2, 4), (3, 4)),#Delivery ad ans
                ('SPAN', (0, 5), (1, 5)), #Tax
                ('SPAN', (2, 5), (3, 5)),#Tax ans
                ('SPAN', (0, 22), (3, 24)), #thai baht

                ('ALIGN', (0, 0), (3, 5), 'LEFT'),
                ('VALIGN', (0, 0), (3, 5), 'TOP'),
                
                ('VALIGN', (4, 0), (5, 5), 'MIDDLE'),#list item code5
                ('ALIGN', (4, 0), (5, 2), 'CENTER'),
                ('ALIGN', (2, 5), (3, 5), 'LEFT'),

                ('FONTSIZE', (2, 4), (3, 4), 8),
                ('FONTNAME', (5, 7), (5, 29), 'ThaiFont'),  # Thai font
                ('FONTSIZE', (5, 7), (5, 29), 12),
                ('FONTSIZE', (0, 6), (4, 29), 8),
                ('VALIGN', (0, 6), (5, 29), 'MIDDLE'),#list item code5
                ('ALIGN', (0, 6), (5, 29), 'CENTER'),
                ('FONTNAME', (0, 22), (3, 24), 'ThaiFont'),
                ('VALIGN', (0, 22), (3, 24), 'MIDDLE'),#list item code5
                ('ALIGN', (0, 22), (3, 24), 'CENTER'),
                ('FONTSIZE', (0, 22), (3, 24), 14),
            ]

            table.setStyle(TableStyle(table_style))

            table.wrapOn(c, 6.81*inch, 5*inch)
            table.drawOn(c, 0.75*inch, 2.5*inch)

            c.save()
            return file_path_3

        def three_page_1_receipt(filename, main, sub):

            file_path_1 = os.path.join(self.path_sum_invoice, f"{filename}_receipt_1.pdf")
            doc_page_1 = SimpleDocTemplate(file_path_1, pagesize=A4)
            styles_page_1 = getSampleStyleSheet()
            normal_style_page_1 = styles_page_1["Normal"]
            content_page_1 = []

            text_page_1 = "This is a sample text for the delivery note."
            paragraph_page_1 = Paragraph(text_page_1, normal_style_page_1)
            content_page_1.append(paragraph_page_1)
            doc_page_1.build(content_page_1)

            #main_data = [company, company_address, main_office, main_address, tax, no_inv, date_inv, total_price, vat_price, total_amount]
            date_text = main[6]

            def add_newlines_to_address(address):
                # Add newline characters to the address string
                modified_address = address.replace(",", ",\n")
                return modified_address
            
            n = 0.1
            pdfmetrics.registerFont(TTFont('ThaiFont', 'image/THSarabunNew Bold.ttf'))

            c = canvas.Canvas(file_path_1)
            logo_path = "image/logo_daikatte.png"
            c.drawImage(logo_path, inch, A4[1] - 1.8*inch- n*inch, width=1.43*inch, height=1.25*inch)

            sign_path = "image\sign_praew.png"
            c.drawImage(sign_path, 1.3*inch, 0.9*inch, width=1.12*inch, height=0.80*inch)

            c.drawString(2.74*inch, A4[1] - 0.67*inch - n*inch, "New Line Production Co., Ltd. (Head Office)")
            c.setFont("Helvetica", 10)
            c.drawString(2.74*inch, A4[1] - 0.98*inch - n*inch, "425 Moo 2, Soi Sukhumvit Rd. 13, Pak Nam, Mueang, Samut Prakan, 10270.")
            c.drawString(2.74*inch, A4[1] - 1.26*inch - n*inch, "TEL : 063-1684278 (Praew) , 081-8350860 (Pair)    E-mail : daikatte@gmail.com")
            c.drawString(2.74*inch, A4[1] - 1.61*inch - n*inch, "TAX : 0115560000981")
            c.setFont("ThaiFont", 14)
            c.drawString(2.5*inch,  2.8*inch , "Note : ใบวางบิลชำระแบบเครดิต 30 วัน นับตั้งแต่วันที่จังส่งใบวางบิล")
            c.drawString(1.8*inch,  2.5*inch , "Term Of Payment : Transfer to account name : New Line Production Co., Ltd.( บจก.นิวไลน์โปรดักชั่น )")
            c.drawString(2.90*inch,  2.3*inch , "Swift code: KASITHBK , NO. 066-3-51404-0, KasikornBank")
            c.drawString(1.61*inch,  1.8*inch , "LSSUED BY")
            c.drawString(1.3*inch,  0.8*inch , f"Date: {date_text}")
            c.drawString(3.83*inch,  1.8*inch , "DELIVERED BY")
            c.drawString(3.44*inch,  0.8*inch , "Date:  ............................... ")
            c.drawString(6.19*inch,  1.8*inch , "RECEIVED BY")
            c.drawString(5.82*inch,  0.8*inch , "Date:  ............................... ")
            c.drawString(7.5*inch,  0.2*inch , "Page 1/3")

            c.setFont("Helvetica", 16)
            c.drawString(3.56*inch, A4[1] - 2.10*inch - n*inch, "Receipt")

            c.setFont("Helvetica", 8)

            Number = []
            delivery_note = []
            description = []
            qty = []
            unit = []
            amount = []
            number = 1
            for index, row in sub.iterrows():
                Number.append(number)
                delivery_note.append(row.get("no_dn", ""))
                description.append("Nama Pudding & Cheese cake can")
                qty.append(row.get("qty", ""))
                unit.append("Bottle")
                amount.append(row.get("total_price_y", ""))
                number = number + 1

            max_number = len(sub)
            if not max_number == 15 :
                for _ in range(10 - max_number):
                    Number.append("")
                    delivery_note.append("")
                    description.append("")
                    qty.append("")
                    unit.append("")
                    amount.append("")

            IV_name = "No. : " + main[5]
            date_text_ = "Date : " + main[6]
            data_list_page_1 = [
                ['Company:', '', main[0],'','Original','' ],
                ['Address:', '', add_newlines_to_address(main[1]),'',IV_name,'' ],
                ['', '', '','',date_text_,'' ],
                ['Delivery Address :', '', main[2],'','Payment terms\n\nCredit : 30 Day','' ],
                ['', '', add_newlines_to_address(main[3]),'','','' ],
                ['Tax :', '', main[4],'','','' ],
                ['No.', 'Delivery Note', 'Description','quantity','Unit','Amount' ],
                [Number[0], delivery_note[0], description[0],qty[0],unit[0],amount[0] ],
                [Number[1], delivery_note[1], description[1],qty[1],unit[1],amount[1] ],
                [Number[2], delivery_note[2], description[2],qty[2],unit[2],amount[2] ],
                [Number[3], delivery_note[3], description[3],qty[3],unit[3],amount[3] ],
                [Number[4], delivery_note[4], description[4],qty[4],unit[4],amount[4] ],
                [Number[5], delivery_note[5], description[5],qty[5],unit[5],amount[5] ],
                [Number[6], delivery_note[6], description[6],qty[6],unit[6],amount[6] ],
                [Number[7], delivery_note[7], description[7],qty[7],unit[7],amount[7] ],
                [Number[8], delivery_note[8], description[8],qty[8],unit[8],amount[8] ],
                [Number[9], delivery_note[9], description[9],qty[9],unit[9],amount[9] ],
                [Number[10], delivery_note[10], description[10],qty[10],unit[10],amount[10] ],
                [Number[11], delivery_note[11], description[11],qty[11],unit[11],amount[11] ],
                [Number[12], delivery_note[12], description[12],qty[12],unit[12],amount[12] ],
                [Number[13], delivery_note[13], description[13],qty[13],unit[13],amount[13] ],
                [Number[14], delivery_note[14], description[14],qty[14],unit[14],amount[14] ]


            ]
            item_heigh = 0.25
            row_heights = [0.3*inch, 0.35*inch, 0.35*inch, 0.24*inch, 0.5*inch
                           , 0.3*inch, 0.35*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch]

            # Create table and set column widths and row heights
            table = Table(data_list_page_1, colWidths=[0.5*inch, 1.35*inch, 2.5*inch, 0.7*inch, 0.7*inch, 1*inch], rowHeights=row_heights)
            table_style = [
                ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Add border to the table
                #('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Bold font for headers
                #('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),  # Background color for headers
                ('SPAN', (0, 0), (1, 0)),#Company
                ('SPAN', (2, 0), (3, 0)),#Company ans
                ('SPAN', (4, 0), (5, 0)),#original
                ('SPAN', (0, 1), (1, 2)),#addres
                ('SPAN', (2, 1), (3, 2)),#addres ans
                ('SPAN', (4, 1), (5, 1)),#No
                ('SPAN', (4, 2), (5, 2)),#Date
                ('SPAN', (0, 3), (1, 4)),#Delivery ad
                ('SPAN', (2, 3), (3, 3)),#main office
                ('SPAN', (4, 3), (5, 5)),#Payment
                ('SPAN', (2, 4), (3, 4)),#Delivery ad ans
                ('SPAN', (0, 5), (1, 5)), #Tax
                ('SPAN', (2, 5), (3, 5)),#Tax ans

                ('ALIGN', (0, 0), (3, 5), 'LEFT'),
                ('VALIGN', (0, 0), (3, 5), 'TOP'),
                
                ('VALIGN', (4, 0), (5, 5), 'MIDDLE'),#list item code5
                ('ALIGN', (4, 0), (5, 2), 'CENTER'),
                ('ALIGN', (2, 5), (3, 5), 'LEFT'),

                ('FONTSIZE', (2, 4), (3, 4), 8),
                ('FONTNAME', (5, 7), (5, 29), 'ThaiFont'),  # Thai font
                ('FONTSIZE', (5, 7), (5, 29), 12),
                ('FONTSIZE', (0, 6), (4, 29), 8),
                ('VALIGN', (0, 6), (5, 29), 'MIDDLE'),#list item code5
                ('ALIGN', (0, 6), (5, 29), 'CENTER'),


            ]

            table.setStyle(TableStyle(table_style))

            table.wrapOn(c, 6.81*inch, 5*inch)
            table.drawOn(c, 0.75*inch, 3.2*inch)

            c.save()
            return file_path_1

        def three_page_2_receipt(filename, main, sub):
            file_path_2 = os.path.join(self.path_sum_invoice, f"{filename}_receipt_2.pdf")
            doc_page_2 = SimpleDocTemplate(file_path_2, pagesize=A4)
            styles_page_2 = getSampleStyleSheet()
            normal_style_page_2 = styles_page_2["Normal"]
            content_page_2 = []

            text_page_2 = "This is a sample text for the delivery note."
            paragraph_page_2 = Paragraph(text_page_2, normal_style_page_2)
            content_page_2.append(paragraph_page_2)
            doc_page_2.build(content_page_2)

            #main_data = [company, company_address, main_office, main_address, tax, no_inv, date_inv, total_price, vat_price, total_amount]
            date_text = main[6]

            def add_newlines_to_address(address):
                # Add newline characters to the address string
                modified_address = address.replace(",", ",\n")
                return modified_address
            
            n = 0.1
            pdfmetrics.registerFont(TTFont('ThaiFont', 'image/THSarabunNew Bold.ttf'))

            c = canvas.Canvas(file_path_2)
            logo_path = "image/logo_daikatte.png"
            c.drawImage(logo_path, inch, A4[1] - 1.8*inch- n*inch, width=1.43*inch, height=1.25*inch)

            sign_path = "image\sign_praew.png"
            c.drawImage(sign_path, 1.3*inch, 0.9*inch, width=1.12*inch, height=0.80*inch)

            c.drawString(2.74*inch, A4[1] - 0.67*inch - n*inch, "New Line Production Co., Ltd. (Head Office)")
            c.setFont("Helvetica", 10)
            c.drawString(2.74*inch, A4[1] - 0.98*inch - n*inch, "425 Moo 2, Soi Sukhumvit Rd. 13, Pak Nam, Mueang, Samut Prakan, 10270.")
            c.drawString(2.74*inch, A4[1] - 1.26*inch - n*inch, "TEL : 063-1684278 (Praew) , 081-8350860 (Pair)    E-mail : daikatte@gmail.com")
            c.drawString(2.74*inch, A4[1] - 1.61*inch - n*inch, "TAX : 0115560000981")
            c.setFont("ThaiFont", 14)
            c.drawString(2.5*inch,  2.8*inch , "Note : ใบวางบิลชำระแบบเครดิต 30 วัน นับตั้งแต่วันที่จังส่งใบวางบิล")
            c.drawString(1.8*inch,  2.5*inch , "Term Of Payment : Transfer to account name : New Line Production Co., Ltd.( บจก.นิวไลน์โปรดักชั่น )")
            c.drawString(2.90*inch,  2.3*inch , "Swift code: KASITHBK , NO. 066-3-51404-0, KasikornBank")
            c.drawString(1.61*inch,  1.8*inch , "LSSUED BY")
            c.drawString(1.3*inch,  0.8*inch , f"Date: {date_text}")
            c.drawString(3.83*inch,  1.8*inch , "DELIVERED BY")
            c.drawString(3.44*inch,  0.8*inch , "Date:  ............................... ")
            c.drawString(6.19*inch,  1.8*inch , "RECEIVED BY")
            c.drawString(5.82*inch,  0.8*inch , "Date:  ............................... ")
            c.drawString(7.5*inch,  0.2*inch , "Page 2/3")

            c.setFont("Helvetica", 16)
            c.drawString(3.56*inch, A4[1] - 2.10*inch - n*inch, "Receipt")

            c.setFont("Helvetica", 8)

            Number = []
            delivery_note = []
            description = []
            qty = []
            unit = []
            amount = []
            number = 1
            for index, row in sub.iterrows():
                Number.append(number)
                delivery_note.append(row.get("no_dn", ""))
                description.append("Nama Pudding & Cheese cake can")
                qty.append(row.get("qty", ""))
                unit.append("Bottle")
                amount.append(row.get("total_price_y", ""))
                number = number + 1

            max_number = len(sub)
            if not max_number == 30 :
                for _ in range(30):
                    Number.append("")
                    delivery_note.append("")
                    description.append("")
                    qty.append("")
                    unit.append("")
                    amount.append("")

            IV_name = "No. : " + main[5]
            date_text_ = "Date : " + main[6]
            data_list_page_1 = [
                ['Company:', '', main[0],'','Original','' ],
                ['Address:', '', add_newlines_to_address(main[1]),'',IV_name,'' ],
                ['', '', '','',date_text_,'' ],
                ['Delivery Address :', '', main[2],'','Payment terms\n\nCredit : 30 Day','' ],
                ['', '', add_newlines_to_address(main[3]),'','','' ],
                ['Tax :', '', main[4],'','','' ],
                ['No.', 'Delivery Note', 'Description','quantity','Unit','Amount' ],
                [Number[15], delivery_note[15], description[15],qty[15],unit[15],amount[15] ],
                [Number[16], delivery_note[16], description[16],qty[16],unit[16],amount[16] ],
                [Number[17], delivery_note[17], description[17],qty[17],unit[17],amount[17] ],
                [Number[18], delivery_note[18], description[18],qty[18],unit[18],amount[18] ],
                [Number[19], delivery_note[19], description[19],qty[19],unit[19],amount[19] ],
                [Number[20], delivery_note[20], description[20],qty[20],unit[20],amount[20] ],
                [Number[21], delivery_note[21], description[21],qty[21],unit[21],amount[21] ],
                [Number[22], delivery_note[22], description[22],qty[22],unit[22],amount[22] ],
                [Number[23], delivery_note[23], description[23],qty[23],unit[23],amount[23] ],
                [Number[24], delivery_note[24], description[24],qty[24],unit[24],amount[24] ],
                [Number[25], delivery_note[25], description[25],qty[25],unit[25],amount[25] ],
                [Number[26], delivery_note[26], description[26],qty[26],unit[26],amount[26] ],
                [Number[27], delivery_note[27], description[27],qty[27],unit[27],amount[27] ],
                [Number[28], delivery_note[28], description[28],qty[28],unit[28],amount[28] ],
                [Number[29], delivery_note[29], description[29],qty[29],unit[29],amount[29] ]


            ]
            item_heigh = 0.25
            row_heights = [0.3*inch, 0.35*inch, 0.35*inch, 0.24*inch, 0.5*inch
                           , 0.3*inch, 0.35*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch]

            # Create table and set column widths and row heights
            table = Table(data_list_page_1, colWidths=[0.5*inch, 1.35*inch, 2.5*inch, 0.7*inch, 0.7*inch, 1*inch], rowHeights=row_heights)
            table_style = [
                ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Add border to the table
                #('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Bold font for headers
                #('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),  # Background color for headers
                ('SPAN', (0, 0), (1, 0)),#Company
                ('SPAN', (2, 0), (3, 0)),#Company ans
                ('SPAN', (4, 0), (5, 0)),#original
                ('SPAN', (0, 1), (1, 2)),#addres
                ('SPAN', (2, 1), (3, 2)),#addres ans
                ('SPAN', (4, 1), (5, 1)),#No
                ('SPAN', (4, 2), (5, 2)),#Date
                ('SPAN', (0, 3), (1, 4)),#Delivery ad
                ('SPAN', (2, 3), (3, 3)),#main office
                ('SPAN', (4, 3), (5, 5)),#Payment
                ('SPAN', (2, 4), (3, 4)),#Delivery ad ans
                ('SPAN', (0, 5), (1, 5)), #Tax
                ('SPAN', (2, 5), (3, 5)),#Tax ans

                ('ALIGN', (0, 0), (3, 5), 'LEFT'),
                ('VALIGN', (0, 0), (3, 5), 'TOP'),
                
                ('VALIGN', (4, 0), (5, 5), 'MIDDLE'),#list item code5
                ('ALIGN', (4, 0), (5, 2), 'CENTER'),
                ('ALIGN', (2, 5), (3, 5), 'LEFT'),

                ('FONTSIZE', (2, 4), (3, 4), 8),
                ('FONTNAME', (5, 7), (5, 29), 'ThaiFont'),  # Thai font
                ('FONTSIZE', (5, 7), (5, 29), 12),
                ('FONTSIZE', (0, 6), (4, 29), 8),
                ('VALIGN', (0, 6), (5, 29), 'MIDDLE'),#list item code5
                ('ALIGN', (0, 6), (5, 29), 'CENTER'),


            ]

            table.setStyle(TableStyle(table_style))

            table.wrapOn(c, 6.81*inch, 5*inch)
            table.drawOn(c, 0.75*inch, 3.2*inch)

            c.save()
            return file_path_2
        
        def three_page_3_receipt(filename, main, sub):
            file_path_3 = os.path.join(self.path_sum_invoice, f"{filename}_receipt_3.pdf")
            doc_page_3 = SimpleDocTemplate(file_path_3, pagesize=A4)
            styles_page_3 = getSampleStyleSheet()
            normal_style_page_3 = styles_page_3["Normal"]
            content_page_3 = []

            text_page_3 = "This is a sample text for the delivery note."
            paragraph_page_3 = Paragraph(text_page_3, normal_style_page_3)
            content_page_3.append(paragraph_page_3)
            doc_page_3.build(content_page_3)

            #main_data = [company, company_address, main_office, main_address, tax, no_inv, date_inv, total_price, vat_price, total_amount]
            date_text = main[6]

            numeric_value = float(re.sub(r'[^\d.]', '', main[9]))
            baht_text = bahttext(numeric_value)


            def add_newlines_to_address(address):
                # Add newline characters to the address string
                modified_address = address.replace(",", ",\n")
                return modified_address
            
            n = 0.1
            pdfmetrics.registerFont(TTFont('ThaiFont', 'image/THSarabunNew Bold.ttf'))

            c = canvas.Canvas(file_path_3)
            logo_path = "image/logo_daikatte.png"
            c.drawImage(logo_path, inch, A4[1] - 1.8*inch- n*inch, width=1.43*inch, height=1.25*inch)

            sign_path = "image\sign_praew.png"
            c.drawImage(sign_path, 1.3*inch, 0.6*inch, width=1.12*inch, height=0.80*inch)

            c.drawString(2.74*inch, A4[1] - 0.67*inch - n*inch, "New Line Production Co., Ltd. (Head Office)")
            c.setFont("Helvetica", 10)
            c.drawString(2.74*inch, A4[1] - 0.98*inch - n*inch, "425 Moo 2, Soi Sukhumvit Rd. 13, Pak Nam, Mueang, Samut Prakan, 10270.")
            c.drawString(2.74*inch, A4[1] - 1.26*inch - n*inch, "TEL : 063-1684278 (Praew) , 081-8350860 (Pair)    E-mail : daikatte@gmail.com")
            c.drawString(2.74*inch, A4[1] - 1.61*inch - n*inch, "TAX : 0115560000981")
            c.setFont("ThaiFont", 14)
            c.drawString(2.5*inch,  2.2*inch , "Note : ใบวางบิลชำระแบบเครดิต 30 วัน นับตั้งแต่วันที่จังส่งใบวางบิล")
            c.drawString(1.8*inch,  2.0*inch , "Term Of Payment : Transfer to account name : New Line Production Co., Ltd.( บจก.นิวไลน์โปรดักชั่น )")
            c.drawString(2.90*inch,  1.8*inch , "Swift code: KASITHBK , NO. 066-3-51404-0, KasikornBank")
            c.drawString(1.61*inch,  1.5*inch , "LSSUED BY")
            c.drawString(1.3*inch,  0.5*inch , f"Date: {date_text}")
            c.drawString(3.83*inch,  1.5*inch , "DELIVERED BY")
            c.drawString(3.44*inch,  0.5*inch , "Date:  ............................... ")
            c.drawString(6.19*inch,  1.5*inch , "RECEIVED BY")
            c.drawString(5.82*inch,  0.5*inch , "Date:  ............................... ")
            c.drawString(7.5*inch,  0.2*inch , "Page 3/3")

            c.setFont("Helvetica", 16)
            c.drawString(3.56*inch, A4[1] - 2.10*inch - n*inch, "Receipt")

            c.setFont("Helvetica", 8)

            Number = []
            delivery_note = []
            description = []
            qty = []
            unit = []
            amount = []
            number = 1
            for index, row in sub.iterrows():
                Number.append(number)
                delivery_note.append(row.get("no_dn", ""))
                description.append("Nama Pudding & Cheese cake can")
                qty.append(row.get("qty", ""))
                unit.append("Bottle")
                amount.append(row.get("total_price_y", ""))
                number = number + 1

            max_number = len(sub)
            if not max_number == 45 :
                for _ in range(45):
                    Number.append("")
                    delivery_note.append("")
                    description.append("")
                    qty.append("")
                    unit.append("")
                    amount.append("")

            IV_name = "No. : " + main[5]
            date_text_ = "Date : " + main[6]
            data_list_page_1 = [
                ['Company:', '', main[0],'','Original','' ],
                ['Address:', '', add_newlines_to_address(main[1]),'',IV_name,'' ],
                ['', '', '','',date_text_,'' ],
                ['Delivery Address :', '', main[2],'','Payment terms\n\nCredit : 30 Day','' ],
                ['', '', add_newlines_to_address(main[3]),'','','' ],
                ['Tax :', '', main[4],'','','' ],
                ['No.', 'Delivery Note', 'Description','quantity','Unit','Amount' ],
                [Number[30], delivery_note[30], description[30],qty[30],unit[30],amount[30] ],
                [Number[31], delivery_note[31], description[31],qty[31],unit[31],amount[31] ],
                [Number[32], delivery_note[32], description[32],qty[32],unit[32],amount[32] ],
                [Number[33], delivery_note[33], description[33],qty[33],unit[33],amount[33] ],
                [Number[34], delivery_note[34], description[34],qty[34],unit[34],amount[34] ],
                [Number[35], delivery_note[35], description[35],qty[35],unit[35],amount[35] ],
                [Number[36], delivery_note[36], description[36],qty[36],unit[36],amount[36] ],
                [Number[37], delivery_note[37], description[37],qty[37],unit[37],amount[37] ],
                [Number[38], delivery_note[38], description[38],qty[38],unit[38],amount[38] ],
                [Number[39], delivery_note[39], description[39],qty[39],unit[39],amount[39] ],
                [Number[40], delivery_note[40], description[40],qty[40],unit[40],amount[40] ],
                [Number[41], delivery_note[41], description[41],qty[41],unit[41],amount[41] ],
                [Number[42], delivery_note[42], description[42],qty[42],unit[42],amount[42] ],
                [Number[43], delivery_note[43], description[43],qty[43],unit[43],amount[43] ],
                [Number[44], delivery_note[44], description[44],qty[44],unit[44],amount[44] ],
                [baht_text, '', '','',"Sub Total", main[7] ],
                ['', '', '','',"Vat 7%", main[8] ],
                ['', '', '','', "Total Amount", main[9] ],


            ]
            item_heigh = 0.25
            row_heights = [0.3*inch, 0.35*inch, 0.35*inch, 0.24*inch, 0.5*inch
                           , 0.3*inch, 0.35*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch
                           , item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch, item_heigh*inch]

            # Create table and set column widths and row heights
            table = Table(data_list_page_1, colWidths=[0.5*inch, 1.35*inch, 2.5*inch, 0.7*inch, 0.7*inch, 1*inch], rowHeights=row_heights)
            table_style = [
                ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Add border to the table
                #('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Bold font for headers
                #('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),  # Background color for headers
                ('SPAN', (0, 0), (1, 0)),#Company
                ('SPAN', (2, 0), (3, 0)),#Company ans
                ('SPAN', (4, 0), (5, 0)),#original
                ('SPAN', (0, 1), (1, 2)),#addres
                ('SPAN', (2, 1), (3, 2)),#addres ans
                ('SPAN', (4, 1), (5, 1)),#No
                ('SPAN', (4, 2), (5, 2)),#Date
                ('SPAN', (0, 3), (1, 4)),#Delivery ad
                ('SPAN', (2, 3), (3, 3)),#main office
                ('SPAN', (4, 3), (5, 5)),#Payment
                ('SPAN', (2, 4), (3, 4)),#Delivery ad ans
                ('SPAN', (0, 5), (1, 5)), #Tax
                ('SPAN', (2, 5), (3, 5)),#Tax ans
                ('SPAN', (0, 22), (3, 24)), #thai baht

                ('ALIGN', (0, 0), (3, 5), 'LEFT'),
                ('VALIGN', (0, 0), (3, 5), 'TOP'),
                
                ('VALIGN', (4, 0), (5, 5), 'MIDDLE'),#list item code5
                ('ALIGN', (4, 0), (5, 2), 'CENTER'),
                ('ALIGN', (2, 5), (3, 5), 'LEFT'),

                ('FONTSIZE', (2, 4), (3, 4), 8),
                ('FONTNAME', (5, 7), (5, 29), 'ThaiFont'),  # Thai font
                ('FONTSIZE', (5, 7), (5, 29), 12),
                ('FONTSIZE', (0, 6), (4, 29), 8),
                ('VALIGN', (0, 6), (5, 29), 'MIDDLE'),#list item code5
                ('ALIGN', (0, 6), (5, 29), 'CENTER'),
                ('FONTNAME', (0, 22), (3, 24), 'ThaiFont'),
                ('VALIGN', (0, 22), (3, 24), 'MIDDLE'),#list item code5
                ('ALIGN', (0, 22), (3, 24), 'CENTER'),
                ('FONTSIZE', (0, 22), (3, 24), 14),
            ]

            table.setStyle(TableStyle(table_style))

            table.wrapOn(c, 6.81*inch, 5*inch)
            table.drawOn(c, 0.75*inch, 2.5*inch)

            c.save()
            return file_path_3

        def merge_pdfs(pdf1_path, pdf2_path, output_name):
                output_path = os.path.join(self.path_sum_invoice, f"{output_name}.pdf")
                # Open the first PDF
                with open(pdf1_path, 'rb') as pdf1_file:
                    pdf1_reader = PyPDF2.PdfReader(pdf1_file)
                    pdf1_pages = pdf1_reader.pages

                    # Open the second PDF
                    with open(pdf2_path, 'rb') as pdf2_file:
                        pdf2_reader = PyPDF2.PdfReader(pdf2_file)

                        # Create a new PDF writer
                        pdf_writer = PyPDF2.PdfWriter()

                        # Add pages from the first PDF
                        for page in pdf1_pages:
                            pdf_writer.add_page(page)

                        # Add pages from the second PDF
                        for page in pdf2_reader.pages:
                            pdf_writer.add_page(page)

                        # Write the merged PDF to the output file
                        with open(output_path, 'wb') as output_file:
                            pdf_writer.write(output_file)
                            print(f"Delivery note PDF created successfully at: {output_path}")
                return output_path

        def merge_3pdfs(pdf1_path, pdf2_path, pdf3_path, output_name):
                output_path = os.path.join(self.path_sum_invoice, f"{output_name}.pdf")
                # Open the first PDF
                with open(pdf1_path, 'rb') as pdf1_file:
                    pdf1_reader = PyPDF2.PdfReader(pdf1_file)
                    pdf1_pages = pdf1_reader.pages

                    # Open the second PDF
                    with open(pdf2_path, 'rb') as pdf2_file:
                        pdf2_reader = PyPDF2.PdfReader(pdf2_file)

                        with open(pdf3_path, 'rb') as pdf3_file:
                            pdf3_reader = PyPDF2.PdfReader(pdf3_file) 

                            # Create a new PDF writer
                            pdf_writer = PyPDF2.PdfWriter()

                            # Add pages from the first PDF
                            for page in pdf1_pages:
                                pdf_writer.add_page(page)

                            # Add pages from the second PDF
                            for page in pdf2_reader.pages:
                                pdf_writer.add_page(page)

                            for page in pdf3_reader.pages:
                                pdf_writer.add_page(page)

                            # Write the merged PDF to the output file
                            with open(output_path, 'wb') as output_file:
                                pdf_writer.write(output_file)
                                print(f"Delivery note PDF created successfully at: {output_path}")
                return output_path

        if len(sub_data) <= 15:

            path_invoice = one_page_invoice(file_name, main_data, sub_data)
            path_receipt = one_page_receipt(file_name, main_data, sub_data)
            merge_pdfs(path_invoice, path_receipt, file_name)
            os.remove(path_invoice)
            os.remove(path_receipt)

        elif len(sub_data) > 15 and len(sub_data) <= 30:

            file_name_invoice = file_name + "_invoice"
            file_name_receipt = file_name + "_receipt"

            path_invoice_1 = two_page_1_invoice(file_name, main_data, sub_data)
            path_invoice_2 = two_page_2_invoice(file_name, main_data, sub_data)
            path_receipt_1 = two_page_1_receipt(file_name, main_data, sub_data)
            path_receipt_2 = two_page_2_receipt(file_name, main_data, sub_data)

            path_merge_invoice = merge_pdfs(path_invoice_1, path_invoice_2, file_name_invoice)
            path_merge_receipt = merge_pdfs(path_receipt_1, path_receipt_2, file_name_receipt)

            merge_pdfs(path_merge_invoice, path_merge_receipt, file_name)
            
            os.remove(path_merge_invoice)
            os.remove(path_merge_receipt)
            os.remove(path_invoice_1)
            os.remove(path_invoice_2)
            os.remove(path_receipt_1)
            os.remove(path_receipt_2)

        elif len(sub_data) > 30 and len(sub_data) <= 45:

            file_name_invoice = file_name + "_invoice"
            file_name_receipt = file_name + "_receipt"

            path_invoice_1 = three_page_1_invoice(file_name, main_data, sub_data)
            path_invoice_2 = three_page_2_invoice(file_name, main_data, sub_data)
            path_invoice_3 = three_page_3_invoice(file_name, main_data, sub_data)
            path_receipt_1 = three_page_1_receipt(file_name, main_data, sub_data)
            path_receipt_2 = three_page_2_receipt(file_name, main_data, sub_data)
            path_receipt_3 = three_page_3_receipt(file_name, main_data, sub_data)

            path_merge_invoice = merge_3pdfs(path_invoice_1, path_invoice_2, path_invoice_3, file_name_invoice)
            path_merge_receipt = merge_3pdfs(path_receipt_1, path_receipt_2, path_receipt_3, file_name_receipt)

            merge_pdfs(path_merge_invoice, path_merge_receipt, file_name)

            os.remove(path_merge_invoice)
            os.remove(path_merge_receipt)
            os.remove(path_invoice_1)
            os.remove(path_invoice_2)
            os.remove(path_invoice_3)
            os.remove(path_receipt_1)
            os.remove(path_receipt_2)
            os.remove(path_receipt_3)

        else:
            print("contact to jel ")