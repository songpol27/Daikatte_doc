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

class Opreate_os:
    def __init__(self):
        self.path_all_folder = "source/all_folder"
        self.path_dn_folder = "source/dn_folder"
        self.path_po_folder = "source/po_folder"
        self.path_inv_folder = "source/inv_folder"

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
