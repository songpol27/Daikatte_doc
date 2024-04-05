import gspread, configparser, logging
import pandas as pd
from googleapiclient.discovery import build
from google.oauth2 import service_account
#gspread -> https://docs.gspread.org/en/v6.0.0/user-guide.html
class manager_google_sheet_data:

    def __init__(self, config_file='config.ini'):
        self.config_file = config_file
        self.setup_logging()
        self.read_config()
        
        try:
            
            sa = gspread.service_account(filename=self.credentials)
            self.sheets = sa.open(self.file_google_sheets)
            
        except Exception as e:
            self.logger.error(f"Error while authenticating with Google Sheets: {e}")
            raise

        self.wks_DN = self.sheets.worksheet("DN")
        self.wks_DN_Items = self.sheets.worksheet("DN_Items")
        self.wks_Company = self.sheets.worksheet("Company")
        self.wks_Products = self.sheets.worksheet("Products")

        self.wks_INV = self.sheets.worksheet("INV")
        self.wks_INV_Items = self.sheets.worksheet("INV_Items")     

        self.read_static_data()

    def setup_logging(self):
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.DEBUG)
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

        fh = logging.FileHandler('error.log')
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(formatter)
        self.logger.addHandler(fh)

        ch = logging.StreamHandler()
        ch.setLevel(logging.INFO)
        ch.setFormatter(formatter)
        self.logger.addHandler(ch)

    def authenticate(self):
        scopes = ['https://www.googleapis.com/auth/drive']
        creds = service_account.Credentials.from_service_account_file(self.credentials, scopes=scopes)
        return creds

    def read_config(self):
        config = configparser.ConfigParser()
        try:
            config.read(self.config_file)
            self.file_google_sheets = config.get('config', 'file_name')
            self.credentials = config.get('config', 'path_credentials')
            self.google_drive_id_po = config.get('config', 'google_drive_id_po')
            self.google_drive_id_invoice = config.get('config', 'google_drive_id_invoice')

        except (configparser.NoSectionError, configparser.NoOptionError, FileNotFoundError) as e:
            self.logger.error(f"Error reading configuration: {e}")
            raise

    def read_static_data(self):
        try:
            raw_data_company = self.wks_Company.get_all_values()
            raw_data_products = self.wks_Products.get_all_values()

            df_company = pd.DataFrame(raw_data_company[1:], columns=raw_data_company[0])
            df_products = pd.DataFrame(raw_data_products[1:], columns=raw_data_products[0])

            return df_company, df_products
        except Exception as e:
            self.logger.error(f"Error reading data from Google Sheets: {e}")
            raise
        
    def read_dynamic_data(self):
        try:
            raw_data_daily_note = self.wks_DN.get_all_values()
            raw_data_daily_note_item = self.wks_DN_Items.get_all_values()

            df_daily_note = pd.DataFrame(raw_data_daily_note[1:], columns=raw_data_daily_note[0])
            df_aily_note_item = pd.DataFrame(raw_data_daily_note_item[1:], columns=raw_data_daily_note_item[0])

            return df_daily_note, df_aily_note_item

        except Exception as e:
            self.logger.error(f"Error reading data from Google Sheets: {e}")
            raise

    def read_invoice(self):
        try:
            raw_data_INV = self.wks_INV.get_all_values()
            raw_data_INV_Items = self.wks_INV_Items.get_all_values()

            df_invoice = pd.DataFrame(raw_data_INV[1:], columns=raw_data_INV[0])
            df_invoice_itemes = pd.DataFrame(raw_data_INV_Items[1:], columns=raw_data_INV_Items[0])

            return df_invoice, df_invoice_itemes
        except Exception as e:
            self.logger.error(f"Error reading data from Google Sheets: {e}")
            raise

    def upload_po_file(self, file_path, namefile):

        creds = self.authenticate()
        service = build('drive', 'v3', credentials=creds)

        file_metadata = {
            'name': namefile,
            'parents': [self.google_drive_id_po]
        }

        try:
            file = service.files().create(
                body=file_metadata,
                media_body=file_path
            ).execute()
            
            file_id = file.get('id')  # Extracting the file ID from the response
            
            return file_id

        except Exception as e:
            self.logger.error(f"Error uploading file to Google Drive: {e}")
            raise
    
    def upload_iv_file(self, file_path, namefile):
        creds = self.authenticate()
        service = build('drive', 'v3', credentials=creds)

        file_metadata = {
            'name': namefile,
            'parents': [self.google_drive_id_invoice]
        }

        try:
            file = service.files().create(
                body=file_metadata,
                media_body=file_path
            ).execute()
            
            file_id = file.get('id')  # Extracting the file ID from the response
            
            return file_id

        except Exception as e:
            self.logger.error(f"Error uploading file to Google Drive: {e}")
            raise

    def add_data_daily_note(self ,id_dn, dn_no, date, id_company, vat, id_iv, id_po):

        number_DN_ID_row = len(self.wks_DN.col_values(1))
        number_addnew_row = number_DN_ID_row + 1 
        self.wks_DN.update_cell(number_addnew_row, 1, id_dn)
        self.wks_DN.update_cell(number_addnew_row, 2, dn_no)
        self.wks_DN.update_cell(number_addnew_row, 3, date)
        self.wks_DN.update_cell(number_addnew_row, 4, id_company)
        self.wks_DN.update_cell(number_addnew_row, 5, vat)
        self.wks_DN.update_cell(number_addnew_row, 6, id_iv)
        self.wks_DN.update_cell(number_addnew_row, 7, id_po)

    def add_data_daily_note_items(self, id_dn, id_product, qty, price_unit):
        number_DN_ID_row = len(self.wks_DN_Items.col_values(1))
        number_addnew_row = number_DN_ID_row + 1 
        #self.wks_DN_Items.update_cell(number_addnew_row, 1, id_dn_item)
        self.wks_DN_Items.update_cell(number_addnew_row, 2, id_dn)
        self.wks_DN_Items.update_cell(number_addnew_row, 3, id_product)
        self.wks_DN_Items.update_cell(number_addnew_row, 4, qty)
        self.wks_DN_Items.update_cell(number_addnew_row, 5, price_unit)


