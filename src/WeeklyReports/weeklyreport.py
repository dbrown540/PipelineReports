import os
import time
import environ
import tkinter as tk
from tkinter import messagebox
from datetime import date
import pandas as pd
from pandas import DataFrame
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

from .webdriver import WebDriverManager

# Load environment variables
env = environ.Env()
environ.Env().read_env()
USERNAME = env("sharepoint_email")
APP_PASSWORD = env("app_password")
SHAREPOINT_SITE = env("sharepoint_url_site")
SHAREPOINT_SITE_NAME = env("sharepoint_site_name")
SHAREPOINT_DOC_LIBRARY = env("sharepoint_doc_library")


class TechnoMileBot(WebDriverManager):
    def open_tm(self, url="https://reli-group.lightning.force.com/"):
        self.driver.get(url)

    def type_username(self, username):
        try:
            username_input = self.wait.until(EC.presence_of_element_located((By.ID, "username")))
            username_input.send_keys(username)
        except Exception as e:
            print(f"Could not find the username box: {e}")

    def type_password(self, password):
        try:
            password_input = self.wait.until(EC.presence_of_element_located((By.ID, "password")))
            password_input.send_keys(password)
        except Exception as e:
            print(f"Could not find the password box: {e}")

    def click_login(self):
        try:
            login_button = self.wait.until(EC.element_to_be_clickable((By.NAME, "Login")))
            login_button.click()
        except Exception as e:
            print(f"Could not find login button or it is not clickable: {e}")

    def handle_mfa(self):
        # Create a popup window to remind the user about completing MFA
        root = tk.Tk()
        root.withdraw()  # Hide the root window
        messagebox.showinfo("Reminder", "Have you completed the MFA?")
        root.destroy()

    def open_daily_pipeline_report(self):
        self.driver.get("https://reli-group.lightning.force.com/lightning/r/sObject/00O8Z000007CWhgUAG/view?queryScope=userFolders")

    def switch_to_iframe(self):
        try:
            iframe = self.wait.until(EC.presence_of_element_located((By.TAG_NAME, 'iframe')))
            self.driver.switch_to.frame(iframe)
            print("Switched to the iframe.")
        except Exception as e:
            print(f"Could not switch to the iframe: {e}")

    def click_dropdown(self):
        try:
            self.switch_to_iframe()
            dropdown_button = self.wait.until(EC.presence_of_element_located((By.XPATH, "//button[@tabindex='0' and contains(@class, 'more-actions-button')]")))
            dropdown_button.click()
            self.driver.switch_to.default_content()
        except Exception as e:
            print(f"Could not find dropdown menu button or it is not clickable: {e}")

    def click_export_dropdown(self):
        try:
            self.switch_to_iframe()
            export_button = self.wait.until(EC.presence_of_element_located((By.XPATH, "//a[@data-index='3' and @role='menuitem']")))
            export_button.click()
            self.driver.switch_to.default_content()
        except Exception as e:
            print(f"Could not click export: {e}")

    def click_details_only(self):
        try:
            self.driver.switch_to.default_content()
            details_only_buttons = self.wait.until(EC.presence_of_all_elements_located((By.XPATH, "//span[text()='Details Only']")))
            details_only_buttons[0].click()
            time.sleep(4)
        except Exception as e:
            print(f"Could not click 'Details Only': {e}")

    def select_excel_format(self):
        try:
            select_element = self.wait.until(EC.presence_of_element_located((By.TAG_NAME, "select")))
            select_object = Select(select_element)
            select_object.select_by_value("xlsx")  # Select the option with value 'xlsx'
            print("Excel format (.xlsx) selected")
        except Exception as e:
            print(f"Failed to select Excel format: {e}")

    def get_latest_downloaded_file(self, download_folder=None):
        if download_folder is None:
            download_folder = os.path.join(os.path.expanduser("~"), "Downloads")

        try:
            # List all files in the download folder and filter out directories
            file_paths = [os.path.join(download_folder, file) for file in os.listdir(download_folder) if os.path.isfile(os.path.join(download_folder, file))]
            if not file_paths:
                raise FileNotFoundError("No files found in the download folder.")

            # Sort files by modification time, descending order
            file_paths.sort(key=os.path.getmtime, reverse=True)
            latest_file = file_paths[0]
            print(latest_file)
            return latest_file
        except Exception as e:
            print(f"Error retrieving the latest downloaded file: {e}")
            return None
        
    def click_export(self):
        try:
            export_button = self.wait.until(EC.presence_of_element_located((
                By.XPATH, "//button[@title='Export']"
            )))
            if export_button:
                export_button.click()
                time.sleep(3)

        except Exception as e:
            print(f"Could not find the export button: {e}")

    def execute(self):
        self.open_tm()
        self.type_username("daniel.brown@religroupinc.com")
        self.type_password("E^piiisnegative1")
        self.click_login()
        self.handle_mfa()
        self.open_daily_pipeline_report()
        self.click_dropdown()
        self.click_export_dropdown()
        self.click_details_only()
        self.select_excel_format()
        self.click_export()
        return self.get_latest_downloaded_file()


class SharePointBot(WebDriverManager):
    def _auth(self):
        conn = ClientContext(SHAREPOINT_SITE).with_credentials(UserCredential(USERNAME, APP_PASSWORD))
        try:
            web = conn.web.get().execute_query()  # Attempt to get the SharePoint web
            print(f"Authentication successful! Connected to SharePoint site: {web.properties['Title']}")
            return conn
        except Exception as e:
            print(f"Authentication failed: {e}")
            return None

    def upload_file(self, conn, local_file_path, sharepoint_folder_url, file_name):
        try:
            with open(local_file_path, "rb") as file_content:
                target_file_url = f"{sharepoint_folder_url}/{file_name}"
                conn.web.get_folder_by_server_relative_url(sharepoint_folder_url).upload_file(target_file_url, file_content).execute_query()
                print(f"File '{file_name}' uploaded successfully to SharePoint.")
        except Exception as e:
            print(f"Failed to upload file: {e}")


class WeeklyPipelineReportCreator(WebDriverManager):
    def fetch_data(self, tm_filepath):
        df_tm = pd.read_excel(tm_filepath)
        print("Technomile file:\n", df_tm.head())
        print("Num of rows in TM download", len(df_tm.index))

        return df_tm

    def replace_data(self, df_tm: DataFrame):
        # Replace the SP Weekly Pipeline report data with the Technomile Data
        df_tm.to_excel("Weekly Pipeline Report.xlsx", index=None)

    def get_date(self):
        return date.today().strftime("%#m/%#d/%Y").lstrip('0')

    def send_email(self, smtp_server="smtp.office365.com", smtp_port=587, 
                   username="daniel.brown@religroupinc.com", 
                   to_email=["david.blackwell@religroupinc.com", "daniel.brown@religroupinc.com"], 
                   body="This is an automated message on behalf of David Blackwell's Market Research Team.\n\nAttached below is this week's Weekly Pipeline Report.\n\nhttps://religroupinc.sharepoint.com/:x:/r/sites/RELIBusinessDevelopment/Shared%20Documents/BD%20Department%20Files/Pipeline/Weekly%20Pipeline%20Report.xlsx?d=wc0516468052840faa34453866ce12e84&csf=1&web=1&e=TYJby6\n\nThanks,\nMr. Bot"):
        try:
            # Generate subject based on the current date
            current_date = self.get_date()
            subject = f"Weekly Pipeline Report for the Week of {current_date}"

            # Create a MIMEMultipart message
            msg = MIMEMultipart()
            msg['From'] = username
            msg['To'] = ', '.join(to_email)
            msg['Subject'] = subject

            # Create the HTML body with Times New Roman, 12pt font
            html_body = f"""
            <html>
                <head>
                    <style>
                        body {{
                            font-family: 'Times New Roman', serif; 
                            font-size: 12pt;
                        }}
                    </style>
                </head>
                <body>
                    <p>This is an automated message on behalf of David Blackwell's Market Research Team.</p>
                    <p>Attached below is this week's Weekly Pipeline Report.</p>
                    <p><a href="https://religroupinc.sharepoint.com/:x:/r/sites/RELIBusinessDevelopment/Shared%20Documents/BD%20Department%20Files/Pipeline/Weekly%20Pipeline%20Report.xlsx?d=wc0516468052840faa34453866ce12e84&csf=1&web=1&e=TYJby6">Weekly Pipeline Report</a></p>
                    <p>Thanks,<br>Mr. Bot</p>
                </body>
            </html>
            """

            # Attach the HTML body to the message
            msg.attach(MIMEText(html_body, 'html'))

            # Connect to the SMTP server
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()  # Upgrade to a secure connection
                server.login(username, APP_PASSWORD)  # Login to your email account
                server.send_message(msg)  # Send the email

            print("Email sent successfully!")
        except Exception as e:
            print(f"Failed to send email: {e}")

    def execute(self, tm_filepath):
        df_tm = self.fetch_data(tm_filepath)  # Converts xlsx into dataframes
        self.replace_data(df_tm)
        self.send_email()
