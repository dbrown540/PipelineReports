import time
import tkinter as tk
from tkinter import messagebox

from .webdriver import WebDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

class TechnoMileBot(WebDriverManager):
    def open_tm(self):
        self.driver.get("https://reli-group.lightning.force.com/")
        
    def type_username(self, username):
        try: 
            username_input = self.wait.until(EC.presence_of_element_located((
                By.ID, "username"
            )))
            username_input.send_keys(username)
        except:
            print("Could not find the username box")

    def type_password(self, password):
        try: 
            password_input = self.wait.until(EC.presence_of_element_located((
                By.ID, "password"
            )))
            password_input.send_keys(password)
        except:
            print("Could not find the password box")

    def click_login(self):
        try:
            login_button = self.wait.until(EC.element_to_be_clickable((
                By.NAME, "Login"
            )))
            login_button.click()
        except:
            print("Could not find login button or it is not clickable")

    def handle_mfa(self):
        # Create a popup window to remind the user about completing MFA
        root = tk.Tk()
        root.withdraw()  # Hide the root window
        messagebox.showinfo("Reminder", "Have you completed the MFA?")
        root.destroy()

    def open_daily_pipeline_report(self):
        self.driver.get("https://reli-group.lightning.force.com/lightning/r/sObject/00O8Z000007CWhgUAG/view?queryScope=userFolders")

    def click_dropdown(self):
        try:
            dropdown_button = self.wait.until(EC.element_to_be_clickable((
                By.CLASS_NAME, "more-actions-button"
            )))
            dropdown_button.click()
        except:
            print("Could not find dropdown menu button or it is not clickable")
