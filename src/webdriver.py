import logging

from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import WebDriverException


class WebDriverManager:
    """
    WebDriverManager provides functionality for managing a 
    WebDriver instance for automated browser testing.

    This class encapsulates methods for initializing and 
    closing a WebDriver instance, with options for
    specifying WebDriver options and service. 
    It integrates with the Selenium library to interact with web browsers.

    Attributes:
        driver_path (str): 
            The path to the WebDriver executable.
        driver (WebDriver): 
            The WebDriver instance initialized by the class.
        wait (WebDriverWait): 
            A WebDriverWait instance associated with the WebDriver instance.

    Methods:
        __init__(): 
            Initializes a WebDriverManager instance 
            and sets up a WebDriver instance with specified options.
        initialize_webdriver(): 
            Initializes a WebDriver instance with specified options and service.
        close_webdriver(): 
            Closes the WebDriver instance associated with the object.

    """

    def __init__(self):
        self.driver_path = 'chromedriver-win64/chromedriver.exe'
        self.driver = self.initialize_webdriver()
        self.wait = WebDriverWait(self.driver, 10)

    def initialize_webdriver(self) -> WebDriver:
        """
        Initializes a WebDriver instance with specified options and service.

        Returns:
            WebDriver: A WebDriver instance with specified options and service.

        Raises:
            WebDriverException: If there is an issue initializing the WebDriver.
            Exception: If an unexpected error occurs while initializing the WebDriver.
        """
        try:
            options = webdriver.ChromeOptions()
            options.add_argument('--disable-extensions')
            options.add_argument("--start-maximized")
            # Add capability to disable infobars
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option('useAutomationExtension', False)
            service = Service(self.driver_path)
            driver = webdriver.Chrome(service=service, options=options)

            # Log the webdriver information
            info_message = (
                f"Service path: {self.driver_path}\nOptions: {options}"
            )
            logging.info(info_message)

            return driver

        except WebDriverException:
            logging.critical('Failed to initialize WebDriver: ', exc_info=True)
            raise

        except Exception:
            # For other exceptions, log the error and provide a generic error message
            error_message = (
                'An unexpected error occurred while initializing WebDriver: '
            )
            logging.critical(error_message, exc_info=True)
            raise

    def close_webdriver(self) -> None:
        """
        Closes the webdriver.

        This method attempts to close the webdriver instance associated with the object. 
        If a webdriver instance exists, it will be closed. 
        If no webdriver instance is found, it prints a message indicating 
        the absence of any instance to close.

        Raises:
            Exception: If an error occurs while attempting to 
                close the webdriver, an exception is raised. This could
                include errors such as WebDriverException or any other 
                exceptions related to webdriver operations.
                The specific error message can be retrieved from the exception object.
        """
        if self.driver:
            self.driver.close()
            logging.info("Webdriver closed successfully.")
        else:
            logging.info("No webdriver instance found to close.")

