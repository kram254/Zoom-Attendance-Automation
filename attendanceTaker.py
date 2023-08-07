import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.firefox import GeckoDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import csv
from datetime import datetime
import PySimpleGUI as sg

class Zoom:
    def __init__(self, selected_driver, meeting_personal_id, username, meeting_password):
        self.meeting_personal_id = meeting_personal_id
        self.selected_driver = selected_driver
        self.username = username
        self.meeting_password = meeting_password
        self.driver = None

    def setup_driver(self):
        if self.selected_driver == "Chrome":
            self.driver = webdriver.Chrome(ChromeDriverManager().install())
        elif self.selected_driver == "FireFox":
            self.driver = webdriver.Firefox(GeckoDriverManager().install())
        elif self.selected_driver == "Edge":
            self.driver = webdriver.Edge(EdgeChromiumDriverManager().install())

    def join_meeting(self):
        self.setup_driver()
        try:
            self.driver.get("https://zoom.us/join")
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "inputname"))
            ).send_keys(self.username)
            self.driver.find_element(By.ID, "inputpasscode").send_keys(self.meeting_password)
            self.driver.find_element(By.ID, "joinBtn").click()

        except NoSuchElementException as e:
            print("Error: Element not found -", e)
        except TimeoutException as e:
            print("Error: Timeout waiting for element -", e)
        except Exception as e:
            print("Error:", e)
        finally:
            self.driver.quit()

    def get_attendees_list(self):
        try:
            WebDriverWait(self.driver, 1000).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="wc-container-left"]/div[3]/div/div[2]/div/div/div[1]/div')
                )
            )
            attendees = []
            i = 1
            while True:
                elem = self.driver.find_element_by_xpath(
                    f'//*[@id="wc-container-left"]/div[3]/div/div[2]/div/div/div[1]/div[{i}]/div'
                )
                attendees.append(elem.text)
                i += 1
        except NoSuchElementException:
            pass
        print(attendees)
        with open("attended.csv", "a") as fp:
            fp.write(f"{datetime.now().isoformat(' ')}\n")
            for attendee in attendees:
                fp.write(f"{attendee}\n")

class GUI:
    def __init__(self):
        self.drivers = ["Chrome", "FireFox", "Edge"]
        self.layout = [
            [sg.Image(filename="rounded_window.png")],
            [sg.Text("Zoom Attendance Check", font=("Helvetica", 25), justification="center")],
            [sg.Text("Meeting ID: "), sg.Input(key="meetingID", size=(20, 1))],
            [sg.Text("Meeting Password"), sg.Input(password_char="*", key="password", size=(20, 1))],
            [sg.Text("Your Name"), sg.Input(key="name", size=(20, 1))],
            [sg.Text("Select Browser Driver:", justification="center")],
            [sg.Listbox(self.drivers, size=(20, 3), key="Driver", default_values="Chrome", pad=(40, 10))],
            [sg.Button("Join Meeting", key="join", size=(15, 2)),
             sg.Button("Get Attendees", key="get_attendees", size=(15, 2))]
        ]

    def create(self):
        sg.theme("LightGrey1")
        self.window = sg.Window("Zoom Attendance", self.layout, element_justification="center",
                                finalize=True, no_titlebar=True, keep_on_top=True)

    def run(self):
        while True:
            event, values = self.window.read()
            if event == sg.WIN_CLOSED or event == "Exit":
                break
            elif event == "join":
                zoom_instance = Zoom(
                    selected_driver=values["Driver"][0],
                    meeting_personal_id=values["meetingID"],
                    username=values["name"],
                    meeting_password=values["password"]
                )
                zoom_instance.join_meeting()
                sg.popup("You have joined the meeting!")
            elif event == "get_attendees":
                zoom_instance = Zoom(
                    selected_driver=values["Driver"][0],
                    meeting_personal_id=values["meetingID"],
                    username=values["name"],
                    meeting_password=values["password"]
                )
                zoom_instance.get_attendees_list()
                sg.popup("Attendees list has been fetched and saved to attended.csv!")

        self.window.close()

if __name__ == "__main__":
    gui = GUI()
    gui.create()
    gui.run()
