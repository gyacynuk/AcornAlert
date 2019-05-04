import json
import requests
import sched
import time
import os
import sys
from datetime import datetime
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys

CONFIG_PATH = 'config.json'
BOT_NAME = 'Acorn Bot'


def perror(*args, **kwargs):
    print(*args, file=sys.stderr, **kwargs)


class AcornAlert:
    def __init__(self):
        # Read in config file
        if not os.path.isfile(CONFIG_PATH):
            perror('Error: No config file found. Make sure config.json is in the same directory as the launcher.')
            exit(1)
        else:
            with open('config.json', 'r') as f:
                self.config = json.load(f)

        self.password = ''
        self.driver = None
        self.scheduler = sched.scheduler(time.time, time.sleep)
        self.initialize_config()

        # Start Scheduler
        self.scheduler.enter(1, 1, self.check_grades)
        self.scheduler.run()

    def initialize_config(self):
        # Get UTor ID if not already in config file
        if 'username' not in self.config:
            self.config['username'] = input('UTor ID: ')
        else:
            print('Found user: {}'.format(self.config['username']))
            new_username = input('Hit enter to continue, or enter a new UTor ID: ')
            if new_username != '':
                self.config['username'] = new_username

        # Get password
        self.password = input('Password (will not be saved for security purposes): ')

        # Get email
        if 'mailing_list' not in self.config:
            self.config['mailing_list'] = [input('Email which you would like to receive updates at: ')]

        # Edit monitored courses
        while True:
            print('Currently monitoring the following courses:')
            self.print_monitoring()
            if input('If you would like to edit this, press "e". Otherwise hit any other key: ') != 'e':
                break

            print('To automatically fetch your IPR courses, enter "auto".\n'
                  'To add courses by code, enter "add".\n'
                  'To remove courses, enter "remove".')
            choice = input('')
            if choice == 'auto':
                self.start_monitoring(*self.auto_find_ipr())
            elif choice == 'add':
                courses = input('Enter course codes separated by spaces '
                                '(include suffixes such as "H1" or "S1"): ').upper()
                self.start_monitoring(*courses.split())
            elif choice == 'remove':
                numbers = input('Enter the corresponding numbers of the courses you want to remove, '
                                'separated by spaces: ')
                numbers = [int(num) for num in numbers.split()]
                self.stop_monitoring(*numbers)
            else:
                print('"{}" is not a valid option. Please check your spelling and try again.'.format(choice))

        # Edit polling interval
        while True:
            print('The bot will check for new grades every {} minutes. Hit enter to keep this setting.'.format(
                self.config['poll_interval']//60))
            interval = input('To change this, enter a new interval in minutes ("60" is recommended): ')
            if interval == '':
                break

            # Ensure interval is valid
            try:
                interval = int(interval)
            except Exception:
                print('Invalid interval. Please enter a positive number.')
                continue
            if interval > 0:
                self.config['poll_interval'] = interval * 60
                break
            else:
                print('Invalid interval. Please enter a positive number.')

        # Save any changes made
        self.update_config()

    def print_monitoring(self):
        if len(self.config['monitoring']) == 0:
            print('\tNone')
        else:
            for num, course in enumerate(self.config['monitoring']):
                print('\t{}. {}'.format(num, course))

    def stop_monitoring(self, *numbers):
        numbers = list(numbers)
        numbers.sort(reverse=True)
        for num in numbers:
            if num < len(self.config['monitoring']):
                self.config['monitoring'].pop(num)

    def stop_monitoring_by_code(self, *codes):
        for code in codes:
            if code in self.config['monitoring']:
                self.config['monitoring'].remove(code)

    def start_monitoring(self, *courses):
        for course in courses:
            if course not in self.config['monitoring']:
                self.config['monitoring'].append(course)

    def update_config(self):
        with open('config.json', 'w') as f:
            json.dump(self.config, f)

    def send_email(self, grades):
        email_body = '\n'.join(
            ["{} - {} ({}) found at {:%Y-%m-%d %H:%M}".format(c, g, m, datetime.now()) for c, g, m in grades])
        return requests.post(
            "https://api.mailgun.net/v3/{0}/messages".format(self.config['mailgun']['domain']),
            auth=("api", self.config['mailgun']['api_key']),
            data={"from": "{} <mailgun@{}>".format(BOT_NAME, self.config['mailgun']['domain']),
                  "to": self.config['mailing_list'],
                  "subject": "Acorn Alert - {0} New Grades".format(len(grades)),
                  "text": email_body})

    def auto_find_ipr(self):
        courses = []
        self.login()

        # Find current semester transcript section
        try:
            elem = self.driver.find_element_by_id('status1')
        except NoSuchElementException:
            elem = self.driver.find_element_by_id('status0')

        # Find IPR courses
        for row in elem.find_elements_by_xpath('.//*'):
            if row.get_attribute('class') == 'courses':
                children = row.find_elements_by_xpath('.//*')
                course, grade = children[0].text, children[4].text
                if grade == 'IPR':
                    courses.append(course)

        # Logout and return
        self.logout()
        return courses

    def login(self):
        # Initialize Driver
        self.driver = webdriver.Chrome(ChromeDriverManager().install())
        self.driver.get(self.config['acorn_url'])

        # Log in
        elem = self.driver.find_element_by_id('username')
        elem.send_keys(self.config['username'])
        elem = self.driver.find_element_by_id('password')
        elem.send_keys(self.password)
        elem.send_keys(Keys.RETURN)

        # Navigate to academic history
        elem = self.driver.find_element_by_link_text('Academic History')
        elem.click()

    def logout(self):
        self.driver.close()

    def check_grades(self):
        updates = []
        self.login()

        # Find current semester transcript section
        try:
            elem = self.driver.find_element_by_id('status1')
        except NoSuchElementException:
            elem = self.driver.find_element_by_id('status0')

        # Read in grades
        for row in elem.find_elements_by_xpath('.//*'):
            if row.get_attribute('class') == 'courses':
                children = row.find_elements_by_xpath('.//*')
                course, grade, mark = children[0].text, children[4].text, children[3].text
                if course in self.config['monitoring'] and grade != 'IPR':
                    updates.append((course, grade, mark))
                    self.stop_monitoring_by_code(course)

        # Logout
        self.logout()

        # Save any changes
        if len(updates) > 0:
            self.update_config()

        # Send email notification
        self.send_email(updates)

        # Reschedule grade checking
        self.scheduler.enter(self.config['poll_interval'], 1, self.check_grades)


if __name__ == '__main__':
    AcornAlert()