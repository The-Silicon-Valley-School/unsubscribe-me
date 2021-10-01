# Name: Unsubscribe-me
# Date: 30-Sep-2021
# Author: Luis Lopez-Echeto
# Version: 0.1
import getpass
import re
import requests
from bs4 import BeautifulSoup
from imap_tools import MailBox
import webbrowser

debug = False

default_providers = [
    {'provider_name':'Office365',
     'server':'outlook.office365.com',
     'port':'993',
     'help':'Username is your full email address. \nPassword is your Office365 password.'
     }
    ]

unsubscribe_words = ['unsubscribe','subscription','optout']

class UnsubscribeMe:
    
    def __init__(self):
        self.credentials = {'user': None, 'password': None}
        self.imap_server = default_providers[0]['server']
        self.imap_port = default_providers[0]['port']
        self.mail_box = None
        self.logged_in = False
        self.words_to_check = []
        for i in range(len(unsubscribe_words)):
            self.words_to_check.append(re.compile(unsubscribe_words[i], re.I))
        self.senders = []
        self.open_links = False
    
    def get_server(self):
        '''Asks user to select provider from preselected list of providers available.'''
        print('Please choose your provider.')
        for pos, provider in enumerate(default_providers):
            print('[{}] {}'.format(pos, provider['provider_name']))
        provider_choice = input('>>> ')
        if provider_choice.isdigit() and int(provider_choice) < len(default_providers):
            self.imap_server = default_providers[int(provider_choice)]['server']
            self.imap_port = default_providers[int(provider_choice)]['port']
        print(self.imap_server)
    
    def login(self):
        '''Connect to the provider IMAP server and attempt to log in.'''
        self.mail_box = MailBox(self.imap_server).login(self.credentials['user'], self.credentials['password'])
        self.logged_in = True

    def logout(self):
        self.mail_box.logout()

    def get_credentials(self):
        '''Get credentials for IMAP server.'''
        self.credentials['user'] = input('User: ')
        if debug:
            self.credentials['password'] = input('Password: ')
        else:
            self.credentials['password'] = getpass.getpass('Password: ')
        
    def get_emails(self):
        '''Search for email messages in selected folder that have the unsubscribe word in their body.'''
        messages = self.mail_box.fetch('TEXT "unsubscribe"')
        
        for message in messages:
            print('#',end='')
            body_html = message.html
            soup = BeautifulSoup(body_html, 'html.parser')
            unsub_url = None
            elements = soup.select('a')
            # For each anchor tag search for possible unsubscribe words on text
            for element in elements:
                print('.', end='')
                for j in range(len(self.words_to_check)):
                    found = self.words_to_check[j].search(element.text)
                    if found != None:
                        unsub_url = element.get('href')
                        break
                if unsub_url != None:
                    break
            if unsub_url != False:
                self.senders.append([message.from_, unsub_url, False, False])
        print()
    
    def display_list_of_senders(self):
        if self.senders != []:
            print('Found the following senders with unsubscribe links:')
            for sender in self.senders:
                print(sender[0], end=' | ')
            print()
    
    def select_links(self):
        '''User selects what unsubscribe links to open.'''
        self.display_list_of_senders()
        for sender_idx, sender in enumerate(self.senders):
            while True:
                open_choice = input('Open link from {} (Y/N): [Y] '.format(sender[0])).lower()
                if open_choice == 'y' or open_choice == '':
                    self.senders[sender_idx][2] = True
                    self.open_links = True
                    break
                else:
                    sender[2] = False
                    break
    
    def open_selected_links(self):
        '''Open up to ten links at a time in the browser.'''
        if self.open_links != True:
            print('No links selected to open.')
        else:
            counter = 0
            for i in range(len(self.senders)):
                if self.senders[i][2] == True:
                    webbrowser.open(self.senders[i][1])
                    counter += 1
                    if counter == 10 or i == len(self.senders) - 1:
                        print('Navigating to links.')
                        input('Press enter to continue.')
                        counter = 0
        
    def get_access(self):
        '''Subroutine to get access to the mail server by asking server choice and credentials from user first.'''
        self.get_server()
        self.get_credentials()
        self.login()
    
    def run(self):
        '''Main function to run the application.'''
        self.get_access()
        if self.logged_in:
            self.get_emails()
            if self.senders != []:
                self.select_links()
                self.open_selected_links()

if __name__ == '__main__':
    app = UnsubscribeMe()
    app.run()

            
