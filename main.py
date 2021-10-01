# Name: Unsubscribe-me
# Date: 30-Sep-2021
# Author: Luis Lopez-Echeto
# Version: 0.1
import getpass
import requests
from imap_tools import MailBox


debug = True

default_providers = [
    {'provider_name':'Office365',
     'server':'outlook.office365.com',
     'port':'993',
     'help':'Username is your full email address. \nPassword is your Office365 password.'
     }
    ]

class UnsubscribeMe:
    credentials = {'user': None, 'password': None}
    imap_server = default_providers[0]['server']
    imap_port = default_providers[0]['port']
    mail_box = None
    logged_in = False
    
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
        self.credentials['password'] =getpass.getpass('Password: ')
        
    def get_emails(self):
        '''Search for email messages in selected folder that have the unsubscribe word in their body.'''
        messages = self.mail_box.fetch('TEXT "unsubscribe"')
        
        for message in messages:
            print(message.from_, ': ', message.subject)

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

if __name__ == '__main__':
    app = UnsubscribeMe()
    app.run()

            
