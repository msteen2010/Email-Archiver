# This script allows for emails to be read from Outlook and stored to the file system for future access
# Created by Mike van der Steen
# Version 1.01
# Last updated 11 June 2022

import glob
import win32com.client
import win32timezone
import win32api
import win32ui
import sys
import os
import re
import time
from datetime import datetime, timedelta
import configparser
import PySimpleGUI as sg
import logging

# List of folders to exclude from Outlook
folder_exclusion_list = [
    'Deleted Items',
    'Outbox',
    'Junk Email',
    'Drafts',
    'Conversation History',
    'Calendar',
    'Contacts',
    'Yammer Root',
    'Sync Issues',
    'Scheduled',
    'Quick Step Settings',
    'PersonMetadata',
    'RSS Subscriptions',
    'MeContact',
    'Archive',
    'Files',
    'Notes',
    'Conversation Action Settings',
    'Tasks',
    'Journal',
    'ExternalContacts',
    'News Feed',
    'Social Activity Notifications',
    'Suggested Contacts',
    'Tools'
]
# Get the current working directory location
directory = os.getcwd()
working_directory = os.path.join(directory, 'Emails')
settings_file_present = False

# Set the logging parameters
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s:%(message)s')
file_handler = logging.FileHandler('email-archiver.log', mode='w', delay=False)
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)

# Check to see if the output director exists, if not create it
if os.path.exists(working_directory):
    logger.info(f'The directory where emails will be stored: {working_directory}')

else:
    os.mkdir(working_directory)
    logger.info('The directory where emails will be stored: {working_directory}')

# Get the maximum age of emails from the settings.ini file
config = configparser.ConfigParser()
try:
    config.read('settings.ini')
    email_max_days = int((config['Settings']['EmailMaxAgeDays']))
    logger.info('Successfully read settings.ini file')
    logger.info('Email max age as defined in the settings.ini file is ' + str(email_max_days) + ' days')
    delete_old_emails = config['Settings']['deleteOldEmails']
    logger.info('Delete old emails as defined in the settings.ini file is set to ' + str(delete_old_emails))
    settings_file_present = True
except Exception as e:
    logger.info('No settings.ini file was present. Terminating the application.')
    settings_file_present = False
    logger.error(e, exc_info=True)


def outlook_running():
    # Check to see if Outlook is running and if not, open it
    try:
        win32ui.FindWindow(None, "Microsoft Outlook")
    except:
        os.startfile("outlook")

def gui():
    # Defining the size and properties of the window
    layout = [
            [sg.Multiline(size=(125, 40), key='-LOG-', autoscroll=True, auto_refresh=True)],
            [sg.Button('Start', bind_return_key=True, key='-START-'), sg.Button('Exit')]
        ]

    window = sg.Window('Log window', layout,
            default_element_size=(30, 2),
            font=('Helvetica', ' 10'),
            default_button_element_size=(8, 2),)

    return window


def gui_status(window, msg):
    # Loop taking in user input and querying queue
    while True:
    # Wake every 100ms and look for work
        event, values = window.read(timeout=100)

        info = 'This is an email archiver application that copies emails from Outlook to the local hard drive\n' \
               'The location of the emails will be stored in a folder called Emails where this application resides\n' \
               'Emails sent/received within x number of days as defined in the settings.ini file will be copied\n' \
               '\n' \
               f'Emails sent/received in the last {email_max_days} days will be processed\n' \
               f'Deletion of mails older than {email_max_days} days is set to: {delete_old_emails}\n' \
               '\n' \
               '--- Press start to begin ---\n'

        window['-LOG-'].update(info)

        if event == '-START-':
            start_processing = True
            window['-START-'].update(disabled=True)
            window['-LOG-'].update('\n', append=True)

            return start_processing

        elif event in (None, 'Exit'):
            start_processing = False
            window.read(timeout=1000)
            return start_processing


def gui_update(window, msg):
    # Update the window with messages from the application
    window['-LOG-'].update(msg + '\n', append=True)
    window.refresh()


def get_outlook_details():
    outlook_namespace = win32com.client.Dispatch('outlook.application').GetNamespace('MAPI')

    counter = 0
    main_account = ''

    for account in outlook_namespace.Accounts:
        # print(account.DeliveryStore.DisplayName)
        main_account = account.DeliveryStore.DisplayName
        counter += 1

    return counter, outlook_namespace, main_account


def get_top_level_folders(window, outlook_namespace, main_account):
    top_level_folders = []

    gui_update(window, '\n--- Starting the process to identify emails to be saved to the file system ---\n')

    msg = 'Found the following Top Level Folders:'
    logger.info(msg)
    gui_update(window, msg)

    for idx, folder in enumerate(outlook_namespace.Folders(main_account).Folders):
        if folder.Name not in folder_exclusion_list:
            msg = folder.Name
            logger.info(msg)
            gui_update(window, msg)
            top_level_folders.append(folder)

    gui_update(window, '')

    return top_level_folders


def process_folders(window, top_level_folders, working_directory):

    # Keep track of all emails processed
    total_emails_processed = 0
    total_new_emails_saved = 0
    total_email_size_mb = 0

    # The top level folder list is passed through to this function
    # Iterate through each top level folder
    for folder_1 in top_level_folders:

        # Eliminate any special characters in the folder name
        folder_name_1 = re.sub('[^A-Za-z0-9]+', ' ', folder_1.Name)

        folder_chain = folder_name_1

        # Create the directory for the folder if it does not exist
        directory_1 = process_directory(window, working_directory, folder_name_1)
        # For each item (message) in the entry (top level folder - level 1), get the message subject
        emails_1, new_emails_1, email_size_1 = get_messages(window, folder_1, directory_1, folder_chain)
        total_emails_processed += emails_1
        total_new_emails_saved += new_emails_1
        total_email_size_mb += email_size_1

        # Check each top level folder to determine if there are other folders nested inside it
        for folder_2 in folder_1.Folders:

            # Eliminate any special characters in the folder name
            folder_name_2 = re.sub('[^A-Za-z0-9]+', ' ', folder_2.Name)

            folder_chain = folder_name_1 + '\\' + folder_name_2

            # Create the directory for the folder if it does not exist
            directory_2 = process_directory(window, directory_1, folder_name_2)
            # For each item (message) in the entry (top level folder - level 2), get the message subject
            emails_2, new_emails_2, email_size_2 = get_messages(window, folder_2, directory_2, folder_chain)
            total_emails_processed += emails_2
            total_new_emails_saved += new_emails_2
            total_email_size_mb += email_size_2

            # Check each top level folder to determine if there are other folders nested inside it
            for folder_3 in folder_2.Folders:

                # Eliminate any special characters in the folder name
                folder_name_3 = re.sub('[^A-Za-z0-9]+', ' ', folder_3.Name)

                folder_chain = folder_name_1 + '\\' + folder_name_2 + '\\' + folder_name_3

                # Create the directory for the folder if it does not exist
                directory_3 = process_directory(window, directory_2, folder_name_3)
                # For each item (message) in the entry (top level folder - level 3), get the message subject
                emails_3, new_emails_3, email_size_3 = get_messages(window, folder_3, directory_3, folder_chain)
                total_emails_processed += emails_3
                total_new_emails_saved += new_emails_3
                total_email_size_mb += email_size_3

                # Check each top level folder to determine if there are other folders nested inside it
                for folder_4 in folder_3.Folders:

                    # Eliminate any special characters in the folder name
                    folder_name_4 = re.sub('[^A-Za-z0-9]+', ' ', folder_4.Name)

                    folder_chain = folder_name_1 + '\\' + folder_name_2 + '\\' + folder_name_3 + '\\' + folder_name_4

                    # No more directories in this 4th level of folders will be processed, only emails
                    directory_4 = process_directory(window, directory_3, folder_name_4)
                    # For each item (message) in the entry (top level folder - level 4), get the message subject
                    emails_4, new_emails_4, email_size_4 = get_messages(window, folder_4, directory_4, folder_chain)
                    total_emails_processed += emails_4
                    total_new_emails_saved += new_emails_4
                    total_email_size_mb += email_size_4

    return total_emails_processed, total_new_emails_saved, total_email_size_mb


def process_directory(window, directory, folder_name):
    folder_directory = os.path.join(directory, folder_name)

    # Check to see if the directory exists, if not create it
    if not os.path.exists(folder_directory):
        os.mkdir(folder_directory)
        msg = f'Created new directory at {folder_directory}'
        logger.info(msg)
        gui_update(window, msg)

    return folder_directory


def get_messages(window, folder, directory, folder_chain):

    # Keep count of the number of messages processed
    emails_processed = 0
    new_emails_saved = 0
    email_size_mb = 0

    # Set the date time format to be used
    pattern = '%d-%m-%Y %H:%M:%S'

    msg = 'Processing emails in: ' + folder_chain
    logger.info(msg)
    gui_update(window, msg)

    # Get today's time and subtract the time period as defined in the config file
    today = datetime.now().strftime(pattern)
    oldest_time = datetime.today() - timedelta(days=email_max_days)
    oldest_time = oldest_time.strftime(pattern)

    # For each item (message) in the folder, get the message subject
    messages = folder.items

    for msg in messages:

        # Check to see if the item is a message and the value of 43 is assigned to an email
        # Refer to Microsoft Outlook Object Class enumeration
        if msg.Class == 43:

            # Keep a track of the emails in each folder
            emails_processed += 1

            try:
                # Get the Sent date for each email
                email_date = msg.SentOn.strftime(pattern)

                # Looking for a positive number, indicating that the email is younger than the oldest time
                # Any email with a negative number of days will be excluded as they were sent after the number of days
                time_delta = datetime.strptime(email_date, pattern) - datetime.strptime(oldest_time, pattern)
                #print(time_delta.days)

                if time_delta.days >= 0:
                    # Get email sent date to be used for setting the modified date
                    date = msg.SentOn.strftime(pattern)
                    epoch_format = int(time.mktime(time.strptime(date, pattern)))

                    # Get subject of the message to be used as the message file name
                    # Check if the subject is blank and if so, assign a generic file name
                    if msg.subject == '':
                        name = 'No Subject Provided'
                    else:
                        name = str(msg.subject)

                    # Eliminate any special characters in the name
                    substitute_characters = re.sub('[^A-Za-z0-9]+', '_', name) + '.msg'

                    # Going to check the length of the directory and filename to ensure it does not exceed 260 characters
                    # Failure of the complete file path exceeding 260 characters will prevent it from being saved
                    directory_length = len(directory)
                    filename_length = len(substitute_characters)

                    # Check if the file path is longer than 230 (leave 30 character buffer, less epoch and file type)
                    if directory_length + filename_length >= 230:
                        max_filename_length = 230 - directory_length
                        filename_shortened = substitute_characters[:max_filename_length]
                        filename = str(epoch_format) + '_' + filename_shortened + '.msg'

                    else:
                        filename = str(epoch_format) + '_' + substitute_characters

                    # Check if the message has already been saved
                    complete_file_path = os.path.join(directory, filename)

                    if not os.path.isfile(complete_file_path):
                        # Save the message in the current working directory

                        try:
                            msg.SaveAs(directory + '//' + filename)
                            print('Saved email to filesystem: ' + directory + '\\' + filename)

                            # Update the modified time to the sent date of the email
                            os.utime(complete_file_path, (epoch_format, epoch_format))

                            # Keep count of the total number of emails processed
                            new_emails_saved += 1

                            # Get the size of the email and add it to the overall tally - size in MB
                            size_bytes = os.path.getsize(complete_file_path)
                            email_size_mb += size_bytes/1000/1000

                        except Exception as e:
                            msg = f'Error writing item to file system: {complete_file_path}\n' \
                                  f'A likely reason could be that the email has been classified as' \
                                  f' restricted or highly restricted'
                            logger.info(msg)
                            gui_update(window, msg)
                            logger.error(e, exc_info=True)

                    else:
                        print('Email previously saved: ' + filename)

            except Exception as e:
                msg = f'Error: Selected item for processing has been identified as Class # {str(msg.Class)} ' \
                      f'(email), however, it could not be processed as one'
                logger.info(msg)
                gui_update(window, msg)
                logger.error(e, exc_info=True)

    msg = f'Processed {str(emails_processed)} emails with {str(new_emails_saved)} new emails discovered ' \
          f'and written to disk'
    logger.info(msg)
    gui_update(window, msg)

    return emails_processed, new_emails_saved, email_size_mb

def remove_old_emails(window, working_directory):

    # Counter to keep track of how many emails are removed
    emails_removed = 0
    removed_email_size_mb = 0

    # If the settings.ini file has delete old emails set to True, then remove the old emails
    if delete_old_emails == 'True':

        gui_update(window, '--- Starting the process to identify old emails to be removed from the file system ---\n')

        # Get today's time and subtract the time period as defined in the config file
        current_epoch = time.time()
        msg = f'The epoch time is now: {current_epoch}'
        logger.info(msg)
        max_epoch = current_epoch - (email_max_days * 86400)
        msg = f'The max number of days epoch time is: {max_epoch}'
        logger.info(msg)

        msg = f'Identifying emails older than {email_max_days} days in the email archive directory on the file system'
        gui_update(window, msg)
        logger.info(msg)

        # Discover all emails stored in the Emails working directory
        email_list = glob.glob(working_directory + '/**/*.msg', recursive=True)
        for email in email_list:

            # Get the modified time of the email file
            email_mtime = os.stat(email).st_mtime

            # If the modified time of the email is older than the max number of days as per the settings.ini file
            # The email will be deleted
            if email_mtime <= max_epoch:
                msg = f'The following email is older than the max number of days: {email}'
                logger.info(msg)

                if os.path.isfile(email):
                    try:
                        # Get the size of the email and add it to the overall tally - size in MB
                        size_bytes = os.path.getsize(email)
                        removed_email_size_mb += size_bytes / 1000 / 1000

                        os.remove(email)
                        msg = f'Successfully deleted old email: {email}'
                        logger.info(msg)
                        print('Removed email from file system: ' + email)
                        emails_removed += 1

                    except Exception as e:
                        msg = f'Failed to delete old email: {email}'
                        logger.info(msg)
                        gui_update(window, msg)
                        logger.error(e, exc_info=True)

    else:
        msg = 'No emails were deleted as it was not defined as True in the settings.ini file'
        logger.info(msg)
        gui_update(window, msg)

    return emails_removed, removed_email_size_mb


# The main application
def main():

    # Check on Outlook running status
    outlook_running()

    # Start the GUI
    window = gui()
    start_processing = gui_status(window, '')

    # If the start button has been pressed and the settings.ini file has been read, then start processing the emails
    if start_processing == True and settings_file_present == True:

        # Get the Outlook details
        account_counter, outlook_namespace, main_account = get_outlook_details()

        # Checking that only 1 Outlook account has been returned
        if account_counter == 1:

            # Display information about Outlook
            msg = f'The directory where emails will be stored: {working_directory}'
            logger.info(msg)
            gui_update(window, msg)
            msg = 'Number of Email Accounts: ' + str(account_counter)
            logger.info(msg)
            gui_update(window, msg)
            msg = 'Discovered email account name: ' + main_account
            logger.info(msg)
            gui_update(window, msg)

            # Identify the top level folders for the account
            top_level_folders = get_top_level_folders(window, outlook_namespace, main_account)

            # Process the emails in each discovered folder
            emails_processed, new_emails_saved, email_size_mb = process_folders(window, top_level_folders, working_directory)
            email_size_mb = format(email_size_mb, '.3f')
            if new_emails_saved > 0:
                msg = f'Processed {str(emails_processed)} emails with {str(new_emails_saved)} new emails discovered ' \
                      f'and written to disk consuming {str(email_size_mb)}MB'
                logger.info(msg)
                gui_update(window, msg)
            else:
                msg = f'Processed {str(emails_processed)} emails with {str(new_emails_saved)} new emails discovered'
                logger.info(msg)
                gui_update(window, msg)
                gui_update(window, '')

            # Remove emails older that the specified number of days
            emails_removed, removed_emails_size_mb = remove_old_emails(window, working_directory)
            msg = f'Finished processing the removal of emails'
            logger.info(msg)
            gui_update(window, msg)
            removed_emails_size_mb = format(removed_emails_size_mb, '.3f')
            if emails_removed > 0:
                msg = f'Successfully removed {emails_removed} emails consuming {removed_emails_size_mb}MB'
                logger.info(msg)
                gui_update(window, msg)
            else:
                msg = f'No emails were removed from the file system'
                logger.info(msg)
                gui_update(window, msg)

            # Print summary at the very end of the log file for easy reference
            gui_update(window, '\n-- Summary of Email Archive Activity ---\n')

            if new_emails_saved > 0:
                msg = f'Processed {str(emails_processed)} emails with {str(new_emails_saved)} new emails discovered ' \
                      f'and written to disk consuming {str(email_size_mb)}MB'
                logger.info(msg)
                gui_update(window, msg)
            else:
                msg = f'Processed {str(emails_processed)} emails with {str(new_emails_saved)} new emails discovered'
                logger.info(msg)
                gui_update(window, msg)

            if emails_removed > 0:
                msg = f'Successfully removed {emails_removed} emails consuming {removed_emails_size_mb}MB'
                logger.info(msg)
                gui_update(window, msg)
            else:
                msg = f'No emails were removed from the file system'
                logger.info(msg)
                gui_update(window, msg)

            # 1 minute delay before the window disappears
            gui_update(window, '')
            gui_update(window, '--- Press Exit to close the application or it will auto close after 1 minute ---')
            window.read(timeout=60000)

        # If no Outlook account or more than 1 is returned, the application will not proceed
        else:
            msg = 'The number of accounts found: ' + str(account_counter)
            logger.info(msg)
            gui_update(window, msg)
            msg = 'Terminating the application as it will only allow for 1 outlook account to be processed.'
            logger.info(msg)
            gui_update(window, msg)

    elif start_processing == False:
        gui_update(window, '')
        msg = 'Exiting the application as the Exit button was pressed'
        logger.info(msg)
        gui_update(window, msg)


if __name__ == '__main__':
    main()

# End of Outlook Archive script
