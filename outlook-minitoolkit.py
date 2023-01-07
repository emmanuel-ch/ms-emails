# -*- coding: utf-8 -*-
"""
@author: Emmanuel Ch.
"""

import win32com.client as win32
import time
import sys

WIDTH_ANNOUNCES = 50
mailboxes_available = {}

def main_menu():
    print('\n' + ' MAIN MENU '.center(WIDTH_ANNOUNCES, '#'))
    print('##  1 . Display list of mailboxes accessible')
    print('##  2 . Export mailbox hierarchy')
    print('##  3 . Move an email')
    print('##  9 . Exit program')
    
    action = 0
    while action not in ['1', '2', '3', '9']:
        action = input('What do you want to do?    ')
    print()
    
    if action == '1':
        print('* Mailbox listing *')
        list_mailboxes()
        main_menu()
    elif action == '2':
        print('* Mailbox hierarchy *')
        export_mailbox_hierarchy()
        main_menu()
    elif action == '3':
        print('* Move an email *')
        move_email()
        main_menu()
    elif action == '9':
        print(''.center(WIDTH_ANNOUNCES, '#'))
        print(' END OF PROGRAM '.center(WIDTH_ANNOUNCES, '#'))
        print(''.center(WIDTH_ANNOUNCES, '#'))
        time.sleep(3)
        sys.exit()
    return True


def list_mailboxes():
    global mailboxes_available
    themailboxes = dict()
    for i, folder in enumerate(win32.Dispatch("Outlook.Application").GetNamespace("MAPI").Folders, 1):
        themailboxes[str(i)] = folder.Name
        print(i, folder.Name)
    mailboxes_available = themailboxes
    return True


def export_mailbox_hierarchy():
    mailbox_name = ask_which_mailbox('Which mailbox hierarchy should we export (Ex: "3")?     ')
    if not mailbox_name:
        return False
    print(f'Exporting hierarchy of mailbox: {mailbox_name} ...')
    
    namespace = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    mailbox = namespace.Folders(mailbox_name)
    
    filename = f'Mailbox {mailbox_name} {time.strftime("%y.%m.%d %Hh%M")}.txt'
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(recursive_menu(mailbox))
    print(f'Mailbox hierarchy exported to: {filename}')
    return True


def ask_which_mailbox(msg):
    if len(mailboxes_available) < 1:
        print('!!! No mailbox listing in memory, generate it first (main menu option #1).')
        return False
    mailbox_num = 0
    while mailbox_num not in mailboxes_available.keys():
        mailbox_num = input(msg)
    return mailboxes_available[mailbox_num]


def move_email():
    # From ...
    src_mailbox_name = ask_which_mailbox('Source mailbox # (ex: "3")?     ')
    if not src_mailbox_name:
        return False
    namespace = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    src_mailbox = namespace.Folders(src_mailbox_name)
    
    mailbox_src_folder = input('In which folder is the email to move (Ex: "2 3 1")?     ')
    src_folder = src_mailbox
    try:
        for subfolder_num in mailbox_src_folder.split():
            src_folder = src_folder.Folders(subfolder_num)
    except:
        print('!!! Problem reaching source folder. Try again.')
        return False
    print(f'Confirmed source folder: {src_folder.Name}.')
    
    print('Printing header of a few emails:')
    print(show_few_emails(src_folder))
    ex_email = src_folder.Items.GetFirst().Subject
    subject_selected = input(f'Subject of email you want to move (ex: "{ex_email}"):     ')
    
    # To ...
    dst_mailbox_name = ask_which_mailbox('Destination mailbox # (ex: "3")?     ')
    dst_namespace = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    dst_mailbox = dst_namespace.Folders(dst_mailbox_name)
    
    mailbox_dst_folder = input('In which folder should emails be moved to (ex: "2 5 1")?     ')
    dst_folder = dst_mailbox
    try:
        for subfolder_num in mailbox_dst_folder.split():
            dst_folder = dst_folder.Folders(subfolder_num)
    except:
        print('!!! Problem reaching destination folder. Try again.')
        return False
    print(f'Confirmed destination folder: {dst_folder.Name}')
    
    # Move it!
    for m in src_folder.Items:
        if m.Subject == subject_selected:
            m.Move(dst_folder)
            print(f'Email "{subject_selected}" moved to: {dst_folder.Name}')
            break
    
    return True

    
def show_few_emails(mailbox_folder, nb_items = 3):
    res = ''
    for i, m in enumerate(mailbox_folder.Items, 1):
        res += f'  {str(m.SentOn)[2:19]} {m.SenderName} | {m.Subject}\n'
        if i>=nb_items:
            break
    return res
    
    
def recursive_menu(outlookFolderItem, indent='', print_block=True, print_email_samples=True):
    res = ''
    if print_block:
        res += f'########## {outlookFolderItem.Name} ##########\n'
    
    # Print emails
    if print_email_samples:
        counter=0
        for m in outlookFolderItem.Items:
            counter+=1
            if print_email_samples:
                res += indent + '(e) ' +  m.Subject + '\n'
            if counter == 3:
                res += indent + '(e) ...\n'
                break
    
    # Print folders and their children
    for i in range(0,30):
        try:
            res += indent + f'{i}: {outlookFolderItem.Folders(i).Name}\n'
            res += recursive_menu(outlookFolderItem.Folders(i), '....'+indent,
                                  False, print_email_samples)
        except:
            pass
    
    if print_block:
        res += '#########################'
    return res


def main():
    print()
    print(''.center(WIDTH_ANNOUNCES, '#'))
    print(' OUTLOOK MINI-TOOLKIT '.center(WIDTH_ANNOUNCES, '#'))
    print(''.center(WIDTH_ANNOUNCES, '#'))
    main_menu()


if __name__ == '__main__':
    main()

