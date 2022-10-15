# -*- coding: utf-8 -*-
"""
@author: Emmanuel Ch.
"""

import win32com.client as win32
import time

def list_mailboxes():
    themailboxes = dict()
    for i, folder in enumerate(win32.Dispatch("Outlook.Application").GetNamespace("MAPI").Folders, 1):
        themailboxes[i] = folder.Name
        print(i, folder.Name)
    return themailboxes


def export_mailbox_hierarchy(mailbox_name):
    namespace = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    mailbox = namespace.Folders(mailbox_name)
    
    filename = 'Mailbox hierarchy.txt'
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(recursive_menu(mailbox))
    
    return filename

    
def export_list_emails(target_folder):
    filename = 'Email list.txt'
    res = ''
    
    for i, m in enumerate(target_folder.Items):
        res += str(m.SentOn) + ' ' + str(m.SenderName) + ' || ' + str(m.Subject) + '\n'
        if i> 100:
            break
    
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(res)
    print(f'List of emails exported to: {filename}')
    

def recursive_menu(outlookFolderItem, indent='', print_block=True, pick_email='', print_email_samples=True, move_to=None):
    res = ''
    
    if print_block:
        res += f'########## {outlookFolderItem.Name} ##########\n'
    
    # Print emails
    if pick_email != '' or print_email_samples:
        counter=0
        for m in outlookFolderItem.Items:
            counter+=1
            if m.Subject == pick_email and pick_email != '':
                res += f'--> EMAIL FOUND Here! ({pick_email}) <--\n'
                if not move_to is None:
                    res += f'--> Moving it to folder: {move_to.Name}\n'
                    m.Move(move_to)
                    res += '--> Email moved\n'
                
            if print_email_samples:
                res += indent + '(e) ' +  m.Subject + '\n'
            if counter == 3:
                res += indent + '(e) ' + '... (Limited to 3 items)' + '\n'
                break
    
    # Print folders and their children
    for i in range(0,30): # Only the first 10 folders are interesting for us here (Inbox, Sent email, etc...)
        try:
            res += indent + f'{i}: ' + outlookFolderItem.Folders(i).Name + '\n'
            res += recursive_menu(outlookFolderItem.Folders(i), '....'+indent, False,
                                  pick_email, print_email_samples, move_to)
        except:
            pass
    
    if print_block:
        res += '#########################'
    
    return res


##########################

def main():
    
    print('\n#########################')
    print('###### EMAIL MOVER ######')
    print('#########################\n')
    
    # Which mailbox?
    print('Generating list of mailboxes accessible:')
    themailboxes = list_mailboxes()
    mailbox_num = input('\nWhich mailbox hierarchy should we export (Ex: "3")?     ')
    mailbox_name = themailboxes[int(mailbox_num)]
    print(f'Exporting hierarchy of mailbox: {mailbox_name} ...')
    mailbox_hier_filename = export_mailbox_hierarchy(mailbox_name)
    print(f'Mailbox hierarchy exported to: {mailbox_hier_filename}. Please check it.')
    
    # Select folder and list emails
    print('\n##### SOURCE FOLDER #####')
    mailbox_target_folder = input('In which folder are emails to move (Ex: "2 3 1")?     ')
    namespace = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    mailbox = namespace.Folders(mailbox_name)
    
    target_folder = mailbox
    for subfolder_num in mailbox_target_folder.split():
        target_folder = target_folder.Folders(subfolder_num)
        
    print('Exporting list of emails...')
    export_list_emails(target_folder)
    
    # Try to move an email
    ex_email = target_folder.Items.GetFirst().Subject
    subject_selected = input(f'\nHave a look at the list of emails. Which email do you want to move (Ex: "{ex_email}")?     ')
    
    print('\n##### DESTINATION FOLDER #####')
    mailbox_num_moveto = input('To which mailbox move the email (Ex: "3")?     ')
    mailbox_name_moveto = themailboxes[int(mailbox_num_moveto)]
    
    mailbox_moveto_folder = input('In which folder should emails be moved to (Ex: "2 5 1")?     ')
    moveto_folder = namespace.Folders(mailbox_name)
    for subfolder_num in mailbox_moveto_folder.split():
        moveto_folder = moveto_folder.Folders(subfolder_num)
    
    for m in target_folder.Items:
        if m.Subject == subject_selected:
            print(f'Found: {subject_selected}')
            m.Move(moveto_folder)
            print(f'Email moved to: {moveto_folder.Name}')
            break
    
    print('\n############################')
    print('###### END OF PROGRAM ######')
    print('############################\n\n')
    time.sleep(3)

if __name__ == '__main__':
    main()

