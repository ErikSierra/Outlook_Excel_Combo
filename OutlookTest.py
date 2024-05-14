# This script tests connection to your outlook inbox

import win32com.client

outlook = win32com.client.Dispatch('outlook.application')
inbox = outlook.GetNamespace("MAPI").GetDefaultFolder(5)
num_emails = inbox.Items.Count
print(f"Number of emails in Inbox: {num_emails}")
