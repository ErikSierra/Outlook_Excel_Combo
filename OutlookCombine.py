import pandas as pd
import win32com.client
import datetime
import os
from colorama import init, Fore, Style

# Initialize colorama
init()

def combine_emails():
    # Connect to Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    print(Fore.GREEN + "Successfully connected to Outlook." + Style.RESET_ALL)
    
    inbox = outlook.GetDefaultFolder(6)  # 6 refers to the inbox folder
    messages = inbox.Items

    # Ensure the datetime format is correct for Outlook
    six_days_ago = (datetime.datetime.now() - datetime.timedelta(days=6)).strftime('%m/%d/%Y %I:%M %p')
    filter_condition = f"[ReceivedTime] >= '{six_days_ago}'"
    print(Fore.BLUE + f"Filter Condition: {filter_condition}" + Style.RESET_ALL)

    try:
        filtered_messages = messages.Restrict(filter_condition)
        print(Fore.BLUE + f"Filtered Messages Count: {filtered_messages.Count}" + Style.RESET_ALL)
    except Exception as e:
        print(Fore.RED + f"Error filtering messages: {e}" + Style.RESET_ALL)
        return pd.DataFrame()

    df_list = []
    for message in filtered_messages:
        if message.Attachments.Count > 0:
            for attachment in message.Attachments:
                if attachment.FileName in ['Manual Systems Not Patching Export.csv', 'Automatic Systems Not Patching Export.csv']:
                    try:
                        # Print email subject and date
                        print(Fore.CYAN + f"Processing email with subject: {message.Subject} and received date: {message.ReceivedTime}" + Style.RESET_ALL)
                        
                        # Saving attachment and reading with pandas
                        temp_file = os.path.join(os.environ['USERPROFILE'], 'Downloads', 'temp_' + attachment.FileName)
                        attachment.SaveAsFile(temp_file)
                        df = pd.read_csv(temp_file)

                        # Add Date and Type columns
                        yesterday = datetime.datetime.now() - datetime.timedelta(days=1)
                        date_string = yesterday.strftime('%m/%d/%y 00:00')
                        df['Date'] = date_string
                        df['Type'] = 'Manual' if 'Manual' in attachment.FileName else 'Automatic'

                        df_list.append(df)
                        os.remove(temp_file)
                        print(Fore.GREEN + f"Successfully loaded data from {attachment.FileName} into DataFrame." + Style.RESET_ALL)
                    except Exception as e:
                        print(Fore.RED + f"Error processing attachment {attachment.FileName}: {e}" + Style.RESET_ALL)

    if df_list:
        df_comb = pd.concat(df_list, ignore_index=True)
        print(Fore.GREEN + "Combined CSV files into a single DataFrame." + Style.RESET_ALL)
        return df_comb
    else:
        print(Fore.RED + "No relevant CSV attachments found." + Style.RESET_ALL)
        return pd.DataFrame()  # Return empty DataFrame if no files were processed

def save_data_frame(df, filename):
    # Saves data frame to an Excel file
    try:
        df.to_excel(filename, index=False)
        print(Fore.GREEN + f"Data successfully saved to file: {filename}" + Style.RESET_ALL)
    except Exception as e:
        print(Fore.RED + f"Error saving DataFrame to file {filename}: {e}" + Style.RESET_ALL)

if __name__ == '__main__':
    df_comb = combine_emails()
    if not df_comb.empty:
        save_data_frame(df_comb, 'combined_data.xlsx')
    else:
        print(Fore.RED + "No data to save." + Style.RESET_ALL)
