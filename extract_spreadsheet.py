# extract_certify_and_email.py
# Certificate_Email_Automation
import csv
import os
from Certificate_Generator.certificate_generator import get_files, get_single_file
from Email_Sender.send_email import check_csv, sort_csv


# === MAIN ENTRY POINT ===
if __name__ == "__main__":
    
    if not os.getcwd()[-9:] == "Resources":
        print("\nPlease change your working directory to the main repository.\n\nExiting...\n")
        exit(1)
        
    EMAIL_SUBJECT = "Subject"

    NEW_FOLDER_DIRECTORY_PATH = os.path.join(os.getcwd(),"New_Folder")
    SPREADSHEET_DIRECTORY_PATH = os.path.join(NEW_FOLDER_DIRECTORY_PATH, "Spreadsheet")
    WORDLIST_DIRECTORY_PATH = os.path.join(NEW_FOLDER_DIRECTORY_PATH, "Wordlist")
    CERTIFICATE_GENERATOR_DIRECTORY_PATH = os.path.join(os.getcwd(), "Certificate_Generator")
    EMAIL_SENDER_DIRECTORY_PATH = os.path.join(os.getcwd(), "Email_Sender")
    
    print("\n" + " Certificate_Email_Automation ".center(40, "-"))
    
    os.makedirs(SPREADSHEET_DIRECTORY_PATH, exist_ok=True)
    os.makedirs(WORDLIST_DIRECTORY_PATH, exist_ok=True)
    
    csv_files = get_files(SPREADSHEET_DIRECTORY_PATH, 'CSV')
    spreadsheet_file = get_single_file('Spreadsheet', SPREADSHEET_DIRECTORY_PATH, 'CSV')
    spreadsheet_file_path = os.path.join(SPREADSHEET_DIRECTORY_PATH, spreadsheet_file)
    
    # Certificate_Generator.certificate_generator.WORDLIST_DIRECTORY_PATH = WORDLIST_DIRECTORY_PATH
    text_files = get_files(WORDLIST_DIRECTORY_PATH, 'TXT')
    if not text_files:
        with open(os.path.join(WORDLIST_DIRECTORY_PATH, 'wordlist.txt'), 'w') as tosend_csv_file:
            pass
        
    wordlist_file_path = os.path.join(WORDLIST_DIRECTORY_PATH, 'wordlist.txt')
    tosend_csv_path = os.path.join(NEW_FOLDER_DIRECTORY_PATH, "tosend.csv")
    certificate_script_path = os.path.join(CERTIFICATE_GENERATOR_DIRECTORY_PATH, "certificate_generator.py")
    email_script_path = os.path.join(EMAIL_SENDER_DIRECTORY_PATH, "send_email.py")
    
    # Open the file and ensure it has the correct contents as needed
    check_csv(spreadsheet_file_path, "Other", "Attendance")
    
    # Process the file
    with open(spreadsheet_file_path, mode='r', encoding="utf-8") as sheet_csv_file:
        reader = csv.DictReader(sheet_csv_file)
        try:
            # Create and write to 'tosend.csv'
            with open(tosend_csv_path, mode='w', newline='') as tosend_csv_file:
                csv_writer = csv.writer(tosend_csv_file)
                csv_writer.writerow(['Full Name', 'Email'])  # Write header row
                
                # Create and write to 'wordlist.txt'
                with open(wordlist_file_path, mode='w') as wordlist_file:
                    for row in reader:
                        if row['Attendance'].strip().upper() == 'TRUE':  # Check Attendance
                            full_name = row['Full Name'].strip().title()
                            csv_writer.writerow([full_name, row['Email']])
                            wordlist_file.write(f"{full_name}\n")
                    print("\n\'Full Name\' column successfully written to wordlist file.")
                    
                print("\'Full Name\' and \'Email\' columns successfully extracted to \'tosend.csv\' file.")
            print("\nSuccessfully extracted the spreadsheet file.\n")
        except:
            print("\nFailed to write to \"tosend.csv\" file.\nEnsure that the file is not open on the system.\n")
            exit(1)

    sort_csv(tosend_csv_path)
    try:
        certificate_script_status = os.system(f"python {certificate_script_path} other_script")
    except KeyboardInterrupt:
        print("\n\nKeyboard Interrupt!\n\nExiting...\n")
        exit(1)
    except Exception as e:
        print(f"\n\nAn error occured while executing the \'certificate_generator script\'.\n{e}\n")
        exit(1)
        
    if certificate_script_status == 0:
        try:
            os.system(f"python {email_script_path} other_script \"{EMAIL_SUBJECT}\"")
        except KeyboardInterrupt:
            print("\n\nKeyboard Interrupt!\n\nExiting...\n")
            exit(1)
        except Exception as e:
            print(f"\n\nAn error occured while executing the \'certificate_generator script\'.\n{e}\n")
            exit(1)
