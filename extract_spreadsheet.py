import csv
import os
from Email_Sender import send_email

# File names

NEW_FOLDER_DIRECTORY_PATH = os.path.join(os.getcwd(),"New_Folder")

input_file = os.path.join(NEW_FOLDER_DIRECTORY_PATH, "spreadsheet.csv")
output_csv = os.path.join(NEW_FOLDER_DIRECTORY_PATH, "tosend.csv")
output_txt = os.path.join(os.path.join(NEW_FOLDER_DIRECTORY_PATH, "Wordlist"), 'wordlist.txt')


# Process the data
with open(input_file, mode='r') as infile:
    reader = csv.DictReader(infile)
    
    # Create and write to 'tosend.csv'
    with open(output_csv, mode='w', newline='') as outfile_csv:
        csv_writer = csv.writer(outfile_csv)
        csv_writer.writerow(['Full Name', 'Email'])  # Write header row
        
        # Create and write to 'wordlist.txt'
        with open(output_txt, mode='w') as outfile_txt:
            for row in reader:
                if row['Attendance'].strip().upper() == 'TRUE':  # Check Attendance
                    full_name_title = row['Full Name'].strip().title()  # Convert to title case
                    csv_writer.writerow([full_name_title, row['Email']])  # Write to CSV
                    outfile_txt.write(f"{full_name_title}\n")  # Write to TXT

certificate_script_path = os.path.join(os.path.join(os.getcwd(), "Certificate_Generator"), "certificate_generator.py")
email_script_path = os.path.join(os.path.join(os.getcwd(), "Email_Sender"), "send_email.py")

os.system(f"python {certificate_script_path} other_script")
os.system(f"python {email_script_path} other_script")
