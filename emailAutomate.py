import win32com.client
import os
import time
import subprocess

def send_outlook_email():
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    recipient_email = input("Enter recipient email: ")
    position = input("Enter the position title: ")
    headless_mode = 'n'

    # Ask the user if they want to call another Python script
    call_another_script = input("Do you want to add a motivation letter? (y/n): ").strip().lower()
    if call_another_script == 'y':
        # Path to the other Python script
        script_path = r"D:\Automate_Word.py"
        if os.path.exists(script_path):  # Check if the script exists
            print(f"Calling another script: {script_path}")
                # Use subprocess to run the script
            subprocess.run(["python", script_path], check=True)
        else:
                print(f"Error: The script '{script_path}' does not exist.")
    elif call_another_script == 'n':
            headless_mode = input("Run in headless mode? (y/n): ").strip().lower()
            print("No additional motivation letter will be created.")
    else:
            print("Invalid input. Please enter 'y' or 'n'.")

    mail.Subject = f"Bewerbung als {position}"
    mail.Body = f"""Sehr geehrte Damen und Herren,

ich bewerbe mich hiermit für die Position {position}. Im Anhang finden Sie meine Bewerbungsunterlagen. Sollten Sie noch weitere Dokumente benötigen, geben Sie mir bitte Bescheid.

Bei Fragen stehe ich gerne zur Verfügung.

Ich freue mich, Sie bei einem persönlichen Gespräch kennenzulernen.

Freundliche Grüße
Marwan Hammad
"""

    mail.To = recipient_email
    mail.Recipients.ResolveAll()

    attachments = [
        r"D:\CV.pdf",
        r"D:\Zeugnisse_Marwan_Hammad.pdf"
    ]
    
    for file_path in attachments:
        if os.path.exists(file_path):
            mail.Attachments.Add(file_path)
        else:
            print(f"Warning: Attachment '{file_path}' not found.")

    if headless_mode == 'y':
        mail.Send()
        print("Email sent successfully (headless mode).")
    elif headless_mode == 'n':
        mail.Display()
        print("Email displayed for review.  Send manually.")  # Clarify manual send
    else:
        print("Invalid input. Please enter 'y' or 'n'.")

send_outlook_email()
