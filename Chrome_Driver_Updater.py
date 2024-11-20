import io
import subprocess
import ctypes
import sys
import os
import shutil
import time
import requests
from zipfile import ZipFile
import contextlib
import tkinter as tk


def hide_console():
    # Hide the console window
    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)


def run_command(dest_link, download_link):
    # check the existing version from the chromedriver.exe file
    drive_path0 = dest_link[0:2].strip()  # scan the drive of the destination path (example: C: or E:)
    command1 = fr"{drive_path0} && cd {dest_link} && chromedriver.exe -v"
    process = subprocess.run(command1, shell=True, stdout=subprocess.PIPE, universal_newlines=True)
    output = process.stdout
    # extract string of the version part only
    version_exist = output[11:28].strip('r( ')
    print("Exist Version: ", version_exist)
    # stable version check via internet
    r1 = requests.get("https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_STABLE")
    version_stable = r1.text
    print("Latest Stable Version: ", version_stable)

    # check the versions
    if version_stable == version_exist:
        # if versions are the same then stop update process
        print("You have the latest version!")
    else:
        # if verisons are different start update process
        print("Update!")
        try:
            # delete existing chromedriver.exe file (older version)
            drive_path1 = dest_link[0:2].strip()
            command2 = fr"{drive_path1} && cd {dest_link} && del chromedriver.exe"
            subprocess.run(command2, shell=True, stdout=subprocess.PIPE, universal_newlines=True)

        except Exception as msg2:
            print(msg2)

        try:
            # Adding stable version number to the url
            r2 = requests.get(f'https://storage.googleapis.com/chrome-for-testing-public/{version_stable}/win64/chromedriver-win64.zip')
            # Download the content zip file
            print("Downloading Stable version!")
            with open(f'{download_link}/chromedriver-win64.zip', 'wb') as f1:
                f1.write(r2.content)
        except Exception as msg3:
            print(msg3)

        filename = f"{download_link}/chromedriver-win64.zip"
        with ZipFile(filename, 'r') as z:
            # Extract the zip file to the root folder
            z.extract('chromedriver-win64/chromedriver.exe', f'{download_link}')
            print(f"{z.namelist()[2]} Extracted!")
            # Move the chromedriver.exe file to the destination folder from the extracted folder
            try:
                print("Moved file to the destination!")
                shutil.move(f'{download_link}/chromedriver-win64/chromedriver.exe', dest_link)

            except Exception as msg:
                print(msg)

        time.sleep(1)
        drive_path2 = download_link[0:2].strip()
        command4 = fr"{drive_path2} && cd {download_link} && del /f /q chromedriver-win64.zip"
        subprocess.run(command4, shell=True, stdout=subprocess.PIPE, universal_newlines=True)
        os.rmdir(f'{download_link}/chromedriver-win64')
        print("Removed Temporary files and folders!")
        print("Done!")


def run_as_admin():

    def button_click():

        # destination path is assigned by entry1
        dest = entry1.get()

        if check_state.get() == 0:
            # if check_state is 0 then download path is assigned by entry2
            d_link = entry2.get()
        else:
            # change the text in the entry2 to the value of entry1
            text = tk.StringVar()
            text.set(entry1.get())
            entry2.config(textvariable=text)
            # if check_state is 1 then download path is assigned by entry1
            d_link = entry1.get()

        if ctypes.windll.shell32.IsUserAnAdmin():
            print("Running with admin privileges!")
            output_buffer = io.StringIO()

            with contextlib.redirect_stdout(output_buffer):
                run_command(dest, d_link)

            out1 = output_buffer.getvalue()
            label3.config(text=out1)

        else:
            # Restart the script with admin privileges
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
            root.destroy()
            sys.exit()

    # Set main configurations of the tkinter GUI
    root = tk.Tk()
    root.overrideredirect(False)
    root.title("ChromeDriver Updater")
    root.geometry('780x295+100+100')
    root.resizable(0, 0)
    root.config(background='#242423')

    # Set configurations for Labels, Entries, and Buttons
    label1 = tk.Label(root, text="Destination path:", font=('Consolus', 10, 'bold'), background='#242423', foreground='#79bd9a')
    label2 = tk.Label(root, text="Download path (Temp):", font=('Consolus', 10, 'bold'), background='#242423', foreground='#79bd9a')
    entry1 = tk.Entry(root, font=('Consolus', 10, 'bold'), width=80, background='#a8dba8', foreground='#0b486b')
    entry2 = tk.Entry(root, font=('Consolus', 10, 'bold'), width=80, background='#a8dba8', foreground='#0b486b')
    button1 = tk.Button(root, text="Check & Update Driver", wraplength=80, height=5, width=10, font=('Consolus', 12), command=button_click, background='#3b8686', foreground='#a8dba8')
    label3 = tk.Label(root, width=70, height=10, relief=tk.SUNKEN, text="", anchor='nw', wraplength=400, justify="left", font=('Consolus', 10, 'bold'), background='#002833', foreground='#29bbac')

    # Set check state variable and configurations for Checkbutton
    check_state = tk.IntVar()
    check1 = tk.Checkbutton(root, text="Use Destination path as Download path", font=('Consolus', 10), background='#242423', foreground='#79bd9a', variable=check_state)

    # Configure the layout of the form
    root.columnconfigure(0, weight=0)
    root.columnconfigure(1, weight=0)

    root.rowconfigure(0, weight=0)
    root.rowconfigure(1, weight=0)
    root.rowconfigure(2, weight=0)
    root.rowconfigure(3, weight=0)

    # Place Labels, Entries, Button and Checkbox
    label1.grid(row=0, column=0, pady=8, padx=10)
    entry1.grid(row=0, column=1, pady=8, padx=0, sticky='w')
    label2.grid(row=1, column=0, pady=5, padx=10)
    entry2.grid(row=1, column=1, pady=5, padx=0, sticky='w')
    check1.grid(row=2, column=1, pady=5, padx=0, sticky='w')
    button1.grid(row=3, column=0, sticky='ew', padx=30)
    label3.grid(row=3, column=1, sticky='news')

    root.mainloop()


if __name__ == '__main__':
    hide_console()
    try:
        run_as_admin()
    except ConnectionError as msg8:
        print(msg8)


