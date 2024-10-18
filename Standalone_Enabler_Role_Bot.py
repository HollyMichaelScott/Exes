import sys
from dateutil import parser
from datetime import datetime
import asyncio
import tkinter as tk
from tkinter import Message, Toplevel, filedialog, messagebox
import pandas as pd
import win32com.client
import time
import pythoncom    
import subprocess
import pyperclip
pythoncom.CoInitialize()


default_arrangements_auth = ['Role Name','Role Description','Auth. Object','Status','Message']



glob_sysname = ''
glob_clientId = ''
glob_username = ''
glob_password = ''
excel_file_path=''
df = pd.DataFrame()

def validate_date(st_end,date_str):

    try:
        date_obj = datetime.strptime(date_str, "%d.%m.%Y")
    except ValueError:

        try:
            date_obj = parser.parse(date_str)

            return False
        except ValueError:
            return False
    

    today = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)

    if st_end =="start":
        if date_obj < today:
            return False
    
    return True





def open_file_dialog():
    global excel_file_path

    root = tk.Tk()
    root.withdraw()
    

    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx")])
    excel_file_path = file_path
    return file_path

def clean_data_frame(data_frame):
    # Remove rows where all elements are NaN
    cleaned_df = data_frame.dropna(how='all')
    return cleaned_df

def read_excel(file_path):
    global df
    try:
        df = pd.read_excel(file_path, sheet_name='Enabler_master')
        df = df.fillna('')
        columns = df.columns.tolist()

        
        print(''.join(columns).lower(),'==',''.join(default_arrangements_auth).lower())
        if ''.join(columns).lower() != ''.join(default_arrangements_auth).lower():
            messagebox.showerror("Column structure mismatch","\nThe column structure does not match template.")
            sys.exit()
        result_dict = {}

# Loop through the DataFrame and populate the dictionary
        for index, row in df.iterrows():
            key = row[0].lower()   # First column as key (lowercased)
            description = row[1]   # Second column as description
            value = row[2]         # Third column as value
            row_num = index        # Get the row number (index)

            # Check if the key already exists in the dictionary
            if key in result_dict:
                result_dict[key][1].append(value)      # Add the value to the list of values
                result_dict[key][2].append(row_num)    # Add the row number to the list of row numbers
            else:
                result_dict[key] = [description, [value], [row_num]]  # Initialize with description, value list, and row number list
  # Initialize with description, value list, and row number
  # Initialize with description and list of one value

# Convert sets back to lists for the final dictionary
        result_dict = {k: list(v) for k, v in result_dict.items()} 
        print("Result_dict",result_dict)
        return result_dict
            
        
    except PermissionError as p:
        messagebox.showerror("Error",p)
        subprocess.run(['taskkill', '/F', '/IM', 'saplogon.exe'], check=True)
        sys.exit()

def process_group(group,row_indices):
    # Initialize lists for each field
    description = ''
    role_type = ''
    transaction_code = []
    object_field = []
    

    for row in group:
        if not description and row[0]:
            description = row[0].lower().strip()
        if not role_type and row[1]:
            role_type = row[1].lower().strip()
        if not transaction_code and row[2]:
            transaction_code = row[2].lower().strip()
        # Collect values for the object field
        if row[3] and row[3] not in object_field:
            object_field.append(row[3].lower().strip())
        result = [description, role_type, transaction_code, object_field] 
        result.append([i+1 for i in row_indices])
    return result

# Process each key in the dictionary




def submit_form(entries,root,download_excel):
    global glob_sysname, glob_clientId, glob_username, glob_password
    
    sysname = entries['System Name'].get()
    clientId = entries['Client ID'].get()
    username = entries['Username'].get()
    password = entries['Password'].get()
    
    if not sysname or not clientId or not username or not password:
        messagebox.showerror("Error", "All fields are required!")
    else:

        glob_sysname = sysname
        glob_clientId = clientId
        glob_username = username
        glob_password = password
        print(download_excel.get())
        if download_excel.get():

            download_excel()

        root.quit()  # Quit the main loop
        root.destroy()  # Destroy the form window



async def main():
    file_path = open_file_dialog()
    
    if file_path:
        data = read_excel(file_path)

        global cleaned_data
        cleaned_data = data
        
        
        await GUI_code()
    else:
        print("No file selected.")


def download_excel():

    
    column_headers = ['Role Name','Role Description', 'Role Type','Type of Object To be Added','Object','Status','Message']


    df = pd.DataFrame(columns=default_arrangements_auth)


    excel_file_path = 'Enabler_Role_Creation.xlsx'

    df.to_excel(excel_file_path, sheet_name="Enabler_master",index=False)

    messagebox.showinfo("Excel file downloaded",f"Empty Excel file created with headers in {excel_file_path}")

def update_excel_row(row_index, message,status):
        try:
            global df, excel_file_path
            print("Row",row_index)
            
            df.loc[row_index, "Message"] = message
            df.loc[row_index, "Status"] = status
            
            
        except PermissionError as p:
            messagebox.showerror("Error","You might have no permission to change the excel file or might have it opened.\n\nPlease ensure you have closed the excel file\n\n")
            subprocess.run(['taskkill', '/F', '/IM', 'saplogon.exe'], check=True)
            sys.exit()

async def GUI_code():
    
    global thisSelect
    thisSelect=""
    #global action
    

    idx=0
    
    
    try:
        sapgui = win32com.client.GetObject("SAPGUI")
        
    except:
        messagebox.showerror("Logon screen unavailable","\nPlease rerun the bot: User cannot select file before SAP Logon GUI window opens.\n\nPlease wait for the SAP logon screen to appears.")
        subprocess.run(['taskkill', '/F', '/IM', 'saplogon.exe'], check=True)
        sys.exit()
    
    application = sapgui.GetScriptingEngine
    print(glob_sysname)
    connection = application.Children(0)


    session = connection.Children(0)

    if session.findById("wnd[0]/sbar").MessageType == 'E':
            error_message = session.findById("wnd[0]/sbar").text
            
            messagebox.showerror("Error", f"{error_message}")
            subprocess.run(['taskkill', '/F', '/IM', 'saplogon.exe'], check=True)
            sys.exit()
    else:
            print("Logon successful")
            
    
    session.findById("wnd[0]/tbar[0]/okcd").text = "PFCG"
    session.findById("wnd[0]").sendVKey(0)
    print("Data set: ",cleaned_data)
    try:
        err=session.findById("wnd[0]/sbar").text
        print("Error in SU01 is:",err)
        if err:
            if err.lower() == "You are not authorized to use transaction PFCG".lower():
                messagebox.showerror("Unauthorized access error",f"\nUser {glob_username} is not authorized to use the transaction PFCG")
                subprocess.run(['taskkill', '/F', '/IM', 'saplogon.exe'], check=True)
                sys.exit()
            else:
                messagebox.showerror("Miscellaneous error",f"\nSome other error occured going into the transaction, please try running")
                subprocess.run(['taskkill', '/F', '/IM', 'saplogon.exe'], check=True)
                sys.exit()

    except Exception as e:
        print("SU01 try catch: ",e)

    keys = list(cleaned_data.keys())
    session.findById("wnd[0]").maximize()
    for key in (keys):
        print("Working On: ",key)
        if(key)=='' or key=='nan':

            for indv in cleaned_data[key][2]:
                update_excel_row(int(indv),'Role ID is missing','Error')
            continue


        try:
            user= key
            print(user)
            # time.sleep(5)
            print("Here-1")
            session.findById("wnd[0]/usr/ctxtAGR_NAME_NEU").text = ''
            session.findById("wnd[0]/usr/ctxtAGR_NAME_NEU").text = f"{user}"
            print("Here0")
            session.findById("wnd[0]/usr/btn%#AUTOTEXT003").press()

            print("Here2")
            session.findById("wnd[0]/usr/txtS_AGR_TEXTS-TEXT").text = cleaned_data[key][0]
            
            print("Here")

        except Exception as e:
            for data in cleaned_data[key][2]:
                update_excel_row(int(data), f"Role {user} already exists","Error")
            continue

        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB5").select()
        session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB5/ssubSUB1:SAPLPRGN_TREE:0350/btnPROFIL1").press()
        session.findById("wnd[1]/tbar[0]/btn[19]").press()
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        
        col=0
        
        #print("\n\n\nedit:::",cleaned_data[key][i])
        for i in range(len(cleaned_data[key][1])):
            print("Col: ",col,"\nI: ",i,"cleaned_data:",cleaned_data[key][1][i])
            session.findById("wnd[0]/tbar[1]/btn[45]").press()
            session.findById(f"wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = cleaned_data[key][1][i]
            session.findById("wnd[0]").sendVKey(0)
            error_message = session.findById("wnd[0]/sbar").text
            if(error_message.lower()!=f'authorization object {cleaned_data[key][1][i]} does not exist'.lower()):
                col+=1
                
            else:
                session.findById(f"wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = ''
                session.findById("wnd[0]").sendVKey(0)
                continue
        


        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        #time.sleep(3)
        print("Hello")
        session.findById("wnd[1]/usr/btnBUTTON_1").press()
        #time.sleep(5)
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        for i in range(len(cleaned_data[key][2])):
            update_excel_row(cleaned_data[key][2][i],'Completed','Done')


    print("DF: ",df)
    try: 
        with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a') as writer:
            writer.book.remove(writer.book['Enabler_master'])
            df.to_excel(writer, index=False, sheet_name="Enabler_master")
    except Exception as e:
        messagebox.showerror("Error",e)
        print(e)
        

    messagebox.showinfo("Complete","Bot execution completed")
    sys.exit()
    
    
    
    


    
                    
def complete_msg():
    root = tk.Tk()
    #top = Toplevel()
    root.title('Success')
    Message(root, text='''Bot execution complete.
            \n\nPlease review the excel file.''',pady=40,padx=40).pack()
    root.after(5000,sys.exit)
    root.mainloop()
    

def on_closing(root):

    root.quit()
    root.destroy() 
    sys.exit()



def show_form():
    root = tk.Tk()
    root.title("Enter Parameters")
    
    labels = ['System Name', 'Client ID', 'Username', 'Password']
    entries = {}
    
    for label in labels:
        frame = tk.Frame(root)
        frame.pack(fill='x')
        
        lbl = tk.Label(frame, text=label, width=15)
        lbl.pack(side='left')
        
        if label.lower() == 'password':
            entry = tk.Entry(frame, show='*')
        else:
            entry = tk.Entry(frame)
        entry.pack(fill='x', expand=True)
        
        entries[label] = entry
    
    # Add a checkbox for downloading the Excel template
    download_template_var = tk.BooleanVar()
    download_template_checkbox = tk.Checkbutton(root, text="Download excel template file", variable=download_template_var,command=download_excel)
    download_template_checkbox.pack(pady=10)
    
    submit_button = tk.Button(root, text="Submit", command=lambda: submit_form(entries,root,download_template_var))
    submit_button.pack(pady=10)
    
    root.protocol("WM_DELETE_WINDOW", lambda: on_closing(root))  # Handle window close

    root.mainloop()           
        

if __name__ == "__main__":
    # Show the checklist message box
    global root
    root = tk.Tk()
    root.withdraw()
    checklist_message = (
        "Please ensure the following before proceeding:\n\n"
        "1. SAP Logon is installed on your system.\n\n\n\n"
        "2. Please make sure, you have the access to the system intended to be used.\n\n\n\n"
        "3. Please ensure the excel file you select is saved and closed.\n\n\n\n"
    )
    messagebox.showinfo("Pre-Execution Checklist", checklist_message)
    choice=messagebox.askokcancel("Attention Required!","By clicking 'OK' bot will close all open SAP Logon instances. \nPlease make sure you save your work to ensure no loss.\n\n")

    if choice:
        try:
            subprocess.run(['taskkill', '/F', '/IM', 'saplogon.exe'], check=True)
            
        except subprocess.CalledProcessError:
            print("No existing SAP Logon processes were found.")

        
        show_form()
        try:
            

            subprocess.check_call(['C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\sapshcut.exe', f'-system={glob_sysname}', f'-client={glob_clientId}', f'-user={glob_username}', f'-pw={glob_password}', '-language=EN'])
            #subprocess.check_call(['C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\sapshcut.exe', f'-system=ds4', f'-client=400', f'-user=TUMOHAN', f'-pw=Master@111', '-language=EN'])
            
            
        
        except:
            try:
                subprocess.check_call(['C:\\Program Files\\SAP\\FrontEnd\\SAPgui\\sapshcut.exe', f'-system={glob_sysname}', f'-client={glob_clientId}', f'-user={glob_username}', f'-pw={glob_password}', '-language=EN'])
            except:
                messagebox.askquestion("Error", "You might have input an incorrect password or not have access to the system, please execute the bot again")
                subprocess.run(['taskkill', '/F', '/IM', 'saplogon.exe'], check=True)
                sys.exit()
                
    else:
        sys.exit()

    asyncio.run(main())