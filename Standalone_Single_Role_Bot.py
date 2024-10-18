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

array_columns = ["Work Center", "Bank Area", "Company Code", "Consolidation unit", "Valuation area", 
                "Aspect", "Condition Area", "Consolidated entity", "View", "Purchasing Group", 
                "Purchasing Organization", "Operating Concern", "FM area", "Business Area", 
                "Maintenance planning plant", "Credit Control Area", "Account Type", 
                "Controlling Area", "Warehouse Number / Warehouse Complex", "Storage type", 
                "Location", "Plan Version", "Profit Centers", "Company", "Division", 
                "Maintenance plant", "Transportation Planning Point", "Sales Office", 
                "Sales Group", "Sales Organization", "Shipping Point", "Distribution Channel", 
                "Plant"]

default_arrangements = ['Role Name','Role Description','Object','Status','Message']

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
        df = pd.read_excel(file_path, sheet_name='role_master')
        df = df = clean_data_frame(df)
        df = df.fillna('')
        columns = df.columns.tolist()

        
        print(''.join(columns).lower(),'==',''.join(default_arrangements).lower())
        if ''.join(columns).lower() != ''.join(default_arrangements).lower():
            messagebox.showerror("Column structure mismatch","\nThe column structure does not match template.")
            sys.exit()



        role_dict = {}


 

# Group by 'Role Name' and aggregate the remaining columns
        for role_name, group in df.groupby('Role Name'):
            role_dict[role_name.lower().strip()] = (group.drop(columns=['Role Name']).values.tolist(), group.index.tolist())

        #print(role_dict)
        result = {key: process_group(value[0],value[1]) for key, value in role_dict.items()}
        print(result)
        return result
            
        
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
        # if not role_type and row[1]:
        # #     role_type = row[1].lower().strip()
        # if not transaction_code and row[1]:
        #     transaction_code = row[1].lower().strip()
        # Collect values for the object field
        if row[1] and row[1] not in object_field:
            object_field.append(row[1].lower().strip())
        result = [description, object_field] 
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


    df = pd.DataFrame(columns=default_arrangements)


    excel_file_path = 'Single_Role_Creation.xlsx'

    df.to_excel(excel_file_path, sheet_name="role_master", index=False)

    messagebox.showinfo("Excel file downloaded",f"Empty Excel file created with headers in {excel_file_path}")

def update_excel_row(row_index, message,status):
        try:
            global df, excel_file_path
            print("Row",row_index)
            
            df.loc[row_index-1, "Message"] = message
            df.loc[row_index-1, "Status"] = status
            
            
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
                update_excel_row(indv,'Role ID is missing','Error')
            continue


        try:
            user= key
            print(user)
            # time.sleep(5)
            print("Here-1")
            session.findById("wnd[0]/usr/ctxtAGR_NAME_NEU").text = ''
            session.findById("wnd[0]/usr/ctxtAGR_NAME_NEU").text = f"{user}"
            print("Here0")
            print("Here1")
            session.findById("wnd[0]/usr/btn%#AUTOTEXT003").press()

            print("Here2")
            session.findById("wnd[0]/usr/txtS_AGR_TEXTS-TEXT").text = cleaned_data[key][0]
            
            print("Here")

        except Exception as e:
            for data in cleaned_data[key][2]:
                update_excel_row(data, f"Role {user} already exists","Error")
            continue

        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB9").select()
        session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB9/ssubSUB1:SAPLPRGN_TREE:0321/cntlTOOL_CONTROL/shellcont/shell").pressButton("TB03")
        
        i=0
        for j,value in enumerate(cleaned_data[key][1]):
            
            session.findById(f"wnd[1]/usr/tblSAPLPRGN_WIZARDCTRL_TCODE/ctxtS_TCODES-TCODE[0,{i}]").text = value
            session.findById("wnd[1]").sendVKey(0)
            err=session.findById("wnd[0]/sbar").text
            if err.lower()==f"transaction {value} does not exist".lower():
                cleaned_data[key][1].pop(j)
                print('Cleaned POP:', cleaned_data[key])
                session.findById(f"wnd[1]/usr/tblSAPLPRGN_WIZARDCTRL_TCODE/ctxtS_TCODES-TCODE[0,{i}]").text = ''
                session.findById("wnd[1]").sendVKey(0)
                print(err,' at index',j)
                update_excel_row(j,err,"Error")
                continue

            else:
                i+=1
                continue




        session.findById("wnd[1]/tbar[0]/btn[19]").press()
        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB5").select()
        session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB5/ssubSUB1:SAPLPRGN_TREE:0350/btnPROFIL1").press()
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/btnBUTTON_1").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        for i in cleaned_data[key][2]:
            print(cleaned_data[key][2])
            update_excel_row(i,"Completed","Done")
        continue
        

    

    try: 
        with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a') as writer:
            writer.book.remove(writer.book['role_master'])
            df.to_excel(writer, index=False, sheet_name="role_master")
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