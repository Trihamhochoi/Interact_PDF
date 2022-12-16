import tkinter as tk
import customtkinter
from tkinter import messagebox
from tkinter import filedialog
import define_function
import os
import time

desktop_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')


# ------------------FUNCTION-----------------------
def choose_folder(entry):
    path = "{}".format(filedialog.askdirectory(title='Choose folder',
                                               initialdir=desktop_path,
                                               mustexist=True))
    entry.delete(0, tk.END)
    entry.insert(0, path)


# ---------------------------------------------------------------#
def openNewWindow(message):
    newWindow = customtkinter.CTkToplevel(window)
    newWindow.title('Information is extracted')
    newWindow.geometry('500x700')
    info_text = customtkinter.CTkTextbox(master=newWindow, width=450, height=600)
    info_text.pack(side=tk.LEFT, expand=True)
    info_text.insert(tk.END, message)
    info_text.configure(state='disabled')
    sb = tk.Scrollbar(master=newWindow)
    sb.pack(side=tk.RIGHT, fill=tk.BOTH)
    info_text.configure(yscrollcommand=sb.set)
    sb.configure(command=info_text.yview)
    newWindow.mainloop()


# ---------------------------------------------------------------#
def get_data(pdfentry, excelentry):
    pdf_path = pdfentry.get()
    excel_path = excelentry.get()
    function = dropbox_function.get()

    # Check whether file exists in system:
    if os.path.exists(pdf_path) and os.path.exists(excel_path):
        is_ok = messagebox.askokcancel(title='Implement', message='Do you would like to continue proceeding?')
        label_loading.configure(text='Please wait, Files are processing...')
        label_loading.update_idletasks()

        if is_ok:
            # Kind 1: General case
            if function == 'General Kind':
                message = define_function.convert_pdffolder_to_excelfolder(pdf_path=pdf_path,
                                                                           xlsx_path=excel_path,
                                                                           function_kind=1)
            # Kind 2:
            elif function == 'Citi Checking Kind':
                message = define_function.convert_pdffolder_to_excelfolder(pdf_path=pdf_path,
                                                                           xlsx_path=excel_path,
                                                                           function_kind=2)
            # Kind 3:
            elif function == 'CheckNo':
                message = define_function.convert_pdffolder_to_excelfolder(pdf_path=pdf_path,
                                                                           xlsx_path=excel_path,
                                                                           function_kind=3)
            # Kind 4:
            elif function == 'BoH/Panacea/ChaseCC/BOW Check/CapitalOne/Hanmi':
                message = define_function.convert_pdffolder_to_excelfolder(pdf_path=pdf_path,
                                                                           xlsx_path=excel_path,
                                                                           function_kind=4)
            else:
                message = "No proper function"

            time.sleep(2)
            print('end sleep')
            label_loading.configure(text='FINISHED!!!')
            openNewWindow(message=message)
        else:
            message = 'No such directory in System.\nYou type wrong the PDF path or Excel. Please choose again'
            messagebox.showinfo(title=' Extract Data',message= message)
            return None


# ----------------------------------------------------------------------------------------------#
def fill_complete(*args):
    pdf_text = PDF_path.get()
    excel_text = Excel_path.get()
    func_text = function_var.get()

    if pdf_text == '' or pdf_text == 'Please fill PDF path':
        button.configure(state='disabled')
    elif excel_text == '' or excel_text == 'Please fill Excel path':
        button.configure(state='disabled')
    elif func_text == '' or func_text == 'Please choose the function':
        button.configure(state='disabled')
    else:
        button.configure(state='normal', text_color='white')


# -------------------------------------SET UP GUI FOR TOOL--------------------------------------#
customtkinter.set_appearance_mode("light")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

window = customtkinter.CTk()
window.title("Tool to convert PDF file to Excel file")
window.geometry('700x300')

# Build Frame
frame_1 = customtkinter.CTkFrame(master=window)
frame_1.pack(pady=20, padx=60, fill="both", expand=True)

# Build text string
PDF_path = tk.StringVar(master=frame_1, value='Please fill PDF path')
Excel_path = tk.StringVar(master=frame_1, value='Please fill Excel path')
function_var = tk.StringVar(master=frame_1, value='Please choose the function')

# Traceback to check validation whether info is filled completely
PDF_path.trace("w", fill_complete)
Excel_path.trace("w", fill_complete)
function_var.trace("w", fill_complete)

# PDF Label
pdf_label = customtkinter.CTkLabel(master=frame_1, text="PDF Folder")
pdf_label.grid(column=1, row=0, pady=15, padx=15)

# PDF Entry
pdf_entry = customtkinter.CTkEntry(master=frame_1, textvariable=PDF_path, width=300)
pdf_entry.grid(column=2, row=0, pady=15,padx=15)

# Add PDF folder path button
pdf_button = customtkinter.CTkButton(master=frame_1,
                                     text_color='white',
                                     text='Add Folder',
                                     command=lambda: choose_folder(pdf_entry),
                                     width=100)

pdf_button.grid(column=3, row=0, padx=10)

# Excel Label
excel_label = customtkinter.CTkLabel(master=frame_1, text="Excel Folder")
excel_label.grid(column=1, row=1)

# Excel Entry
excel_entry = customtkinter.CTkEntry(master=frame_1, textvariable=Excel_path, width=300)
excel_entry.grid(column=2, row=1)

# Add PDF folder path button
excel_button = customtkinter.CTkButton(master=frame_1,
                                       text_color='white',
                                       text='Add Folder',
                                       command=lambda: choose_folder(excel_entry),
                                       width=100)
excel_button.grid(column=3, row=1, padx=10)

# Dropbox to choose function
dropbox_function = customtkinter.CTkComboBox(frame_1,
                                             width=200,
                                             values=["General Kind",
                                                     "Citi Checking Kind",
                                                     "CheckNo",
                                                     'BoH/Panacea/ChaseCC/BOW Check/CapitalOne/Hanmi'],
                                             variable=function_var)

dropbox_function.grid(column=2, row=2, pady=15, padx=10)

# Button
button = customtkinter.CTkButton(master=frame_1,
                                 text='Export',
                                 command=lambda: get_data(pdfentry=pdf_entry, excelentry=excel_entry),
                                 state='disabled'
                                 )

button.grid(column=2, row=4, pady=15, padx=10)

# Loading Label
label_loading = customtkinter.CTkLabel(master=frame_1, text='')
label_loading.grid(column=2, row=3)

window.mainloop()
