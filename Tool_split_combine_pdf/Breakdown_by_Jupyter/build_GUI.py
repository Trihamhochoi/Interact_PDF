import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from define_function import Split_document, Combine_document
import os
import time

desktop_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')


# --------------------------------------- DEFINE FUNCTION -------------------------------------#
def choose_files_folder(entry, kind: int):
    if kind == 3:
        entry.configure(state='normal')
        files_list = list(filedialog.askopenfilenames(title='Choose file',
                                                      initialdir=desktop_path,
                                                      filetypes=[("pdf files", "*.pdf"),
                                                                 ("image", "*.jpeg"),
                                                                 ("image", "*.png"),
                                                                 ("image", "*.jpg"),
                                                                 ('all files', '.*')])
                          )
        path = "{}".format('\n'.join(files_list))
        entry.delete('1.0', tk.END)
        entry.insert('end', path)

    elif kind == 2:
        entry.configure(state='normal')
        path = "{}".format(filedialog.askopenfilename(title='Choose file',
                                                      initialdir=desktop_path,
                                                      filetypes=[("pdf files", "*.pdf")])
                           )
        entry.delete(0, tk.END)
        entry.insert(0, path)
    elif kind == 1:
        entry.configure(state='normal')
        path = "{}".format(filedialog.askdirectory(title='Choose folder',
                                                   initialdir=desktop_path,
                                                   mustexist=True))
        entry.delete(0, tk.END)
        entry.insert(0, path)

    entry.configure(state='disabled')


# -----------------------------------------------------------------------------------------------#
def split_fill_complete(*args):
    pdf_text = pdf_path.get()
    output_text = output_path.get()
    split_from_text = split_from.get()
    split_to_text = split_to.get()

    # Validation Rule
    if pdf_text == 'Please select PDF path' or output_text == 'Please select Output location':
        split_button.configure(state='disabled')
    elif pdf_text == '' or output_text == '' or split_from_text == '' or split_to_text == '':
        split_button.configure(state='disabled')
    else:
        split_button.configure(state='normal', text_color='white')


def combine_fill_complete(*args):
    combined_pdfs_text = multiple_files_text.get(index1='1.0', index2=tk.END)
    print(combined_pdfs_text.isspace())
    filename_text = combined_file_name.get()
    location_text = location_path.get()

    # Validation Rule
    if filename_text == 'Please type the file name' or location_text == 'Please select Output Location':
        combine_button.configure(state='disabled')
    elif filename_text == '' or location_text == '':
        combine_button.configure(state='disabled')
    else:
        combine_button.configure(state='normal', text_color='white')


# -----------------------------------------------------------------------------------------------#
def openNewWindow(message):
    newWindow = ctk.CTkToplevel(root)
    newWindow.title('Information is extracted')
    newWindow.geometry('700x400')
    info_text = ctk.CTkTextbox(master=newWindow,
                               width=600,
                               height=300)
    info_text.pack(side=tk.LEFT, expand=True)
    info_text.insert(tk.END, message)
    info_text.configure(state='disabled')
    sb = tk.Scrollbar(master=newWindow)
    sb.pack(side=tk.RIGHT, fill=tk.BOTH)
    info_text.configure(yscrollcommand=sb.set)
    sb.configure(command=info_text.yview)
    newWindow.mainloop()


# -----------------------------------------------------------------------------------------------#
def split_pdf(pdf_enter,
              output_enter,
              split_from_number,
              split_to_number):
    file_pdf_path = pdf_enter.get()
    save_location_path = output_enter.get()
    from_number = int(split_from_number.get())
    to_number = int(split_to_number.get())

    # Check whether file exists in system:
    if os.path.exists(save_location_path):
        is_ok = messagebox.askokcancel(title='Implement', message='Do you would like to continue proceeding?')
        label_loading.configure(text='Please wait, Files are processing...')
        label_loading.update_idletasks()

        if is_ok:
            message_dict = Split_document(source_doc=file_pdf_path,
                                          start_page=from_number,
                                          end_page=to_number,
                                          save_location=save_location_path)
            time.sleep(2)
            print('end sleep')
            label_loading.configure(text='FINISHED!!!')

            if 'Error' not in message_dict.keys():
                message = f"File name: {message_dict['pdf_name']}" \
                          f"\n\nSplit from: page {message_dict['split_from']}" \
                          f"\n\nSplit to: page {message_dict['split_to']}" \
                          f"\n\nSave Location: {message_dict['Save_location']}" \
                          f"\n\nPDF Split name: {message_dict['pdf_split_name']}"
            else:
                message = ''
                for k, v in message_dict.items():
                    message += f'{k}: {v}'

            openNewWindow(message=message)

        else:
            message = 'No such directory in System.\nYou type wrong the PDF path or Excel. Please choose again'
            messagebox.showinfo(title=' Extract Data', message=message)
            return None


def combine_pdf(pdf_enter,
                location_enter,
                file_name):
    file_lists = pdf_enter.get(index1='1.0', index2=tk.END).split("\n")
    files_pdf_path = [file for file in file_lists if os.path.isfile(file)]
    print(files_pdf_path)
    save_location_path = location_enter.get()
    comb_file_name = file_name.get()

    # Check whether file exists in system:
    if os.path.exists(save_location_path):
        is_ok = messagebox.askokcancel(title='Implement', message='Do you would like to continue proceeding?')
        combined_label_loading.configure(text='Please wait, Files are processing...')
        combined_label_loading.update_idletasks()

        if is_ok:
            message_dict = Combine_document(files_list=files_pdf_path,
                                            output_filename=comb_file_name,
                                            save_location=save_location_path)
            time.sleep(2)
            print('end sleep')
            combined_label_loading.configure(text='FINISHED!!!')

            if 'Error' not in message_dict.keys():
                message = f"Combined Filename: {message_dict['combine_filename']}" \
                          f"\n\nNumber Page: {message_dict['page_num']} pages" \
                          f"\n\nSave Location: {message_dict['Save_location']}" \
                          f"\n\nTotal combine PDF files: {message_dict['Total_file']} files"
            else:
                message = ''
                for k, v in message_dict.items():
                    message += f'{k}: {v}'

            openNewWindow(message=message)

        else:
            message = 'No such directory in System.\nYou type wrong the PDF path or Excel. Please choose again'
            messagebox.showinfo(title=' Extract Data', message=message)
            return None


###################################BUID TOOL##################################################

ctk.set_appearance_mode("light")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

root = ctk.CTk()
root.title("Tool to convert PDF file to Excel file")
root.geometry('600x350')

s = ttk.Style()
s.configure('TNotebook.Tab', font=('URW Gothic L', '14', 'bold'))

# Build Notebook
notebook = ttk.Notebook(master=root)
notebook.pack(pady=15, padx=10)

# Build 2 Frame
combine_frame = ctk.CTkFrame(master=notebook,
                             width=600,
                             height=300,
                             corner_radius=40)

split_frame = ctk.CTkFrame(master=notebook,
                           width=550,
                           height=300,
                           corner_radius=40)

combine_frame.pack(fill='both', expand=True)
split_frame.pack(fill='both', expand=True)

# set size of the window and add row and column
# split_frame.rowconfigure(4)
# split_frame.columnconfigure(3)
# Add frame to notebook
notebook.add(combine_frame, text="Combine Function")
notebook.add(split_frame, text="Split Function")

# ------------------------------- SPLIT PDF -----------------------------------------#
# Build text string
pdf_path = tk.StringVar(master=split_frame, value='Please select PDF path')
output_path = tk.StringVar(master=split_frame, value='Please select Output location')
split_from = tk.StringVar(master=split_frame)
split_to = tk.StringVar(master=split_frame)

# Traceback to check validation whether info is filled completely
pdf_path.trace("w", split_fill_complete)
output_path.trace("w", split_fill_complete)
split_from.trace("w", split_fill_complete)
split_to.trace("w", split_fill_complete)

# PDF Label
pdf_label = ctk.CTkLabel(master=split_frame,
                         text="PDF File")
pdf_label.grid(column=0,
               row=0,
               pady=15,
               padx=15,
               )

# PDF Entry
pdf_entry = ctk.CTkEntry(master=split_frame,
                         textvariable=pdf_path,
                         width=250,
                         )
pdf_entry.grid(column=1,
               row=0,
               pady=15,
               padx=15,
               )

# Button to select a PDF file
pdf_button = ctk.CTkButton(master=split_frame,
                           text_color='white',
                           text='Add file',
                           command=lambda: choose_files_folder(pdf_entry, 2),
                           width=100)

pdf_button.grid(column=2,
                row=0,
                padx=10,
                )

# Output Label
output_label = ctk.CTkLabel(master=split_frame, text="Output Folder")
output_label.grid(column=0,
                  row=1,
                  )

# Output Entry
output_entry = ctk.CTkEntry(master=split_frame,
                            textvariable=output_path,
                            width=250
                            )
output_entry.grid(column=1,
                  row=1,
                  )

# Button to select Output folder
output_button = ctk.CTkButton(master=split_frame,
                              text_color='white',
                              text='Add Folder',
                              command=lambda: choose_files_folder(output_entry, 1),
                              width=100)
output_button.grid(column=2,
                   row=1,
                   padx=10,
                   )

# Split from Label
split_from_label = ctk.CTkLabel(master=split_frame, text="Split From")
split_from_label.grid(column=0,
                      row=2,
                      pady=15,
                      padx=15,
                      )

# Split from Entry
split_from_entry = ctk.CTkEntry(master=split_frame, textvariable=split_from, width=50)
split_from_entry.grid(column=1,
                      row=2,
                      pady=15,
                      padx=15,
                      sticky='W'
                      )

# Split to Label
split_to_label = ctk.CTkLabel(master=split_frame, text="Split To")
split_to_label.grid(column=0,
                    row=3,
                    pady=15,
                    padx=15,
                    sticky="NESW")

# Split to Entry
split_to_entry = ctk.CTkEntry(master=split_frame, textvariable=split_to, width=50)
split_to_entry.grid(column=1,
                    row=3,
                    pady=15,
                    padx=15,
                    sticky='W'
                    )

# Build Split Button
split_button = ctk.CTkButton(master=split_frame,
                             text='Split PDF',
                             command=lambda: split_pdf(pdf_enter=pdf_path,
                                                       output_enter=output_path,
                                                       split_from_number=split_from,
                                                       split_to_number=split_to),
                             state='disabled'
                             )
split_button.grid(column=1,
                  row=5,
                  pady=15,
                  padx=10,
                  sticky='WE'
                  )
# Loading Label
label_loading = ctk.CTkLabel(master=split_frame, text='')
label_loading.grid(column=1, row=4)

# ------------------------------- COMBINE PDF TOOLS-----------------------------------------#
# Build text string
location_path = tk.StringVar(master=combine_frame,
                             value='Please select Output Location')
combined_pdfs_path = tk.StringVar(master=combine_frame,
                                  value='Please select PDF files to combine')
combined_file_name = tk.StringVar(master=combine_frame,
                                  value='Please type the file name')

# Traceback to check validation whether info is filled completely
location_path.trace("w", combine_fill_complete)
combined_pdfs_path.trace("w", combine_fill_complete)
combined_file_name.trace("w", combine_fill_complete)

# ------------------------- Location Folder --------------------------#
# Location Label
location_label = ctk.CTkLabel(master=combine_frame, text="Output Folder")
location_label.grid(column=0,
                    row=0,
                    padx=10,
                    pady=10,
                    sticky='w'
                    )

# Location Entry
location_entry = ctk.CTkEntry(master=combine_frame,
                              textvariable=location_path,
                              width=250
                              )
location_entry.grid(column=1,
                    row=0,
                    padx=10,
                    pady=10
                    )

# Button to select Location folder
location_button = ctk.CTkButton(master=combine_frame,
                                text_color='white',
                                text='Add Folder',
                                command=lambda: choose_files_folder(location_entry, 1),
                                width=100)
location_button.grid(column=2,
                     row=0,
                     padx=20,
                     pady=10,
                     sticky='W'
                     )

# ------------------------- Combine File name--------------------------#
filename_label = ctk.CTkLabel(master=combine_frame, text="Output Name")
filename_label.grid(column=0,
                    row=1,
                    padx=10,
                    pady=10,
                    sticky='w'
                    )

# Location Entry
filename_entry = ctk.CTkEntry(master=combine_frame,
                              textvariable=combined_file_name,
                              width=250
                              )
filename_entry.grid(column=1,
                    row=1,
                    padx=10,
                    pady=10
                    )

# ------------------------- Mutiple files PDF--------------------------#
multiple_files_label = ctk.CTkLabel(master=combine_frame, text='PDF files')
multiple_files_label.grid(column=0,
                          row=2,
                          padx=10,
                          pady=10,
                          sticky='NW'
                          )

multiple_files_text = ctk.CTkTextbox(master=combine_frame,
                                     width=250,
                                     height=100)
multiple_files_text.grid(column=1,
                         row=2,
                         padx=10,
                         pady=10,
                         )
multiple_files_button = ctk.CTkButton(master=combine_frame,
                                      text_color='white',
                                      text='Add Files',
                                      command=lambda: choose_files_folder(multiple_files_text, 3),
                                      width=100)

multiple_files_button.grid(column=2,
                           row=2,
                           padx=15,
                           pady=10,
                           sticky='NW'
                           )
# Build Combine Button
combine_button = ctk.CTkButton(master=combine_frame,
                               text='Combine PDF',
                               command=lambda: combine_pdf(pdf_enter=multiple_files_text,
                                                           location_enter=location_path,
                                                           file_name=combined_file_name),
                               state='disabled'
                               )
combine_button.grid(column=1,
                    row=4,
                    pady=15,
                    padx=10,
                    sticky='WE'
                    )
# Loading Label
combined_label_loading = ctk.CTkLabel(master=combine_frame, text='')
combined_label_loading.grid(column=1, row=3)

root.mainloop()
