import fitz
import re
import os
import pandas as pd
import numpy as np
import warnings
from tkinter import *
import sys
from tkinter import filedialog
from PIL import Image

warnings.filterwarnings('ignore')


##################### COMBINE FUNCTION ##############################
def Combine_document(files_list: list,
                     output_filename: str,
                     save_location: str = r'C:\Users\trilnd\Desktop\QBO\File PDF\Checks_to_combine\Combine\Output'):
    """
    files_list is group of file user chose before with type is List
    output_filename is name of combined filename the user will name
    save_location is Location include combine-pdf file
    """
    info_dict = dict()
    try:
        # Connect with files_list and check whether file is PDF or not
        for i in range(len(files_list)):
            file = files_list[i]

            # Get filepath and extension
            f, e = os.path.splitext(file)
            if e[1:].lower() != 'pdf':
                # If file is PNG or JPEG will conv
                img = Image.open(file)
                img = img.convert(mode='RGB')
                path_pdf_file = f'{f}.pdf'
                img.save(path_pdf_file, save_all=True)
                print(path_pdf_file)

                # update files_list:
                files_list[i] = path_pdf_file

        # Add all files to summary file
        blank_doc = fitz.open()
        for file in files_list:
            # Parse file PDF
            document = fitz.open(filename=file)
            base_name = os.path.basename(file)[:-4]
            # Insert document
            blank_doc.insert_pdf(docsrc=document)

        # Save the summaried file to PDF file
        combine_file_name = f'combine_{output_filename}.pdf'
        location = os.path.join(save_location, combine_file_name)

        # Store new document in the relevant location
        blank_doc.save(filename=location)

        # Save info of doc to dict
        info_dict['Total_file'] = len(files_list)
        info_dict['combine_filename'] = combine_file_name
        info_dict['page_num'] = blank_doc.page_count
        info_dict['Save_location'] = save_location
        return info_dict

    except Exception as e:
        info_dict['Error'] = e
        return info_dict


##################### SPLIT FUNCTION ##############################
def Split_document(source_doc,
                   start_page: int,
                   end_page: int,
                   save_location):
    """
    - Source document is the location of PDF file,
    Ex: C:/Users/trilnd/Desktop/QBO/File PDF/Bank Statement/BofA.pdf

    - Start_page is first page number user want to be splited
    - End_page is last page number user want to be splited
    - Save location: the folder contains splited PDF file
    Ex: C:/Users/trilnd/Desktop/QBO/File PDF/Bank Statement
    """
    info_dict = dict()

    try:
        # Connect with soure_doc
        document = fitz.open(filename=source_doc)
        base_name = os.path.basename(source_doc)[:-4]
        split_file_name = f'split_{base_name}.pdf'
        location = os.path.join(save_location, split_file_name)
        num_pages = document.page_count

        # Create blank document
        blank_doc = fitz.open()

        # Insert relevant page to new doc
        blank_doc.insert_pdf(docsrc=document, from_page=start_page - 1, to_page=end_page-1)

        # Save new document in the metioned location
        blank_doc.save(filename=location)

        # Save info of doc to dict
        info_dict['pdf_name'] = base_name
        info_dict['page_num'] = num_pages
        info_dict['split_from'] = start_page
        info_dict['split_to'] = end_page
        info_dict['Save_location'] = save_location
        info_dict['pdf_split_name'] = split_file_name
    except Exception as e:
        info_dict['Error'] = e
        return info_dict
    else:
        return info_dict
