import pandas as pd
from PyPDF2 import PdfReader
import re
import os
import warnings
import fitz
from PyPDF2.errors import PdfReadError
import openpyxl

warnings.filterwarnings('ignore')
pd.set_option('display.width', 1000,
              'display.max_columns', 10,
              'display.max_colwidth', 100)


# Kind 1: Apply general case except Citi checking
def convert_pdf_to_csv(pdf_path: str, excel_path='C:/Users/trilnd/Desktop/QBO/Test_file/'):
    """
    pdf_path: the directory of PDF file. For example: 'C:/Users/trilnd/Desktop/QBO/File PDF/Citi CC.pdf'
    Kind_extract_pattern is Format of Trans date and Post date.
    Will be divide into 2 kinds:
        pattern 1: Abbreviations of Month name and 2 digits o day number (For example: Sep 12, Sep 05)
        pattern 2: Date format (dd/mm or mm/dd)
    If the user choose pattern 1: type 1. Pattern 2: type 2
    excel_name: Name the excel file which was converted from PDF file
    """

    # Function to get description

    def get_description(trans, len_upper, len_lower):
        if re.search('www.53.com/businessbanking', all_text_page):
            return trans[len_upper + len_lower + 1:]
        else:
            # check space:
            check_space = pat + r'(?=\s)'
            if re.search(check_space, trans) is not None:
                return trans[len_upper + 1:(len(trans) - len_lower)]
            else:
                return trans[len_upper:(len(trans) - len_lower)]

    pattern_1 = r"^([A-Z][a-z]{2}\s?\d{1,2})"
    pattern_2 = r'^((\d{1,2}/\d{1,2})(/\d{2})?)'
    pattern_3 = r'^((\d{1,2}-\d{1,2})(-\d{2})?)'
    base_name = os.path.basename(pdf_path)[0:-4]
    pdf_dict = {base_name: {'Status': '',
                            'Transaction': '',
                            'Dataframe': ''}}

    # 1: Check whether the path is correct or wrong
    try:
        with open(pdf_path, 'rb') as file:
            print(f"{base_name}: Connect file successful")
            pdf_dict[base_name]["Status"] = "Connect file successful"
    except IOError:
        print("Can not find the file directory or file name does not exist")
        pdf_dict[base_name]["Status"] = "Can not find the file directory or file name does not exist"
        return pdf_dict

    # 2: Implement to convert file to pdf variable and print total number pages
    else:
        try:
            pdf_file = PdfReader(pdf_path, strict=True)
        except Exception as e:
            pdf_dict[base_name]["Status"] = f"Error: {e}"
            return pdf_dict
        else:
            number_page = pdf_file.getNumPages()
            # print("Total number pages of File PDF are",number_page)

            # 3: Convert pdf file to the string
            all_text_page = ''
            for page_num in range(0, pdf_file.getNumPages()):
                try:
                    text = pdf_file.getPage(page_num).extractText()
                except PdfReadError as e:
                    pdf_dict[base_name]["Status"] = e
                    return pdf_dict
                else:
                    all_text_page += (text + '\n')

            # Check if the file contain text or numeric?
            if re.search(r'.', all_text_page) is None:
                print("This PDF file is Scan data")
                pdf_dict[base_name]["Status"] = "This PDF file is Scan data"
                return pdf_dict
            else:
                all_text_list = [line.strip() for line in all_text_page.split('\n')]

            # Select amount pattern if this PDF is AMEX or Citi or capital-one => amount contains $
            if re.search('www.americanexpress.com',
                         all_text_page) or re.search('www.citicards.com',
                                                     all_text_page) or re.search('capitalone.com',
                                                                                 all_text_page):
                num_pat = r'(-?\s?\$\d{1,3}(?:,\d{3})*\.\d{2})'
            else:
                num_pat = r'(-?\s?\$?\d{1,3}(?:,\d{3})*\.\d{2})'

            # 4 Select proper date pattern for PDF file
            count_p1 = sum([1 for line in all_text_list if re.search(pattern_1, line)])
            count_p2 = sum([1 for line in all_text_list if re.search(pattern_2, line)])
            count_p3 = sum([1 for line in all_text_list if re.search(pattern_3, line)])

            if count_p1 > count_p2 and count_p1 > count_p3:
                kind_extract_pattern = 1
                pat = pattern_1

            elif count_p2 > count_p1 and count_p2 > count_p3:
                kind_extract_pattern = 2
                pat = pattern_2

            elif count_p3 > count_p1 and count_p3 > count_p2:
                kind_extract_pattern = 2
                pat = pattern_3
            else:
                print("ERROR: There is not proper pattern for this file\n\n")
                pdf_dict[base_name]["Status"] = "ERROR: There is not proper pattern for this file"
                return pdf_dict

            # 5: Extract Transaction from PDF
            extract_list = []
            # APPLY FOR BANK OF WEST CC
            if re.search('www.bankofthewestcorporaterewards.com', all_text_page.replace('\n', '')):
                num_pat_BOW = r'(^\d{1,3}(?:,\d{3})*\.\d{2}$)'
                for i in range(len(all_text_list)):
                    try:
                        line = all_text_list[i]
                        next_line = all_text_list[i + 1]
                        if re.compile(pattern=pat).search(line) and re.compile(pattern=pat).search(next_line):
                            transaction = all_text_list[i:i + 2]

                        if re.search(pattern=num_pat_BOW, string=next_line):
                            amount = [line, next_line]
                            transaction.extend(amount)
                            line = ' '.join(transaction)
                            extract_list.append(line)
                    except IndexError:
                        break
            else:
                # APPLY FOR THE OTHER BANK
                for i in range(len(all_text_list)):
                    line = all_text_list[i]

                    if re.search(pat, line) and re.search(num_pat, line):
                        extract_list.append(line.strip())

                    elif re.search(pat, line):
                        extract_list.append(line)

                    elif re.search(num_pat, line):
                        span = re.search(num_pat, line).span()
                        only_numeric = line[span[0]:span[1]]
                        extract_list.append(only_numeric)

            # 6: GET ALL TRANSACTION START WITH DATE OR COMBINE ROW (START WITH DATE) AND ROW NUMBER
            filter_list = []
            # Filter with 53bank
            if re.search('www.53.com/businessbanking', all_text_page):
                for i in range(len(extract_list)):
                    line = extract_list[i]
                    if re.search(pat, line) and re.search(num_pat, line):
                        filter_list.append(line.strip())
            else:
                for i in range(len(extract_list)):
                    line = extract_list[i]
                    # if this line contains both of date and amount => append to filter list
                    if re.search(pat, line) and re.search(num_pat, line):
                        span = re.search(num_pat, line).span()
                        line = line[0:span[1]]
                        filter_list.append(line.strip())

                    # if this line is date = > combine with next line and append to next list
                    elif re.search(pat, line):
                        next_line = extract_list[i + 1]

                        # if next line is date get next second line
                        if re.search(pat, next_line) is not None:
                            new_line = line + ' ' + extract_list[i + 2]
                            filter_list.append(new_line)
                        else:
                            new_line = line + ' ' + next_line
                            filter_list.append(new_line)

                    # if this line is number: => continue
                    elif re.search(num_pat, line):
                        continue

            # 7: Convert list to dataFrame
            trans_df = pd.DataFrame(filter_list, columns=['Transaction'])

            # Get only transaction which contains amount
            trans_df = trans_df[~trans_df.Transaction.apply(lambda x: re.search(num_pat, x)).isnull()].reset_index(
                drop=True)

            # Print total transaction in file pdf
            print(f"File {base_name}.pdf: {trans_df.shape[0]} transactions\n\n")
            pdf_dict[base_name]["Transaction"] = f'File {base_name}.pdf: {trans_df.shape[0]} transactions'

            # Create 2 columns Post date and Trans date
            if kind_extract_pattern == 2:

                # Define pattern
                kind_2_pat_2d = r'(\d{1,2}[/-]\d{1,2})\s?(\d{1,2}[/-]\d{1,2})'
                kind_2_pat_1d = r'(\d{1,2}[/-]\d{1,2})'
                kind_2_pat_has_year = r'(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})'

                # Create filter
                has_year = trans_df.Transaction.str.contains(pat=kind_2_pat_has_year).sum() / trans_df.shape[0]
                has_2date = trans_df.Transaction.str.contains(pat=kind_2_pat_2d).sum() / trans_df.shape[0]

                # Check case 1: Has year ?
                if has_year > 0.7:
                    trans_df['Post_date'] = trans_df.Transaction.str.extract(pat=kind_2_pat_has_year)
                    trans_df['Trans_date'] = [''] * trans_df.shape[0]

                # Check case 2: Has two date? Post date and Trans date
                elif has_2date > 0.7:
                    trans_df[['Post_date', 'Trans_date']] = trans_df.Transaction.str.extract(pat=kind_2_pat_2d)

                    # If there are some transaction having 1 date => get this date to Post day
                    outlier = (trans_df['Transaction'].str.contains(pat=kind_2_pat_2d) == False)
                    trans_df.loc[outlier, ['Post_date']] = trans_df[outlier].Transaction.str.extract(pat=kind_2_pat_1d)

                else:
                    trans_df['Post_date'] = trans_df.Transaction.str.extract(pat=kind_2_pat_1d)
                    trans_df['Trans_date'] = [''] * trans_df.shape[0]

            if kind_extract_pattern == 1:
                # Define pattern
                kind_1_pat_2d = r'([A-Z][a-z]{2}\s?\d{1,2})\s?([A-Z][a-z]{2}\s?\d{1,2})'
                kind_1_pat_1d = r'([A-Z][a-z]{2}\s?\d{1,2})'

                # Create filter
                has_2date = trans_df.Transaction.str.contains(pat=kind_1_pat_2d).sum() / trans_df.shape[0]

                # Check case: Has two date? Post date and Trans date
                if has_2date > 0.7:
                    trans_df[['Post_date', 'Trans_date']] = trans_df.Transaction.str.extract(pat=kind_1_pat_2d)

                    # If there are some transaction having 1 date => get this date to Post day
                    outlier = (trans_df['Transaction'].str.contains(pat=kind_1_pat_2d) == False)
                    trans_df.loc[outlier, ['Post_date']] = trans_df[outlier].Transaction.str.extract(pat=kind_1_pat_1d)

                else:
                    trans_df['Post_date'] = trans_df.Transaction.str.extract(pat=kind_1_pat_1d)
                    trans_df['Trans_date'] = [''] * trans_df.shape[0]

            # Fill NA
            trans_df = trans_df.fillna('')

            # Create Amount columns
            trans_df['Amount'] = trans_df.Transaction.str.extract(pat=num_pat, expand=False).str.replace(pat=r'[$\s]',
                                                                                                         repl='',
                                                                                                         regex=True)

            # Create the Description columns
            trans_df['len_upper'] = trans_df[['Trans_date', 'Post_date']].apply(
                lambda x: len(x['Trans_date']) + len(x['Post_date']), axis=1)
            trans_df['len_lower'] = trans_df.Transaction.str.extract(num_pat, expand=False).apply(lambda x: len(x))

            trans_df['Description'] = trans_df.apply(lambda x: get_description(trans=x['Transaction'],
                                                                               len_upper=x['len_upper'],
                                                                               len_lower=x['len_lower']),
                                                     axis=1).str.strip()

            # Replace the multiple space to one multiple space
            trans_df['Description'] = trans_df.Description.str.replace(regex=True, pat=r'\s+', repl=' ')

            # Convert to the excel file
            trans_convert = trans_df[['Transaction', 'Post_date', 'Trans_date', 'Description', 'Amount']]

            excel_file_path = os.path.join(excel_path, f'{base_name}.xlsx')
            trans_convert.to_excel(excel_writer=excel_file_path, index=False)
            pdf_dict[base_name]["Dataframe"] = trans_convert
            return pdf_dict


# Kind 2: Tool to get transaction from citi bank because citi bank has specific format
def citi_bank_to_csv(pdf_path: str,
                     excel_path='C:/Users/trilnd/Desktop/QBO/Test_file/'):
    pattern_2 = r'^(\d{1,2}/\d{1,2})'

    base_name = os.path.basename(pdf_path)[0:-4]
    pdf_dict = {base_name: {'Status': '',
                            'Transaction': '',
                            'Dataframe': ''}}

    # 1: Check whether the path is correct or wrong
    try:
        with open(pdf_path, 'rb') as file:
            print("Connect file successful")
            pdf_dict[base_name]['Status'] = 'Connect file successful'
    except IOError:
        print("Can not find the file directory or file name does not exist")
        pdf_dict[base_name]['Status'] = 'Can not find the file directory or file name does not exis'

    # 2: Implement to convert file to pdf variable and print total number pages
    else:
        try:
            pdf_file = PdfReader(pdf_path, strict=True)
        except Exception as e:
            raise e
        else:

            # 3: Convert pdf file to the string
            all_text_page = ''
            for page_num in range(0, pdf_file.getNumPages()):
                all_text_page += (pdf_file.getPage(page_num).extractText() + '\n')
            all_text_list = [line.strip() for line in all_text_page.split('\n')]

            # : Check if the file contain text or numeric?
            if re.search(r'.', all_text_page) is None:
                print("This PDF file is Scan data")
                pdf_dict[base_name]['Status'] = "This PDF file is Scan data"
                return None

            # 4 Select proper pattern for PDF file
            count_p2 = sum([1 for line in all_text_list if re.search(pattern_2, line)])
            if count_p2 > 1:
                pat = pattern_2
                # print("Chose the patern 2")
            else:
                print("ERROR: There is not proper pattern for this file")
                pdf_dict[base_name]['Status'] = "ERROR: There is not proper pattern for this file"
                return pdf_dict

            # 5: Extract Transaction from PDF
            extract_list = []
            for i in range(len(all_text_list)):
                line = all_text_list[i]
                if re.search(pat, line):
                    list_trans = [line for line in all_text_list[i:(i + 6)] if line != '']
                    extract_list.append(list_trans)

                # Check whether last value is date value? if yes, replace None
            for line in extract_list:
                line[-1] = re.sub(pattern=pat, repl='', string=line[-1])

            # 6: Convert trans to Data frame
            trans_df = pd.DataFrame(extract_list,
                                    columns=['Date', 'Description', 'Amount', 'Balance', 'Sub_description_1',
                                             'Sub_description_2'])
            # Print total transaction in DF
            print(f"File {base_name}.pdf: {trans_df.shape[0]} transactions\n\n")
            pdf_dict[base_name]['Transaction'] = f"File {base_name}.pdf: {trans_df.shape[0]} transactions"

            trans_df['Description'] = trans_df.Description.str.replace(regex=True, pat=r'\s+', repl=' ')
            trans_df['Sub_description_1'] = trans_df.Sub_description_1.str.replace(regex=True, pat=r'\s+', repl=' ')
            trans_df['Sub_description_2'] = trans_df.Sub_description_2.str.replace(regex=True, pat=r'\s+', repl=' ')

            trans_convert = trans_df[['Date', 'Description', 'Sub_description_1', 'Sub_description_2', 'Amount']]

            # 7: Convert to excel file
            excel_file_path = os.path.join(excel_path, f'{base_name}.xlsx')
            trans_convert.to_excel(excel_writer=excel_file_path, index=False)
            pdf_dict[base_name]['Dataframe'] = trans_convert
            return pdf_dict


# Kind 3: Function to get Check No and transaction in Bank of Hope and Panacea Statement
def hope_bank_to_csv(pdf_path: str,
                     excel_path='C:/Users/trilnd/Desktop/QBO/Test_file/'):
    # Function to get description

    def get_description(trans, len_upper, len_lower):
        # APPLY FOR BANK OF WEST
        if re.search('bankofthewest.com', all_text_page):
            return trans[len_upper + len_lower + 1:]
        else:
            # check space:
            check_space = pat + '(?=\s)'
            if re.search(check_space, trans) != None:
                return trans[len_upper + 1:(len(trans) - len_lower)]
            else:
                return trans[(len_upper):(len(trans) - len_lower)]

    pattern_1 = r"^([A-Z][a-z]{2}\s?\d{1,2})"
    pattern_2 = r'^((\d{1,2}/\d{1,2})(/\d{4})?)'

    base_name = os.path.basename(pdf_path)[0:-4]
    pdf_dict = {base_name: {'Status': '',
                            'Transaction': '',
                            'Dataframe': ''}}

    # 1: Check whether the path is correct or wrong
    try:
        with open(pdf_path, 'rb') as file:
            print(f"{base_name}: Connect file successful")
            pdf_dict[base_name]["Status"] = "Connect file successful"

    except IOError:
        print("Can not find the file directory or file name does not exist")
        pdf_dict[base_name]["Status"] = "Can not find the file directory or file name does not exist"
        return pdf_dict

    # 2: Implement to convert file to pdf variable and print total number pages
    else:
        try:
            pdf_file = fitz.open(pdf_path)
        except Exception as e:
            raise e
        else:
            number_page = pdf_file.page_count
            #################################### GET TRANSACTION ###########################################

            # 3: Convert pdf file to the string
            all_text_page = ''
            for page in pdf_file:
                text = page.get_text('text')
                all_text_page += (text + '\n')

            if re.search(r'.', all_text_page) is None:
                raise Exception("This file without text data")
            else:
                all_text_list = [line.strip() for line in all_text_page.split('\n')]

            # : Check if the file contain text or numeric?
            if re.search(r'.', all_text_page) is None:
                print("This PDF file is Scan data")
                pdf_dict[base_name]["Status"] = "This PDF file is Scan data"
                return pdf_dict

                # Define num pattern
            if re.search('www.Hanmi.com', all_text_page) or re.search('capitalone.com', all_text_page):
                num_pat = r'(-?\s?\$\d{1,3}(?:,\d{3})*\.\d{2})'
            else:
                num_pat = r'(-?\s?\$?\d{1,3}(?:,\d{3})*\.\d{2})'

            # 4 Select proper date pattern for PDF file
            count_p1 = sum([1 for line in all_text_list if re.search(pattern_1, line)])
            count_p2 = sum([1 for line in all_text_list if re.search(pattern_2, line)])
            if count_p1 > count_p2:
                kind_extract_pattern = 1
                pat = pattern_1
            elif count_p1 < count_p2:
                kind_extract_pattern = 2
                pat = pattern_2
            else:
                print("ERROR: There is not proper pattern for this file\n\n")
                pdf_dict[base_name]["Status"] = "ERROR: There is not proper pattern for this file"
                return pdf_dict

            # 5: Extract Transaction from PDF
            extract_list = []
            if re.search('capitalone.com', all_text_page):
                for i, line in enumerate(all_text_list):
                    if re.search(pat, line) and re.search(num_pat, line):
                        extract_list.append(line.strip())

                    elif re.search(pat, line) and re.search(num_pat, all_text_list[i + 3]):
                        trans = all_text_list[i:i + 4]
                        extract_list.append(' '.join(trans))
            else:
                for i, line in enumerate(all_text_list):
                    if re.search(pat, line) and re.search(num_pat, line):
                        extract_list.append(line.strip())

                    elif re.search(pat, line):
                        extract_list.append(line + ' ' + all_text_list[i + 1])

                    elif re.search(num_pat, line):
                        span = re.search(num_pat, line).span()
                        only_numeric = line[span[0]:span[1]]
                        extract_list.append(only_numeric)

            # 6: GET ALL TRANSACTION START WITH DATE OR COMBINE ROW (START WITH DATE) AND ROW NUMBER
            filter_list = []

            # APPLY FOR BANK OF WEST CHECKING AND CAPITAL ONE
            if re.search('bankofthewest.com', all_text_page) or re.search('capitalone.com', all_text_page):
                for i, line in enumerate(extract_list):
                    if re.search(pat, line) and re.search(num_pat, line):
                        filter_list.append(line.strip())
            else:
                for i, line in enumerate(extract_list):
                    if len(extract_list) - 1 > i:
                        # if this line contains both of date and amount => append to filter list
                        if re.search(pat, line) and re.search(num_pat, line):
                            span = re.search(num_pat, line).span()
                            line = line[0:span[1]]
                            filter_list.append(line.strip())

                        # if this line is date = > combine with next line and append to next list
                        elif re.search(pat, line):
                            next_line = extract_list[i + 1]

                            # if next line is date get next second line
                            if re.search(pat, next_line) != None:
                                try:
                                    new_line = line + ' ' + extract_list[i + 2]
                                    filter_list.append(new_line)
                                except IndexError:
                                    pass
                            else:
                                new_line = line + ' ' + next_line
                                filter_list.append(new_line)

                        # if this line is number: => continue
                        elif re.search(num_pat, line):
                            continue

            #################################### GET CHECK NUMBER ###########################################

            # Pattern define:
            pat_check_no = r'\b\d+\*?\b'
            pat_date = r'((\d{1,2}/\d{1,2})(/\d{4})?)'
            pat_num = r'(-?\s?\$\d{1,3}(?:,\d{3})*\.\d{2})'

            # 1: Get Page and Coordinator contain "Checks" word
            String = 'Checks'
            Pagelist = []
            for i in range(0, number_page):
                PageObj = pdf_file[i]
                Text = PageObj.get_text('text')
                ReSearch = re.search(String, Text, flags=re.IGNORECASE)
                if ReSearch is not None:
                    # print(re.findall(String,Text))
                    info = (i, ReSearch.span()[0])
                    Pagelist.append(info)

            # 2: Convert string to list
            checkno_text_page = ''
            for i in Pagelist:
                checkno_text_page += (pdf_file[i[0]].get_text('text')[i[1]:] + '\n')

            # 3: Get line without Alphabelt digit
            filter_checkno_list = [line.strip() for line in checkno_text_page.split('\n') if
                                   re.search(r'[a-zA-Z+]', line) is None]

            # 4: Get only checkno line
            check_no_list = []
            for i in range(len(filter_checkno_list)):
                if i < len(filter_checkno_list) - 1:
                    line = filter_checkno_list[i]
                    next_line = filter_checkno_list[i + 1]
                    if re.search(pat_check_no, line) and re.search(pat_date, line):
                        if re.search(pat_num, next_line):
                            check_no_line = line + ' ' + next_line
                            check_no_list.append(check_no_line.split())

            ################################### CONVERT TO DATAFRAME AND EXPORT TO FILE CSV ###############################################

            #################### TRANSACTION DATAFRAME #########################
            trans_df = pd.DataFrame(filter_list, columns=['Transaction'])

            # Get only transaction which contaims amount
            trans_df = trans_df[~trans_df.Transaction.apply(lambda x: re.search(num_pat, x)).isnull()].reset_index(
                drop=True)

            # Print total transaction in file pdf
            print(f"File {base_name}.pdf: {trans_df.shape[0]} transactions\n\n")
            pdf_dict[base_name]["Transaction"] = f'File {base_name}.pdf: {trans_df.shape[0]} transactions'

            # Create 2 columns Post date and Trans date
            if kind_extract_pattern == 2:

                # Define pattern
                kind_2_pat_2d = r'(\d{1,2}[/-]\d{1,2})\s?(\d{1,2}[/-]\d{1,2})'
                kind_2_pat_1d = r'(\d{1,2}[/-]\d{1,2})'
                kind_2_pat_has_year = r'(\d{1,2}[/-]\d{1,2}[/-]\d{4})'

                # Create filter
                has_year = trans_df.Transaction.str.contains(pat=kind_2_pat_has_year).sum() / trans_df.shape[0]
                has_2date = trans_df.Transaction.str.contains(pat=kind_2_pat_2d).sum() / trans_df.shape[0]

                # Check case 1: Has year ?
                if has_year > 0.7:
                    trans_df['Post_date'] = trans_df.Transaction.str.extract(pat=kind_2_pat_has_year)
                    trans_df['Trans_date'] = [''] * trans_df.shape[0]

                # Check case 2: Has two date? Post date and Trans date
                elif has_2date > 0.7:
                    trans_df[['Post_date', 'Trans_date']] = trans_df.Transaction.str.extract(pat=kind_2_pat_2d)

                    # If there are some transaction having 1 date => get this date to Post day
                    outlier = (trans_df['Transaction'].str.contains(pat=kind_2_pat_2d) == False)
                    trans_df.iloc[outlier, [2]] = trans_df[outlier].Transaction.str.extract(pat=kind_2_pat_1d)

                else:
                    trans_df['Post_date'] = trans_df.Transaction.str.extract(pat=kind_2_pat_1d)
                    trans_df['Trans_date'] = [''] * trans_df.shape[0]

            if kind_extract_pattern == 1:
                # Define pattern
                kind_1_pat_2d = r'([A-Z][a-z]{2}\s?\d{1,2})\s?([A-Z][a-z]{2}\s?\d{1,2})'
                kind_1_pat_1d = r'([A-Z][a-z]{2}\s?\d{1,2})'
                kind_1_pat_has_year = r'([A-Z][a-z]{2}\s?\d{1,2},?\s?\d{2,4})'

                # Create filter
                has_2date = trans_df.Transaction.str.contains(pat=kind_1_pat_2d).sum() / trans_df.shape[0]
                has_year = trans_df.Transaction.str.contains(pat=kind_1_pat_has_year).sum() / trans_df.shape[0]

                # Check case 1: Has year ?
                if has_year > 0.7:
                    trans_df['Post_date'] = trans_df.Transaction.str.extract(pat=kind_1_pat_has_year)
                    trans_df['Trans_date'] = [''] * trans_df.shape[0]

                # Check case: Has two date? Post date and Trans date
                elif has_2date > 0.7:
                    trans_df[['Post_date', 'Trans_date']] = trans_df.Transaction.str.extract(pat=kind_1_pat_2d)

                    # If there are some transaction having 1 date => get this date to Post day
                    outlier = (trans_df['Transaction'].str.contains(pat=kind_1_pat_2d) == False)
                    trans_df.iloc[outlier, [2]] = trans_df[outlier].Transaction.str.extract(pat=kind_1_pat_1d)

                else:
                    trans_df['Post_date'] = trans_df.Transaction.str.extract(pat=kind_1_pat_1d)
                    trans_df['Trans_date'] = [''] * trans_df.shape[0]
                    # Fill NA
            trans_df = trans_df.fillna('')

            # Create Amount columns
            trans_df['Amount'] = trans_df.Transaction.str.extract(pat=num_pat, expand=False).str.replace(pat=r'[$\s]',
                                                                                                         repl='',
                                                                                                         regex=True)

            # Create the Description columns
            trans_df['len_upper'] = trans_df[['Trans_date', 'Post_date']].apply(
                lambda x: len(x['Trans_date']) + len(x['Post_date']), axis=1)
            trans_df['len_lower'] = trans_df.Transaction.str.extract(num_pat, expand=False).apply(lambda x: len(x))

            trans_df['Description'] = trans_df.apply(lambda x: get_description(trans=x['Transaction'],
                                                                               len_upper=x['len_upper'],
                                                                               len_lower=x['len_lower']),
                                                     axis=1).str.strip()

            # Replace the multiple space to one multiple space
            trans_df['Description'] = trans_df.Description.str.replace(regex=True, pat=r'\s+', repl=' ')
            columns = ['Transaction', 'Trans_date', 'Post_date', 'Description', 'Amount']

            # Convert to the excel file
            excel_file_path = os.path.join(excel_path, f'{base_name}.xlsx')
            trans_df[columns].to_excel(excel_writer=excel_file_path, index=False)
            pdf_dict[base_name]["Dataframe"] = trans_df[columns]

            # ---------------------------- CHECK NO DATAFRAME -------------------------------- #
            check_no_df = pd.DataFrame(check_no_list, columns=['Check_no', 'Date', 'Amount'])
            check_no_df['Amount'] = check_no_df['Amount'].str.replace('$', '', regex=False)
            # 6: Update file excel created before with new sheet "Check No"

            # if check path is valid
            if os.path.exists(excel_file_path):
                with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    check_no_df.to_excel(writer, sheet_name='Check_no', index=False)
            else:
                check_no_df.to_excel(excel_file_path, sheet_name='Check_no', index=False)

            pdf_dict[base_name]["DF_Check_No"] = check_no_df

            return pdf_dict


# Kind 4: Get Check No in General Kind
def get_check_no(pdf_path: str,
                 excel_path: str = 'C:/Users/trilnd/Desktop/QBO/Test_file/'):
    pattern_1 = r"^([A-Z][a-z]{2}\s?\d{1,2})"
    pattern_2 = r'^((\d{1,2}/\d{1,2})(/\d{2})?)'
    num_pat = r'(-?\s?\$?\d+,?\d*\.\d*)'
    pattern_check_no = r'\b\d{4,5}\b'
    base_name = os.path.basename(pdf_path)[:-4]
    pdf_dict = {base_name: {'Status': '',
                            'Transaction': '',
                            'Dataframe': ''}}

    try:
        with open(pdf_path, 'rb') as file:
            print("Connect file successful")
    except IOError:
        print("Can not find the file directory or file name does not exist")
        pdf_dict['Status'] = "Can not find the file directory or file name does not exist"

    # 2: Implement to convert file to pdf variable and print total number pages
    else:
        try:
            pdf_file = PdfReader(pdf_path, strict=True)
        except Exception as e:
            raise e
        else:
            number_page = pdf_file.getNumPages()

            # Get Page and Coodinator contain "Checks" word
            String = 'Checks'
            Pagelist = []
            for i in range(0, number_page):
                PageObj = pdf_file.getPage(i)
                Text = PageObj.extractText()
                ReSearch = re.search(String, Text, flags=re.IGNORECASE)
                if ReSearch != None:
                    # print(re.findall(String,Text))
                    info = (i, ReSearch.span()[0])
                    Pagelist.append(info)

            # 3: Convert pdf file to the string
            all_text_page = ''
            for i in Pagelist:
                all_text_page += (pdf_file.getPage(i[0]).extract_text()[i[1]:] + '\n')

            all_text_list = [line.strip() for line in all_text_page.split('\n')]

            # 4 Select proper pattern for PDF file
            list_check_no = [line for line in all_text_list if
                             re.search(pattern=pattern_check_no, string=line) and re.search(num_pat, line)]

            # 5: Create DF
            df = pd.DataFrame({"Check no": list_check_no})
            print(f"File {base_name}.pdf: {df.shape[0]} Checks")
            pdf_dict[base_name]['Transaction'] = f"File {base_name}.pdf: {df.shape[0]} Checks"

            # 6: Update file excel created before with new sheet "Check No"
            path = os.path.join(excel_path, f'{base_name}.xlsx')

            # if check path is valid
            if os.path.exists(path):
                with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name='Check_no', index=False)
            else:
                df.to_excel(path, sheet_name='Check_no', index=False)
                pdf_dict[base_name]['Dataframe'] = df
            return pdf_dict


####################### TOOLS GET ALL PDF IN A PDF FOLDER THEN CONVERTING TO EXCEL FILES#######################

def convert_pdffolder_to_excelfolder(pdf_path: str, xlsx_path: str, function_kind: int):
    function_list = [convert_pdf_to_csv,
                     citi_bank_to_csv,
                     get_check_no,
                     hope_bank_to_csv]
    function = function_list[function_kind - 1]
    list_pdfs = []
    list_file = [os.path.join(pdf_path, f) for f in os.listdir(pdf_path) if f.endswith('.pdf')]
    # kind 1: Popular type: convert PDF to excel (except Citi Checking)
    if len(list_file) == 0:
        print("There is no file PDF. Please choose again")
    else:
        if function != get_check_no:
            list_file = [file for file in list_file if
                         os.path.exists(os.path.join(xlsx_path, os.path.basename(file)[:-4] + '.xlsx')) is False]
            for file in list_file:
                try:
                    dict_pdf = function(pdf_path=file, excel_path=xlsx_path)
                except Exception:
                    continue
                else:
                    list_pdfs.append(dict_pdf)
        else:
            for file in list_file:
                try:
                    dict_pdf = get_check_no(pdf_path=file, excel_path=xlsx_path)
                except Exception:
                    continue
                else:
                    list_pdfs.append(dict_pdf)

        # Create message to display in GUI
        status_pdf = [info['Status'] for dict_pdf in list_pdfs for (key, info) in dict_pdf.items()]
        trans_pdf = [info['Transaction'] for dict_pdf in list_pdfs for (key, info) in dict_pdf.items()]
        filename_pdf = [key for dict_pdf in list_pdfs for (key, info) in dict_pdf.items()]
        message = '\n'.join([f'{n}.pdf\n{s}\n{t}\n'
                             for n, s, t in zip(filename_pdf,
                                                status_pdf,
                                                trans_pdf)])

        # Get total Transaction in dataframe
        valid_dfs = [info['Dataframe']
                     for dict_pdf in list_pdfs
                     for (key, info) in dict_pdf.items()
                     if isinstance(info['Dataframe'], pd.DataFrame)]

        # sum transaction of total PDF file
        total = sum([df.shape[0] for df in valid_dfs])
        summary = f"In {len(list_file)} files,There are {len(valid_dfs)} valid files with Total {total} transactions"

        return f'{message}\n{summary}'
