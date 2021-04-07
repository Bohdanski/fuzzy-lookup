import os
import sys
import csv
import xlsxwriter

from fuzzywuzzy import fuzz
from fuzzywuzzy import process


# User input

base = "tblTopsMatch.csv"
match = "tblWegmansMatch.csv"
base_field = "topsDesc"
match_field = "wegmansDesc"
method = "sort"
threshold = 60


def fuzzy_match(base, match, method):
    """
    
    """
    if method == "ratio":
        return fuzz.ratio(base.lower(), match.lower())
    elif method == "pratio":
        return fuzz.partial_ratio(base.lower(), match.lower())
    elif method == "sort":
        return fuzz.token_sort_ratio(base, match)
    elif method == "set":
        return fuzz.token_set_ratio(base, match)
    else:
        print("ERROR: Invalid match method.")
        raise


def main():
    try:
        data_dir = ".\\excel\\data\\"
        archive_dir = ".\\excel\\archive\\"

        # Open base file 
        with open(data_dir + base, "r") as file:
            base_file = csv.DictReader(file)
            base_lst = []
            header_lst = []
            
            # Copy dictionary rows into a list, and extract headers into a list
            for row in base_file:
                for key in row:
                    if key not in header_lst:
                        header_lst.append(key)
                base_lst.append(row)
        
        # Open file to match records against the base file
        with open(data_dir + match, "r") as file:
            match_file = csv.DictReader(file)
            match_lst = []
            
            # Copy dictionary rows into a list
            for row in match_file:
                match_lst.append(row)
        
        # For dictionary row in the base file list...
        write_lst = []
        for base_row in base_lst:
            best_match = ("No Match", 0)
            row_lst = []         
            # Match each row against the current base row
            for match_row in match_lst:
                match_ratio = fuzzy_match(base_row[base_field], match_row[match_field], method)                
                # If the match ratio is less than threshold, skip record
                if match_ratio < threshold:
                    continue
                # Else, assign the highest ratio (linear search)
                elif match_ratio > best_match[1]:
                    best_match = (match_row[match_field], match_ratio)
            print(f"[{base_row[base_field]} | {best_match[0]}] Match Ratio: {best_match[1]}")
            
            # For each row bring in all additional fields from base file, using header as key
            for header in header_lst:
                row_lst.append(base_row[header])    
                
            # For each row, create a list of values to be appeneded to a master write list
            row_lst.extend(list(best_match))
            write_lst.append(row_lst)
        
        # Create a new workbook and worksheet
        workbook = xlsxwriter.Workbook(archive_dir + f"match-{method}.xlsx")
        worksheet = workbook.add_worksheet()
        # Write coloumn headers as the first row
        for col_num, data in enumerate(header_lst):
            worksheet.write(0, col_num, data)
        # Write each row from the write list to workbook
        for row_num, row_data in enumerate(write_lst):
            if row_num == 0:
                continue
            for col_num, col_data in enumerate(row_data):
                worksheet.write(row_num, col_num, col_data)
        workbook.close()
    except:
        if not os.path.exists(archive_dir):
            os.makedirs(archive_dir)
        if not os.path.exists(data_dir):
            os.makedirs(data_dir)


if if __name__ == "__main__":
    main()
