import pdfplumber
import re
import pandas as pd
from tqdm import tqdm

# Regular Expressions for parsing the PDF file
doc_num_regEx = re.compile(r'4\d{7}')
line_item_regEx = re.compile(r'(\d+)\s(CS|EA|cs|ea)\s(\w+)\s(.+)\s(\d+\.\d+)\s(\$\d+,?(?:\d+)?\.?\d{2}?)\sT?')
sales_tax_regEx = re.compile(r'STATE\sSALES\sTAX\s(\$\d+,?\d+\.?\d{2})')
ship_to_regEx = re.compile(r'2\d{7}')

# List to store all of our data that will be written to disk
records = list()

def mfg_docs(path_to_pdf: str) -> None:
    # Iterate over the pages in the PDF file
    with pdfplumber.open(path_or_fp=path_to_pdf) as pdf:
        for i in tqdm(range(len(pdf.pages))):
            page = pdf.pages[i]
            text = page.extract_text(x_tolerance=3, y_tolerance=3)

            inv_num = re.search(pattern=doc_num_regEx, string=text).group()
            line_items = re.findall(pattern=line_item_regEx, string=text)
            ship_to = re.search(pattern=ship_to_regEx, string=text).group()

            try:
                sales_tax = re.search(pattern=sales_tax_regEx, string=text).group(1)
            except Exception:
                sales_tax = '-'    

            for j in range(len(line_items)):
                records.append((i+1, inv_num, ship_to, line_items[j][0], line_items[j][1], line_items[j][2], line_items[j][3], line_items[j][4], line_items[j][5], '-'))

            if sales_tax != '-':
                    records.append((i+1, inv_num, ship_to, '-', '-', '-', '-', '-', '-', sales_tax))

    # Create pandas data frame    
    df = pd.DataFrame(records, columns=['Page Number',
                                        'Invoice Number',
                                        'Ship To Number',
                                        'Quantity Shipped', 
                                        'UOM', 
                                        'Catalog Number', 
                                        'Description', 
                                        'Unit Price', 
                                        'Extended Amount', 
                                        'Sales Tax'])

    # Write the data frame to Excel
    df.to_excel(r'C:\Users\REDACTED\Desktop\Document_Data.xlsx', sheet_name='Customer Data', index=False, freeze_panes=(1,0))

# Call the function
mfg_docs(path_to_pdf=r'C:\Users\REDACTED\Desktop\Documents\Tax\Assignments\REDACTED\01.19.22 - Data Extraction from MFG PRO Documents\Document Numbers.pdf')
