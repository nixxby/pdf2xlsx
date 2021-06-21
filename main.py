import pandas as pd
import pdfplumber
import re
from datetime import datetime

def change_date_format(date_str):
    """
    Change date string type (Ex. 02 Mar 2019) to date() object (Ex. 2019-03-04)
    """
    return datetime.strptime(date_str, '%d %b %Y').date()

#Open the pdf file
with pdfplumber.open("test_input.pdf") as pdf:
    val_dates, ord_dates, descriptions, currencies, debit, credit = ([] for _ in range(6))
    #Start text extraction for each page. 
    for page in pdf.pages:
        page_text = page.extract_text(x_tolerance=2, y_tolerance=0)
        
        #Read the extract line by line
        for line in page_text.split('\n'):
            #RegEx for dates and currencies.
            rows = re.compile(r'\b(\d{1,2} ?(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) \d{4})')  
            curr_rows = re.compile(r'\bDeposits ')
            
            #Find Dates, Description, Currency and Amounts of Credit/Debit from each line
            if re.match(curr_rows,line):
                curr = re.compile(r'[A-Z]{3}')
                currency = re.findall(curr,line)

            if re.match(rows,line):
                dates = re.findall(rows, line)
                modified_dates = [change_date_format(date_str) for date_str in dates]
                amount = list(line.split(" "))[-1]

                val_dates.append(modified_dates[0])
                ord_dates.append(modified_dates[1])

                description = ' '.join(line.split(" ")[6:-1])
                descriptions.append(description)
                currencies.append(currency[0])
                
                if amount[0]=='(':
                    debit.append('{:.2f}'.format(round(float(amount[1:-1].replace(',', '')),2)))
                    credit.append(None)
                else:
                    credit.append('{:.2f}'.format(round(float(amount.replace(',', '')),2)))
                    debit.append(None)


            #Each extended Description in newline is in all Upper caps. 
            #Otherwise we can adopt another strategy to extract extended descriptions,
            #using the position of the characters and aligning them with description line just above.

            elif line.isupper():
                descriptions[-1] = ''.join([descriptions[-1],line])


#Assign Extracted Data to a DataFrame and export to an Excel file.
data = {
        'Value Date': val_dates, 
        'Order Date': ord_dates,
        'Description': descriptions, 
        'Currency': currencies, 
        'Debit': debit, 
        'Credit': credit,
}
df = pd.DataFrame(data)
df['Debit'] = df['Debit'].astype(float)
df['Credit'] = df['Credit'].astype(float)

df = df.sort_values(by = ["Value Date","Order Date"],ignore_index=True)
df.to_excel('test_output.xlsx', index=False)
print("Extraction Successful👍")