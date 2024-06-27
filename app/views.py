from django.shortcuts import render
from django.http.response import Http404
import pandas as pd
from io import BytesIO
from app.tasks import email_background_worker
from django.core.mail import EmailMessage
from django.conf import settings
from openpyxl.styles import Font
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.styles import Border, Side




def home(request):
    """
    column-list:
    sr_no	gst_no	vendor_name	invoice	tax_amt	mobile_no	email	remark	date	summary
    """
    # step 1; column clean -> remove white space and trim and reassign
    # step 2; read email and invoice row. 
    # step 3; convert row to excel and attach to email 
    # step 4; send email via background tasks
    
    context = {}
    if request.method == "POST":
        try:
            files = request.FILES

            excel = files.get('file')
            if not excel:
                raise Http404('Invalid excel file')
            
            df = pd.read_excel(excel)
            # step 1
            df.columns = df.columns.str.lower().str.strip()
            
            
            # step 2
            for index, row in df.iterrows():
                heading = [
                        "SHREE GANESH PRESS N COAT INDUSTRIES PRIVATE LIMITED.\n",
                        f"{row['gst_no']}.\n"
                        "Report Name: Not in GSTR 2B.\n"
                        "April 23 to April 24.\n\n"
                        ]
                # row.to_excel('Invoice.xlsx', header=heading)  # => proper working, but heading is not displayed in mobile devices, and downloaded file
                
                # my_heads = ['sr_no', 'gst_no', 'vendor_name', 'invoice', 'tax_amt', 'mobile_no', 'email', 'remark', 'date', 'summary']###
                # row.to_excel('Invoice.xlsx', header=my_heads, startrow=5, startcol=3, index=False)###

                excel_buffer = row_to_excel(row, heading)###
                
                with open('invoice.xlsx', 'wb') as f:
                    f.write(excel_buffer.getbuffer())

                message = get_email_template(row)
                email_background_worker(row['email'], message)

            context["message"] = "Email sending in progress"
        except Exception as e:
            import traceback
            traceback.print_exc()
            context["message"] = str(e)

    return render(request, 'pages/index.html', context)


def get_email_template(row):
    
    msg = f"""
            Dear Sir/Madam,\n
        
            During the reconciliation of GSTR-2B for FY 2023-2025, we found following discrepancies 
            related to invoices/ Debit Note/ Credit Note not uploaded by you on the GST portal 
            during the April 23 to April 24 F.Y. 23-25. Due to this we are unable to take input tax credit 
            of said invoices/ Debit Note/ Credit Notes.  These discrepancies have been worked out on 
            the basis of 2B downloaded up to 30th April 2024. \n

            Request to please resolve these issues and submit necessary documents at the earliest for 
            said discrepancies as otherwise any reversal of Input tax credit on account of mismatch in 
            record will be debited to you along with interest and penalty. This may be due to an error in 
            uploading the details of Invoices/ Debit Note/ Credit Note in Name of SHREE GANESH 
            PRESS N COAT INDUSTRIES PRIVATE LIMITED on GST No. 27AAFCS1275F1ZE.\n
            
            The summary of discrepancies and invoice wise detail is annexed herewith.\n

            In case of any doubt/ clarification please feel free to write to us.\n

            Manesh G Sheolikar
            Shree Ganesh Press N Coat Industries Pvt Ltd
            Account Department
            M-152,Waluj-Aurangabad
            M-Number-9881621443
            """
    return msg

def row_to_excel(row, heading):
    # Create a DataFrame from the row
    df = pd.DataFrame([row])

    # Create an in-memory buffer
    excel_buffer = BytesIO()

    # Write the DataFrame to the buffer
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        # my_heads = ['sr_no', 'gst_no', 'vendor_name', 'invoice', 'tax_amt', 'mobile_no', 'email', 'remark', 'date', 'summary']
        writer.sheets['Sheet1'] = writer.book.create_sheet('Sheet1')###
        sheet = writer.sheets['Sheet1']

        wb = Workbook()
        ws = wb.active 
         
        for idx, line in enumerate(heading, start=1):
            sheet.cell(row=idx, column=10, value=line.strip()).alignment = openpyxl.styles.Alignment(horizontal='center')#for making data horrizantal from vertical.
            sheet.cell(row=idx, column=10).font = Font(bold=True)
        
        fill_color = PatternFill(fill_type='solid', fgColor='00FFCC99')#=>Set color to row
        row_number = 5 #number of row
        
        # max_col = df.shape[0]
        
        # Apply the fill color to the entire row
        # for col in range(1, max_col + 1):
        #     sheet.cell(row=row_number, column=6).fill = fill_color
        
        for col in sheet.iter_rows(min_col=6, max_col=14 + df.shape[0], min_row=row_number, max_row=5):
            for cell in col:
                cell.fill = fill_color


        # Define a border style
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        # Apply the border to all cells in the DataFrame
        for row in sheet.iter_rows(min_row=1, max_row=5 + df.shape[0], min_col=6, max_col=15):
            for cell in row:
                cell.border = thin_border

        startrow = len(heading)+1
        df.to_excel(writer, index=False, startrow=startrow, startcol=5)

    # Seek to the beginning of the stream
    excel_buffer.seek(0)

    return excel_buffer

def bold_text(text):
  """Applies bold formatting to a string."""
  font = Font(bold=True)
  return font + text

            # This is testing email Please Ignore It !!!
        
            # Hello {row['vendor_name']}, 

            # Your last invoice is pending!
            # The amount is {row['invoice']}.\n\n
            # Regards,
            # Zencon Infotech Pvt Ltd.

