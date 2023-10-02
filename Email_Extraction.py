import imaplib
import email
import pandas as pd
import streamlit as st
from datetime import datetime
import time
from io import BytesIO


today = datetime.today().date()

# Set Streamlit page configuration
st.set_page_config(layout = 'wide', page_title='Email Extraction', page_icon='📧')
st.title('Email Extraction')

# Input fields for email address and app password
add = st.text_input('Your Email Address')
pas = st.text_input('Your App Password', type='password')  # Use type='password' to hide the input
mailbox = ['Inbox', 'Starred', 'Important', 'Sent']
box = st.selectbox('Select Mail Box', mailbox)
start_date = st.date_input("Select a start date:", key="start", max_value=today)
end_date = st.date_input("Select an end date:", key="end", max_value=today)


# Check if both email address and password are provided
if add and pas: 
    
    try:
        # Connect to the Gmail IMAP server securely
        mail = imaplib.IMAP4_SSL('imap.gmail.com', 993)

        # Login to the email account
        mail.login(add, pas)
        mail.select(box)
        result, data = mail.search(None, f'SINCE {start_date.strftime("%d-%b-%Y")}', f'BEFORE {end_date.strftime("%d-%b-%Y")}')

        # Parse the email messages and extract information
        email_list = []
        for i in data[0].split():
            result, data = mail.fetch(i, '(RFC822)')
            raw_email = data[0][1]
            email_message = email.message_from_bytes(raw_email)
            email_subject = email_message['Subject']
            email_sender = email_message['From']
            email_date = email_message['Date']
            
            # Get the email body
            if email_message.is_multipart():
                for part in email_message.walk():
                    if part.get_content_type() == "text/plain":
                        email_body = part.get_payload(decode=True).decode("utf-8")
                        break
            else:
                email_body = email_message.get_payload(decode=True).decode("utf-8")

            email_list.append([email_subject, email_body, email_sender, email_date])
           


        # Create a dataframe of email data
        create_df = st.button('Get Emails')
        if create_df:         
            # progress_text = "Operation in progress. Please wait."
            # my_bar = st.progress(0, text=progress_text)

            # for percent_complete in range(100):
            #     time.sleep(0.01)
            #     my_bar.progress(percent_complete + 1, text=progress_text)
            # time.sleep(1)
            # my_bar.empty()
            df = pd.DataFrame(email_list, columns=['Subject', 'Email Body','Sender', 'Date' ])
            st.write(f'Your emails from {start_date} to {end_date}')
            #st.dataframe(df)
            st.write(df)
        
            if df.empty:
                st.write('Dataframe is empty')

        
            def convert_df(df):
                return df.to_csv(index=False).encode('utf-8')
            csv = convert_df(df)

            st.download_button(
                "Press to Download",
                csv,
                "file.csv",
                "text/csv",
                key='download-csv'
                )
                

            # flnme = st.text_input('Enter Excel file name (e.g. email_data.xlsx)')
            # if flnme != "":
            #     if flnme.endswith(".xlsx") == False:  # add file extension if it is forgotten
            #         flnme = flnme + ".xlsx"

            #     buffer = BytesIO()
            #     with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            #         df.to_excel(writer, sheet_name='Report')

            #     st.download_button(label="Download Excel workbook", data=buffer.getvalue(), file_name=flnme, mime="application/vnd.ms-excel")
            # else:
            #     st.write('Click download button to download as Excel Sheet')
            #     download_df = st.button('Download')
            #     if download_df:
            #         file_name = st.text_input('Enter a file name for the Excel file (e.g., email_data.xlsx)')
            #         if file_name:
            #             file_path = f"{file_name}.xlsx"
            #             df.to_excel(file_path, index=False)
            #             st.success(f"File '{file_path}' has been downloaded successfully.")


            # if df is not None:
            #     file_name = st.text_input('Enter a file name for the Excel file (e.g., email_data.xlsx)')
            #     download = st.download_button(label='Download Excel', data=df.to_excel(index=False, header=True), key='download')

            #     if file_name and download:
            #         with open(file_name, "wb") as f:
            #             f.write(download)
            #         st.success(f"File '{file_name}' has been downloaded successfully.")


            # if df is not None:
            #     file_name = st.text_input('Enter a file name for the Excel file (e.g., email_data.xlsx)')
    
            #     if st.button('Download Excel') and file_name:
            #         with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
            #             df.to_excel(writer, sheet_name='Sheet1', index=False, header=True)
            #         st.success(f"File '{file_name}' has been downloaded successfully.")


         
    except imaplib.IMAP4.error as e:
        st.error('Authentication failed. Please check your email address and password.')
        st.error(f"Error: {str(e)}")

else:
    st.warning('Please enter your email address and password.')



#mail.login('sidratulmuntaha135@gmail.com', 'bfsh faer jpkz diox')

# Select the mailbox and search for emails

# mail.select('inbox')
# result, data = mail.search(None, 'SINCE "01-SEP-2023"')
# 
     #Parse the email messages and extract information
# 
# email_list = []
# for i in data[0].split():
    # result, data = mail.fetch(i, '(RFC822)')
    # raw_email = data[0][1]
    # email_message = email.message_from_bytes(raw_email)
    # email_subject = email_message['Subject']
    # email_sender = email_message['From']
    # email_date = email_message['Date']
    # email_list.append([email_subject, email_sender, email_date])
# 
     #Create a data frame and export to Excel
# 
# df = pd.DataFrame(email_list, columns=['Subject', 'Sender', 'Date'])
# df.to_excel('email_data.xlsx', index=False)
# 
       #Disconnect from the IMAP server
# 
# mail.close()
# mail.logout()
# 
# df.head()