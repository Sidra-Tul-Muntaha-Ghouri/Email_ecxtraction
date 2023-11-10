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
st.header('Email Navigator')

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
        
        #result is a variable that will store the result of the search operation. This variable typically contains the search status.
        '''he possible values that result can take after using IMAP4.search are as follows:
        OK: The search operation was successful, and the requested email messages have been found. This typically means that the search criteria were matched.

NO: The search operation was not successful, and no email messages matching the search criteria were found.

BAD: The search operation was not recognized or understood by the server, indicating an issue with the search query.

BYE: The server has closed the connection, which might be due to various reasons, such as an error or the server's unavailability.'''
        #data is a variable that will store the UIDs (unique identifiers) of the email messages that match the search criteria.

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
                
         
    except imaplib.IMAP4.error as e:
        st.error('Authentication failed. Please check your email address and password.')
        st.error(f"Error: {str(e)}")

else:
    st.warning('Please enter your email address and password.')


