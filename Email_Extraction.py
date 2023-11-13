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
st.title('Email Extraction App')
#st.header('Email Navigator')
st.subheader(':grey[Thank you for using this application!]\n\n :grey[**To enhance your experience, I kindly encourage you to read the instructions available in the sidebar**].',divider='rainbow')
with st.sidebar:
  
    st.title('Introduction')

    # Adding content to the sidebar
    st.write('Email Extraction is a simple Web application, developed for extracting emails from Gmail accounts and saving them as a .csv file. The project is designed to work seamlessly with Gmail accounts and offers a user-friendly Streamlit interface for ease of use.')
    
#2nd heading in the sidebar
    st.title('How to use?')
    st.write("""
    - Enter your Gmail **Email address** and **App password**.
    - You can then select a **mailbox** from the dropdown list.
    - Set a date range, and click the **Get Emails** button to initiate the extraction process.
    - The extracted email data will be displayed in a tabular format and can be downloaded via **Download** button. 
    - Remember to **Logout** after using the app.    
    \\
    (Note that the **App password** is different from the usual passwords. Click [here](https://medium.com/@sidratulmuntahaghouri/get-your-emails-in-excel-b33f4e8b28cc) to get step by step process of generating app password.)""")
    st.title('Objective')
    st.write("""The main objective of this project is to download emails as .csv file, to be stored as record or be used for further analysis.""")
  #st.write("""The extracted email data will be displayed in a tabular format within the Streamlit web application. You can explore, analyze, and download the data as needed. """)

    st.write("**Let's Connect!** [Linkedin](https://www.linkedin.com/in/sidra-tul-muntaha-ghouri/) & [Medium](https://sidratulmuntahaghouri.medium.com/)")          
# Input fields for email address and app password
add = st.text_input('Your Email Address')
pas = st.text_input('Your App Password', type='password')  # Using type='password' to hide the input
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
        # The possible values that result can take after using IMAP4.search are as follows:
        # OK: The search operation was successful, and the requested email messages have been found. This typically means that the search criteria were matched.
        # NO: The search operation was not successful, and no email messages matching the search criteria were found.
        # BAD: The search operation was not recognized or understood by the server, indicating an issue with the search query.
        # BYE: The server has closed the connection, which might be due to various reasons, such as an error or the server's unavailability.
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
                "Download",
                csv,
                "Emails.csv",
                "text/csv",
                key='download-csv'
                )
        lgout = st.button('Logout')
        if lgout:
            add = None  # Set add to None to clear the email address
            password = None  # Set password to None to clear the password.
            st.write('Thanks For Using 🙂 Have a nice day!')
            mail.close()
            mail.logout()        
            st.balloons()
         
    except imaplib.IMAP4.error as e:
        st.error('Authentication failed. Please check your email address and password.')
        st.error(f"Error: {str(e)}")

else:
    st.warning('Please enter your email address and password.')


