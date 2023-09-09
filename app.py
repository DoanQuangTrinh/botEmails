from flask import Flask, render_template, request
import smtplib
import ssl
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import email.utils as eut
import time
import os
import re

app = Flask(__name__)

@app.route("/")
def hello():
    return render_template('index.html', title='Docker Python', name='James')

@app.route('/send_email', methods=['GET','POST'])
def send_email():
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    email = request.form['email']
    password = request.form['password']
    sheet_link = request.form['sheet_link']
    file_paths = request.files.getlist('file_paths')
    output_df = pd.DataFrame(columns=['name', 'email', 'sent time'])

    try:
        sheet_id = sheet_link.split('/')[5]
        url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx'
        df = pd.read_excel(url)
    except Exception as e:
        print(f"Error reading Google Sheet: {e}")
        return "Invalid Google Sheet link", 400

    interval = int(request.form['interval']) if request.form['interval'] else 0

    for index, row in df.iterrows():
        if '@' in str(row['Email']) and '.' in str(row['Email']).split('@')[1]:
            email_to = [row['Email']]
            email_subject = row['Subject']
            email_message = row['Message'].replace('@name', row['Name'])

            lines = email_message.split('\n')
            formatted_lines = []
            for line in lines:
                if line.endswith(':'):
                    formatted_lines.append(line + '\n')
                else:
                    formatted_lines.append(line)
            email_message = '<br>'.join(formatted_lines)
            signature = row['Signature']
            email_message += "<br>" + signature

            for recipient in email_to:
                msg = MIMEMultipart()
                if file_paths:
                    for file_path in file_paths.split(', '):
                        with open(file_path, 'rb') as f:
                            attachment = MIMEApplication(f.read(), _subtype='octet-stream')
                            attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file_path))
                            msg.attach(attachment)
                msg.attach(MIMEText(email_message, 'html'))

                msg['To'] = eut.formataddr((row['Name'], recipient))
                msg['From'] = eut.formataddr(('Henry Universes', email))
                msg['Subject'] = email_subject

                try:
                    server = smtplib.SMTP(smtp_server, smtp_port)
                    server.starttls()
                    server.login(email, password)
                    server.sendmail(email, recipient, msg.as_string())
                    server.quit()
                    df.at[index, 'Status'] = 'Sent'
                    sent_time = time.strftime('%Y-%m-%d %H:%M:%S')
                    output_df = output_df.append({'name': row['Name'], 'email': recipient, 'sent time': sent_time}, ignore_index=True)
                except Exception as e:
                    df.at[index, 'Status'] = 'Failed'
                    
                time.sleep(interval)
        else:
            df.at[index, 'Status'] = 'Failed'
            output_df = output_df.append({'name': row['Name'], 'email': 'null', 'sent time': 'null'}, ignore_index=True)
    output_df.to_excel('output.xlsx', sheet_name='Sheet1', index=False)
    return "All emails sent successfully", 200

if __name__ == '__main__':
    app.run(debug=True)
