# from flask import Flask, render_template, request 
# import smtplib
# import ssl
# import pandas as pd
# from email.mime.text import MIMEText
# from email.mime.multipart import MIMEMultipart
# from email.mime.application import MIMEApplication
# import email.utils as eut
# import time
# import os
# import re
# import schedule
# import threading
# from werkzeug.utils import secure_filename



# app = Flask(__name__)

# UPLOAD_FOLDER = 'uploads'  # Thư mục tạm thời để lưu trữ tệp tải lên
# app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# def send_email(email, password, sheet_link, interval, file_paths=[]):
#         smtp_server = 'smtp.gmail.com'
#         smtp_port = 587
#         output_df = pd.DataFrame(columns=['name', 'email', 'sent time'])

#         try:
#             sheet_id = sheet_link.split('/')[5]
#             url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx'
#             df = pd.read_excel(url)
#         except Exception as e:
#             print(f"Error reading Google Sheet: {e}")
#             return "Invalid Google Sheet link", 400

#         for index, row in df.iterrows():
#             if '@' in str(row['Email']) and '.' in str(row['Email']).split('@')[1]:
#                 email_to = [row['Email']]
#                 email_subject = row['Subject']
#                 email_message = row['Message'].replace('@name', row['Name'])

#                 lines = email_message.split('\n')
#                 formatted_lines = []
#                 for line in lines:
#                     if line.endswith(':'):
#                         formatted_lines.append(line + '\n')
#                     else:
#                         formatted_lines.append(line)
#                 email_message = '<br>'.join(formatted_lines)
#                 signature = row['Signature']
#                 email_message += "<br>" + signature

#                 for recipient in email_to:
#                     msg = MIMEMultipart()
#                     for file_path in file_paths:
#                         if os.path.isfile(file_path):
#                             with open(file_path, 'rb') as f:
#                                 attachment = MIMEApplication(f.read(), _subtype='octet-stream')
#                                 attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file_path))
#                                 msg.attach(attachment)

#                     msg.attach(MIMEText(email_message, 'html'))

#                     msg['To'] = eut.formataddr((row['Name'], recipient))
#                     msg['From'] = eut.formataddr(('Henry Universes', email))
#                     msg['Subject'] = email_subject

#                     try:
#                         server = smtplib.SMTP(smtp_server, smtp_port)
#                         server.starttls()
#                         server.login(email, password)
#                         server.sendmail(email, recipient, msg.as_string())
#                         server.quit()
#                         df.at[index, 'Status'] = 'Sent'
#                         sent_time = time.strftime('%Y-%m-%d %H:%M:%S')
#                         output_df = output_df.append({'name': row['Name'], 'email': recipient, 'sent time': sent_time}, ignore_index=True)
#                     except Exception as e:
#                         df.at[index, 'Status'] = 'Failed'
                        
#                     time.sleep(interval)
#             else:
#                 df.at[index, 'Status'] = 'Failed'
#                 output_df = output_df.append({'name': row['Name'], 'email': 'null', 'sent time': 'null'}, ignore_index=True)
#         output_df.to_excel('output.xlsx', sheet_name='Sheet1', index=False)
#         return "All emails sent successfully", 200

# @app.route("/")
# def hello():
#     return render_template('index.html')

# @app.route('/send_email', methods=['POST'])
# def send_email_route():
#     email = request.form['email']
#     password = request.form['password']
#     sheet_link = request.form['sheet_link']
#     interval = int(request.form['interval']) if request.form['interval'] else 0
#     file_paths = request.files.getlist('file_paths')

#     if not os.path.exists(app.config['UPLOAD_FOLDER']):
#         os.makedirs(app.config['UPLOAD_FOLDER'])
#     file_paths = []
#     for uploaded_file in request.files.getlist('file_paths'):
#         if uploaded_file.filename != '':
#             filename = secure_filename(uploaded_file.filename)
#             uploaded_file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
#             file_paths.append(os.path.join(app.config['UPLOAD_FOLDER'], filename))

#     # Thực hiện việc gửi email bằng cách gọi hàm send_email với các tham số truyền vào
#     result, status_code = send_email(email, password, sheet_link, interval, file_paths)
#     for file_path in file_paths:
#         os.remove(file_path)
#     return result, status_code

# # def send_daily_email():
# #     print('batdau')
# #     email = "dtrinhb3@gmail.com"
# #     password = "sdaj guch abps xovd"
# #     sheet_link = "https://docs.google.com/spreadsheets/d/1h26VjAo2DBcdUsNh6yu2buz-6nZ77WtSL17WWVjj-S4/edit#gid=0"
# #     interval = 0
# #     send_email(email, password, sheet_link, interval, [])
# #     print("thanh cong")


# # Lên lịch thực hiện hàm send_daily_email vào mỗi ngày lúc 08:00 AM
# # schedule.every().day.at("13:41").do(send_daily_email)

# # def run_schedule():
# #     while True:
# #         schedule.run_pending()
# #         time.sleep(1)



# if __name__ == '__main__':
#     # schedule_thread = threading.Thread(target=run_schedule)
#     # schedule_thread.daemon = True
#     # schedule_thread.start()

#     app.run(debug=True)


from flask import Flask, render_template, request, jsonify, session
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
import schedule
import threading
from werkzeug.utils import secure_filename

app = Flask(__name__)

app.secret_key = 'trinh2001'


# Các biến toàn cục để lưu trữ thời gian được lên lịch
scheduled_time = None
scheduled_job = None

UPLOAD_FOLDER = 'uploads'  # Thư mục tạm thời để lưu trữ tệp tải lên
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def send_email(email, password, sheet_link, interval, file_paths=[]):
        smtp_server = 'smtp.gmail.com'
        smtp_port = 587
        output_df = pd.DataFrame(columns=['name', 'email', 'sent time'])

        try:
            sheet_id = sheet_link.split('/')[5]
            url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx'
            df = pd.read_excel(url)
        except Exception as e:
            print(f"Error reading Google Sheet: {e}")
            return "Invalid Google Sheet link", 400

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
                    for file_path in file_paths:
                        if os.path.isfile(file_path):
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


@app.route("/")
def hello():
    return render_template('index.html')


@app.route('/send_email', methods=['POST'])
def send_email_route():
    email = request.form['email']
    password = request.form['password']
    sheet_link = request.form['sheet_link']
    interval = int(request.form['interval']) if request.form['interval'] else 0
    file_paths = request.files.getlist('file_paths')

    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        os.makedirs(app.config['UPLOAD_FOLDER'])
    file_paths = []
    for uploaded_file in request.files.getlist('file_paths'):
        if uploaded_file.filename != '':
            filename = secure_filename(uploaded_file.filename)
            uploaded_file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            file_paths.append(os.path.join(app.config['UPLOAD_FOLDER'], filename))

    # Thực hiện việc gửi email bằng cách gọi hàm send_email với các tham số truyền vào
    result, status_code = send_email(email, password, sheet_link, interval, file_paths)
    for file_path in file_paths:
        os.remove(file_path)
    return result, status_code


def send_daily_email(email, password, sheet_link, interval, scheduled_time):
    # Thực hiện logic gửi email hàng ngày với các thông tin được truyền vào
    print('Bat dau gui email...')
    send_email(email, password, sheet_link, interval, [])
    print("Sending daily email...")


def schedule_job(email, password, sheet_link, interval, scheduled_time):
    global scheduled_job
    if scheduled_time is not None:
        scheduled_job = schedule.every().day.at(scheduled_time).do(send_daily_email, email, password, sheet_link, interval, scheduled_time)


def run_schedule():
    while True:
        schedule.run_pending()
        time.sleep(1)

@app.route("/setime")
def index():
    return render_template("main.html")

@app.route("/schedule", methods=["POST"])
def schedule_email():
    email = request.form.get("email")
    password = request.form.get("password")
    sheet_link = request.form.get("sheet_link")
    interval = int(request.form.get("interval"))
    scheduled_time = request.form.get("schedule_time")
    
    # if scheduled_job:
    #     scheduled_job.cancel()
    
    schedule_job(email, password, sheet_link, interval, scheduled_time)  # Truyền các giá trị vào hàm
    return jsonify({"message": f"Email scheduled for {scheduled_time}"}), 200


if __name__ == "__main__":
    schedule_thread = threading.Thread(target=run_schedule)
    schedule_thread.daemon = True
    schedule_thread.start()

    app.run(debug=True)