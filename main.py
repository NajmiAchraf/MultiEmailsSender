from email.message import EmailMessage
import smtplib
import pandas as pd
import time


class RecruitmentSenderForMoroccanCompanies:
    def __init__(self, csv_file, pdf_list_file, email_user, email_pass):
        self.data = pd.read_csv(csv_file)

        self.documents = pdf_list_file
        self.user_email = email_user
        self.pass_email = email_pass
        self.count = 0

        try:
            self.mail = smtplib.SMTP_SSL("smtp.gmail.com", 465)
            self.Connect()
        except Exception:
            self.Connect()

    def Connect(self):
        print('Connecting ...')
        try:
            self.mail = smtplib.SMTP_SSL("smtp.gmail.com", 465)
            print('Connected\nLogin ...')
            self.mail.login(self.user_email, self.pass_email)
            print('Login success')
            self.Send()
        except Exception:
            print('Field to connect, check if your are turn on less secure apps on : https://myaccount.google.com/lesssecureapps')
            print('1 minute to reconnect')
            time.sleep(60)
            self.Connect()

    def Send(self):
        for self.staff in range(self.count, 612):
            receiver_email = self.data.iloc[self.staff, 0]
            try:
                msg = EmailMessage()
                msg['Subject'] = "Demande/Request"
                msg['From'] = self.user_email
                msg['To'] = receiver_email
                msg.set_content("""TEXT""")

                for file in self.documents:
                    with open(file, 'rb') as f:
                        file_data = f.read()
                        file_name = f.name

                    msg.add_attachment(file_data,
                                       maintype='application',
                                       subtype='octect-stream',
                                       filename=file_name)

                self.mail.send_message(msg)
                print('Message sent to :', self.staff, receiver_email)
            except Exception:
                self.count = self.staff
                print('Disconnect')
                print('Field to send the message to', self.count, receiver_email)
                self.Connect()
        return print('Sending Complete Successfully !!')


if __name__ == '__main__':
    # https://myaccount.google.com/lesssecureapps
    pdf = ["PDF_FILE_1.pdf", "PDF_FILE_2.pdf"]

    RecruitmentSenderForMoroccanCompanies(csv_file="recruitment-email-without-double-quote.csv",
                                          pdf_list_file=pdf,
                                          email_user="email@gmail.com",
                                          email_pass="Password")
