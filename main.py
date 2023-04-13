import datetime
import email
import imaplib
import re
import smtplib
import sys
import time

import gensim.downloader as api
import nltk
import numpy as np
import pytz
from PySide2.QtWidgets import *
from nltk import word_tokenize
from nltk.corpus import stopwords, wordnet
from sklearn.metrics.pairwise import cosine_similarity

from PySide2.QtWidgets import QApplication, QTextEdit, QPushButton, QVBoxLayout, QWidget, \
    QLineEdit, QLabel

word_vectors = api.load("glove-wiki-gigaword-100")
nltk.download('stopwords')
nltk.download('punkt')
nltk.download('wordnet')
stop_words = set(stopwords.words('english'))


def preprocess_text(text):
    tokens = re.findall(r'\b\w+\b', text)
    tokens = [token.lower() for token in tokens if token.lower() not in stop_words]
    vectors = [word_vectors[token] for token in tokens if token in word_vectors]
    if len(vectors) > 0:
        return np.mean(vectors, axis=0)
    else:
        return np.zeros(100)


# log in window that directs to main window when you're logged in successfully
class LoginWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('Log in')
        self.setGeometry(300, 300, 400, 100)
        self.username_input = QLineEdit()
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        self.login_button = QPushButton('Log in')
        self.login_button.clicked.connect(self.login)

        layout = QVBoxLayout()
        layout.addWidget(QLabel('Username:'))
        layout.addWidget(self.username_input)
        layout.addWidget(QLabel('Password:'))
        layout.addWidget(self.password_input)

        button_layout = QHBoxLayout()
        button_layout.addStretch()
        button_layout.addWidget(self.login_button)
        button_layout.addStretch()
        layout.addLayout(button_layout)

        self.setLayout(layout)

    def login(self):
        username = self.username_input.text()
        password = self.password_input.text()
        try:
            name, rest = username.split("@")
            service, domain = rest.split(".")
        except ValueError:
            QMessageBox.information(self, "Log in failed", "Incorrect email or password")

        imap_host = 'poczta.interia.pl'
        smtp_host = 'poczta.interia.pl'
        imap_port = 993
        smtp_port = 587

        # checking which email service is used
        if service == "interia":
            imap_host = 'poczta.interia.pl'
            smtp_host = 'poczta.interia.pl'
            imap_port = 993
            smtp_port = 587
        elif service == "gmail":
            imap_host = 'imap.gmail.com'
            smtp_host = 'smtp.gmail.com'
            imap_port = 993
            smtp_port = 587
        elif service == "wp":
            imap_host = 'imap.wp.pl'
            smtp_host = 'smtp.wp.pl'
            imap_port = 993
            smtp_port = 465

        imap_server = imaplib.IMAP4_SSL(imap_host, imap_port)

        try:
            imap_server.login(username, password)
            imap_server.select("INBOX")
            status, messages = imap_server.search(None, "ALL")

            if status == 'OK':
                self.close()
                self.main_window = EmailApplication(username, password, imap_host, imap_port, smtp_host, smtp_port)
                self.main_window.show()

        except imaplib.IMAP4.error:
            QMessageBox.information(self, "Log in failed", "Incorrect email or password")


# main window
class EmailApplication(QWidget):
    def __init__(self, user_from_login, password_from_login, imap_host_form_login, imap_port_from_login,
                 smtp_host_from_login, smtp_port_from_login):
        super().__init__()
        self.setWindowTitle('Email')
        self.setGeometry(300, 300, 800, 600)

        self.tabs = QTabWidget(self)
        self.send_tab = QWidget()
        self.receive_tab = QWidget()
        self.search_tab = QWidget()

        self.tabs.addTab(self.send_tab, 'Send Email')
        self.tabs.addTab(self.receive_tab, 'Received')
        self.tabs.addTab(self.search_tab, 'Search by key word')

        self.to_label = QLabel('To:')
        self.to_line_edit = QLineEdit()
        self.subject_label = QLabel('Subject:')
        self.subject_line_edit = QLineEdit()
        self.body_label = QLabel('Body:')
        self.body_text_edit = QTextEdit()
        self.send_button = QPushButton('Send')

        send_layout = QVBoxLayout()
        form_layout = QFormLayout()
        form_layout.addRow(self.to_label, self.to_line_edit)
        form_layout.addRow(self.subject_label, self.subject_line_edit)
        send_layout.addLayout(form_layout)
        send_layout.addWidget(self.body_label)
        send_layout.addWidget(self.body_text_edit)
        send_layout.addWidget(self.send_button)
        self.send_tab.setLayout(send_layout)

        self.send_button.clicked.connect(self.send_email)

        self.email_list = QListWidget()
        self.email_text_edit = QTextEdit()
        self.email_text_edit.setReadOnly(True)
        self.email_display = QTextEdit()
        self.refresh_button = QPushButton('Refresh')

        receive_layout = QVBoxLayout()
        receive_layout.addWidget(self.email_list)
        receive_layout.addWidget(self.email_text_edit)
        receive_layout.addWidget(self.refresh_button)
        self.receive_tab.setLayout(receive_layout)

        self.refresh_button.clicked.connect(self.refresh_emails)
        self.email_list.itemClicked.connect(self.display_email)

        search_layout = QVBoxLayout()
        search_form = QFormLayout()

        self.search_button = QPushButton('Search')
        self.search_button.clicked.connect(self.search_emails)
        self.search_emails_edit = QLineEdit()
        self.search_label = QLabel('Key words:')
        self.found_emails = QListWidget()
        self.email_body = QTextEdit()
        self.email_body.setReadOnly(True)

        search_form.addRow(self.search_label, self.search_emails_edit)
        search_form.addWidget(self.search_button)
        search_form.addWidget(self.found_emails)
        search_form.addWidget(self.email_body)
        self.found_emails.itemClicked.connect(self.display_email)

        search_layout.addLayout(search_form)

        self.search_tab.setLayout(search_layout)

        layout = QVBoxLayout()
        layout.addWidget(self.tabs)
        self.setLayout(layout)

        # data to log in
        self.smtp_host = smtp_host_from_login
        self.smtp_port = smtp_port_from_login
        self.imap_host = imap_host_form_login
        self.imap_port = imap_port_from_login
        self.user = user_from_login
        self.password = password_from_login

    def send_email(self):
        to = self.to_line_edit.text()
        subject = self.subject_line_edit.text()
        body = self.body_text_edit.toPlainText()

        message = f"Subject: {subject}\n\n{body}"

        server_smtp = smtplib.SMTP(self.smtp_host, self.smtp_port)
        server_smtp.ehlo()
        server_smtp.starttls()
        server_smtp.ehlo()

        server_smtp.login(self.user, self.password)
        server_smtp.sendmail(self.user, to, message)
        server_smtp.quit()

        QMessageBox.information(self, "Email Sent", "Your email has been sent!")
        self.to_line_edit.clear()
        self.subject_line_edit.clear()
        self.body_text_edit.clear()

    def refresh_emails(self):
        self.email_list.clear()

        # getting the date of last refresh to auto-respond to new messages
        try:
            with open('last_refresh.txt', 'r') as f:
                date_last = datetime.datetime.strptime(f.read(), '%Y-%m-%d %H:%M:%S')
        except FileNotFoundError:
            date_last = datetime.datetime.now(pytz.timezone('Europe/Warsaw')) - datetime.timedelta(days=1)
            f = open('last_refresh.txt', 'w')

        date_last_strf = date_last.strftime('%Y-%m-%d %H:%M:%S')

        imap_server = imaplib.IMAP4_SSL(self.imap_host, self.imap_port)
        imap_server.login(self.user, self.password)
        imap_server.select("INBOX")

        status, messages = imap_server.search(None, "ALL")

        if status == 'OK' and messages[0]:
            for message in messages[0].split():
                _, msg = imap_server.fetch(message, "(RFC822)")
                email_message = email.message_from_bytes(msg[0][1])

                subject = email_message['Subject']
                from_address = email_message["From"]
                date_tuple = email.utils.parsedate_tz(email_message['Date'])
                if date_tuple is None:
                    continue

                date_str = datetime.datetime.fromtimestamp(email.utils.mktime_tz(date_tuple)).strftime(
                    '%Y-%m-%d %H:%M:%S')

                item_text = f"{from_address}: {subject} /{date_str}"
                self.email_list.addItem(item_text)

                # checking if emails were received before or after last refresh
                if date_str > date_last_strf:
                    response_email = email.message.EmailMessage()
                    response_email['From'] = self.user
                    response_email['To'] = from_address
                    response_email['Subject'] = "Auto-Response to " + subject
                    response_email.set_content(
                        "Thank you for your email. We have received it and will get back to you shortly.")
                    server_smtp = smtplib.SMTP(self.smtp_host, self.smtp_port)
                    server_smtp.ehlo()
                    server_smtp.starttls()
                    server_smtp.ehlo()
                    server_smtp.login(self.user, self.password)

                    # sending response email
                    server_smtp.sendmail(self.user, from_address, response_email.as_string())
                    server_smtp.quit()

        # writing time of refresh to a file
        with open('last_refresh.txt', 'w') as f:
            now = datetime.datetime.now(pytz.timezone('Europe/Warsaw'))
            f.write(now.strftime('%Y-%m-%d %H:%M:%S'))

    def display_email(self):
        server_imap = imaplib.IMAP4_SSL(self.imap_host, self.imap_port)
        server_imap.login(self.user, self.password)
        server_imap.select("inbox")

        selected_item = self.email_list.currentItem()
        selected_index = self.tabs.currentIndex()
        selected_tab = self.tabs.tabText(selected_index)

        # checking in what tab we are
        if selected_tab == "Received":
            selected_item = self.email_list.currentItem()
        elif selected_tab == "Search by key word":
            selected_item = self.found_emails.currentItem()

        if selected_item is not None:
            selected_text = selected_item.text()
            from_address, subject = selected_text.split(": ")
            subject, date = subject.split(" /")
            email_pattern = r'<([^>]+)>'

            if "<" in from_address:
                match = re.search(email_pattern, from_address)
                if match:
                    email_address = match.group(1)
                else:
                    email_address = ""
            else:
                email_address = from_address

            # searching for correct email to display
            status, messages = server_imap.search(None, f'(FROM "{email_address}" SUBJECT "{subject}")')

            for message in messages[0].split():
                status, msg = server_imap.fetch(message, "(BODY.PEEK[])")
                if status == 'OK':
                    email_message = email.message_from_bytes(msg[0][1])
                    message_subject = email_message['Subject']
                    date_tuple = email.utils.parsedate_tz(email_message['Date'])
                    if date_tuple is None:
                        continue

                    date_str = datetime.datetime.fromtimestamp(email.utils.mktime_tz(date_tuple)).strftime(
                        '%Y-%m-%d %H:%M:%S')
                    if subject == message_subject and date == date_str:
                        # getting the message content
                        if email_message.is_multipart():
                            message_content = ""
                            for part in email_message.get_payload():
                                if part.get_content_type() == 'text/plain':
                                    message_content += part.get_payload(decode=True).decode('utf-8')
                        else:
                            message_content = email_message.get_payload(decode=True).decode('utf-8')

                        if selected_tab == "Received":
                            self.email_text_edit.setText(f"From: {email_message['From']}\n"
                                                         f"To: {email_message['To']}\n"
                                                         f"Subject: {message_subject}\n"
                                                         f"Date: {date_str}\n\n"
                                                         f"{message_content}")
                        elif selected_tab == "Search by key word":
                            self.email_body.setText(f"From: {email_message['From']}\n"
                                                    f"To: {email_message['To']}\n"
                                                    f"Subject: {message_subject}\n"
                                                    f"Date: {date_str}\n\n"
                                                    f"{message_content}")

            server_imap.close()
            server_imap.logout()

    def search_emails(self):
        self.found_emails.clear()
        self.email_body.clear()
        text_from_search = self.search_emails_edit.text()

        # splitting the keywords
        if "," in text_from_search:
            keywords = text_from_search.split(',')
        else:
            keywords = text_from_search.split(' ')

        server_imap = imaplib.IMAP4_SSL(self.imap_host, self.imap_port)
        server_imap.login(self.user, self.password)

        # getting the synonyms to keywords
        synonyms = []
        for word in keywords:
            word_synonyms = wordnet.synsets(word)
            for syn in word_synonyms:
                for lemma in syn.lemmas():
                    if lemma.name() not in synonyms:
                        synonyms.append(lemma.name())

        stop_words_eng = set(stopwords.words('english'))
        mails = []

        server_imap.select("inbox")
        status, messages = server_imap.search(None, "ALL")

        if status == 'OK' and messages[0]:
            for message in messages[0].split():
                status, msg = server_imap.fetch(message, '(RFC822)')
                if status == 'OK':
                    email_message = email.message_from_bytes(msg[0][1])
                    if email_message.is_multipart():
                        message_content = ""
                        for part in email_message.get_payload():
                            if part.get_content_type() == 'text/plain':
                                message_content += part.get_payload(decode=True).decode('utf-8')
                    else:
                        email_m = email_message.get_payload(decode=True).decode('utf-8')
                    tokens = word_tokenize(email_m.lower())
                    from_address = email_message['From']
                    subject = email_message['Subject']

                    # not checking in auto-response emails
                    if "auto" in subject and "response" in  subject:
                        continue

                    t_sub = word_tokenize(subject.lower())
                    date_tuple = email.utils.parsedate_tz(email_message['Date'])
                    date_str = time.strftime('%Y-%m-%d %H:%M:%S', time.gmtime(time.mktime(date_tuple[:9])))
                    item_text = f"{from_address}: {subject} /{date_str}"

                    filtered = [word for word in tokens if word not in stop_words_eng and word.isalnum()]
                    mails.append(email_m)
                    f_sub = [word for word in t_sub if word not in stop_words_eng and word.isalnum()]

                    # calculating cosine similarities to find emails based on the context
                    keyword_vector = preprocess_text(text_from_search)
                    email_vector = preprocess_text(email_m)
                    sim = cosine_similarity(keyword_vector.reshape(1, -1), email_vector.reshape(1, -1))[0][0]

                    # accepting only emails where contextual similarity is higher than 0.6
                    if sim > 0.6:
                        self.found_emails.addItem(item_text)

                    # also checking based on synonyms
                    elif any(word.lower() in synonyms for word in filtered) or any(
                            word.lower() in synonyms for word in f_sub):
                        self.found_emails.addItem(item_text)


if __name__ == "__main__":
    app = QApplication([])
    email_application = LoginWindow()
    email_application.show()
    app.exec_()
