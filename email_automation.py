import os
import smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import PySimpleGUI as sg
from email.message import EmailMessage
from O365 import Message


class EmailAutomation:
    port = 465
    context = ssl.create_default_context()
    username = None
    password = None

    def initializeCredentials(self, username: str, password: str) -> None:
        self.username = username
        self.password = password

        try:
            with smtplib.SMTP_SSL("smtp.gmail.com", self.port, context=self.context) as server:
                server.login(self.username, self.password)
        except Exception as e:
            print(e)
            raise Exception("Username or password is incorrect")

    def openEmailLayout(self) -> sg.Window:

        information = [
            [sg.Frame('Information',[[
                sg.Text(
                    "Please enter the email ccs, candidates' emails, names, and positions separated by a comma.\nExample: test@gmail.com, next@gmail.com.\n"
                    "\nFor the email message, indicate where you want to include the individual's name and positions using"
                    " {name} and {position}.\nExample: Dear {name}, thank you for applying to the {position}",
                    size=(100, 5), font=("Helvetica", 14), pad=(10,10))]], font=("Helvetica",18))],
            [sg.Text('')],

            [sg.Text("Email Attachment", size=(20, 1), font=("Helvetica", 18))],

            [sg.Input(key='attachment', font=("Helvetica", 18), size=(68, 1)),
             sg.FileBrowse(font=("Helvetica", 15))],

            [sg.Text('')],

            [sg.Text("Email Subject (required)", size=(35, 1), font=("Helvetica", 18), justification="left"),
             sg.Text(''),
             sg.Text(''),
             sg.Text("Email CC", size=(35, 1), font=("Helvetica", 18), justification="left")],

            [sg.Input(key='subject', font=("Helvetica", 18), size=(35, 1), justification="left"),
             sg.Text(''),
             sg.Text(''),
             sg.Input(key='cc', font=("Helvetica", 18), size=(35, 1), justification="left")],

            [sg.Text('')],

            [sg.Text("Email Message (required)", size=(35, 1), font=("Helvetica", 18)),
             sg.Text(''),
             sg.Text(''),
             sg.Text("Candidates' Emails (required)", size=(35, 1), font=("Helvetica", 18), justification="left")],

            [sg.Multiline(size=(34, 3), font=("Helvetica", 18), key="message", justification="left"),
             sg.Text(''),
             sg.Text(''),
             sg.Multiline(size=(34, 3), font=("Helvetica", 18), key="emails", justification="right")],

            [sg.Text('')],

            [sg.Text("Candidates' Names", size=(35, 1), font=("Helvetica", 18)),
             sg.Text(''),
             sg.Text(''),
             sg.Text("Candidates' Positions", size=(35, 1), font=("Helvetica", 18))],

            [sg.Multiline(size=(34, 3), font=("Helvetica", 18), key="names"),
             sg.Text(''),
             sg.Text(''),
             sg.Multiline(size=(34, 3), font=("Helvetica", 18), key="positions")],

            [sg.Text('')],
        ]

        buttons = [
             [sg.Button('Send emails', key="send", font=("Helvetica", 15))],
        ]

        layout = [
            [sg.Column(information)],
            [sg.Column(buttons, justification="right")],
        ]

        return sg.Window("Email Information",
                         layout,
                         margins=(100,100),
                         finalize=True,
                         )

    def openLoginLayout(self) -> sg.Window:
        loginInput = [
            [sg.Text("Email", size=(10, 1), font=("Helvetica", 18)),
             sg.Input(key='email', font=("Helvetica", 18), size=(30, 1))],
            [sg.Text("Password", size=(10, 1), font=("Helvetica", 18)),
             sg.Input(key='password', font=("Helvetica", 18), size=(30, 1), password_char="*")],
        ]

        loginButtons = [
            [sg.Button('Login', font=("Helvetica", 15), bind_return_key=True)]
        ]

        layout = [
            [sg.Column(loginInput)],
            [sg.Column(loginButtons, justification="right")],

        ]

        return sg.Window("Login",
                         layout,
                         margins=(50, 50),
                         finalize=True,
                         )

    def openErrorLayout(self, errorMessage: str) -> sg.Window:
        errorInformation = [
            [sg.Text("Error: " + errorMessage, size=(30, 5), text_color="red", font=("Helvetica", 18))]
        ]

        exitButtons = [
            [sg.Button('Exit', font=("Helvetica", 15))]
        ]

        layout = [
            [sg.Column(errorInformation)],
            [sg.Column(exitButtons, justification="right")],

        ]

        return sg.Window("Error",
                         layout,
                         margins=(50, 50),
                         finalize=True,
                         text_justification="center",
                         )

    def openSuccessLayout(self) -> sg.Window:
        successInformation = [
            [sg.Text("Your emails have successfully sent!", size=(30, 5), font=("Helvetica", 18))]
        ]

        exitButtons = [
            [sg.Button('New session', key="clear", font=("Helvetica", 15)), sg.Button('Exit', font=("Helvetica", 15))],
        ]

        layout = [
            [sg.Column(successInformation)],
            [sg.Column(exitButtons, justification="right")],

        ]

        return sg.Window("Error",
                         layout,
                         margins=(50, 50),
                         finalize=True,
                         text_justification="center",
                         )

    def sendEmails(self, names: str, emails: str, positions: str, subject: str, cc: str, message: str, attachment: str) -> None:

        names = names.replace("\n", "")
        positions = positions.replace("\n", "")
        emails = emails.replace("\n", "")
        cc = cc.replace("\n", "")

        if emails == "" or message == "" or subject == "":
            raise Exception("You have empty fields, please fill them in")

        names = names.split(",")
        positions = positions.split(",")
        emails = emails.split(",")

        try:
            with smtplib.SMTP_SSL("smtp.gmail.com", self.port, context=self.context) as server:
                server.login(self.username, self.password)

                if attachment != "":
                    filename = os.path.basename(attachment)
                    file = open(attachment, "rb")
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(file.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f"attachment; filename={filename}")

                for i in range(len(names)):
                    msg = MIMEMultipart()

                    # Dear, {name}
                    # Thanks
                    # for applying to the {position}.
                    # Here
                    # Here's a link to our discord: https://discord.gg/VNAUv26y

                    msg['Subject'] = subject
                    msg['From'] = self.username
                    msg['To'] = emails[i]
                    msg['CC'] = cc
                    msg.attach(MIMEText(message.format(name=names[i], position=positions[i]), 'plain'))

                    if attachment != "":
                        msg.attach(part)

                    server.send_message(msg)
                    print(f"Email sent to {emails[i]}")



        except Exception as e:
            print(e)
            raise Exception("One of the recipient addresses are invalid")


if __name__ == "__main__":
    tool = EmailAutomation()
    sg.theme('DarkAmber')

    mainWindow = tool.openLoginLayout()
    errorWindow = None
    successWindow = None
    clearFields = ["subject", "cc", "emails", "names", "positions", "message", "attachment"]

    while True:
        window, event, values = sg.read_all_windows()
        if event == "Exit" or event == sg.WIN_CLOSED:
            window.Close()
            if window == errorWindow:
                errorWindow = None
            else:
                mainWindow = None
                break
        elif event == "Login":
            try:
                tool.initializeCredentials(values['email'], values['password'])
                window.close()
                mainWindow = tool.openEmailLayout()
            except Exception as e:
                errorWindow = tool.openErrorLayout(str(e))
        elif event == "send":
            try:
                tool.sendEmails(values['names'], values['emails'], values['positions'], values['subject'],
                                values['cc'], values['message'], values['attachment'])
                successWindow = tool.openSuccessLayout()
            except Exception as e:
                errorWindow = tool.openErrorLayout(str(e))

        elif event == "clear":
            if window == successWindow:
                window.Close()
                successWindow = None

            window = mainWindow
            for key in clearFields:
                window[key]('')


