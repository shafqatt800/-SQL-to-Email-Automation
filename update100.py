rom kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.textinput import TextInput
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.scrollview import ScrollView
from kivy.uix.popup import Popup
from kivy.uix.label import Label as PopupLabel
import smtplib
import pyodbc
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

class MyApp(App):
    def build(self):
        # Main Layout
        main_layout = BoxLayout(orientation="vertical", padding=20, spacing=10)

        # Scrollable Content
        scroll_view = ScrollView(size_hint=(1, 1))
        content_layout = GridLayout(cols=2, spacing=10, size_hint_y=None, row_default_height=50)
        content_layout.bind(minimum_height=content_layout.setter("height"))

        # Label Style
        label_style = {
            "size_hint_x": 0.3,    
            "font_size": 18,
        }

        # TextInput Style
        input_style = {
            "multiline": False,
            "size_hint_x": 0.7,
            "font_size": 16,     
        }

        # SQL Configuration Section
        content_layout.add_widget(Label(text="Server", **label_style))
        self.server_input = TextInput(**input_style)
        content_layout.add_widget(self.server_input)

        content_layout.add_widget(Label(text="Database", **label_style))
        self.database_input = TextInput(**input_style)
        content_layout.add_widget(self.database_input)

        content_layout.add_widget(Label(text="UID", **label_style))
        self.uid_input = TextInput(**input_style)
        content_layout.add_widget(self.uid_input)

        content_layout.add_widget(Label(text="Password", **label_style))
        self.password_input = TextInput(password=True, **input_style)
        content_layout.add_widget(self.password_input)

        # SQL Query with Scroll Bar
        content_layout.add_widget(Label(text="SQL Query", **label_style))
        query_scroll = ScrollView(size_hint=(0.7, None), height=150)
        self.query_input = TextInput(size_hint_y=None, height=150, multiline=True, font_size=16)
        query_scroll.add_widget(self.query_input)
        content_layout.add_widget(query_scroll)

        # Email Configuration Section
        content_layout.add_widget(Label(text="Email Header", **label_style))
        self.header_input = TextInput(**input_style)
        content_layout.add_widget(self.header_input)

        content_layout.add_widget(Label(text="Email Body", **label_style))
        email_body_scroll = ScrollView(size_hint=(0.7, None), height=150)
        self.body_input = TextInput(size_hint_y=None, height=150, multiline=True, font_size=16)
        email_body_scroll.add_widget(self.body_input)
        content_layout.add_widget(email_body_scroll)

        content_layout.add_widget(Label(text="Sender Email", **label_style))
        self.sender_input = TextInput(**input_style)
        content_layout.add_widget(self.sender_input)

        content_layout.add_widget(Label(text="Password", **label_style))
        self.email_password_input = TextInput(password=True, **input_style)
        content_layout.add_widget(self.email_password_input)

        content_layout.add_widget(Label(text="Recipient", **label_style))
        self.recipient_input = TextInput(**input_style)
        content_layout.add_widget(self.recipient_input)

        # Add content to scroll view
        scroll_view.add_widget(content_layout)
        main_layout.add_widget(scroll_view)

        # Buttons
        buttons_layout = BoxLayout(orientation="horizontal", size_hint_y=None, height=50, spacing=10)
        submit_button = Button(text="Submit", background_color=(0, 1, 0, 1), size_hint=(0.5, 1), font_size=16)
        submit_button.bind(on_press=self.submit_action)
        cancel_button = Button(text="Cancel", background_color=(1, 0, 0, 1), size_hint=(0.5, 1), font_size=16)
        cancel_button.bind(on_press=self.cancel_action)
        buttons_layout.add_widget(submit_button)
        buttons_layout.add_widget(cancel_button)
        main_layout.add_widget(buttons_layout)

        return main_layout

    def submit_action(self, instance):
        # Fetch values from input fields
        server = self.server_input.text
        database = self.database_input.text
        uid = self.uid_input.text
        password = self.password_input.text
        query = self.query_input.text
        header = self.header_input.text
        body = self.body_input.text
        sender = self.sender_input.text
        email_password = self.email_password_input.text
        recipient = self.recipient_input.text

        # Check for empty fields
        if not all([server, database, uid, password, query, header, body, sender, email_password, recipient]):
            self.show_popup("Error", "All fields must be filled!")
            return

        # try:
        #     conn_str = f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={uid};PWD={password}'
        #     conn = pyodbc.connect(conn_str)
        #     cursor = conn.cursor()
        #     cursor.execute(query)
        #     rows = cursor.fetchall()
        #     conn.close()

            
        data = self.fetch_data_from_sql_server(server, database, uid, password, query) 
            
        if data:
            # Convert data to HTML table format
            html_content = "<table border='1'><tr><th>Column1</th><th>Column2</th></tr>"
            for row in data:
                html_content += f"<tr><td>{row[0]}</td><td>{row[1]}</td></tr>"
            html_content += "</table>"
            self.send_email(header, html_content, recipient, sender, email_password)
            # self.send_email(sender, email_password, recipient, html_content)
            
            

            
            # Format email content
            # email_content = f"Subject: {header}\n\n{body}\n\nSQL Query Results:\n"
            # for row in rows:
            #     email_content += str(row) + "\n"        
    
           # Send email
            # self.send_email(sender, email_password, recipient, email_content)
            self.show_popup("Success", "Email sent successfully!")

        # except Exception as e:
        #     self.show_popup("Error", f"An error occurred: {str(e)}")
    
    def fetch_data_from_sql_server(self, server, database, uid, pwd, query):
        conn_str = f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={uid};PWD={pwd}'
        try:
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()
            cursor.execute(query)
            data = cursor.fetchall()
            conn.close()
            return data
        except Exception as e:
            self.show_popup("Error", f"Failed to fetch data: {e}")
            return []

    # def send_email(self, sender, email_password, recipient, content):
    #     try:
    #         with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
    #             server.login(sender, email_password)
    #             server.sendmail(sender, recipient, content)
    #     except Exception as e:
    #         self.show_popup("Error", f"Failed to send email: {str(e)}")
    
    def send_email(self, subject, body_html, recipient_emails, sender_email, sender_password):
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['Subject'] = subject
        msg['To'] = ', '.join(recipient_emails) if isinstance(recipient_emails, list) else recipient_emails
        msg.attach(MIMEText(body_html, 'html'))

        try:
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
                server.login(sender_email, sender_password)
                server.sendmail(sender_email, recipient_emails, msg.as_string())
                self.show_popup("Success", "Email sent successfully!")
        except Exception as e:
            self.show_popup("Error", f"Failed to send email: {e}")

    def cancel_action(self, instance):
        # Clear all input fields
        self.server_input.text = ""
        self.database_input.text = ""
        self.uid_input.text = ""
        self.password_input.text = ""
        self.query_input.text = ""
        self.header_input.text = ""
        self.body_input.text = ""
        self.sender_input.text = ""
        self.email_password_input.text = ""
        self.recipient_input.text = ""

    def show_popup(self, title, message):
        popup = Popup(title=title, content=PopupLabel(text=message), size_hint=(None, None), size=(400, 200))
        popup.open()


if __name__ == "__main__":
    MyApp().run()
