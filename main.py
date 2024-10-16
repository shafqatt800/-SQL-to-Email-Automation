import pyodbc
  
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Step 1: Function to fetch data from SQL S
def send_email(subject, body_html, recipient_emails):
    sender_email = "efgfoodvally@gmail.com"
    sender_password = "lvctovnjpwtqbirg"  # Replace with app-specific password
    
    # Set up the MIME
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['Subject'] = subject
    
    # If recipient_emails is a list, join them into a single string for email header
    if isinstance(recipient_emails, list):
        msg['To'] = ', '.join(recipient_emails)
    else:
        msg['To'] = recipient_emails

    # Attach the HTML body
    msg.attach(MIMEText(body_html, 'html'))

    # Establish a secure session with Gmail's SMTP server using SSL
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender_email, sender_password)
            # If recipient_emails is a list, send to all
            if isinstance(recipient_emails, list):
                server.sendmail(sender_email, recipient_emails, msg.as_string())
            else:
                server.sendmail(sender_email, [recipient_emails], msg.as_string())
            print("Email sent successfully!")
    except Exception as e:
        print(f"Failed to send email: {e}")

# Step 1: Function to fetch data from SQL Server
def fetch_data_from_sql_server():
    # Connect to the SQL Server database
    conn_str = (
        'DRIVER={SQL Server Native Client 11.0};'
        'SERVER=DESKTOP-MM6V0P3;'  # No space between SERVER= and server name
        'DATABASE=AdventureWorks2017;'
        'UID=sa;'  
        'PWD=erp;' 
        # Remove 'Trusted_connection={yes}' when using UID and PWD
    )    

    # Create a connection
    conn = pyodbc.connect(conn_str)

    # Create a cursor to execute the query
    cursor = conn.cursor()

    # Write your SQL query
    query = "SELECT * FROM HumanResources.Department;"
    
    # Execute the query
    cursor.execute(query)

    # Fetch all rows and print them
    data = []
    for row in cursor.fetchall():
        data.append((row.DepartmentID, row.Name))
        print(row.DepartmentID, row.Name)

    # Close the connection
    conn.close()

    # Return the fetched data
    return data



# Step 4: Main function to fetch data and send email
def main():
    # Fetch data from SQL Server
    # Fetch data from SQL Server
    data = fetch_data_from_sql_server()

    # Check if data is fetched successfully
    if data:
        print("Data fetched successfully")
        
    # Convert the fetched data to HTML format
    html_content = "<table border='1'><tr><th>DepartmentID</th><th>Name</th></tr>"
    for row in data:
        html_content += f"<tr><td>{row[0]}</td><td>{row[1]}</td></tr>"
    html_content += "</table>"

    # Print the HTML content (for testing)
    print(html_content)
    send_email(subject="test",body_html=html_content,recipient_emails=["shafqattoky@gmail.com","sabrina.afrin@efgt.eurofoods-bd.com","akhlaqur.nsu0077@gmail.com"])
if __name__ == "__main__":
    main()
    