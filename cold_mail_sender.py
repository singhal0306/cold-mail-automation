import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import pandas as pd
from mail_templates import mern_cold_mail

sender_password = ""

smtp_server = 'smtp.gmail.com'
smtp_port = 465

path_to_resume = 'Suryansh_Singhal_Resume1.pdf'

sender_email = "singhalsuryansh.0306@gmail.com"

your_name = "Suryansh Singhal"
job_role = "Software Development"

subject = f"Seeking Opportunities in {job_role}"
years_of_experience = 1
phone_number = "+91-9837490573"

def create_mail(company_name, recipient_email):
    message = MIMEMultipart()

    message['Subject'] = subject
    message['From'] = sender_email
    message['To'] = recipient_email
    body = mern_cold_mail

    formatted_body = body.format(
        employee_name="sir/madam",
        your_name=your_name,
        company_name=company_name,
        years_of_experience=years_of_experience,
        phone_number=phone_number
    )
    body_part = MIMEText(formatted_body)
    message.attach(body_part)

    # Attach Resume File with body.
    with open(path_to_resume,'rb') as file:
        message.attach(
            MIMEApplication(
                file.read(), 
                Name=f'{your_name.upper().replace(" ", "_")}_RESUME.pdf'
            )
        )
        
    return message.as_string()    

def main():
    hr_contact = pd.read_csv("hr_list.csv")
    
    df = pd.DataFrame(columns=["Company Name", "Email", "Remark"])
    sent_df = pd.DataFrame(columns=["Company Name", "Email", "Remark"])
    error_df = pd.DataFrame(columns=["Company Name", "Email", "Remark"])
    
    with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
        server.login(sender_email, sender_password)
        for index, row in hr_contact.iterrows():
            email_list = [email.strip() for email in row["Email ID"].split(",")]
            employee_name = []
            for email in email_list:
                try:
                    mail_body = create_mail(row['Name of Company'], email)
                    print(mail_body)
                    # server.sendmail(sender_email, recipient_email, mail_body)
                except Exception as e:
                    error_df = pd.concat([
                        error_df, 
                        pd.DataFrame([{
                                "Company Name": row['Name of Company'], 
                                "Email": email, 
                                "Remark": str(e)
                            }
                        ]
                    )], ignore_index=True)
                else:
                    df = pd.concat([
                        df, 
                        pd.DataFrame([{
                                "Company Name": row['Name of Company'], 
                                "Email": email, 
                                "Remark": "email sent"
                            }
                        ]
                    )], ignore_index=True)

    sent_df.to_csv("emails_sent.csv", index=False)
    error_df.to_csv("email_errors.csv", index=False)
    

if __name__ == "__main__":
    main()

# hr_contact.iloc[26:].head()
# df.to_csv(f"logs{random_number}.csv",index=False)


