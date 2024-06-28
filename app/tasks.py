from background_task import background
from django.core.mail import EmailMessage
from django.conf import settings


@background(schedule=None)
def email_background_worker(send_to, message, party_name, cc_list=None):
    print(f'Sending email to {send_to}')

    subject = f"Subject: GSTR-2B RECO F.Y. 2023-2025-April-23-April-24 SGPNC-{party_name}"
    
    email = EmailMessage(
        subject=subject,
        body=message,
        from_email=settings.EMAIL_HOST_USER,
        to=[send_to,],
        cc=cc_list,
    )

    # Attach the Excel data
    email.attach_file('Invoice.xlsx')

    email.content_subtype = 'html'  # Set the email to HTML

    # Attach the image
    with open('app/media/save tree.jpg', 'rb') as img:
        email.attach('app/media/save tree.jpg', img.read(), 'image/jpeg')
    
    # Send the email
    email.send()
    # context['success'] = 'Email Sent Successfully to {send_to}'
    print(f'Sent email to {send_to} with cc :{cc_list}')
            
