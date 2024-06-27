from background_task import background
from django.core.mail import EmailMessage
from django.conf import settings


@background(schedule=None)
def email_background_worker(send_to, message):
    print(f'Sending email to {send_to}')

    # context = {}
    # Email configuration (replace with your actual settings)
    email = EmailMessage(
        subject='Invoice',
        body=message,
        from_email=settings.EMAIL_HOST_USER,
        to=[send_to,],
    )

    # Attach the Excel data
    email.attach_file('Invoice.xlsx')

    # Send the email
    email.send()
    # context['success'] = 'Email Sent Successfully to {send_to}'
    print(f'Sent email to {send_to}!')
