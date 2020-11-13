# Phishing From MS365 Graph API
A Couple of Python Scripts Leveraging MS365's GraphAPI to Send Custom Calendar Events / Emails from Cheap O365 Accounts

This project uses O365 Python project @ https://pypi.org/project/O365/#calendar. I created these for a Phishing engagement I was on, where I was asked to send customized calendar meeting invites as well as emails using a template to over 2,000 individuals. The invites and emails needed to be from one sender to one recipient and had to include unique links for tracking/analytics purposes. These scripts both leverage the Microsoft/Office 365 Graph API. I used GoDaddy for my domains and it was easy to sign up for a cheap O365 account through them for this purpose.

The scripts will loop through each row of a spreadsheet and, depending on which script, will either send a calendar event set to the nearest 30 minutes (default) or email with an HTML body to each recipient on the sheet. It can also pull in other columns, such as the recipients name, and can replace the HTML template in memory to customize the experience for spear-phishing purposes. All in all it's a pretty simple set of scripts, but I'm sharing in case it helps others with a simliar task!

I also have one script, "O365SendSingleEmailFormBody.py" that uses a new phishing technique I discovered where you can embed an HTML form for phishing credentials directly into the body of an email. Instead of mass-mailings this script is meant to send to a single recipient, but can be altered easily by looking at the others as examples.

Thanks to the Black Hills Security folks (Beau Bullock & Michael Felch) for the inspiration with their GSuite Mailsniper Calendar Injection tool and techniqe!

## Sending Calendar Events / Emails

Clone this repo and edit the relevant python variables. Update the Template spreadsheet with targets.

Authentication instruction taken from  https://pypi.org/project/O365/#authentication:

1) Login at Azure Portal (App Registrations) @ https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade

2) Create an app. Set a name.

3) In Supported account types choose "Accounts in any organizational directory and personal Microsoft accounts (e.g. Skype, Xbox, Outlook.com)", if you are using a personal account.

4) Set the redirect uri (Web) to: https://login.microsoftonline.com/common/oauth2/nativeclient and click register. This needs to be inserted into the "Redirect URI" text box as simply checking the check box next to this link seems to be insufficent. This is the default redirect uri used by this library, but you can use any other if you want.

5) Write down the Application (client) ID. You will need this value.

6) Under "Certificates & secrets", generate a new client secret. Set the expiration preferably to never. Write down the value of the client secret created now. It will be hidden later on.

7) Under Api Permissions:

    When authenticating "on behalf of a user":
            Add the delegated permissions for Microsoft Graph you want (see scopes).
            
            It is highly recommended to add "offline_access" permission. If not the user you will have to re-authenticate every hour.
            
    When authenticating "with your own identity":
            Add the application permissions for Microsoft Graph you want.
            
            Click on the Grant Admin Consent button (if you have admin permissions) or wait until the admin has given consent to your application.
            
 8) To read / send emails AND inject calendar events use:

        Microsoft Graph - Calendars - Calendars.ReadWrite
        Microsoft Graph - Mail - Mail.Send
        Microsoft Graph - Mail - Mail.ReadWrite
        Microsoft Graph - User - User.Read

 9) Then you need to login for the first time to get the access token that will grant access to the user resources.
