import msal
import os
from office365.graph_client import GraphClient
from dotenv import load_dotenv

load_dotenv("test.env")

client_id = os.getenv('app_id')
client_secret = os.getenv('client_secret')
tenant_id = os.getenv('azure_active_directory_id')

username = os.getenv('email')
password = os.getenv('password')


def acquire_token():
    """
    Acquire token via MSAL
    """
    global client_id, client_secret, tenant_id
    authority_url = f'https://login.microsoftonline.com/{tenant_id}'
    app = msal.ConfidentialClientApplication(
        authority=authority_url,
        client_id='{client_id}',
        client_credential='{client_secret}'
    )
    token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return token

def acquire_token_func():
    """
    Acquire token via MSAL
    """
    global client_id, client_secret, tenant_id
    authority_url = f'https://login.microsoftonline.com/{tenant_id}'
    app = msal.ConfidentialClientApplication(
        authority=authority_url,
        client_id='{client_id}',
        client_credential='{client_secret}'
    )
    token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return token

client = GraphClient(acquire_token_func)

client.me.send_mail(
    subject="Meet for lunch?",
    body="The new cafeteria is open.",
    to_recipients=["shanto@usc.edu"]
).execute_query()

# client = GraphClient(acquire_token)
client = GraphClient(acquire_token_func)



files = {}
with open("./test.html", 'rb') as f, \
    open("./test.png", 'rb') as img_f, \
    open("./test.pdf", 'rb') as pdf_f:

    files["imageBlock1"] = img_f
    files["fileBlock1"] = pdf_f

    page = client.me.onenote.pages.add(presentation_file=f, attachment_files=files).execute_query()
