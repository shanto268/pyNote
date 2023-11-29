import requests
from msal import ConfidentialClientApplication
from msal import PublicClientApplication
from dotenv import load_dotenv
import io
import base64


# Read credentials from environment variables

class NoteBook:
    def __init__(self, name, page, title):
        self.name = name
        self.page = page
        self.title = title
        self.client_id = os.getenv("CLIENT_ID")
        self.client_secret = os.getenv("CLIENT_SECRET")
        self.tenant_id = os.getenv("TENANT_ID")
        self.token = self.authenticate()
        self.nb_id, self.section_id = self.get_or_create_nb_and_section()

    def authenticate(self):
        authority = f'https://login.microsoftonline.com/{self.tenant_id}'
        app = ConfidentialClientApplication(
            self.client_id,
            authority=authority,
            client_credential=self.client_secret,
        )

        token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        return token

    def make_request(self, method, url, data=None, json=None):
        headers = {
            "Authorization": f"Bearer {self.token['access_token']}",
            "Content-Type": "application/json"
        }
        response = requests.request(method, url, headers=headers, json=json)
        response.raise_for_status()
        return response.json()

    def get_or_create_nb_and_section(self):
        nb_id, section_id = None, None

        # Get or create notebook
        url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks"
        notebooks = self.make_request("GET", url).get('value', [])
        for nb in notebooks:
            if nb["displayName"] == self.name:
                nb_id = nb["id"]
                break
        if nb_id is None:
            data = {"displayName": self.name}
            response = self.make_request("POST", url, json=data)
            nb_id = response["id"]

        # Get or create section
        url = f"https://graph.microsoft.com/v1.0/me/onenote/notebooks/{nb_id}/sections"
        sections = self.make_request("GET", url).get('value', [])
        for section in sections:
            if section["displayName"] == self.page:
                section_id = section["id"]
                break
        if section_id is None:
            data = {"displayName": self.page}
            response = self.make_request("POST", url, json=data)
            section_id = response["id"]

        return nb_id, section_id

    def print(self, content):
        url = f"https://graph.microsoft.com/v1.0/me/onenote/sections/{self.section_id}/pages"

        html_content = f"""
        <!DOCTYPE html>
        <html>
            <head>
                <title>{self.title}</title>
                <style>
                    pre {{
                        font-family: 'Courier New', monospace;
                    }}
                </style>
            </head>
            <body>
                <pre>{content}</pre>
            </body>
        </html>
        """

        data = {"presentation": {"contentType": "text/html", "content": html_content}}
        self.make_request("POST", url, json=data)


    def savefig(self, fig, file_name):
        # Save the Matplotlib figure to a bytes buffer
        buf = io.BytesIO()
        fig.savefig(buf, format='png')
        buf.seek(0)
        img_base64 = base64.b64encode(buf.read()).decode('utf-8')

        url = f"https://graph.microsoft.com/v1.0/me/onenote/sections/{self.section_id}/pages"

        html_content = f"""
        <!DOCTYPE html>
        <html>
            <head>
                <title>{file_name}</title>
            </head>
            <body>
                <img src="data:image/png;base64,{img_base64}" />
            </body>
        </html>
        """

        data = {"presentation": {"contentType": "text/html", "content": html_content}}
        self.make_request("POST", url, json=data)


    def acquire_token(self):
            # Create a Public Client Application
            self.app = PublicClientApplication(
                self.client_id,
                authority=f"https://login.microsoftonline.com/{self.tenant_id}"
            )

            # Try to acquire a token silently
            accounts = self.app.get_accounts()
            if accounts:
                self.token = self.app.acquire_token_silent(
                    ["https://graph.microsoft.com/.default"],
                    account=accounts[0]
                )

            # If silent acquisition fails, acquire a token interactively
            if not self.token:
                self.token = self.app.acquire_token_interactive(
                    ["https://graph.microsoft.com/.default"]
                )

            # Raise an exception if token acquisition still fails
            if not self.token:
                raise Exception("Failed to acquire token.")
