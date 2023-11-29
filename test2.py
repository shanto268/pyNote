import os
from azure_api import *
from dotenv import load_dotenv

load_dotenv("test.env")

def main():
    app_id = os.getenv('app_id')
    client_secret = os.getenv('client_secret')
    tenant_id = os.getenv('azure_active_directory_id')

    username = os.getenv('email')
    password = os.getenv('password')

    client = OneNoteAPI(app_id, client_secret, tenant_id)
    client.authorize(username, password)

    for i in range(1, 3):
        notebook = client.create_notebook(f'Notebook_{i}')
        try:
            notebook_id = notebook['id']
        except KeyError:
            client.refresh_access_token()
            notebook = client.create_notebook(f'Notebook_{i}')
            print(notebook)
            notebook_id = notebook['id']
        print(f'Notebook {i} created')
        for j in range (1, 3):
            section = client.create_section(notebook_id, f'Note_{i}_Section_{j}')
            try:
                section_id = section['id']
            except KeyError:
                client.refresh_access_token()
                section = client.create_section(notebook_id, f'Note_{i}_Section_{j}')
                print(section)
                section_id = section['id']
            print(f'Notebook {i} Section {j} created')
            for k in range(1, 3):
                page_data = f'<html><head> <title>N_{i}_Sec_{j}_Page_{k}</title> <style>div{{position: absolute; left: 50%; transform: translate(-50%,0); text-align: center;}}h3, button{{text-align: center; font-size: 2rem;}}button{{margin:auto;}}</style></head><body> <div> <h3>Without any modifications</h3> <button>Hello</button> <h3>With padding</h3> <button style="padding:30px;">Hello</button> <h3>With margin</h3> <button style="margin:30px;">Hello</button> <h3>With margin and padding</h3> <button style="margin:30px; padding:30px;">Hello</button> </div></body><html>'
                response = client.create_page(section_id, page_data)
                try:
                    if response.status_code == 401:
                        client.refresh_access_token()
                        print(client.create_page(section_id, page_data))
                except AttributeError:
                    pass
                print(f'Notebook {i} Section {j} Page {k} created')

if __name__ == '__main__':
    main()

