import requests
import json


class OneNoteAPI:
    def __init__(self, app_id, client_secret, tenant_id):
        self.app_id = app_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.access_token = None
        self.refresh_token = None


    def authorize(self, username, password):
        url = f'https://login.microsoftonline.com/{self.tenant_id}/oauth2/token'

        token_data = {
            'grant_type': 'password',
            'client_id': self.app_id,
            'client_secret': self.client_secret,
            'resource': 'https://graph.microsoft.com',
            'scopes': ['onedrive.readwrite', 'offline_access'],
            'username': username,
            'password': password,
        }
        token_r = requests.post(url, data=token_data)
        self.access_token = token_r.json().get('access_token')
        self.refresh_token = token_r.json().get('refresh_token')

    def refresh_access_token(self):
        url = f'https://login.microsoftonline.com/{self.tenant_id}/oauth2/token'

        token_data = {
            'grant_type': 'refresh_token',
            'client_id': self.app_id,
            'client_secret': self.client_secret,
            'resource': 'https://graph.microsoft.com',
            'scopes': ['onedrive.readwrite', 'offline_access'],
            'refresh_token': self.refresh_token,
        }
        token_r = requests.post(url, data=token_data)
        self.access_token = token_r.json().get('access_token')
        self.refresh_token = token_r.json().get('refresh_token')


    def create_notebook(self, notebook_name):
        url = 'https://graph.microsoft.com/v1.0/me/onenote/notebooks'
        data = {
            'displayName': notebook_name
        }
        return self._post(url, json=data)

    def create_section(self, notebook_id, section_name):
        url = f'https://graph.microsoft.com/v1.0/me/onenote/notebooks/{notebook_id}/sections'
        data = {
            'displayName': section_name
        }
        return self._post(url, json=data)

    def create_page(self, section_id, data, **kwargs):
        url = f'https://graph.microsoft.com/v1.0/me/onenote/sections/{section_id}/pages'
        kwargs['headers'] = {'Content-Type': 'application/xhtml+html'}
        return self._post(url, params=data, **kwargs)


    def _get(self, url, **kwargs):
        return self._request('GET', url, **kwargs)

    def _post(self, url, **kwargs):
        return self._request('POST', url, **kwargs)

    def _put(self, url, **kwargs):
        return self._request('PUT', url, **kwargs)

    def _patch(self, url, **kwargs):
        return self._request('PATCH', url, **kwargs)

    def _delete(self, url, **kwargs):
        return self._request('DELETE', url, **kwargs)


    def _request(self, method, url, headers=None, **kwargs):
        _headers = {
            'Accept': 'application/json',
        }
        _headers['Authorization'] = 'Bearer ' + self.access_token
        if headers:
            _headers.update(headers)
        if 'files' not in kwargs:
            # If you use the 'files' keyword, the library will set the Content-Type to multipart/form-data
            # and will generate a boundary.
            _headers['Content-Type'] = 'application/json'
        return self._parse(requests.request(method, url, headers=_headers, **kwargs))

    def _parse(self, response):
        if 'application/json' in response.headers['Content-Type']:
            return response.json()
        else:
            return response
