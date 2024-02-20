import requests
from azure.identity import ClientSecretCredential


def api_azure_token_generation(tenant_id: str, client_id: str, client_secret: str, scope: str) -> str:
    auth = ClientSecretCredential(authority='https://login.microsoftonline.com/',
                                  tenant_id=tenant_id, client_id=client_id,
                                  client_secret=client_secret)
    access_token = 'Bearer ' + auth.get_token(scope).token
    return access_token


def api_get_response(url: str, access_token: str, apim_key: str) -> str:
    headers = {
        'Cache-Control': 'no-cache',
        'Authorization': access_token,
        'Ocp-Apim-Subscription-Key': apim_key,
    }
    query_result = requests.get(url=url, headers=headers)
    return query_result.json()


def api_get_query(url: str, body: str, access_token: str, apim_key: str) -> str:
    headers = {
        'Cache-Control': 'no-cache',
        'Authorization': access_token,
        'Ocp-Apim-Subscription-Key': apim_key,
    }
    query_result = requests.post(url=url, headers=headers, json=body)
    return query_result.json()
