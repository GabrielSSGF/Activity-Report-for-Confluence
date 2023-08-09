from collections import Counter
import datetime as dt
from datetime import datetime
import requests
from requests.auth import HTTPBasicAuth
import pandas as pd


auth = HTTPBasicAuth("email", "token")

headers = {
    'Accept': 'application/json'
}

def filterYear(data):
    this_year = dt.datetime.now().year
    users_this_year = [
        page['authorId']
        for page in data['results']
        if dt.datetime.fromisoformat(page['createdAt'].replace('Z', '+00:00')).year == this_year
    ]
    return users_this_year

def contar_ocorrencias(lista):
    return Counter(lista)

def getUserName(UserID, cache={}):
    if UserID in cache:
        return cache[UserID]

    base_url = f'https://your-space.atlassian.net/wiki/rest/api/user?accountId={UserID}'
    params = {
        'cql': f"created >= startOfYear()",
        'limit': 250,
        'start': 0
    }
    try:
        response = requests.get(base_url, headers=headers, auth=auth, params=params)
        data = response.json()
        if 'publicName' in data:
            nomeUsuario = data['publicName']
        else:
            nomeUsuario = 'Usuário Desconhecido'
        cache[UserID] = nomeUsuario
        return nomeUsuario
    except Exception as e:
        return f'Erro ao buscar informações do usuário: {e}'

def getUsersID(ID):
    base_url = f'https://your-space.atlassian.net/wiki/api/v2/spaces/{ID}/pages'
    params = {
        'cql': f"created >= startOfYear()",
        'limit': 250,
        'start': 0
    }
    response = requests.get(base_url, headers=headers, auth=auth, params=params)
    data = response.json()
    user_ID = filterYear(data)
    return user_ID

def getPagesID():
    base_url = 'https://your-space.atlassian.net/wiki/api/v2/spaces/'
    params = {
        'cql': f"created >= startOfYear()",
        'limit': 250,
        'start': 0
    }
    try:
        response = requests.get(base_url, headers=headers, params=params, auth=auth)
        data = response.json()
        page_ID = [page['id'] for page in data.get('results', [])]
        return page_ID
    except Exception as e:
        print(f"Error fetching page IDs: {e}")
        return []

def getUserPagesAmount():
    try:
        pageIDs = getPagesID()
        userNamesList = []

        for page_id in pageIDs:
            userIDs = getUsersID(page_id)
            userNames = [getUserName(user_id) for user_id in userIDs]
            userNamesList.extend(userNames)

        contagem_por_pessoa = contar_ocorrencias(userNamesList)
        df = pd.DataFrame(contagem_por_pessoa.items(), columns=['Analista', 'Total Criado 2023'])
        df_ordenado = df.sort_values(by='Total Criado 2023', ascending=False)
        return escreverPaginaConfluence(df_ordenado)
    except Exception as e:
        print(f"Error processing user pages amount: {e}")
        return None
    
def padrao_de_colunas(worksheet):
    worksheet.column_dimensions['A'].width = 35
    worksheet.column_dimensions['B'].width = 35

def exportacao_xlsx(data_frames):
    writer = pd.ExcelWriter("INSERT THE ABSOLUTE PATH FOR THE FILE", engine='openpyxl')
    data_frames.to_excel(writer, sheet_name='Sheet')
    workbook = writer.book
    worksheet = workbook['Sheet']
    padrao_de_colunas(worksheet)
    writer._save()

def escreverPaginaConfluence(df):
    table_html = df.to_html(index=False)
    url = f"https://your-space.atlassian.net/wiki/rest/api/content/INSERTPAGEIDHERE"
    response = requests.get(url, headers=headers, auth=auth)
    dados = response.json()
    versaoPage = dados['version']['number']

    update_data = {
        'version': {
            'number': versaoPage + 1
        },
        'title': '[INSERT YOUR PAGE TITLE]',
        'type': 'page',
        'body': {
            'storage': {
                'value': f'<html><body>{table_html}</body></html>',
                'representation': 'storage'
            }
        }
    }
    print(requests.put(url, json=update_data, auth=auth, headers=headers))

getUserPagesAmount()

