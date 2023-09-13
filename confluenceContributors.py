from collections import Counter
import datetime as dt
from datetime import datetime
import requests
from requests.auth import HTTPBasicAuth
import pandas as pd
import json

def main():
    global AUTH, HEADERS, configData
    configData = getDataFromJson()
    AUTH, HEADERS = setRequestsData()
    getUserPagesAmount()

def getDataFromJson():
    with open('configData.json', 'r') as arquivoJson:
        configData = json.load(arquivoJson)
    return configData

def setRequestsData():
    auth = HTTPBasicAuth(configData["Email"], configData["APIToken"])
    headers = {
        'Accept': 'application/json'
    }
    return auth, headers

def getUserPagesAmount():
    pageIDs = getPagesID()
    userNamesList = []

    for pageID in pageIDs:
        userIDs = getUsersID(pageID)
        userNames = [getUserName(userID) for userID in userIDs]
        userNamesList.extend(userNames)

    contagemPorPessoa = contarOcorrencias(userNamesList)
    dataFrame = pd.DataFrame(contagemPorPessoa.items(), columns=['Analista', 'Total Criado Neste Ano'])
    dataFrameOrdenado = dataFrame.sort_values(by='Total Criado Neste Ano', ascending=False)
    exportacaoXLSX(dataFrameOrdenado)
    return escreverPaginaConfluence(dataFrameOrdenado)

def getPagesID():
    baseURL = f'https://{configData["CompanyDomain"]}.atlassian.net/wiki/api/v2/spaces/'
    params = {
        'cql': f"created >= startOfYear()",
        'limit': 250,
        'start': 0
    }
    
    response = requests.get(baseURL, headers=HEADERS, params=params, auth=AUTH)
    data = response.json()
    pageID = [page['id'] for page in data.get('results', [])]
    return pageID

def getUsersID(ID):
    baseURL = f'https://{configData["CompanyDomain"]}.atlassian.net/wiki/api/v2/spaces/{ID}/pages'
    params = {
        'cql': f"created >= startOfYear()",
        'limit': 250,
        'start': 0
    }
    response = requests.get(baseURL, headers=HEADERS, auth=AUTH, params=params)
    data = response.json()
    userID = filterYear(data)
    return userID

def filterYear(data):
    thisYear = dt.datetime.now().year
    usersThisYear = [
        page['authorId']
        for page in data['results']
        if dt.datetime.fromisoformat(page['createdAt'].replace('Z', '+00:00')).year == thisYear
    ]
    return usersThisYear

def getUserName(UserID, cache={}):
    if UserID in cache:
        return cache[UserID]

    baseURL = f'https://{configData["CompanyDomain"]}.atlassian.net/wiki/rest/api/user?accountId={UserID}'
    params = {
        'cql': f"created >= startOfYear()",
        'limit': 250,
        'start': 0
    }
    
    response = requests.get(baseURL, headers=HEADERS, auth=AUTH, params=params)
    data = response.json()
    if 'publicName' in data:
        nomeUsuario = data['publicName']
    else:
        nomeUsuario = 'Usuário Desconhecido'
    cache[UserID] = nomeUsuario
    return nomeUsuario

def contarOcorrencias(lista):
    return Counter(lista)

def exportacaoXLSX(dataFrame):
    dataAtual = datetime.now()
    diaAtual = dataAtual.strftime('%d')
    mesAtual = dataAtual.strftime("%m")
    anoAtual = dataAtual.year
    
    writer = pd.ExcelWriter(f'{configData["PathToFile"]}_{diaAtual}-{mesAtual}-{anoAtual}.xlsx', engine='openpyxl')
    dataFrame.to_excel(writer, sheet_name='Sheet')
    workbook = writer.book
    worksheet = workbook['Sheet']
    padraoDeColunas(worksheet)
    writer._save()

def padraoDeColunas(worksheet):
    worksheet.column_dimensions['A'].width = 35
    worksheet.column_dimensions['B'].width = 35

def escreverPaginaConfluence(df):
    tableHTML = df.to_html(index=False)
    url = f'https://{configData["CompanyDomain"]}.atlassian.net/wiki/rest/api/content/{configData["PageID"]}' 
    response = requests.get(url, headers=HEADERS, auth=AUTH)
    dados = response.json()
    versaoPage = dados['version']['number']

    updateData = {
        'version': {
            'number': versaoPage + 1
        },
        'title': f'{configData["PageTitle"]}',
        'type': 'page',
        'body': {
            'storage': {
                'value': f'<html><body>{tableHTML}</body></html>',
                'representation': 'storage'
            }
        }
    }
    requests.put(url, json=updateData, auth=AUTH, headers=HEADERS)

if __name__ == '__main__':
    main()
