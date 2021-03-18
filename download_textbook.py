import pandas as pd
import requests
import webbrowser

df = pd.read_excel("2016.xlsx")
for index, row in df.iterrows():
    # loop through the excel list
    file_name = f"{row.loc['Book Title']}_{row.loc['Author']}".replace('/', '-').replace(':', '-')
    url = f"{row.loc['OpenURL']}"
    print(url)
    r = requests.get(url)
    download_url = f"{r.url.replace('book', 'content/pdf').replace('%2F', '/')}.pdf"
    webbrowser.open(download_url, new=2)

df = pd.read_excel("2017.xlsx")
for index, row in df.iterrows():
    # loop through the excel list
    file_name = f"{row.loc['Book Title']}_{row.loc['Author']}".replace('/', '-').replace(':', '-')
    url = f"{row.loc['OpenURL']}"
    print(url)
    r = requests.get(url)
    download_url = f"{r.url.replace('book', 'content/pdf').replace('%2F', '/')}.pdf"
    webbrowser.open(download_url, new=2)

df = pd.read_excel("2018.xlsx")
for index, row in df.iterrows():
    # loop through the excel list
    file_name = f"{row.loc['Book Title']}_{row.loc['Author']}".replace('/', '-').replace(':', '-')
    url = f"{row.loc['OpenURL']}"
    print(url)
    r = requests.get(url)
    download_url = f"{r.url.replace('book', 'content/pdf').replace('%2F', '/')}.pdf"
    webbrowser.open(download_url, new=2)

df = pd.read_excel("2019.xlsx")
for index, row in df.iterrows():
    # loop through the excel list
    file_name = f"{row.loc['Book Title']}_{row.loc['Author']}".replace('/', '-').replace(':', '-')
    url = f"{row.loc['OpenURL']}"
    print(url)
    r = requests.get(url)
    download_url = f"{r.url.replace('book', 'content/pdf').replace('%2F', '/')}.pdf"
    webbrowser.open(download_url, new=2)
