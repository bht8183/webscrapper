import requests
from datetime import datetime, timedelta
from excel import insert_data_into_excel


def get_data_from_api(api_url, payload):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36 Edg/127.0.0.0',
        'Content-Type': 'application/json; charset=utf-8',
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Accept-Encoding': 'gzip, deflate, br, zstd',
        'Accept-Language': 'en-US,en;q=0.9',
        'Connection': 'keep-alive',
        'Cookie': 'breadcrumbroot=Home; ARRAffinity=3ef195cf5a67ac460b90f11efad550e2a7ea82cb73c3f5805474d7edee0abe98; ARRAffinitySameSite=3ef195cf5a67ac460b90f11efad550e2a7ea82cb73c3f5805474d7edee0abe98; .AspNetCore.Antiforgery.9fXoN5jHCXs=CfDJ8C54liaoNOdPhfXaWtgUt4b4XLt7UY-BsBdyH5AoiD9fs_7Sf0I_ic73tvwCub2vWDKNxsrxzgC8O_FvNZTtwLumQJ47iPcO1A85Mipqz2Fv2h6CSTWcbNde4Rc4b-Xq0eF47zWIChKlwfYzyyrMoyc',
        'Host': '232app.azurewebsites.net',
        'Origin': 'https://232app.azurewebsites.net',
        'Referer': 'https://232app.azurewebsites.net/',
        'RequestVerificationToken': 'CfDJ8C54liaoNOdPhfXaWtgUt4Y0FqBTKODhouLi7PAJwemT_YdeoJAObCNB-qdIgToV2dAb4QQ1hZh9z7CTXlhCxNRmut6ECLnmGG-D-nxBOiSwAA6Tfq-JtEEhhz64zYMhRcK4nQKpf0OdzTc6WZAh96k',
        'Sec-CH-UA': '"Not)A;Brand";v="99", "Microsoft Edge";v="127", "Chromium";v="127"',
        'Sec-CH-UA-Mobile': '?0',
        'Sec-CH-UA-Platform': '"Windows"',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'same-origin',
        'X-Requested-With': 'XMLHttpRequest'
    }

    try:
        response = requests.post(api_url, headers=headers, json=payload)
        if response.status_code == 200:
            data = response.json()
            return data['data']
        else:
            print(f"Failed to retrieve the data. Status code: {response.status_code}")
            print("Response headers:", response.headers)
            print("Response content:", response.text)
            return []
    except requests.exceptions.RequestException as e:
        print(f"An error occurred: {e}")
        return []
    


def loop_date(start_date: datetime, end_date: datetime, material):
    current_date = start_date
    while current_date <= end_date:

        formatted_date = f"{current_date.month}/{current_date.day}/{current_date.year}"
        print(formatted_date)


        api_url = 'https://232app.azurewebsites.net/?handler=SummaryView'
        payload = {
            'draw': 1,
            'columns': [
                {'data': 0, 'name': 'ID', 'searchable': True, 'orderable': True, 'search': {'value': '', 'regex': False}},
                {'data': 1, 'name': 'Company', 'searchable': True, 'orderable': True, 'search': {'value': '', 'regex': False}},
                {'data': 2, 'name': 'Product', 'searchable': True, 'orderable': True, 'search': {'value': '', 'regex': False}},
                {'data': 3, 'name': 'HTSUSCode', 'searchable': True, 'orderable': True, 'search': {'value': '', 'regex': False}},
                {'data': 4, 'name': 'PublicStatus', 'searchable': True, 'orderable': True, 'search': {'value': '', 'regex': False}},
                {'data': 5, 'name': 'WindowClose', 'searchable': True, 'orderable': True, 'search': {'value': '', 'regex': False}},
                {'data': 6, 'name': 'PublishDate', 'searchable': True, 'orderable': True, 'search': {'value': formatted_date, 'regex': False}},
                {'data': 7, 'name': '', 'searchable': True, 'orderable': False, 'search': {'value': '', 'regex': False}}
            ],
            'order': [{'column': 0, 'dir': 'desc'}],
            'start': 0,
            'length': 1000,
            'search': {'value': material, 'regex': False}
        }

        data = get_data_from_api(api_url, payload)
        print(data)
        insert_data_into_excel(data)

        current_date += timedelta(days=1)


