from __future__ import print_function

import os.path
import pandas as pd

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# ID исходной Google таблицы
SPREADSHEET_ID_BASE = '165sp-lWd1L4qWxggw25DJo_njOCvzdUjAd414NSE8co'

# Имя таблицы и ячеек с данными 
Data_Range_Copy = "tz_data!A1:H230"

def main():
    
    credentials = None

    # Файл token.json хранит информацию пользователя и обновление токена,
    # создается автоматически при первой авторизации
    if os.path.exists('token.json'):
        credentials = Credentials.from_authorized_user_file("token.json", SCOPES)

    # Если нет файла с учетными данными, дает возможность пользователю войти в систему
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("Part one\credentials.json", SCOPES)
            credentials = flow.run_local_server(port=0)

        # Создает файл token.json с учетными данными для следующего раза
        with open("token.json","w") as token:
            token.write(credentials.to_json())
    
    try:
        service = build("sheets","v4", credentials=credentials)
        # Вызов Sheets API
        sheets = service.spreadsheets()

        result = sheets.values().get(spreadsheetId=SPREADSHEET_ID_BASE, range=Data_Range_Copy).execute()

        values = result.get("values", [])
        
        # Создает DataFrame из наших данных
        df = pd.DataFrame(values)

        # Присваивает первую строку как заголовки столбцов 
        df.columns = df.iloc[0]

        # Удаляет первую строку, так как она стала заголовки
        df = df[1:]

        # Убирает ненужную колонки
        df = df[["area", "cluster", "cluster_name", "keyword", "count", "x", "y"]]

        # Удаляет строки, где значения пустые
        df = df.dropna()

        # Удаляет точные дубликаты из DataFrame
        df = df.drop_duplicates()

        # Удаляет дубликаты словосочитаний в одной области
        df = df.drop_duplicates(subset=["area","keyword"])
        # Убирает все строки, где числовые значения не являются таковыми
        df = df[df["count"].apply(lambda x: x.isdigit())]
        df = df[df["cluster"].apply(lambda x: x.isdigit())]

        # Изменяем String значения на Int
        df["count"] = df["count"].astype(int)
        df["cluster"] = df["cluster"].astype(int)
        

        # Функция проверяет вещественность числа. Возвращает True/False
        def is_float(val):
            try:
                float(val)
                return True
            except ValueError:
                return False
        
        df = df[df["x"].apply(is_float)]
        df = df[df["y"].apply(is_float)]

        # Цвета из палитры TableUA
        TableUA_color = ["#9edae5", "#dbdb8d", "#c7c7c7", "#f7b6d2", "#c49c94", "#c5b0d5", "#ff9896", "#98df8a", "#ffbb78", "#aec7e8"]

        #Список хранит все области в том порядке в котором они приведены
        ar_list = df["area"].unique().tolist()

        #Функция присвоения цветов
        def assign_color(row):
            # Переменая используются для смещение по списку TableUA_color для присвоения разных цветов
            # Переменой присваевается значение по следующему правилу 
            # Значение = порядковый номер первого появления области + номер кластера
            temp = ar_list.index(str(row["area"])) + int(row["cluster"])
            # Цикл вычитает из переменной смещения размер списка цветов
            # Это убирает возможность указание за пределы списка цветов и позволяет последовательно проходить по списку 
            while temp >= len(TableUA_color):
                temp -= len(TableUA_color)
            # Возвращает значение из списка цветов с индексом перемменой смещения
            return TableUA_color[temp]

        # Добавляет и заполняет колонку "color"
        df["color"] = df.apply(assign_color, axis=1)

        # Сортировка по area, cluster, cluster_name, count
        df = df.sort_values(by=["area", "cluster", "cluster_name", "count"], ascending=[True,True,True,False])

        # Переменная хранящая финальные данные
        result = []

        # Преобразует DataFrame обратно в список списков
        result = df.values.tolist()
        
        # Добавляет заголовки как первый элемент списка
        result.insert(0,df.columns.tolist())

        # Обновляет индексы DataFrame после всех операций
        df = df.reset_index(drop=True)

        # Создает .txt файл для вывода DataFrame 
        with open("output.txt", "w", encoding="utf-8") as f:
            f.write(df.to_string())

        #ID финальной Google Таблицы
        SPREADSHEET_ID_FINAL = "1DTK-vR7hEb5cXZScVuFfskA19iYEt8g69qfZxnxun2Y"

        # Диапозон для записи данных
        Data_Range_Write = "MainTable!A1:H230"

        # Обновляем ячейки таблицы (Выводим результат)
        sheets.values().update(spreadsheetId=SPREADSHEET_ID_FINAL, range=Data_Range_Write, 
                               valueInputOption="RAW", body={"values":result}).execute()

        # Функция переводит HEX в RGB, так как google API работает с RGB цветами
        def hex_to_rgb(hex_color):
            hex_color = hex_color.lstrip('#')
            return {
                "red": int(hex_color[0:2], 16) / 255,
                "green": int(hex_color[2:4], 16) / 255,
                "blue": int(hex_color[4:6], 16) / 255
            }

        # Список хранит все запросы по смене цвета ячеек
        update_requests = []

        # Цикл групирует набор строк по значению color
        for color, group_df in df.groupby("color"):
            # Вложенный цикл который проходит по наборам и создает уникальные запросы
            for index, row in group_df.iterrows():
                request = {
                    "updateCells": {
                        "range": {
                            "sheetId": 0,
                            "startRowIndex": index + 1,
                            "endRowIndex": index + 2,
                            "startColumnIndex": group_df.columns.get_loc("color"),
                            "endColumnIndex": group_df.columns.get_loc("color") + 1
                        },
                        "rows": [{"values": [{"userEnteredFormat": {"backgroundColor": hex_to_rgb(color)}}]}],
                        "fields": "userEnteredFormat.backgroundColor"
                    }
                }
                # Добавляет в список запрос на смену цветая ячейки 
                update_requests.append(request)
        
        # Запрос на добавление фильтров
        request = {
            "setBasicFilter": {
                "filter": {
                    "range": {
                        "sheetId": 0,
                        "startRowIndex": 0,
                        "endRowIndex": 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": df.shape[1],
                    }
                }
            }
        }

        # Добавляет в список запрос на добавление фильтров
        update_requests.append(request)

        # Запрос на закрепление заголовка
        request = {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": 0,
                    "gridProperties": {
                        "frozenRowCount": 1
                    }
                },
                "fields": "gridProperties.frozenRowCount"
            }
        }

        # Добавляет в список запрос на закрепление заголовка
        update_requests.append(request)

        try:
            # Отправляет запрос на выполнения всех запросов
            service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID_FINAL, body={"requests": update_requests}).execute()
        except HttpError as error:
            print(error)


    except HttpError as error:
        print(error)

# Точка входа (Gate) 
if __name__ == "__main__":
    main()