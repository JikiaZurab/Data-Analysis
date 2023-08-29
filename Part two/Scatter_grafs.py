import pandas as pd
import os
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import plotly.graph_objects as go

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# ID Google таблицы
SPREADSHEET_ID = "1DTK-vR7hEb5cXZScVuFfskA19iYEt8g69qfZxnxun2Y"

# Имя таблицы и ячеек с данными 
Data_Range = "MainTable!A1:H230"

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
            flow = InstalledAppFlow.from_client_secrets_file("../Part one/credentials.json", SCOPES)
            credentials = flow.run_local_server(port=0)

        # Создает файл token.json с учетными данными для следующего раза
        with open("token.json","w") as token:
            token.write(credentials.to_json())
    
    try:
        service = build("sheets","v4", credentials=credentials)

        # Вызов Sheets API
        sheets = service.spreadsheets()

        result = sheets.values().get(spreadsheetId=SPREADSHEET_ID, range=Data_Range).execute()
        
        # Создает DataFrame из наших данных
        df = pd.DataFrame(result.get("values", []))

        # Присваивает первую строку как заголовки столбцов 
        df.columns = df.iloc[0]

        # Удаляет первую строку, так как она стала заголовки
        df = df[1:]

        # Преобразование столбца "count" в числовой формат
        df["count"] = pd.to_numeric(df["count"])

        # Цикл работает с каждой area отдельно
        for area, group_df in df.groupby('area'):

            fig = go.Figure()

            # Цикл проходит по уникальным значениям столбца "cluster_name" в group_df
            for cl_name in group_df["cluster_name"].unique():
                
                # Создает DF, содержащий только строки с текущим "cluster_name".
                color_group = group_df[group_df["cluster_name"] == cl_name]

                trace = go.Scatter(
                    y=color_group["keyword"],
                    x=color_group["count"],
                    mode="markers",                              # Режим отображения - точки             
                    marker=dict(
                        size=color_group["count"] / 15,
                        color=color_group["color"],
                        line=dict(                               # Обводка точек
                            width=2,
                            color="grey"
                        )
                    ),
                    name=f"{cl_name}"
                )

                # Добавляет созданный объект к объекту графика
                fig.add_trace(trace)

            # Создает Scatter объект с текст поверх точек
            text_scatter = go.Scatter(
                y = group_df["keyword"],
                x = group_df["count"],
                mode = "text",
                text = group_df["keyword"].apply(lambda x: x if len(x) < 15 else x[:5] + "\\" + x[-5:]),
                textposition='top center',
                showlegend = False
            )

            # Добавляет текст поверх точек
            fig.add_trace(text_scatter)

            # Изменяем параметры таблицы
            fig.update_layout(
                title = "Область - " + area,
                xaxis = dict(showgrid=False, showticklabels=False),     # Выключает сетку и надпись
                yaxis = dict(showgrid=False, showticklabels=False),     # Выключает сетку и надпись
                plot_bgcolor = "white",                                 # Меняет цвет таблицы на white
                annotations = [                                         # Footer
                    dict(
                        x=0.95,                                         # Координата x для центра по горизонтали
                        y=0.005,                                        # Координата y для подвала (нужно подбирать под свой график)
                        text= "Footer для тестовое задание",            # Текст, который вы хотите добавить
                        showarrow=False,                                # Отключаем стрелку
                        xref="paper",                                   # Ссылка на координату x в процентах (0-1)
                        yref="paper",                                   # Ссылка на координату y в процентах (0-1)
                        font=dict(size=24, color="gray")                # Настройки шрифта и цвета
                    )
                ]
            )

            fig.show()
            
            # Создает директорию, если она не существуе
            if not os.path.exists("images"):
                os.mkdir("images")

            try:
                #Убирает из названия областей \, чтобы не было проблемы с выстраиванием пути 
                area = area.replace("\\","-")
                fig.write_image(f"images\{area}.png", format='png', width=1500, height=1500)
            except TypeError as error:
                print(error)
        

    except HttpError as error:
        print(error)


if __name__ == "__main__":
    main()