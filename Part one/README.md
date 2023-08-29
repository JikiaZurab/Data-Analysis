# Работа с данными
## Подготовка к работе с Google API

Для работы с Google Sheets и создания надстроек необходимо создать [проект Google Cloud](https://developers.google.com/workspace/guides/create-project?hl=ru).

![1](https://github.com/JikiaZurab/Data-Analysis/assets/22364092/69257cf8-6586-4497-b173-9c60141eafb6)


После создания проекта Google Cloud необходимо подключить [Google Sheets API](https://console.cloud.google.com/apis/library/sheets.googleapis.com?authuser=2&hl=ru&project=pygsheets-397019). Для нужного API можно зайти либо в библиотеку API, либо нажав на кнопку Enable Apis and Services, после чего либо в ручную искать нужный элемент, либо воспользоваться поисковой строкой. 

![2](https://github.com/JikiaZurab/Data-Analysis/assets/22364092/730a36d2-0637-4025-80f3-2289859d570e)


Далее необходимо перейти во вкладку OAuth consent screen. На странице выбора User Type выбираем External. В следующем окне необходимо заполнить поля App name, User support email и Developer contact Information (в оба поляможно вписать один email). __ВАЖНО!__ на этапе Scopes обязательно необходимо добавить Google Sheets API нажав на кнопку ADD OR REMOVE SCOPES, который работает и на чтение и на запись (Scope __БЕЗ__ надписи __.readonly__). На этапе Test Users нужно добавить как минимум себя. Тестогово пользователя можно добавить вписав его электронную почту.

![3](https://github.com/JikiaZurab/Data-Analysis/assets/22364092/76bcfdf7-b169-44e9-841b-403a995dfdbf)


Выбираем вкладку Credentials и нажимаем кнопу Create Credentials выбираем __OAuth client ID__. Тип приложение выставляем на Descktop App и задаем имя. Скачиваем __.json файл__ и задаем ему имя __credentials.json__. Это необходимо для удобства при работе непосредственно с Python кодом.

![4](https://github.com/JikiaZurab/Data-Analysis/assets/22364092/4efcb708-d781-474d-b2a5-f3f499a383f6)


## Работа с файлом Main.py

Перед началом работы с кодом нужно установить все библиотеки. Это можно сделать просто вызвав командную строку из каталога Part One и исполнив следующую команду:

```
pip install -r requirements.txt
```

Также необходимо добавить файл __credentials.json__ в каталог Part one.

В самом коде необходимо поменять всего две переменные __SPREADSHEET_ID_FINAL и Data_Range_Write__. Первая переменная хранит в себе ID Google таблицы, в которую будут записаны готовые данные. Таблицу Вы должны создать сами заранее. ID таблицы можно скопировать из его ссылки, оно начинается после "..d/" и заканчивается символом "/". Не забудьте в настройках дать возможно редактировать таблицу всем у кого есть ссылка. Вторая переменная указывает на нужную таблицу и диапfзон ячеек в которые будет выводиться результат. Обязательно укажите название правильной таблицы ее название можно увидеть снизу в Google Sheets, диапазон менять необязательно.

![5](https://github.com/JikiaZurab/Data-Analysis/assets/22364092/b5afb207-aece-4b03-9e9c-98350d85d067)

Можно запускать исходный код.
