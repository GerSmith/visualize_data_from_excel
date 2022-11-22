#! python
# -*- coding: utf-8 -*-

"""Визуализация данных (выборок) из файлов Excel."""

import os
from os.path import isfile, join

import pandas as pd  # pip install pandas xlrd
from plotly.subplots import make_subplots
import plotly.graph_objects as go  # pip install plotly==5.11.0

# Для записи визуализации в графические файлы надо сделать:
# pip install -U kaleido

# Максимальная шкала приборов (нужно для графиков)
MAX_SENSOR_SCALE = 6
# Директория где хранятся данные для анализа
XLS_PATH = ".\\xls"

# Получаем список файлов из директории "xls"
xls_files = [
    join(XLS_PATH, f)
    for f in os.listdir(XLS_PATH)
    if isfile(join(XLS_PATH, f))
]

# Создаём папку где будем хранить картинки
if not os.path.exists("images"):
    os.mkdir("images")

# Пробегаем по файлам из директории
for data_file in xls_files:
    # читаем датафрейм и собираем заголовки
    df = pd.read_excel(data_file)
    # Формируем "фигуру" с графиками для отрисовки
    fig = make_subplots(
        rows=2,
        cols=1,
        vertical_spacing=0.08,
        subplot_titles=("Котёл-утилизатор №1", "Котёл-утилизатор №2"),
    )
    # Добавялем линии
    fig.add_trace(
        go.Scatter(
            x=df[df.columns[0]], y=df[df.columns[1]], name=df.columns[2]
        ),
        row=1,
        col=1,
    )
    fig.add_trace(
        go.Scatter(
            x=df[df.columns[0]], y=df[df.columns[2]], name=df.columns[2]
        ),
        row=1,
        col=1,
    )
    fig.add_trace(
        go.Scatter(
            x=df[df.columns[0]], y=df[df.columns[3]], name=df.columns[3]
        ),
        row=1,
        col=1,
    )
    fig.add_trace(
        go.Scatter(
            x=df[df.columns[0]], y=df[df.columns[5]], name=df.columns[5]
        ),
        row=2,
        col=1,
    )
    fig.add_trace(
        go.Scatter(
            x=df[df.columns[0]], y=df[df.columns[6]], name=df.columns[6]
        ),
        row=2,
        col=1,
    )
    fig.add_trace(
        go.Scatter(
            x=df[df.columns[0]], y=df[df.columns[7]], name=df.columns[7]
        ),
        row=2,
        col=1,
    )
    # Вычисляем период расчёта для заголовка (берем из имени файла)
    p = data_file[-26:-5]
    # Общие настройки отображения
    fig.update_layout(
        title=f"P дымовых газов на входе в котёл-утилизатор за {p}",
        showlegend=True,
    )
    # Настройки осей (нужны т.к. у нас мульти-чарт)
    fig.update_yaxes(
        title_text="Давление, МПа", range=[0, MAX_SENSOR_SCALE], row=1, col=1
    )
    fig.update_yaxes(
        title_text="Давление, МПа", range=[0, MAX_SENSOR_SCALE], row=2, col=1
    )
    fig.update_xaxes(title_text="Время", row=2, col=1)
    # Отображем графику
    fig.show()
    # Записываем красоту в файл
    fig.write_image(f"images/{p}.png", width=1920, height=1200, scale=3)
    # Чистим датафрейм
    df = pd.DataFrame(None)
