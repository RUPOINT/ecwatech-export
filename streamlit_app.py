# streamlit_app.py
# -*- coding: utf-8 -*-

import os, sys, time
import pandas as pd
import streamlit as st

# Гарантируем, что текущая папка видна для импорта локального модуля
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ecwatech_export import (
    load_exhibitors_any,
    save_excel,
    save_csv,
    save_json,
)

st.set_page_config(page_title="EcwaTech Export", layout="wide")
st.title("Экспорт экспонентов → единая таблица (без логотипов)")

uploaded = st.file_uploader(
    "Загрузите целиковый файл (search.json или DOCX с JSON внутри)",
    type=["json", "docx"]
)

col1, col2, col3 = st.columns(3)
with col1:
    preview = st.number_input("Размер предпросмотра", min_value=10, max_value=5000, value=100, step=10)
with col2:
    do_xlsx = st.checkbox("Сохранить Excel", value=True)
with col3:
    do_csv = st.checkbox("Сохранить CSV", value=False)
do_json_out = st.checkbox("Сохранить очищенный JSON", value=False)

if uploaded:
    # Сохраняем аплоад во временный файл внутри контейнера
    tmp_path = f"/tmp/{int(time.time())}_{uploaded.name}"
    with open(tmp_path, "wb") as f:
        f.write(uploaded.read())

    try:
        rows, n = load_exhibitors_any(tmp_path)
        st.success(f"Найдено записей: {n}")

        df = pd.DataFrame(rows)
        st.dataframe(df.head(int(preview)))

        # Кнопки скачивания
        ts = int(time.time())

        if do_xlsx:
            fn_xlsx = f"Экватэк_таблица_{ts}.xlsx"
            save_excel(rows, fn_xlsx)
            with open(fn_xlsx, "rb") as f:
                st.download_button(
                    "⬇️ Скачать Excel",
                    f,
                    file_name=fn_xlsx,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        if do_csv:
            fn_csv = f"Экватэк_таблица_{ts}.csv"
            save_csv(rows, fn_csv)
            with open(fn_csv, "rb") as f:
                st.download_button(
                    "⬇️ Скачать CSV",
                    f,
                    file_name=fn_csv,
                    mime="text/csv",
                )

        if do_json_out:
            fn_json = f"Экватэк_таблица_{ts}.json"
            save_json(rows, fn_json)
            with open(fn_json, "rb") as f:
                st.download_button(
                    "⬇️ Скачать JSON",
                    f,
                    file_name=fn_json,
                    mime="application/json",
                )

    except Exception as e:
        st.error(f"Ошибка: {e}")
        st.exception(e)
else:
    st.info("Загрузите файл `search.json` (или .docx с JSON внутри), чтобы получить таблицу и выгрузки.")
