# streamlit_app.py
# -*- coding: utf-8 -*-

import os, sys, time, io, json
import pandas as pd
import streamlit as st

# чтобы видеть локальные модули
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ecwatech_export import (
    load_exhibitors_any,
    save_excel,
    save_csv,
    save_json,
)

# необязательно, но если хочешь режим по URL:
try:
    from scrape_textilexpo import (
        get_html,
        discover_api_candidates,
        try_fetch_all_from_api,
        normalize,
        parse_cards_from_html,
    )
    HAS_URL_MODE = True
except Exception:
    HAS_URL_MODE = False

st.set_page_config(page_title="Экспорт экспонентов", layout="wide")
st.title("Экспорт экспонентов → единая таблица (без логотипов)")

mode = st.radio("Режим", ["Из файла", "По URL (витрина/экспозиция)"], index=0 if not HAS_URL_MODE else 0,
                help="Из файла — search.json или DOCX c JSON внутри; По URL — страница экспозиции (например TextileExpo).")

col_top1, col_top2, col_top3 = st.columns(3)
with col_top1:
    preview = st.number_input("Размер предпросмотра", min_value=10, max_value=10000, value=100, step=10)
with col_top2:
    do_xlsx = st.checkbox("Сохранить Excel", value=True)
with col_top3:
    do_csv = st.checkbox("Сохранить CSV", value=False)
do_json_out = st.checkbox("Сохранить очищенный JSON", value=False)

def _offer_downloads(rows, base_filename_prefix="Экспорт"):
    ts = int(time.time())
    df = pd.DataFrame(rows)
    st.dataframe(df.head(int(preview)), use_container_width=True)

    if do_xlsx:
        fn_xlsx = f"{base_filename_prefix}_{ts}.xlsx"
        save_excel(rows, fn_xlsx)
        with open(fn_xlsx, "rb") as f:
            st.download_button(
                "⬇️ Скачать Excel", f,
                file_name=fn_xlsx,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    if do_csv:
        fn_csv = f"{base_filename_prefix}_{ts}.csv"
        save_csv(rows, fn_csv)
        with open(fn_csv, "rb") as f:
            st.download_button(
                "⬇️ Скачать CSV", f,
                file_name=fn_csv,
                mime="text/csv",
            )
    if do_json_out:
        fn_json = f"{base_filename_prefix}_{ts}.json"
        save_json(rows, fn_json)
        with open(fn_json, "rb") as f:
            st.download_button(
                "⬇️ Скачать JSON", f,
                file_name=fn_json,
                mime="application/json",
            )

if mode == "Из файла":
    uploaded = st.file_uploader(
        "Загрузите целиковый файл (search.json или DOCX с JSON внутри)",
        type=["json", "docx"]
    )
    if uploaded:
        tmp_path = f"/tmp/{int(time.time())}_{uploaded.name}"
        with open(tmp_path, "wb") as f:
            f.write(uploaded.read())
        try:
            rows, n = load_exhibitors_any(tmp_path)
            st.success(f"Найдено записей: {n}")
            _offer_downloads(rows, base_filename_prefix="Экспоненты")
        except Exception as e:
            st.error(f"Ошибка: {e}")
            st.exception(e)
    else:
        st.info("Загрузите файл `search.json` (или .docx с JSON внутри), чтобы получить таблицу и выгрузки.")

else:
    if not HAS_URL_MODE:
        st.warning("Модуль парсинга по URL не найден. Добавь `scrape_textilexpo.py` и зависимости `beautifulsoup4`, `requests` в requirements.txt.")
    url = st.text_input(
        "URL страницы экспозиции",
        value="https://spring2025.textilexpo.ru/expositions/exposition/5724.html",
        placeholder="https://.../expositions/exposition/XXXX.html"
    )
    cookie = st.text_input(
        "Cookie (если требуется авторизация, можно оставить пустым)",
        value="",
        type="password",
        help="DevTools → Network → любой запрос → Headers → Request Headers → Cookie"
    )
    go = st.button("Выгрузить")

    @st.cache_data(show_spinner=False)
    def _extract_from_url(url: str, cookie: str):
        # cookie через окружение для requests.Session из модуля
        if cookie:
            os.environ["COOKIE"] = cookie
        html = get_html(url)
        api_urls, embedded = discover_api_candidates(url, html)

        # 1) пытаемся найти API, выкачать все
        records = None
        for u in api_urls:
            items = try_fetch_all_from_api(u)
            if items:
                records = items
                break

        # 2) если нет — пытаемся взять встроенный JSON
        if records is None and embedded:
            for blob in embedded:
                for v in blob.values():
                    if isinstance(v, list) and v and isinstance(v[0], dict):
                        records = v
                        break
                if records is not None:
                    break

        # 3) если и этого нет — парсим HTML карточки
        base = f"{'https' if url.startswith('https') else 'http'}://{url.split('/')[2]}"
        if records is None:
            data = parse_cards_from_html(base, html)
        else:
            data = normalize(records, base)
        return data

    if go and url.strip():
        try:
            with st.spinner("Тяну данные со страницы…"):
                data = _extract_from_url(url.strip(), cookie.strip())
            if not data:
                st.warning("Не нашёл ни одной карточки. Проверь URL/куки.")
            else:
                st.success(f"Готово! Найдено записей: {len(data)}")
                _offer_downloads(data, base_filename_prefix="Экспорт_URL")
        except Exception as e:
            st.error(f"Ошибка: {e}")
            st.exception(e)

