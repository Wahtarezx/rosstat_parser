import os
import requests
from bs4 import BeautifulSoup

BASE_URL = "https://rosstat.gov.ru"
START_URL = BASE_URL + "/"
TARGET_PAGE = "/statistics/turizm"

TARGETS = [
    "Оценка туристского потока",
    "месяцы (с 2022 г.)",
    "годы (с 2002 г.)",
]

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/120.0.0.0 Safari/537.36",
    "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
}

def download_rosstat_tables(save_dir="downloads"):
    os.makedirs(save_dir, exist_ok=True)

    session = requests.Session()
    session.headers.update(headers)

    r1 = session.get(START_URL, verify=False, timeout=30)
    r1.raise_for_status()

    page_url = BASE_URL + TARGET_PAGE
    r2 = session.get(page_url, verify=False, timeout=30, headers={"Referer": START_URL})
    r2.raise_for_status()
    r2.encoding = r2.apparent_encoding

    soup = BeautifulSoup(r2.text, "html.parser")
    print("TITLE:", soup.title.text)

    downloaded = {}

    items = soup.select(".document-list__item")
    for item in items:
        title_block = item.select_one(".document-list__item-title")
        link = item.select_one("a[href]")

        if not title_block or not link:
            continue

        title = title_block.get_text(strip=True)
        href = link["href"]

        for target in TARGETS:
            if target in title:
                file_url = BASE_URL + href
                filename = os.path.basename(href)
                filepath = os.path.join(save_dir, filename)

                file_resp = session.get(file_url, verify=False, timeout=30, headers={"Referer": page_url})
                file_resp.raise_for_status()
                with open(filepath, "wb") as f:
                    f.write(file_resp.content)

                downloaded[target] = filepath

    return downloaded
