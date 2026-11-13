import io
import re
import json
import math
import urllib.parse
import datetime
from datetime import timezone, timedelta

import pandas as pd
import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

# ==========================================
# 1. 網頁基本設定
# ==========================================
st.set_page_config(
    page_title="大阪旅遊規劃共用小幫手",
    page_icon="🌸",
    layout="wide"
)

DEFAULT_COLS = ["網址", "分類", "區域", "自訂名稱", "自動抓取標題", "最近地鐵站", "備註", "新增者", "新增時間"]
CAT_OPTIONS = ["美食", "購物", "景點與地標", "未分類"]
REGION_OPTIONS = ["難波/心齋橋", "梅田", "天王寺/阿倍野", "環球影城/港區", "京都/神戶/奈良", "其他/未指定"]

# 常見大阪連鎖店關鍵字
COMMON_CHAINS = [
    "一蘭", "鳥貴族", "唐吉訶德", "唐吉軻德", "驚安殿堂", "激安殿堂",
    "大國藥妝", "松本清", "鶴羽藥妝", "Sundrug", "尚都樂客", "杉藥局",
    "薩莉亞", "松屋", "吉野家", "食其家", "Sukiya", "藏壽司", "壽司郎", "濱壽司",
    "力丸", "國產牛", "串炸達摩", "金龍拉麵", "神座拉麵", "四天王拉麵", "一風堂",
    "利久", "敘敘苑", "麥當勞", "星巴克", "藍瓶", "羅多倫", "客美多", "Komeda",
    "業務超市", "玉出", "LIFE", "蟹道樂", "鶴橋風月", "千房", "551", "蓬萊",
    "老爺爺", "Rikuro", "HARBS", "幸福鬆餅", "CoCo壹番屋",
    "無印良品", "UNIQLO", "GU", "Bic Camera", "Yodobashi", "愛電王", "EDION"
]

# 分店或明確地標/商圈字眼
BRANCH_INDICATORS = [
    "店", "館", "大樓", "百貨", "市場", "商店街", "園區", "機場", "站",
    "本社", "本店", "別館", "分店", "號店",
    "難波", "なんば", "心齋橋", "心斎橋", "道頓堀", "梅田", "天王寺", "阿倍野",
    "日本橋", "新世界", "通天閣", "黑門", "京橋", "鶴橋", "中崎町", "北濱",
    "淀屋橋", "本町", "堀江", "美國村", "環球", "天保山", "海遊館", "大阪城",
    "勝尾寺", "箕面", "京都", "神戶", "奈良", "堺", "吹田", "枚方", "高槻", "臨空", "桜塚"
]

# ==========================================
# 2. 連接 Google 試算表 (永久資料庫)
# ==========================================
@st.cache_resource
def get_worksheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds_dict = json.loads(st.secrets["gcp_json"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    sheet = client.open_by_url(st.secrets["sheet_url"])
    return sheet.sheet1

def load_data():
    ws = get_worksheet()
    records = ws.get_all_values()
    if not records:
        ws.append_row(DEFAULT_COLS)
        return pd.DataFrame(columns=DEFAULT_COLS)

    headers = records[0]
    rows = records[1:]
    df = pd.DataFrame(rows, columns=headers)

    added = False
    for col in DEFAULT_COLS:
        if col not in df.columns:
            if col == "分類":
                df[col] = "未分類"
            elif col == "區域":
                df[col] = "其他/未指定"
            else:
                df[col] = ""
            added = True

    df = df[DEFAULT_COLS].fillna("")
    if added:
        save_data(df)
    return df

def save_data(df):
    ws = get_worksheet()
    df_clean = df[DEFAULT_COLS].fillna("").astype(str)
    ws.clear()
    ws.update([df_clean.columns.values.tolist()] + df_clean.values.tolist())

# ==========================================
# 3. 精準座標解析與全鐵道(地鐵+電鐵+JR)查詢引擎
# ==========================================
def is_google_maps_url(url):
    return any(domain in url.lower() for domain in [
        'maps.app.goo.gl', 'goo.gl/maps', 'google.com/maps', 'maps.google.com'
    ])

def is_valid_kansai_coords(lat, lon):
    """確認座標位於關西圈合理範圍內，並排除大阪府廳預設中心點誤判"""
    if not ((33.0 <= lat <= 36.2) and (134.0 <= lon <= 137.0)):
        return False
    # 排除大阪府廳預設中心點 (34.6863, 135.5197) 避免假天滿橋站
    if abs(lat - 34.6863) < 0.0012 and abs(lon - 135.5197) < 0.0012:
        return False
    return True

def extract_coords_from_text(text):
    """從已解碼的網址或網頁內容提取精準經緯度（優先抓地標座標 !3d!4d）"""
    patterns = [
        r'!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)',
        r'@(-?\d+\.\d+),(-?\d+\.\d+)',
        r'[?&](?:q|ll|center)=(-?\d+\.\d+),(-?\d+\.\d+)',
        r'\[null,null,(-?\d+\.\d+),(-?\d+\.\d+)\]',
        r'"latitude"\s*:\s*"?(-?\d+\.\d+)"?\s*,\s*"longitude"\s*:\s*"?(-?\d+\.\d+)"?'
    ]
    for pat in patterns:
        for m in re.finditer(pat, text):
            lat, lon = float(m.group(1)), float(m.group(2))
            if is_valid_kansai_coords(lat, lon):
                return lat, lon
    return None, None

def resolve_google_maps_url(url):
    """穿透 Google Consent 擋牆並追蹤完整轉跳紀錄，解析 Google 地圖連結的店家名稱與經緯度"""
    title = ""
    lat, lon = None, None
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
            'Accept-Language': 'ja,zh-TW;q=0.9,en;q=0.8'
        }
        cookies = {
            "CONSENT": "PENDING+987",
            "SOCS": "CAISHAgBEhJnd3NfMjAyMzA4MTAtMF9SQzIaAmphIAEaBgiA_LyaBg"
        }

        raw_decoded = urllib.parse.unquote(url)
        lat, lon = extract_coords_from_text(raw_decoded)

        res = requests.get(url, headers=headers, cookies=cookies, timeout=6, allow_redirects=True)
        candidate_urls = [r.url for r in res.history] + [res.url]
        for u in candidate_urls:
            decoded_u = urllib.parse.unquote(u)
            cont_match = re.search(r'[?&]continue=([^&]+)', decoded_u)
            if cont_match:
                decoded_u += " " + urllib.parse.unquote(cont_match.group(1))

            if not title:
                place_match = re.search(r'/place/([^/?&]+)', decoded_u)
                if place_match:
                    title = place_match.group(1).replace('+', ' ')

            if lat is None:
                lat, lon = extract_coords_from_text(decoded_u)

        if lat is None:
            lat, lon = extract_coords_from_text(res.text)

        if not title:
            soup = BeautifulSoup(res.text, 'html.parser')
            if soup.title and soup.title.string:
                t = soup.title.string.strip()
                for suffix in ['- Google Maps', '- Google 地圖', 'Google Maps', 'Google 地圖']:
                    t = t.replace(suffix, '')
                if t.strip() and "Before you continue" not in t:
                    title = t.strip()
    except Exception:
        pass
    return title, lat, lon

def scrape_webpage_title_and_coords(url):
    """從一般網頁（食記、IG、官網）抓取標題與內嵌地圖座標"""
    title = ""
    lat, lon = None, None
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'
        }
        res = requests.get(url, headers=headers, timeout=4, allow_redirects=True)
        soup = BeautifulSoup(res.text, 'html.parser')
        if soup.title and soup.title.string and soup.title.string.strip():
            t = soup.title.string.strip()
            if 'youtube.com' in url or 'youtu.be' in url:
                t = t.replace('- YouTube', '').strip()
            title = t
        lat, lon = extract_coords_from_text(res.text)
    except Exception:
        pass
    return title, lat, lon

def has_specific_location_clue(name, region="", note=""):
    combined = f"{name} {note}".lower()
    if any(k in note for k in ["很多分店", "多家分店", "分店很多", "連鎖"]):
        return False
    if region and region != "其他/未指定":
        return True
    if any(ind.lower() in combined for ind in BRANCH_INDICATORS):
        return True
    return False

def is_ambiguous_chain(name, region="", note=""):
    if any(k in note for k in ["很多分店", "多家分店", "分店很多"]):
        return True
    if has_specific_location_clue(name, region, note):
        return False
    return any(chain.lower() in name.lower() for chain in COMMON_CHAINS)

def extract_jp_address_from_web(query_str):
    """透過 DuckDuckGo HTML 搜尋擷取日本真實地址（大阪府/京都府/兵庫縣/奈良縣 + 丁目/番地）"""
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'
        }
        res = requests.post(
            "https://html.duckduckgo.com/html/",
            data={"q": f"{query_str} 住所 大阪"},
            headers=headers,
            timeout=5
        )
        # 匹配日本完整門牌地址（例如：大阪府大阪市中央区心斎橋筋1-4-29）
        addr_match = re.search(
            r'((?:大阪府|京都府|兵庫県|奈良県)[^\s<>"\'、。]{4,25}[0-9０-９一二三四五六七八九十\-－丁目番地号]+)',
            res.text
        )
        if addr_match:
            return addr_match.group(1)
    except Exception:
        pass
    return None

def search_coordinates_by_clues(place_name, region="", note=""):
    """
    精準座標反查（嚴格過濾行政區中心點，絕不亂給假座標）：
    1. 若有設定 Google Places API 金鑰則優先使用
    2. 透過日本國土地理院 (GSI) 查詢具體地標或反查到的日本門牌地址
    3. 透過 OpenStreetMap Nominatim 查詢（嚴格排除 city/administrative 行政區中心點）
    """
    if not place_name or place_name in ["未知地點 (需手動確認)", "YouTube 推薦影片"]:
        return None, None

    # 引擎 0（選配）：Google Places API
    api_key = st.secrets.get("google_maps_api_key", "").strip()
    if api_key:
        try:
            reg_hint = region.split("/")[0] if (region and region != "其他/未指定") else "大阪"
            text_q = f"{reg_hint} {place_name} {note}".strip()
            res = requests.post(
                "https://places.googleapis.com/v1/places:searchText",
                headers={
                    "Content-Type": "application/json",
                    "X-Goog-Api-Key": api_key,
                    "X-Goog-FieldMask": "places.location"
                },
                json={"textQuery": text_q, "languageCode": "zh-TW"},
                timeout=5
            )
            places = res.json().get("places", [])
            if places and "location" in places[0]:
                lat = float(places[0]["location"]["latitude"])
                lon = float(places[0]["location"]["longitude"])
                if is_valid_kansai_coords(lat, lon):
                    return lat, lon
        except Exception:
            pass

    reg_first = region.split("/")[0] if (region and region != "其他/未指定") else ""
    clean_name = re.sub(r'[（(].*?[）)]', '', place_name).strip()

    # 組合搜尋詞（避免過度拆詞導致抓到整個區公所）
    queries = []
    if reg_first:
        queries.append(f"{clean_name} {reg_first}")
    queries.append(clean_name)

    # 嘗試從網路搜尋擷取該店真實日本門牌地址，再用國土地理院轉成精準座標
    jp_addr = extract_jp_address_from_web(f"{clean_name} {reg_first}".strip())
    if jp_addr:
        queries.insert(0, jp_addr)

    # 引擎 1：日本國土地理院 (GSI) 官方地址/設施 API
    for q in queries:
        try:
            res = requests.get(
                "https://msearch.gsi.go.jp/address-search/AddressSearch",
                params={"q": q},
                timeout=4
            )
            results = res.json()
            if isinstance(results, list) and results:
                for item in results:
                    title_prop = item.get("properties", {}).get("title", "")
                    # 排除只匹配到整個「大阪市」或「大阪府」的寬泛結果
                    if title_prop in ["大阪府", "大阪府大阪市", "京都府", "兵庫県", "奈良県"]:
                        continue
                    coords = item.get("geometry", {}).get("coordinates", [])
                    if len(coords) == 2:
                        lon, lat = float(coords[0]), float(coords[1])
                        if is_valid_kansai_coords(lat, lon):
                            return lat, lon
        except Exception:
            continue

    # 引擎 2：OpenStreetMap Nominatim（嚴格過濾行政區中心點）
    for q in queries:
        try:
            search_q = q if any(k in q for k in ["大阪", "京都", "神戶", "奈良"]) else f"{q} 大阪"
            res = requests.get(
                "https://nominatim.openstreetmap.org/search",
                params={"q": search_q, "format": "json", "limit": 3, "countrycodes": "jp"},
                headers={"User-Agent": "OsakaTripPlannerApp/4.0"},
                timeout=4
            )
            data = res.json()
            if data:
                for item in data:
                    # 關鍵防呆：如果回傳類別是行政區邊界(boundary/administrative)，代表沒找到店面，直接跳過！
                    if item.get("class") == "boundary" or item.get("type") in ["administrative", "city", "prefecture"]:
                        continue
                    lat, lon = float(item["lat"]), float(item["lon"])
                    if is_valid_kansai_coords(lat, lon):
                        return lat, lon
        except Exception:
            continue

    return None, None

def format_walking_time(dist_str):
    """將車站直線距離換算為實際預估步行分鐘數（日本標準 80m/分 × 1.2 街區係數）"""
    try:
        num_match = re.search(r'(\d+(?:\.\d+)?)', str(dist_str))
        if not num_match:
            return f"約 {dist_str}"
        val = float(num_match.group(1))
        meters = val * 1000 if "km" in str(dist_str).lower() else val
        walk_mins = max(1, round((meters * 1.2) / 80))
        return f"步行約 {walk_mins} 分鐘 / {dist_str}"
    except Exception:
        return f"約 {dist_str}"

def clean_line_name(raw_line):
    """美化日本鐵道路線名稱"""
    return (
        raw_line
        .replace("大阪市高速電気軌道", "Osaka Metro ")
        .replace("大阪メトロ", "Osaka Metro ")
        .replace("大阪", "Osaka Metro ")
        .strip()
    )

def query_nearest_stations_all(lat, lon):
    """
    查詢距離最近的車站（包含 Osaka Metro 地鐵、JR、阪急、阪神、南海、近鐵、京阪全部電鐵），
    並自動合併同站多條交會路線，避免捨近求遠。
    """
    try:
        res = requests.get(
            "https://express.heartrails.com/api/json",
            params={"method": "getStations", "x": lon, "y": lat},
            timeout=5
        )
        stations = res.json().get("response", {}).get("station", [])
        if not stations:
            return "查無鄰近車站"

        # 直接取距離最近的第一個車站（不論地鐵或電鐵，絕不捨近求遠走50分鐘）
        nearest = stations[0]
        nearest_name = nearest.get("name", "")
        walk_text = format_walking_time(nearest.get("distance", ""))

        # 收集距離該點 350 公尺內的所有車站與路線（涵蓋梅田/難波等共構與鄰近轉乘站）
        nearby_lines = []
        for s in stations:
            s_name = s.get("name", "")
            dist_str = str(s.get("distance", ""))
            num_m = re.search(r'(\d+(?:\.\d+)?)', dist_str)
            meters = 9999
            if num_m:
                v = float(num_m.group(1))
                meters = v * 1000 if "km" in dist_str.lower() else v

            # 同站名一律收錄所有路線；若在 250m 內的其他同名/共構站也一併列出
            if s_name == nearest_name or meters <= 250:
                line_label = clean_line_name(s.get("line", ""))
                entry = line_label if s_name == nearest_name else f"{s_name}站-{line_label}"
                if entry not in nearby_lines:
                    nearby_lines.append(entry)

        lines_str = " / ".join(nearby_lines[:4])  # 最多顯示 4 條主要路線避免過長
        icon = "🚇" if "Osaka Metro" in lines_str else "🚃"
        return f"{icon} {nearest_name}站（{lines_str}，{walk_text}）"
    except Exception:
        return "車站查詢逾時（可稍後補抓）"

def resolve_place_and_subway(url, custom_name="", region="", note=""):
    final_title = custom_name.strip()
    lat, lon = None, None

    if is_google_maps_url(url):
        gmap_title, g_lat, g_lon = resolve_google_maps_url(url)
        if not final_title:
            final_title = gmap_title if gmap_title else "未知地點 (需手動確認)"
        lat, lon = g_lat, g_lon
    else:
        web_title, w_lat, w_lon = scrape_webpage_title_and_coords(url)
        if not final_title:
            final_title = web_title if web_title else ("YouTube 推薦影片" if ('youtube.com' in url or 'youtu.be' in url) else "未知地點 (需手動確認)")
        lat, lon = w_lat, w_lon

    if lat is None and is_ambiguous_chain(final_title, region, note):
        return final_title, "⏭️ 疑有多家分店（未指定分店或區域，不自動匹配車站）"

    if lat is None or lon is None:
        lat, lon = search_coordinates_by_clues(final_title, region, note)

    if lat is not None and lon is not None:
        subway_info = query_nearest_stations_all(lat, lon)
    else:
        subway_info = "⚠️ 未取得精準座標（建議貼 Google 地圖分享連結，或於修改頁手動輸入車站）"

    return final_title, subway_info

# ==========================================
# 4. 產生美化 Excel 報表 (記憶體串流版)
# ==========================================
def generate_styled_excel(df_links):
    all_list = []
    food_list, shop_list, spot_list, other_list = [], [], [], []
    sheet_updated = False

    progress_bar = st.progress(0, text="🔍 正在解析座標、最近車站(地鐵/電鐵)與步行時間...")
    total_rows = len(df_links)

    for index, row in df_links.iterrows():
        url = str(row["網址"]).strip()
        category = str(row["分類"]).strip()
        region = str(row["區域"]).strip()
        custom_name = str(row["自訂名稱"]).strip()
        note = str(row["備註"]).strip()
        adder = str(row["新增者"]).strip()
        add_time = str(row["新增時間"]).strip()
        cached_title = str(row["自動抓取標題"]).strip()
        subway_info = str(row["最近地鐵站"]).strip()

        need_retry = (not subway_info) or subway_info.startswith(("查無", "未取得", "地鐵查詢逾時", "車站查詢逾時", "非地圖連結", "⚠️"))
        if (not cached_title and not custom_name) or need_retry:
            resolved_title, resolved_subway = resolve_place_and_subway(url, custom_name, region, note)
            final_title = custom_name if custom_name else (cached_title if cached_title else resolved_title)
            if not cached_title:
                df_links.at[index, "自動抓取標題"] = final_title
                sheet_updated = True
            if resolved_subway != subway_info:
                subway_info = resolved_subway
                df_links.at[index, "最近地鐵站"] = subway_info
                sheet_updated = True
        else:
            final_title = custom_name if custom_name else cached_title

        title_search = final_title
        if not any(k in title_search.lower() for k in ["大阪", "osaka"]):
            title_search = f"大阪 {title_search}"

        gmap_query = urllib.parse.quote(title_search)
        gmap_url = f"https://www.google.com/maps/search/?api=1&query={gmap_query}"

        spot_data = {
            "區域/商圈": region,
            "名稱/標題": final_title,
            "最近車站(地鐵/電鐵)": subway_info,
            "網址": url,
            "Google地圖導航": gmap_url,
            "備註": note,
            "推薦人": adder,
            "新增時間": add_time
        }

        total_spot_data = {
            "分類": category,
            "區域/商圈": region,
            "名稱/標題": final_title,
            "最近車站(地鐵/電鐵)": subway_info,
            "備註": note,
            "Google地圖導航": gmap_url,
            "網址": url,
            "推薦人": adder,
            "新增時間": add_time
        }
        all_list.append(total_spot_data)

        if category == "美食":
            food_list.append(spot_data)
        elif category == "購物":
            shop_list.append(spot_data)
        elif category == "景點與地標":
            spot_list.append(spot_data)
        else:
            other_list.append(spot_data)

        progress_bar.progress((index + 1) / total_rows, text=f"🔍 處理中 ({index + 1}/{total_rows})：{final_title[:25]}")

    progress_bar.empty()

    if sheet_updated:
        save_data(df_links)

    def sort_df_by_region(df_temp):
        if df_temp.empty:
            return df_temp
        df_temp["區域_index"] = df_temp["區域/商圈"].apply(
            lambda x: REGION_OPTIONS.index(x) if x in REGION_OPTIONS else len(REGION_OPTIONS)
        )
        df_temp = df_temp.sort_values(by=["區域_index", "名稱/標題"]).drop(columns=["區域_index"])
        df_temp.insert(0, "編號", range(1, len(df_temp) + 1))
        return df_temp

    df_all = sort_df_by_region(pd.DataFrame(all_list))
    df_food = sort_df_by_region(pd.DataFrame(food_list))
    df_shop = sort_df_by_region(pd.DataFrame(shop_list))
    df_spot = sort_df_by_region(pd.DataFrame(spot_list))
    df_other = sort_df_by_region(pd.DataFrame(other_list))

    color_palette = {
        "難波/心齋橋": "FFE4E1",
        "梅田": "E6F2FF",
        "天王寺/阿倍野": "E6FFE6",
        "環球影城/港區": "FFF0F5",
        "京都/神戶/奈良": "FFFFE0",
        "其他/未指定": "F5F5F5"
    }

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        sheets_written = []
        sheet_map = [
            (df_all, "📋 總表"),
            (df_food, "🍔 美食指南"),
            (df_shop, "🛍️ 購物商圈"),
            (df_spot, "🏯 景點地標"),
            (df_other, "📝 未分類/其他"),
        ]
        for df_item, s_name in sheet_map:
            if not df_item.empty:
                df_item.to_excel(writer, sheet_name=s_name, index=False)
                sheets_written.append(s_name)

        workbook = writer.book
        for sheet_name in sheets_written:
            ws = workbook[sheet_name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            region_col_idx = None
            url_cols = []
            for col in range(1, ws.max_column + 1):
                col_name = ws.cell(row=1, column=col).value
                if col_name == "區域/商圈":
                    region_col_idx = col
                if col_name in ["網址", "Google地圖導航"]:
                    url_cols.append((col, col_name))

            header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
            header_font = Font(bold=True, size=11)
            header_align = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 26

            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=1, column=col)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = header_align

            for r in range(2, ws.max_row + 1):
                for c in range(1, ws.max_column + 1):
                    ws.cell(row=r, column=c).alignment = Alignment(wrap_text=True, vertical="center")

                if region_col_idx:
                    region_val = str(ws.cell(row=r, column=region_col_idx).value).strip()
                    fill_color = color_palette.get(region_val, None)
                    if fill_color:
                        my_fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
                        for c in range(1, ws.max_column + 1):
                            ws.cell(row=r, column=c).fill = my_fill

                for c_idx, col_name in url_cols:
                    orig_url = ws.cell(row=r, column=c_idx).value
                    if orig_url and str(orig_url).startswith("http"):
                        disp_text = "🔗 原始連結" if col_name == "網址" else "📍 開啟地圖"
                        ws.cell(row=r, column=c_idx).value = disp_text
                        ws.cell(row=r, column=c_idx).hyperlink = orig_url
                        ws.cell(row=r, column=c_idx).font = Font(color="0563C1", underline="single")

            for col in ws.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col:
                    val_str = str(cell.value or "")
                    char_len = sum(2 if ord(ch) > 127 else 1 for ch in val_str)
                    if char_len > max_len:
                        max_len = char_len
                adjusted_width = min(max(max_len + 4, 12), 50)
                ws.column_dimensions[col_letter].width = adjusted_width

    output.seek(0)
    return output

# ==========================================
# 5. 網頁主介面
# ==========================================
st.title("🌸 大阪旅遊規劃共用小幫手")

col_user, col_refresh = st.columns([3, 1])
with col_user:
    user_id = st.text_input("👤 你的 ID / 暱稱（新增或修改時會自動帶入）", placeholder="例如：小明、小華").strip()
    if not user_id:
        user_id = "匿名夥伴"
with col_refresh:
    st.write("")
    st.write("")
    if st.button("🔄 重新同步雲端清單", use_container_width=True):
        st.rerun()

try:
    df_links = load_data()
except Exception as e:
    st.error(f"❌ 無法連接 Google 試算表，請檢查 Secrets 設定與試算表共用權限。\n錯誤訊息：{e}")
    st.stop()

st.caption(f"📊 目前共收錄 **{len(df_links)}** 筆景點與美食 ｜ [📄 直接開啟 Google 雲端試算表]({st.secrets['sheet_url']})")

tab1, tab2, tab3, tab4 = st.tabs([
    "📋 瀏覽共用清單",
    "➕ 新增連結",
    "✏️ 修改 / 刪除",
    "📊 匯出精美 Excel"
])

# --- 分頁 1：瀏覽共用清單 ---
with tab1:
    if df_links.empty:
        st.info("目前清單空空如也，快切換到「➕ 新增連結」加入第一個大阪景點吧！")
    else:
        btn_col1, btn_col2 = st.columns(2)
        with btn_col1:
            def needs_subway_lookup(val):
                s = str(val).strip()
                return (not s) or s.startswith(("查無", "未取得", "地鐵查詢逾時", "車站查詢逾時", "非地圖連結", "⏭️", "⚠️")) or ("天滿橋" in s)

            retry_mask = df_links["最近地鐵站"].apply(needs_subway_lookup)
            retry_count = int(retry_mask.sum())
            if retry_count > 0:
                if st.button(f"🚇 修正/補查 {retry_count} 筆資料的最近車站", type="primary", use_container_width=True):
                    prog = st.progress(0, text="🚇 正在重新精準反查最近車站（地鐵/電鐵/JR）...")
                    retry_indices = df_links[retry_mask].index.tolist()
                    for i, idx in enumerate(retry_indices):
                        row = df_links.loc[idx]
                        c_name = str(row["自訂名稱"]).strip() or str(row["自動抓取標題"]).strip()
                        resolved_title, resolved_subway = resolve_place_and_subway(
                            str(row["網址"]).strip(),
                            c_name,
                            str(row["區域"]).strip(),
                            str(row["備註"]).strip()
                        )
                        if not str(row["自動抓取標題"]).strip():
                            df_links.at[idx, "自動抓取標題"] = c_name if c_name else resolved_title
                        df_links.at[idx, "最近地鐵站"] = resolved_subway
                        prog.progress((i + 1) / len(retry_indices))
                    prog.empty()
                    save_data(df_links)
                    st.success("✅ 已完成車站與步行時間更新！")
                    st.rerun()

        with btn_col2:
            if st.button("🔄 全部重新查詢最近車站（地鐵+電鐵）", type="secondary", use_container_width=True):
                prog = st.progress(0, text="🚇 正在全部重新解析座標與最近車站...")
                for i, idx in enumerate(df_links.index.tolist()):
                    row = df_links.loc[idx]
                    c_name = str(row["自訂名稱"]).strip() or str(row["自動抓取標題"]).strip()
                    resolved_title, resolved_subway = resolve_place_and_subway(
                        str(row["網址"]).strip(),
                        c_name,
                        str(row["區域"]).strip(),
                        str(row["備註"]).strip()
                    )
                    if not str(row["自動抓取標題"]).strip():
                        df_links.at[idx, "自動抓取標題"] = c_name if c_name else resolved_title
                    df_links.at[idx, "最近地鐵站"] = resolved_subway
                    prog.progress((i + 1) / len(df_links))
                prog.empty()
                save_data(df_links)
                st.success("✅ 全部資料已重新更新完成！")
                st.rerun()

        f_col1, f_col2 = st.columns(2)
        with f_col1:
            sel_cats = st.multiselect("🏷️ 篩選分類", CAT_OPTIONS, default=CAT_OPTIONS)
        with f_col2:
            sel_regions = st.multiselect("🗺️ 篩選區域", REGION_OPTIONS, default=REGION_OPTIONS)

        filtered_df = df_links[
            df_links["分類"].isin(sel_cats) & df_links["區域"].isin(sel_regions)
        ].copy()

        filtered_df["顯示名稱"] = filtered_df.apply(
            lambda r: r["自訂名稱"] if str(r["自訂名稱"]).strip() else (r["自動抓取標題"] if str(r["自動抓取標題"]).strip() else "(尚未解析名稱)"),
            axis=1
        )
        filtered_df["最近車站(地鐵/電鐵)"] = filtered_df["最近地鐵站"].apply(
            lambda x: x if str(x).strip() else "(尚未查詢)"
        )

        display_df = filtered_df[["分類", "區域", "顯示名稱", "最近車站(地鐵/電鐵)", "備註", "網址", "新增者"]]
        st.dataframe(
            display_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "網址": st.column_config.LinkColumn("網址", display_text="🔗 開啟連結")
            }
        )

# --- 分頁 2：新增連結 ---
with tab2:
    with st.form("add_link_form", clear_on_submit=True):
        st.subheader("新增景點 / 餐廳 / 購物連結")
        raw_input = st.text_input("🔗 貼上網址（支援 Google Maps、IG、食記或一般網頁）*", placeholder="https://...")
        c_name = st.text_input("✏️ 自訂中文/日文名稱（選填；系統會自動搜尋地址並計算最近地鐵與電鐵）", placeholder="例如：一蘭拉麵 道頓堀店")

        c1, c2 = st.columns(2)
        with c1:
            chosen_cat = st.selectbox("🏷️ 主分類", CAT_OPTIONS)
        with c2:
            chosen_region = st.selectbox("🗺️ 區域 / 商圈", REGION_OPTIONS)

        note_text = st.text_input("📝 備註（選填）", placeholder="例如：OPA百貨 五樓、必點牛舌")
        submitted = st.form_submit_button("✅ 加入共用清單", type="primary")

        if submitted:
            if not raw_input.strip():
                st.warning("⚠️ 請先輸入網址喔！")
            else:
                url_match = (
                    re.search(r'https?://[^\u4e00-\u9fff\s]+', raw_input)
                    or re.search(r'https?://[^\u3000-\u303F\s]+', raw_input)
                    or re.search(r'https?://[^\n\s]+', raw_input)
                )
                if not url_match:
                    st.error("⚠️ 找不到有效的 http 或 https 網址，請確認後再試！")
                else:
                    extracted_url = url_match.group(0).strip()
                    norm_new = extracted_url.lower().split("?")[0].rstrip("/")

                    dup_row = None
                    for _, r in df_links.iterrows():
                        norm_exist = str(r["網址"]).lower().split("?")[0].rstrip("/")
                        if norm_new == norm_exist:
                            dup_row = r
                            break

                    if dup_row is not None:
                        dup_name = dup_row["自訂名稱"] if str(dup_row["自訂名稱"]).strip() else "(無自訂名稱)"
                        st.warning(f"⚠️ 此網址已存在清單中！（名稱：{dup_name} ｜ 推薦人：{dup_row['新增者']} 於 {dup_row['新增時間']} 新增）")
                    else:
                        with st.spinner("🔍 正在解析地點座標並計算最近車站(地鐵/電鐵)與步行時間..."):
                            fetched_title, subway_result = resolve_place_and_subway(
                                extracted_url,
                                c_name.strip(),
                                chosen_region,
                                note_text.strip()
                            )

                        tw_tz = timezone(timedelta(hours=8))
                        now_time = datetime.datetime.now(tw_tz).strftime("%Y-%m-%d %H:%M:%S")
                        new_row = pd.DataFrame([{
                            "網址": extracted_url,
                            "分類": chosen_cat,
                            "區域": chosen_region,
                            "自訂名稱": c_name.strip(),
                            "自動抓取標題": fetched_title,
                            "最近地鐵站": subway_result,
                            "備註": note_text.strip(),
                            "新增者": user_id,
                            "新增時間": now_time
                        }])
                        df_updated = pd.concat([df_links, new_row], ignore_index=True)
                        save_data(df_updated)
                        st.success(f"✅ 新增成功！「{fetched_title}」｜{subway_result}")
                        st.rerun()

# --- 分頁 3：修改 / 刪除 ---
with tab3:
    if df_links.empty:
        st.info("目前清單沒有任何資料可以修改或刪除。")
    else:
        options_list = []
        for idx, r in df_links.iterrows():
            title_preview = r["自訂名稱"] if str(r["自訂名稱"]).strip() else (r["自動抓取標題"] if str(r["自動抓取標題"]).strip() else r["網址"][:35])
            options_list.append(f"[{idx}] {title_preview} ({r['分類']} | {r['區域']} | 推薦：{r['新增者']})")

        selected_label = st.selectbox("請選擇要修改或刪除的項目：", options_list)
        target_idx = int(selected_label.split("]")[0].replace("[", ""))
        target_row = df_links.iloc[target_idx]

        with st.form("edit_form"):
            edit_url = st.text_input("🔗 網址", value=str(target_row["網址"]))
            edit_name = st.text_input("✏️ 自訂名稱", value=str(target_row["自訂名稱"]))
            ec1, ec2 = st.columns(2)
            with ec1:
                cat_idx = CAT_OPTIONS.index(target_row["分類"]) if target_row["分類"] in CAT_OPTIONS else 3
                edit_cat = st.selectbox("🏷️ 主分類", CAT_OPTIONS, index=cat_idx)
            with ec2:
                reg_idx = REGION_OPTIONS.index(target_row["區域"]) if target_row["區域"] in REGION_OPTIONS else 5
                edit_region = st.selectbox("🗺️ 區域 / 商圈", REGION_OPTIONS, index=reg_idx)

            edit_subway = st.text_input("🚇 最近車站（可手動填寫，或勾選下方重新反查）", value=str(target_row["最近地鐵站"]))
            recheck_subway = st.checkbox("🔄 儲存時重新反查最近車站（地鐵+電鐵）")
            edit_note = st.text_input("📝 備註", value=str(target_row["備註"]))

            save_edit = st.form_submit_button("💾 儲存修改", type="primary")
            if save_edit:
                clean_url = edit_url.strip()
                new_name = edit_name.strip()

                if recheck_subway:
                    resolved_title, resolved_subway = resolve_place_and_subway(
                        clean_url,
                        new_name,
                        edit_region,
                        edit_note.strip()
                    )
                    df_links.at[target_idx, "自動抓取標題"] = new_name if new_name else resolved_title
                    final_subway = resolved_subway
                else:
                    if new_name != str(target_row["自訂名稱"]).strip():
                        df_links.at[target_idx, "自動抓取標題"] = new_name if new_name else str(target_row["自動抓取標題"])
                    final_subway = edit_subway.strip()

                df_links.at[target_idx, "網址"] = clean_url
                df_links.at[target_idx, "自訂名稱"] = new_name
                df_links.at[target_idx, "分類"] = edit_cat
                df_links.at[target_idx, "區域"] = edit_region
                df_links.at[target_idx, "最近地鐵站"] = final_subway
                df_links.at[target_idx, "備註"] = edit_note.strip()
                df_links.at[target_idx, "新增者"] = user_id
                save_data(df_links)
                st.success("✅ 修改已儲存並同步至雲端！")
                st.rerun()

        st.divider()
        if st.button("🗑️ 刪除此筆連結", type="secondary"):
            df_links = df_links.drop(target_idx).reset_index(drop=True)
            save_data(df_links)
            st.success("✅ 已成功刪除該筆連結！")
            st.rerun()

# --- 分頁 4：匯出精美 Excel ---
with tab4:
    st.write("點擊下方按鈕，系統會自動補查尚未解析的最近車站（地鐵/電鐵）、依區域上色分類，並產生包含「總表」與各分類分頁的 Excel 檔案。")
    if df_links.empty:
        st.info("目前沒有資料可以匯出喔！")
    else:
        if st.button("🚀 開始統整並產生 Excel", type="primary"):
            excel_data = generate_styled_excel(df_links.copy())
            tw_tz = timezone(timedelta(hours=8))
            ts = datetime.datetime.now(tw_tz).strftime("%Y%m%d_%H%M%S")
            st.session_state["excel_bytes"] = excel_data
            st.session_state["excel_filename"] = f"大阪旅遊統整表_{ts}.xlsx"
            st.success("✨ 統整大功告成！請點擊下方按鈕下載檔案：")

        if "excel_bytes" in st.session_state:
            st.download_button(
                label=f"⬇️ 下載 {st.session_state['excel_filename']}",
                data=st.session_state["excel_bytes"],
                file_name=st.session_state["excel_filename"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
