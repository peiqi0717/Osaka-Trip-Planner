import io
import re
import json
import math
import hashlib
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
# 1. 網頁基本設定與常數
# ==========================================
st.set_page_config(
    page_title="大阪旅遊規劃共用小幫手",
    page_icon="🌸",
    layout="wide"
)

DEFAULT_COLS = ["網址", "分類", "區域", "自訂名稱", "自動抓取標題", "最近地鐵站", "備註", "新增者", "新增時間"]
USER_COLS = ["暱稱", "密碼雜湊", "建立時間"]
EXPENSE_COLS = ["日期時間", "項目名稱", "幣別", "金額", "付款人", "分攤成員", "備註", "記錄者"]
SHOPPING_COLS = ["擁有者", "商品名稱", "分類", "購買地點", "日幣金額", "幫誰買", "已買到", "備註"]
PACKING_COLS = ["擁有者", "分類", "物品名稱", "數量", "已打包"]

CAT_OPTIONS = ["美食", "購物", "景點與地標", "未分類"]
REGION_OPTIONS = ["難波/心齋橋", "梅田", "天王寺/阿倍野", "環球影城/港區", "京都/神戶/奈良", "其他/未指定"]
SHOP_CAT_OPTIONS = ["藥妝美妝", "伴手禮/零食", "服飾/精品", "動漫/雜貨", "親友代購", "其他"]
PACK_CAT_OPTIONS = ["重要證件/錢包", "3C電子/網路", "衣物穿搭", "盥洗/常備藥品", "隨身小物", "其他"]

DEFAULT_PACKING_TEMPLATE = [
    ("重要證件/錢包", "護照（確認效期 6 個月以上）", "1"),
    ("重要證件/錢包", "日幣現金 ＆ 信用卡（海外回饋高）", "1"),
    ("重要證件/錢包", "Visit Japan Web 截圖 / 機票憑證", "1"),
    ("重要證件/錢包", "ICOCA / 西瓜卡（或手機綁定）", "1"),
    ("3C電子/網路", "手機充電線 ＆ 行動電源（務必隨身登機）", "1"),
    ("3C電子/網路", "日本上網卡 / eSIM", "1"),
    ("衣物穿搭", "換洗衣物 ＆ 貼身衣物", "5"),
    ("衣物穿搭", "好走的運動鞋（大阪每天走 1.5 萬步起跳）", "1"),
    ("盥洗/常備藥品", "個人常備藥（腸胃藥、止痛藥、OK繃）", "1"),
    ("隨身小物", "折疊傘 / 輕便購物袋 / 夾鏈袋", "1"),
]

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

BRANCH_INDICATORS = [
    "店", "館", "大樓", "百貨", "市場", "商店街", "園區", "機場", "站",
    "本社", "本店", "別館", "分店", "號店",
    "難波", "なんば", "心齋橋", "心斎橋", "道頓堀", "梅田", "天王寺", "阿倍野",
    "日本橋", "新世界", "通天閣", "黑門", "京橋", "鶴橋", "中崎町", "北濱",
    "淀屋橋", "本町", "堀江", "美國村", "環球", "天保山", "海遊館", "大阪城",
    "勝尾寺", "箕面", "京都", "神戶", "奈良", "堺", "吹田", "枚方", "高槻", "臨空", "桜塚"
]

# ==========================================
# 2. 連接 Google 試算表 (多分頁永久資料庫)
# ==========================================
@st.cache_resource
def get_spreadsheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds_dict = json.loads(st.secrets["gcp_json"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    return client.open_by_url(st.secrets["sheet_url"])

def get_or_create_ws(title, default_cols):
    sh = get_spreadsheet()
    try:
        ws = sh.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=500, cols=len(default_cols) + 2)
        ws.append_row(default_cols)
    return ws

def load_sheet_df(ws, default_cols, defaults_map=None):
    records = ws.get_all_values()
    if not records:
        ws.append_row(default_cols)
        return pd.DataFrame(columns=default_cols)

    headers = records[0]
    rows = records[1:]
    df = pd.DataFrame(rows, columns=headers)

    added = False
    for col in default_cols:
        if col not in df.columns:
            df[col] = defaults_map.get(col, "") if defaults_map else ""
            added = True

    df = df[default_cols].fillna("")
    if added:
        save_sheet_df(ws, df, default_cols)
    return df

def save_sheet_df(ws, df, default_cols):
    df_clean = df[default_cols].fillna("").astype(str)
    ws.clear()
    ws.update([df_clean.columns.values.tolist()] + df_clean.values.tolist())

def get_main_worksheet():
    return get_spreadsheet().sheet1

def load_data():
    ws = get_main_worksheet()
    return load_sheet_df(ws, DEFAULT_COLS, {"分類": "未分類", "區域": "其他/未指定"})

def save_data(df):
    ws = get_main_worksheet()
    save_sheet_df(ws, df, DEFAULT_COLS)

# ==========================================
# 3. 帳號驗證與分帳演算法工具
# ==========================================
def hash_pin(nickname, pin_str):
    raw = f"osaka_trip_{nickname.strip()}_{pin_str.strip()}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()

def get_tw_now_str():
    tw_tz = timezone(timedelta(hours=8))
    return datetime.datetime.now(tw_tz).strftime("%Y-%m-%d %H:%M:%S")

def calculate_settlements(df_exp, currency):
    """
    針對指定幣別計算每個人：先墊總額、應分攤總額、淨差額，
    並使用貪婪演算法產生「最少轉帳次數」的還款建議清單。
    """
    sub = df_exp[df_exp["幣別"] == currency].copy()
    if sub.empty:
        return pd.DataFrame(), []

    paid_map = {}
    owed_map = {}

    for _, row in sub.iterrows():
        try:
            amt = float(row["金額"])
        except Exception:
            continue
        if amt <= 0:
            continue

        payer = str(row["付款人"]).strip()
        participants = [p.strip() for p in str(row["分攤成員"]).split(",") if p.strip()]
        if not participants:
            participants = [payer]

        paid_map[payer] = paid_map.get(payer, 0.0) + amt
        share = amt / len(participants)
        for p in participants:
            owed_map[p] = owed_map.get(p, 0.0) + share

    all_people = sorted(list(set(list(paid_map.keys()) + list(owed_map.keys()))))
    summary_rows = []
    debtors = []   # 欠錢的人 (net < 0)
    creditors = [] # 該收錢的人 (net > 0)

    for person in all_people:
        paid = paid_map.get(person, 0.0)
        owed = owed_map.get(person, 0.0)
        net = paid - owed
        summary_rows.append({
            "成員": person,
            "實際先墊金額": round(paid, 1),
            "個人應分攤金額": round(owed, 1),
            "淨收支 (+應收 / -應付)": round(net, 1)
        })
        if net < -0.5:
            debtors.append([person, -net])
        elif net > 0.5:
            creditors.append([person, net])

    # 貪婪演算法配對最少轉帳筆數
    debtors.sort(key=lambda x: x[1], reverse=True)
    creditors.sort(key=lambda x: x[1], reverse=True)

    transfers = []
    i, j = 0, 0
    while i < len(debtors) and j < len(creditors):
        d_name, d_amt = debtors[i]
        c_name, c_amt = creditors[j]
        pay_amt = min(d_amt, c_amt)
        if pay_amt >= 1:
            transfers.append((d_name, c_name, round(pay_amt)))

        debtors[i][1] -= pay_amt
        creditors[j][1] -= pay_amt
        if debtors[i][1] < 0.5:
            i += 1
        if creditors[j][1] < 0.5:
            j += 1

    return pd.DataFrame(summary_rows), transfers

# ==========================================
# 4. Google Places API (New) 座標與全鐵道解析
# ==========================================
def is_google_maps_url(url):
    return any(domain in url.lower() for domain in [
        'maps.app.goo.gl', 'goo.gl/maps', 'google.com/maps', 'maps.google.com'
    ])

def is_valid_kansai_coords(lat, lon):
    if not ((33.0 <= lat <= 36.2) and (134.0 <= lon <= 137.0)):
        return False
    if abs(lat - 34.6863) < 0.0012 and abs(lon - 135.5197) < 0.0012:
        return False
    return True

def extract_coords_from_text(text):
    patterns = [
        r'!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)',
        r'@(-?\d+\.\d+),(-?\d+\.\d+)',
        r'[?&](?:q|ll|center)=(-?\d+\.\d+),(-?\d+\.\d+)',
        r'\[null,null,(-?\d+\.\d+),(-?\d+\.\d+)\]',
        r'"latitude"\s*:\s*"?(-?\d+\.\d+)"?\s*,\s*"longitude"\s*:\s*"?(-?\d+\.\d+)"?',
        r'center=(-?\d+\.\d+)%2C(-?\d+\.\d+)'
    ]
    for pat in patterns:
        for m in re.finditer(pat, text):
            lat, lon = float(m.group(1)), float(m.group(2))
            if is_valid_kansai_coords(lat, lon):
                return lat, lon
    return None, None

def resolve_google_maps_url(url):
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

def query_google_places_new(place_name, region="", note=""):
    api_key = st.secrets.get("google_maps_api_key", "").strip()
    if not api_key:
        return None, None, "", "未設定 google_maps_api_key"

    reg_first = region.split("/")[0] if (region and region != "其他/未指定") else ""
    area_prefix = reg_first if reg_first else "大阪"

    clean_note = ""
    if note and not any(k in note for k in ["很多分店", "多家分店", "燒肉店", "咖哩飯"]):
        clean_note = note.strip()[:25]

    text_query = f"{place_name} {area_prefix} {clean_note}".strip()
    if not any(k in text_query for k in ["大阪", "京都", "神戶", "奈良", "Osaka", "Kyoto", "関西", "關西"]):
        text_query = f"大阪 {text_query}"

    try:
        res = requests.post(
            "https://places.googleapis.com/v1/places:searchText",
            headers={
                "Content-Type": "application/json",
                "X-Goog-Api-Key": api_key,
                "X-Goog-FieldMask": "places.displayName,places.location,places.formattedAddress"
            },
            json={
                "textQuery": text_query,
                "languageCode": "zh-TW",
                "regionCode": "JP"
            },
            timeout=6
        )
        data = res.json()
        if "error" in data:
            err_msg = data["error"].get("message", "未知 API 錯誤")
            return None, None, "", err_msg

        if data.get("places"):
            for p in data["places"]:
                loc = p.get("location", {})
                lat = float(loc.get("latitude", 0))
                lon = float(loc.get("longitude", 0))
                if is_valid_kansai_coords(lat, lon):
                    disp_name = p.get("displayName", {}).get("text", "")
                    return lat, lon, disp_name, ""

        return None, None, "", ""
    except Exception as e:
        return None, None, "", str(e)

def format_walking_time(dist_str):
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
    line = raw_line.strip()
    if line.startswith("大阪市高速電気軌道"):
        return line.replace("大阪市高速電気軌道", "Osaka Metro ", 1)
    if line.startswith("大阪メトロ"):
        return line.replace("大阪メトロ", "Osaka Metro ", 1)
    if line.startswith("大阪") and not line.startswith("大阪モノレール"):
        return "Osaka Metro " + line[2:]
    return line

def get_station_details(lat, lon):
    try:
        res = requests.get(
            "https://express.heartrails.com/api/json",
            params={"method": "getStations", "x": lon, "y": lat},
            timeout=5
        )
        stations = res.json().get("response", {}).get("station", [])
        if not stations:
            return "查無鄰近車站", "", []

        nearest = stations[0]
        nearest_name = nearest.get("name", "")
        walk_text = format_walking_time(nearest.get("distance", ""))

        nearby_lines_display = []
        raw_lines_set = []
        for s in stations:
            s_name = s.get("name", "")
            dist_str = str(s.get("distance", ""))
            num_m = re.search(r'(\d+(?:\.\d+)?)', dist_str)
            meters = 9999
            if num_m:
                v = float(num_m.group(1))
                meters = v * 1000 if "km" in dist_str.lower() else v

            if s_name == nearest_name or meters <= 250:
                line_label = clean_line_name(s.get("line", ""))
                if line_label not in raw_lines_set:
                    raw_lines_set.append(line_label)
                entry = line_label if s_name == nearest_name else f"{s_name}站-{line_label}"
                if entry not in nearby_lines_display:
                    nearby_lines_display.append(entry)

        lines_str = " / ".join(nearby_lines_display[:4])
        icon = "🚇" if "Osaka Metro" in lines_str else "🚃"
        disp = f"{icon} {nearest_name}站（{lines_str}，{walk_text}）"
        return disp, f"{nearest_name}站", raw_lines_set
    except Exception:
        return "車站查詢逾時（可稍後補抓）", "", []

def resolve_place_full(url="", custom_name="", region="", note=""):
    final_title = custom_name.strip()
    lat, lon = None, None

    if url:
        if is_google_maps_url(url):
            gmap_title, g_lat, g_lon = resolve_google_maps_url(url)
            if not final_title and gmap_title:
                final_title = gmap_title
            lat, lon = g_lat, g_lon
        else:
            web_title, w_lat, w_lon = scrape_webpage_title_and_coords(url)
            if not final_title and web_title:
                final_title = web_title
            lat, lon = w_lat, w_lon

    if not final_title:
        final_title = "YouTube 推薦影片" if (url and ('youtube.com' in url or 'youtu.be' in url)) else "未知地點 (需手動確認)"

    if lat is None and is_ambiguous_chain(final_title, region, note):
        return final_title, "⏭️ 疑有多家分店（未指定分店或區域，不自動匹配車站）", None, None, "", []

    api_err = ""
    if (lat is None or lon is None) and final_title not in ["未知地點 (需手動確認)", "YouTube 推薦影片"]:
        g_lat, g_lon, official_name, api_err = query_google_places_new(final_title, region, note)
        if g_lat is not None and g_lon is not None:
            lat, lon = g_lat, g_lon
            if not custom_name.strip() and official_name and final_title == "未知地點 (需手動確認)":
                final_title = official_name

    if lat is not None and lon is not None:
        subway_info, st_name, st_lines = get_station_details(lat, lon)
        return final_title, subway_info, lat, lon, st_name, st_lines
    else:
        if api_err:
            subway_info = f"⚠️ Google API 提示：{api_err}"
        else:
            subway_info = "⚠️ 查無此店座標（請確認店名或貼完整電腦版地圖網址）"
        return final_title, subway_info, None, None, "", []

def resolve_free_text_or_url(raw_text):
    raw_b = raw_text.strip()
    url_m = re.search(r'https?://[^\s]+', raw_b)
    if url_m:
        b_url = url_m.group(0)
        b_custom = raw_b.replace(b_url, "").strip()
    else:
        b_url = ""
        b_custom = raw_b

    title, sub, lat, lon, st_name, st_lines = resolve_place_full(url=b_url, custom_name=b_custom)
    clean_title = title if title != "未知地點 (需手動確認)" else (b_custom or "指定地點")
    return {
        "title": clean_title,
        "subway": sub,
        "lat": lat,
        "lon": lon,
        "st_name": st_name,
        "st_lines": st_lines
    }

def resolve_place_and_subway(url, custom_name="", region="", note=""):
    title, subway_info, _, _, _, _ = resolve_place_full(url, custom_name, region, note)
    return title, subway_info

def calc_haversine_km(lat1, lon1, lat2, lon2):
    R = 6371.0
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return R * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))

def estimate_travel_duration(dist_km, is_direct_line=False):
    if dist_km <= 1.2:
        walk_mins = max(1, round((dist_km * 1000 * 1.25) / 80))
        return f"🚶 步行約 {walk_mins} 分鐘"

    buffer_mins = 8 if is_direct_line else 14
    if dist_km <= 15.0:
        ride_mins = round((dist_km * 1.25 / 30.0) * 60)
    else:
        ride_mins = round((dist_km * 1.2 / 55.0) * 60)

    total_mins = max(10, ride_mins + buffer_mins)
    low = max(8, round(total_mins / 5) * 5 - 3)
    high = low + 8
    return f"🚃 搭車約 {low}～{high} 分鐘"

def build_route_comparison_row(orig_obj, dest_obj):
    dist_km = calc_haversine_km(orig_obj["lat"], orig_obj["lon"], dest_obj["lat"], dest_obj["lon"])
    common_lines = list(set(orig_obj["st_lines"]) & set(dest_obj["st_lines"]))
    is_direct = bool(common_lines)
    est_time_str = estimate_travel_duration(dist_km, is_direct_line=is_direct)

    orig_st = orig_obj["st_name"]
    dest_st = dest_obj["st_name"]

    if dist_km <= 1.2:
        transit_hint = "🚶 距離很近，散步即可抵達（免搭電車）"
    elif orig_st and dest_st and orig_st == dest_st:
        transit_hint = f"🚶 兩地皆在【{orig_st}】生活圈"
    elif common_lines:
        transit_hint = f"✨ 可搭【{common_lines[0]}】一線直達免轉乘（{orig_st} ➔ {dest_st}）"
    elif orig_st and dest_st:
        transit_hint = f"🔄 需轉乘：{orig_st} ➔ {dest_st}"
    else:
        transit_hint = "📍 建議點擊右側查看 Google Maps 詳細轉乘"

    gmap_dir_url = (
        f"https://www.google.com/maps/dir/?api=1"
        f"&origin={orig_obj['lat']},{orig_obj['lon']}"
        f"&destination={dest_obj['lat']},{dest_obj['lon']}"
        f"&travelmode=transit"
    )

    return {
        "起點 ➔ 終點": f"{orig_obj['title']} ➔ {dest_obj['title']}",
        "直線距離": f"{dist_km:.2f} km",
        "預估移動時間": est_time_str,
        "電車轉乘建議": transit_hint,
        "終點最近車站": dest_obj["subway"],
        "實際路線導航": gmap_dir_url
    }

# ==========================================
# 5. 產生美化 Excel 報表
# ==========================================
def generate_styled_excel(df_links):
    all_list = []
    food_list, shop_list, spot_list, other_list = [], [], [], []
    sheet_updated = False

    progress_bar = st.progress(0, text="🔍 正在透過 Google API 解析座標與最近車站...")
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
# 6. 頁面狀態與三層式導覽路由
# ==========================================
if "auth_user" not in st.session_state:
    st.session_state["auth_user"] = None
if "workspace_area" not in st.session_state:
    st.session_state["workspace_area"] = None  # None (選擇大廳) / "shared" (共用區) / "private" (私人區)

# 檢查試算表連線
try:
    _ = get_spreadsheet()
except Exception as e:
    st.error(f"❌ 無法連接 Google 試算表，請檢查 Secrets 設定與試算表共用權限。\n錯誤訊息：{e}")
    st.stop()

# ==============================================================================
# 【第一層】登入與註冊大廳（暱稱 ＋ 4 位數密碼）
# ==============================================================================
if not st.session_state["auth_user"]:
    st.title("🌸 大阪旅遊規劃共用小幫手")
    st.markdown("歡迎來到專屬旅遊規劃空間！請先輸入你的**暱稱**與 **4 位數數字密碼**以進入大廳。")

    col_left, col_mid, col_right = st.columns([1, 2, 1])
    with col_mid:
        with st.container(border=True):
            st.subheader("🔐 夥伴身分驗證")
            st.caption("💡 首次使用請直接輸入想要的暱稱與 4 位數密碼，系統會自動幫你建立專屬帳號；下次回來輸入相同資料即可登入！")

            with st.form("login_register_form"):
                input_nick = st.text_input("👤 你的暱稱", placeholder="例如：Pei、小明、小華").strip()
                input_pin = st.text_input("🔢 4 位數數字密碼", max_chars=4, type="password", placeholder="請輸入 4 個數字，例如：1234").strip()
                submit_auth = st.form_submit_button("🚀 登入 / 註冊進入系統", type="primary", use_container_width=True)

                if submit_auth:
                    if not input_nick:
                        st.warning("⚠️ 請輸入你的暱稱喔！")
                    elif not re.fullmatch(r'\d{4}', input_pin):
                        st.warning("⚠️ 密碼必須剛好是「4 位數純數字」（例如：0520、1234）！")
                    else:
                        with st.spinner("正在驗證身分..."):
                            ws_users = get_or_create_ws("使用者帳號", USER_COLS)
                            df_users = load_sheet_df(ws_users, USER_COLS)
                            hashed = hash_pin(input_nick, input_pin)

                            matched = df_users[df_users["暱稱"].str.strip() == input_nick]
                            if not matched.empty:
                                stored_hash = str(matched.iloc[0]["密碼雜湊"]).strip()
                                if stored_hash == hashed:
                                    st.session_state["auth_user"] = input_nick
                                    st.session_state["workspace_area"] = None
                                    st.rerun()
                                else:
                                    st.error("❌ 此暱稱已被註冊，且密碼不正確！請重新確認你的 4 位數密碼。")
                            else:
                                new_u = pd.DataFrame([{
                                    "暱稱": input_nick,
                                    "密碼雜湊": hashed,
                                    "建立時間": get_tw_now_str()
                                }])
                                df_users = pd.concat([df_users, new_u], ignore_index=True)
                                save_sheet_df(ws_users, df_users, USER_COLS)
                                st.session_state["auth_user"] = input_nick
                                st.session_state["workspace_area"] = None
                                st.success(f"🎉 歡迎新夥伴「{input_nick}」！帳號已自動建立。")
                                st.rerun()
    st.stop()

# 已登入的使用者資訊
current_user = st.session_state["auth_user"]

# ==============================================================================
# 【第二層】功能區塊選擇大廳（共用區 vs 私人空間）
# ==============================================================================
if st.session_state["workspace_area"] is None:
    top_c1, top_c2 = st.columns([4, 1])
    with top_c1:
        st.title(f"🌸 嗨，{current_user}！請選擇要進入的空間")
    with top_c2:
        st.write("")
        if st.button("🚪 登出 / 切換帳號", use_container_width=True):
            st.session_state["auth_user"] = None
            st.session_state["workspace_area"] = None
            st.rerun()

    st.write("")
    card_col1, card_col2 = st.columns(2)

    with card_col1:
        with st.container(border=True):
            st.subheader("🌍 大家共用區（團隊協作）")
            st.write("所有同行夥伴共同編輯與查詢的旅遊大本營：")
            st.markdown("""
            * 📋 **共用旅遊清單**：瀏覽所有美食、購物、景點與最近車站
            * 🧭 **路線與距離試算**：飯店/機場雙向路線、任意兩定點對比
            * ➕ **新增 / ✏️ 修改景點**：自動解析 Google 地圖與鐵道轉乘
            * 💰 **共同支出記帳＆分帳計算機**：記錄代墊款項，一鍵精算誰給誰錢
            * 📊 **匯出精美 Excel**：自動按區域上色分類下載
            """)
            st.write("")
            if st.button("🌍 進入【大家共用區】", type="primary", use_container_width=True):
                st.session_state["workspace_area"] = "shared"
                st.rerun()

    with card_col2:
        with st.container(border=True):
            st.subheader(f"🔒 {current_user} 的私人空間")
            st.write("專屬於你個人的私密筆記，其他夥伴登入後不會看到：")
            st.markdown("""
            * 🛍️ **私人購物＆代購清單**：記錄想買的藥妝、伴手禮、親友代購
            * 💴 **自動預算＆台幣換算**：勾選「已買到 ✅」，即時統計花費
            * 🧳 **個人行李打包清單**：一鍵匯入日本自由行必帶清單並勾選確認
            """)
            st.write("")
            st.write("")
            st.write("")
            if st.button(f"🔒 進入【{current_user} 的私人空間】", type="primary", use_container_width=True):
                st.session_state["workspace_area"] = "private"
                st.rerun()

    st.stop()

# ==============================================================================
# 【第三層】頂部導覽返回列（隨時可返回選擇大廳或一鍵切換區塊）
# ==============================================================================
nav_c1, nav_c2, nav_c3, nav_c4 = st.columns([2, 2, 3, 1])
with nav_c1:
    if st.button("⬅️ 返回選擇區（大廳）", use_container_width=True):
        st.session_state["workspace_area"] = None
        st.rerun()
with nav_c2:
    if st.session_state["workspace_area"] == "shared":
        if st.button(f"🔒 切換至【{current_user} 的私人區】", use_container_width=True):
            st.session_state["workspace_area"] = "private"
            st.rerun()
    else:
        if st.button("🌍 切換至【大家共用區】", use_container_width=True):
            st.session_state["workspace_area"] = "shared"
            st.rerun()
with nav_c3:
    area_label = "🌍 大家共用區" if st.session_state["workspace_area"] == "shared" else f"🔒 {current_user} 的私人空間"
    st.markdown(f"<div style='padding-top:8px;'>目前位置：<b>{area_label}</b> ｜ 登入者：<b>{current_user}</b></div>", unsafe_allow_html=True)
with nav_c4:
    if st.button("🚪 登出", use_container_width=True):
        st.session_state["auth_user"] = None
        st.session_state["workspace_area"] = None
        st.rerun()

st.divider()

# ==============================================================================
# 區塊 A：🌍 大家共用區
# ==============================================================================
if st.session_state["workspace_area"] == "shared":
    df_links = load_data()
    has_gmap_key = bool(st.secrets.get("google_maps_api_key", "").strip())
    key_badge = "✅ Google 地圖 API 已連線" if has_gmap_key else "⚠️ 尚未偵測到 Google 地圖 API 金鑰"
    st.caption(f"📊 目前共收錄 **{len(df_links)}** 筆景點與美食 ｜ {key_badge} ｜ [📄 開啟 Google 雲端試算表]({st.secrets['sheet_url']})")

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "📋 共用旅遊清單",
        "🧭 路線與距離試算",
        "➕ 新增連結",
        "✏️ 修改 / 刪除",
        "💰 共同支出＆分帳計算機",
        "📊 匯出精美 Excel"
    ])

    # --- 共用分頁 1：瀏覽共用清單 ---
    with tab1:
        if df_links.empty:
            st.info("目前清單空空如也，快切換到「➕ 新增連結」加入第一個大阪景點吧！")
        else:
            btn_col1, btn_col2 = st.columns(2)
            with btn_col1:
                def needs_subway_lookup(val):
                    s = str(val).strip()
                    return (not s) or s.startswith(("查無", "未取得", "地鐵查詢逾時", "車站查詢逾時", "非地圖連結", "⏭️", "⚠️")) or ("JROsaka" in s)

                retry_mask = df_links["最近地鐵站"].apply(needs_subway_lookup)
                retry_count = int(retry_mask.sum())
                if retry_count > 0:
                    if st.button(f"🚇 修正/補查 {retry_count} 筆資料的最近車站", type="primary", use_container_width=True):
                        prog = st.progress(0, text="🚇 正在透過 Google API 精準反查最近車站...")
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

    # --- 共用分頁 2：路線與距離試算 ---
    with tab2:
        st.subheader("🧭 路線距離與電車轉乘試算")
        calc_mode = st.radio(
            "請選擇試算模式：",
            [
                "🏨 基準定點（飯店/機場）↔ 📋 清單景點（批量雙向比對）",
                "📍 任意兩定點直接對比（例如：比對機場與候選飯店距離）"
            ],
            horizontal=True
        )

        if calc_mode.startswith("🏨"):
            with st.container(border=True):
                st.markdown("**步驟 1️⃣：設定基準定點（例如：飯店、關西機場、難波車站）**")
                base_col1, base_col2 = st.columns([3, 1])
                with base_col1:
                    base_input = st.text_input(
                        "貼上 Google Maps 網址，或直接輸入飯店 / 機場名稱",
                        value=st.session_state.get("base_raw_input", ""),
                        placeholder="例如：https://maps.app.goo.gl/... 或 難波東方酒店 / 關西國際機場"
                    )
                with base_col2:
                    st.write("")
                    st.write("")
                    lock_base_btn = st.button("📍 確認設定定點", type="primary", use_container_width=True)

                if lock_base_btn and base_input.strip():
                    with st.spinner("🔍 正在解析基準定點座標與最近車站..."):
                        parsed = resolve_free_text_or_url(base_input)
                        if parsed["lat"] is not None and parsed["lon"] is not None:
                            st.session_state["base_raw_input"] = base_input.strip()
                            st.session_state["base_info"] = parsed
                        else:
                            st.error(f"❌ 無法解析基準定點位置：{parsed['subway']}")

                if "base_info" in st.session_state:
                    bi = st.session_state["base_info"]
                    st.success(f"🏨 目前基準定點：**{bi['title']}** ｜ 最近車站：{bi['subway']}")

            if df_links.empty:
                st.info("目前清單尚無景點，請先新增資料，或切換上方為「📍 任意兩定點直接對比」！")
            else:
                with st.container(border=True):
                    st.markdown("**步驟 2️⃣：選擇路線方向與要比對的景點**")
                    dir_mode = st.radio(
                        "行進方向：",
                        [
                            "🏨 從【基準定點】出發 ➔ 前往【所選景點】",
                            "🎯 從【所選景點】出發 ➔ 返回【基準定點】"
                        ],
                        horizontal=True
                    )

                    spot_options_map = {}
                    for idx, r in df_links.iterrows():
                        s_name = r["自訂名稱"] if str(r["自訂名稱"]).strip() else (r["自動抓取標題"] if str(r["自動抓取標題"]).strip() else r["網址"][:30])
                        label = f"[{idx}] {s_name}（{r['區域']}｜{r['分類']}）"
                        spot_options_map[label] = idx

                    selected_spot_labels = st.multiselect(
                        "選擇要比對的景點（可單選，也可多選一次比較多個景點距離）：",
                        options=list(spot_options_map.keys()),
                        default=list(spot_options_map.keys())[:1] if spot_options_map else []
                    )

                    calc_route_btn = st.button("🚀 開始試算距離與轉乘路線", type="primary")

                if calc_route_btn:
                    if "base_info" not in st.session_state:
                        st.warning("⚠️ 請先在上方「步驟 1」輸入並確認基準定點（飯店或機場）喔！")
                    elif not selected_spot_labels:
                        st.warning("⚠️ 請至少選擇一個要比對的景點！")
                    else:
                        bi = st.session_state["base_info"]
                        is_from_base = dir_mode.startswith("🏨")
                        results_rows = []

                        prog_r = st.progress(0, text="🧭 正在計算直線距離、預估乘車時間與轉乘線索...")
                        for i, lbl in enumerate(selected_spot_labels):
                            idx = spot_options_map[lbl]
                            row = df_links.loc[idx]
                            s_name = str(row["自訂名稱"]).strip() or str(row["自動抓取標題"]).strip()
                            s_url = str(row["網址"]).strip()
                            s_reg = str(row["區域"]).strip()
                            s_note = str(row["備註"]).strip()

                            _, s_sub, s_lat, s_lon, s_st_name, s_st_lines = resolve_place_full(
                                url=s_url, custom_name=s_name, region=s_reg, note=s_note
                            )

                            if s_lat is not None and s_lon is not None:
                                spot_obj = {
                                    "title": s_name,
                                    "subway": s_sub,
                                    "lat": s_lat,
                                    "lon": s_lon,
                                    "st_name": s_st_name,
                                    "st_lines": s_st_lines
                                }
                                if is_from_base:
                                    results_rows.append(build_route_comparison_row(bi, spot_obj))
                                else:
                                    results_rows.append(build_route_comparison_row(spot_obj, bi))
                            else:
                                results_rows.append({
                                    "起點 ➔ 終點": f"{bi['title']} ↔ {s_name}",
                                    "直線距離": "無法計算",
                                    "預估移動時間": "-",
                                    "電車轉乘建議": f"⚠️ 該景點尚未取得座標（{s_sub}）",
                                    "終點最近車站": s_sub,
                                    "實際路線導航": ""
                                })

                            prog_r.progress((i + 1) / len(selected_spot_labels))
                        prog_r.empty()

                        if results_rows:
                            st.dataframe(
                                pd.DataFrame(results_rows),
                                use_container_width=True,
                                hide_index=True,
                                column_config={
                                    "實際路線導航": st.column_config.LinkColumn(
                                        "Google Maps 轉乘導航",
                                        display_text="🗺️ 查看時刻與轉乘"
                                    )
                                }
                            )

        else:
            with st.container(border=True):
                st.markdown("**📍 輸入任意兩個地點進行距離與轉乘對比（免加入清單即可查詢）**")

                q_cols = st.columns(5)
                quick_presets = [
                    ("✈️ 關西國際機場", "關西國際機場"),
                    ("🚉 難波車站", "難波車站"),
                    ("🚉 大阪梅田車站", "大阪梅田車站"),
                    ("🎢 日本環球影城", "日本環球影城"),
                    ("🦌 京都車站", "京都車站")
                ]
                for idx_q, (btn_lbl, preset_val) in enumerate(quick_presets):
                    with q_cols[idx_q]:
                        if st.button(f"起點設為 {btn_lbl}", use_container_width=True):
                            st.session_state["p2p_point_a"] = preset_val
                            st.rerun()

                col_a, col_swap, col_b = st.columns([5, 1, 5])
                with col_a:
                    point_a_input = st.text_input(
                        "🚩 起點 A（貼 Google Maps 網址或輸入名稱）",
                        value=st.session_state.get("p2p_point_a", "關西國際機場"),
                        placeholder="例如：關西國際機場 或 https://maps.app.goo.gl/..."
                    )
                with col_swap:
                    st.write("")
                    st.write("")
                    if st.button("⇄ 對調", use_container_width=True, help="對調起點 A 與終點 B"):
                        old_a = point_a_input
                        old_b = st.session_state.get("p2p_point_b", "")
                        st.session_state["p2p_point_a"] = old_b
                        st.session_state["p2p_point_b"] = old_a
                        st.rerun()
                with col_b:
                    point_b_input = st.text_input(
                        "🏁 終點 B（貼 Google Maps 網址或輸入名稱）",
                        value=st.session_state.get("p2p_point_b", ""),
                        placeholder="例如：難波東方酒店 / 阪急龍仕柏飯店"
                    )

                st.session_state["p2p_point_a"] = point_a_input
                st.session_state["p2p_point_b"] = point_b_input

                compare_p2p_btn = st.button("🔍 立即比對兩定點距離與轉乘", type="primary")

            if compare_p2p_btn:
                if not point_a_input.strip() or not point_b_input.strip():
                    st.warning("⚠️ 請同時輸入「起點 A」與「終點 B」喔！")
                else:
                    with st.spinner("🔍 正在解析兩定點座標、最近車站與轉乘路線..."):
                        obj_a = resolve_free_text_or_url(point_a_input)
                        obj_b = resolve_free_text_or_url(point_b_input)

                    if obj_a["lat"] is None:
                        st.error(f"❌ 無法解析起點 A（{point_a_input}）：{obj_a['subway']}")
                    elif obj_b["lat"] is None:
                        st.error(f"❌ 無法解析終點 B（{point_b_input}）：{obj_b['subway']}")
                    else:
                        info_c1, info_c2 = st.columns(2)
                        with info_c1:
                            st.info(f"🚩 **起點 A：{obj_a['title']}**\n\n最近車站：{obj_a['subway']}")
                        with info_c2:
                            st.info(f"🏁 **終點 B：{obj_b['title']}**\n\n最近車站：{obj_b['subway']}")

                        p2p_rows = [
                            build_route_comparison_row(obj_a, obj_b),
                            build_route_comparison_row(obj_b, obj_a)
                        ]
                        st.dataframe(
                            pd.DataFrame(p2p_rows),
                            use_container_width=True,
                            hide_index=True,
                            column_config={
                                "實際路線導航": st.column_config.LinkColumn(
                                    "Google Maps 轉乘導航",
                                    display_text="🗺️ 查看時刻與轉乘"
                                )
                            }
                        )

    # --- 共用分頁 3：新增連結 ---
    with tab3:
        with st.form("add_link_form", clear_on_submit=True):
            st.subheader("新增景點 / 餐廳 / 購物連結")
            raw_input = st.text_input("🔗 貼上網址（支援 Google Maps、IG、食記或一般網頁）*", placeholder="https://...")
            c_name = st.text_input("✏️ 自訂中文/日文名稱（選填；系統會透過 Google API 自動搜尋地址與最近車站）", placeholder="例如：一蘭拉麵 道頓堀店")

            c1, c2 = st.columns(2)
            with c1:
                chosen_cat = st.selectbox("🏷️ 主分類", CAT_OPTIONS)
            with c2:
                chosen_region = st.selectbox("🗺️ 區域 / 商圈", REGION_OPTIONS)

            note_text = st.text_input("📝 備註（選填）", placeholder="例如：OPA百貨 五樓、必點牛舌")
            submitted = st.form_submit_button(f"✅ 以【{current_user}】身分加入共用清單", type="primary")

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
                            with st.spinner("🔍 正在透過 Google API 解析地點座標並計算最近車站與步行時間..."):
                                fetched_title, subway_result = resolve_place_and_subway(
                                    extracted_url,
                                    c_name.strip(),
                                    chosen_region,
                                    note_text.strip()
                                )

                            new_row = pd.DataFrame([{
                                "網址": extracted_url,
                                "分類": chosen_cat,
                                "區域": chosen_region,
                                "自訂名稱": c_name.strip(),
                                "自動抓取標題": fetched_title,
                                "最近地鐵站": subway_result,
                                "備註": note_text.strip(),
                                "新增者": current_user,
                                "新增時間": get_tw_now_str()
                            }])
                            df_updated = pd.concat([df_links, new_row], ignore_index=True)
                            save_data(df_updated)
                            st.success(f"✅ 新增成功！「{fetched_title}」｜{subway_result}")
                            st.rerun()

    # --- 共用分頁 4：修改 / 刪除 ---
    with tab4:
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
                    df_links.at[target_idx, "新增者"] = current_user
                    save_data(df_links)
                    st.success("✅ 修改已儲存並同步至雲端！")
                    st.rerun()

            st.divider()
            if st.button("🗑️ 刪除此筆連結", type="secondary"):
                df_links = df_links.drop(target_idx).reset_index(drop=True)
                save_data(df_links)
                st.success("✅ 已成功刪除該筆連結！")
                st.rerun()

    # --- 共用分頁 5：💰 共同支出記帳＆自動分帳計算機 ---
    with tab5:
        st.subheader("💰 共同支出記帳區 ＆ 自動分帳計算機")
        st.caption("記錄誰先代墊了機票、住宿、門票、餐費或交通費，系統會自動分幣別（日幣 JPY / 台幣 TWD）精算每個人該付或該收多少錢！")

        ws_users = get_or_create_ws("使用者帳號", USER_COLS)
        df_users = load_sheet_df(ws_users, USER_COLS)
        registered_members = [u.strip() for u in df_users["暱稱"].tolist() if u.strip()]
        if current_user not in registered_members:
            registered_members.append(current_user)

        ws_exp = get_or_create_ws("共同支出", EXPENSE_COLS)
        df_exp = load_sheet_df(ws_exp, EXPENSE_COLS)

        # 1. 新增一筆共同支出
        with st.container(border=True):
            st.markdown("**➕ 新增一筆共同支出 / 代墊款項**")
            with st.form("add_expense_form", clear_on_submit=True):
                ec1, ec2, ec3 = st.columns([3, 1, 2])
                with ec1:
                    exp_title = st.text_input("支出項目名稱 *", placeholder="例如：燒肉力丸晚餐、環球影城門票、飯店住宿費")
                with ec2:
                    exp_curr = st.selectbox("幣別", ["JPY (日幣)", "TWD (台幣)"])
                with ec3:
                    exp_amt = st.number_input("總金額 *", min_value=0.0, step=100.0, format="%.0f")

                ec4, ec5, ec6 = st.columns([2, 3, 2])
                with ec4:
                    default_payer_idx = registered_members.index(current_user) if current_user in registered_members else 0
                    exp_payer = st.selectbox("誰先付的錢（付款人）*", registered_members, index=default_payer_idx)
                with ec5:
                    exp_split = st.multiselect(
                        "這筆錢由誰平分（預設全選）*",
                        options=registered_members,
                        default=registered_members
                    )
                with ec6:
                    exp_note = st.text_input("備註（選填）", placeholder="例如：含飲料、刷卡匯率0.215")

                submit_exp = st.form_submit_button("✅ 記錄這筆支出", type="primary")
                if submit_exp:
                    if not exp_title.strip():
                        st.warning("⚠️ 請填寫支出項目名稱！")
                    elif exp_amt <= 0:
                        st.warning("⚠️ 金額必須大於 0！")
                    elif not exp_split:
                        st.warning("⚠️ 請至少選擇一位分攤成員！")
                    else:
                        curr_code = "JPY" if "JPY" in exp_curr else "TWD"
                        new_exp = pd.DataFrame([{
                            "日期時間": get_tw_now_str(),
                            "項目名稱": exp_title.strip(),
                            "幣別": curr_code,
                            "金額": str(int(exp_amt)) if exp_amt.is_integer() else str(exp_amt),
                            "付款人": exp_payer,
                            "分攤成員": ", ".join(exp_split),
                            "備註": exp_note.strip(),
                            "記錄者": current_user
                        }])
                        df_exp = pd.concat([df_exp, new_exp], ignore_index=True)
                        save_sheet_df(ws_exp, df_exp, EXPENSE_COLS)
                        st.success(f"✅ 已新增支出：{exp_title.strip()}（{curr_code} {int(exp_amt):,}）")
                        st.rerun()

        # 2. 顯示所有帳目明細與刪除功能
        if df_exp.empty:
            st.info("目前還沒有任何共同支出紀錄，第一筆代墊由你開始記錄吧！")
        else:
            st.markdown("### 📜 所有共同支出明細")
            st.dataframe(df_exp, use_container_width=True, hide_index=True)

            # 刪除錯誤帳目
            with st.expander("🗑️ 刪除錯誤的帳目紀錄"):
                del_exp_opts = [
                    f"[{i}] {r['日期時間']}｜{r['項目名稱']}（{r['幣別']} {r['金額']}，付款：{r['付款人']}）"
                    for i, r in df_exp.iterrows()
                ]
                sel_del_exp = st.selectbox("選擇要刪除的帳目：", del_exp_opts)
                if st.button("確認刪除此筆帳目", type="secondary"):
                    del_i = int(sel_del_exp.split("]")[0].replace("[", ""))
                    df_exp = df_exp.drop(del_i).reset_index(drop=True)
                    save_sheet_df(ws_exp, df_exp, EXPENSE_COLS)
                    st.success("✅ 已刪除該筆帳目！")
                    st.rerun()

            st.divider()

            # 3. 自動分帳結算區（JPY & TWD 分開結算）
            st.markdown("### 🧮 自動分帳結算結果（誰該給誰多少錢）")
            settle_col1, settle_col2 = st.columns(2)

            with settle_col1:
                st.markdown("#### 💴 日幣（JPY）結算")
                jpy_sum_df, jpy_transfers = calculate_settlements(df_exp, "JPY")
                if jpy_sum_df.empty:
                    st.caption("尚無日幣支出紀錄")
                else:
                    st.dataframe(jpy_sum_df, use_container_width=True, hide_index=True)
                    if not jpy_transfers:
                        st.success("🎉 日幣帳目已完全打平，無需互相轉帳！")
                    else:
                        st.markdown("**👉 日幣轉帳建議（最少轉帳次數）：**")
                        for debtor, creditor, amt in jpy_transfers:
                            st.warning(f"💸 **{debtor}** ➔ 轉給 **{creditor}**：**¥ {amt:,} JPY**")

            with settle_col2:
                st.markdown("#### 💵 台幣（TWD）結算")
                twd_sum_df, twd_transfers = calculate_settlements(df_exp, "TWD")
                if twd_sum_df.empty:
                    st.caption("尚無台幣支出紀錄")
                else:
                    st.dataframe(twd_sum_df, use_container_width=True, hide_index=True)
                    if not twd_transfers:
                        st.success("🎉 台幣帳目已完全打平，無需互相轉帳！")
                    else:
                        st.markdown("**👉 台幣轉帳建議（最少轉帳次數）：**")
                        for debtor, creditor, amt in twd_transfers:
                            st.warning(f"💸 **{debtor}** ➔ 轉給 **{creditor}**：**NT$ {amt:,} TWD**")

    # --- 共用分頁 6：匯出精美 Excel ---
    with tab6:
        st.write("點擊下方按鈕，系統會自動補查尚未解析的最近車站（地鐵/電鐵）、依區域上色分類，並產生包含「總表」與各分類分頁的 Excel 檔案。")
        if df_links.empty:
            st.info("目前沒有資料可以匯出喔！")
        else:
            if st.button("🚀 開始統整並產生 Excel", type="primary"):
                excel_data = generate_styled_excel(df_links.copy())
                ts = datetime.datetime.now(timezone(timedelta(hours=8))).strftime("%Y%m%d_%H%M%S")
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

# ==============================================================================
# 區塊 B：🔒 個人專屬私人空間（只有該暱稱登入才看得到）
# ==============================================================================
else:
    st.title(f"🔒 {current_user} 的私人專屬空間")
    st.caption("這裡的所有清單都只綁定你的個人帳號，其他夥伴登入後不會看到喔！")

    ptab1, ptab2 = st.tabs([
        "🛍️ 私人購物 ＆ 代購清單",
        "🧳 個人行李打包清單"
    ])

    # --- 私人分頁 1：🛍️ 私人購物 ＆ 代購清單 ---
    with ptab1:
        ws_shop = get_or_create_ws("私人購物清單", SHOPPING_COLS)
        df_shop_all = load_sheet_df(ws_shop, SHOPPING_COLS)
        my_shop = df_shop_all[df_shop_all["擁有者"].str.strip() == current_user].copy()

        # 匯率設定與預算儀表板
        rate_col, m1, m2, m3 = st.columns([1.2, 1, 1, 1])
        with rate_col:
            jpy_rate = st.number_input("💱 參考匯率 (1 日幣 = ? 台幣)", min_value=0.100, max_value=0.500, value=0.215, step=0.005, format="%.3f")

        total_jpy = 0.0
        bought_jpy = 0.0
        for _, r in my_shop.iterrows():
            try:
                p = float(r["日幣金額"])
            except Exception:
                p = 0.0
            total_jpy += p
            if str(r["已買到"]).strip() in ["True", "true", "TRUE", "1", "是", "✅"]:
                bought_jpy += p

        with m1:
            st.metric("🛒 預計購物總預算", f"¥ {int(total_jpy):,}", f"約 NT$ {int(total_jpy * jpy_rate):,}")
        with m2:
            st.metric("✅ 已買到花費", f"¥ {int(bought_jpy):,}", f"約 NT$ {int(bought_jpy * jpy_rate):,}")
        with m3:
            rem_jpy = max(0.0, total_jpy - bought_jpy)
            st.metric("⏳ 尚未購買預算", f"¥ {int(rem_jpy):,}", f"約 NT$ {int(rem_jpy * jpy_rate):,}")

        # 新增私人購物項目
        with st.container(border=True):
            st.markdown("**➕ 新增想買的商品 / 親友代購**")
            with st.form("add_private_shop_form", clear_on_submit=True):
                sc1, sc2, sc3 = st.columns([3, 2, 2])
                with sc1:
                    s_item = st.text_input("商品名稱 *", placeholder="例如：合利他命 EX Plus、呼吸巧克力、Le Labo 香水")
                with sc2:
                    s_cat = st.selectbox("商品分類", SHOP_CAT_OPTIONS)
                with sc3:
                    s_price = st.number_input("預估單價/總價 (日幣 ¥)", min_value=0, step=100)

                sc4, sc5, sc6 = st.columns([2, 2, 3])
                with sc4:
                    s_loc = st.text_input("預計購買地點", placeholder="例如：心齋橋大國藥妝、關西機場免稅店")
                with sc5:
                    s_for = st.text_input("買給誰", value="自己", placeholder="例如：自己、媽媽、同事小美")
                with sc6:
                    s_note = st.text_input("備註（色號/尺寸/數量）", placeholder="例如：270錠兩罐、要湊免稅")

                if st.form_submit_button("✅ 加入我的購物清單", type="primary"):
                    if not s_item.strip():
                        st.warning("⚠️ 請輸入商品名稱喔！")
                    else:
                        new_s = pd.DataFrame([{
                            "擁有者": current_user,
                            "商品名稱": s_item.strip(),
                            "分類": s_cat,
                            "購買地點": s_loc.strip(),
                            "日幣金額": str(int(s_price)),
                            "幫誰買": s_for.strip() or "自己",
                            "已買到": "FALSE",
                            "備註": s_note.strip()
                        }])
                        df_shop_all = pd.concat([df_shop_all, new_s], ignore_index=True)
                        save_sheet_df(ws_shop, df_shop_all, SHOPPING_COLS)
                        st.success(f"✅ 已加入「{s_item.strip()}」！")
                        st.rerun()

        # 互動式勾選編輯表格
        if my_shop.empty:
            st.info("你的私人購物清單目前是空的，快把想買的藥妝、伴手禮或代購記下來吧！")
        else:
            st.markdown("### 📝 我的購物與代購清單（可直接勾選「已買到」或修改金額後按儲存）")
            edit_shop_df = my_shop.drop(columns=["擁有者"]).copy()
            edit_shop_df["已買到"] = edit_shop_df["已買到"].apply(lambda x: str(x).strip().upper() in ["TRUE", "1", "是", "✅"])
            edit_shop_df["日幣金額"] = pd.to_numeric(edit_shop_df["日幣金額"], errors="coerce").fillna(0).astype(int)

            edited_shop = st.data_editor(
                edit_shop_df,
                use_container_width=True,
                hide_index=True,
                num_rows="dynamic",
                column_config={
                    "已買到": st.column_config.CheckboxColumn("已買到 ✅", default=False),
                    "分類": st.column_config.SelectboxColumn("分類", options=SHOP_CAT_OPTIONS),
                    "日幣金額": st.column_config.NumberColumn("日幣金額 (¥)", min_value=0, step=100, format="¥ %d")
                }
            )

            if st.button("💾 儲存購物清單變更（含打勾與刪除）", type="primary"):
                others_shop = df_shop_all[df_shop_all["擁有者"].str.strip() != current_user].copy()
                saved_my_shop = edited_shop.copy()
                saved_my_shop.insert(0, "擁有者", current_user)
                saved_my_shop["已買到"] = saved_my_shop["已買到"].apply(lambda b: "TRUE" if b else "FALSE")
                saved_my_shop["日幣金額"] = saved_my_shop["日幣金額"].fillna(0).astype(int).astype(str)

                merged_shop = pd.concat([others_shop, saved_my_shop], ignore_index=True)
                save_sheet_df(ws_shop, merged_shop, SHOPPING_COLS)
                st.success("✅ 私人購物清單已同步儲存！")
                st.rerun()

    # --- 私人分頁 2：🧳 個人行李打包清單 ---
    with ptab2:
        ws_pack = get_or_create_ws("個人行李清單", PACKING_COLS)
        df_pack_all = load_sheet_df(ws_pack, PACKING_COLS)
        my_pack = df_pack_all[df_pack_all["擁有者"].str.strip() == current_user].copy()

        # 進度條與一鍵匯入範本
        if my_pack.empty:
            st.info("你的個人行李清單目前是空的！你可以一鍵匯入「日本自由行常用行李範本」，再依個人需求增減：")
            if st.button("✨ 一鍵匯入日本自由行常用行李範本（10 項必備品）", type="primary"):
                tpl_rows = [
                    {
                        "擁有者": current_user,
                        "分類": cat,
                        "物品名稱": item_n,
                        "數量": qty,
                        "已打包": "FALSE"
                    }
                    for cat, item_n, qty in DEFAULT_PACKING_TEMPLATE
                ]
                df_pack_all = pd.concat([df_pack_all, pd.DataFrame(tpl_rows)], ignore_index=True)
                save_sheet_df(ws_pack, df_pack_all, PACKING_COLS)
                st.success("🎉 已匯入 10 項日本旅遊必備物品！")
                st.rerun()
        else:
            total_items = len(my_pack)
            packed_items = sum(1 for v in my_pack["已打包"] if str(v).strip().upper() in ["TRUE", "1", "是", "✅"])
            pct = packed_items / total_items if total_items > 0 else 0
            st.progress(pct, text=f"🧳 行李打包進度：已完成 **{packed_items} / {total_items}** 項（{int(pct * 100)}%）")

        # 新增單項行李
        with st.container(border=True):
            st.markdown("**➕ 新增個人行李物品**")
            with st.form("add_packing_form", clear_on_submit=True):
                pc1, pc2, pc3 = st.columns([2, 3, 1])
                with pc1:
                    p_cat = st.selectbox("行李分類", PACK_CAT_OPTIONS)
                with pc2:
                    p_name = st.text_input("物品名稱 *", placeholder="例如：隱形眼鏡、髮捲、變壓器、摺疊衣架")
                with pc3:
                    p_qty = st.text_input("數量", value="1")

                if st.form_submit_button("✅ 加入行李清單", type="primary"):
                    if not p_name.strip():
                        st.warning("⚠️ 請輸入物品名稱喔！")
                    else:
                        new_p = pd.DataFrame([{
                            "擁有者": current_user,
                            "分類": p_cat,
                            "物品名稱": p_name.strip(),
                            "數量": p_qty.strip() or "1",
                            "已打包": "FALSE"
                        }])
                        df_pack_all = pd.concat([df_pack_all, new_p], ignore_index=True)
                        save_sheet_df(ws_pack, df_pack_all, PACKING_COLS)
                        st.success(f"✅ 已加入「{p_name.strip()}」！")
                        st.rerun()

        if not my_pack.empty:
            st.markdown("### 📋 我的行李檢查表（直接勾選「已打包」或增減列後按儲存）")
            edit_pack_df = my_pack.drop(columns=["擁有者"]).copy()
            edit_pack_df["已打包"] = edit_pack_df["已打包"].apply(lambda x: str(x).strip().upper() in ["TRUE", "1", "是", "✅"])

            edited_pack = st.data_editor(
                edit_pack_df,
                use_container_width=True,
                hide_index=True,
                num_rows="dynamic",
                column_config={
                    "已打包": st.column_config.CheckboxColumn("已打包 ✅", default=False),
                    "分類": st.column_config.SelectboxColumn("分類", options=PACK_CAT_OPTIONS)
                }
            )

            if st.button("💾 儲存行李清單變更", type="primary"):
                others_pack = df_pack_all[df_pack_all["擁有者"].str.strip() != current_user].copy()
                saved_my_pack = edited_pack.copy()
                saved_my_pack.insert(0, "擁有者", current_user)
                saved_my_pack["已打包"] = saved_my_pack["已打包"].apply(lambda b: "TRUE" if b else "FALSE")

                merged_pack = pd.concat([others_pack, saved_my_pack], ignore_index=True)
                save_sheet_df(ws_pack, merged_pack, PACKING_COLS)
                st.success("✅ 個人行李清單已同步儲存！")
                st.rerun()
