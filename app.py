import io
import re
import time
import json
import math
import hashlib
import urllib.parse
import datetime
from datetime import timezone, timedelta

import pandas as pd
import pydeck as pdk
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
ITINERARY_COLS = ["天數", "日期", "時間", "行程名稱", "分類", "最近車站", "網址", "備註", "排程者"]
HOTEL_COLS = ["飯店名稱", "網址", "入住日期", "入住時間", "退房日期", "退房時間", "最近車站", "訂房備註", "新增者"]
FLIGHT_COLS = ["航段類型", "航空與班次", "出發機場航廈", "起飛日期", "起飛時間", "抵達機場航廈", "抵達日期", "抵達時間", "訂位與行李備註", "新增者"]
NOTE_COLS = ["置頂", "標籤", "標題", "內文", "相關連結", "作者", "更新時間"]
META_COLS = ["設定鍵", "設定值"]

CAT_OPTIONS = ["美食", "購物", "景點與地標", "交通/住宿", "未分類"]
REGION_OPTIONS = ["難波/心齋橋", "梅田", "天王寺/阿倍野", "環球影城/港區", "京都/神戶/奈良", "其他/未指定"]
SHOP_CAT_OPTIONS = ["藥妝美妝", "伴手禮/零食", "服飾/精品", "動漫/雜貨", "親友代購", "其他"]
PACK_CAT_OPTIONS = ["重要證件/錢包", "3C電子/網路", "衣物穿搭", "盥洗/常備藥品", "隨身小物", "其他"]
NOTE_TAG_OPTIONS = ["📌 行前攻略", "🚃 交通票券", "🎬 必看影片", "🍜 美食筆記", "🛍️ 購物退稅", "⚠️ 注意事項", "📝 其他筆記"]
FLIGHT_LEG_OPTIONS = ["🛫 去程航班", "🛬 回程航班", "✈️ 轉機 / 國內線"]
WEEKDAY_ZH = ["週一", "週二", "週三", "週四", "週五", "週六", "週日"]

SHARED_TAB_KEYS = [
    "itin",
    "flight_hotel",
    "map",
    "list",
    "route",
    "notes",
    "add",
    "edit",
    "expense",
    "excel"
]

SHARED_TAB_LABELS = {
    "itin": "📅 每日行程 ＆ AI 顧問",
    "flight_hotel": "✈️ 機票 ＆ 住宿整理",
    "map": "🗺️ 互動地圖總覽",
    "list": "📋 共用旅遊清單",
    "route": "🧭 路線距離試算",
    "notes": "📝 共用筆記整理",
    "add": "➕ 新增連結",
    "edit": "✏️ 修改 / 刪除",
    "expense": "💰 共同支出＆分帳",
    "excel": "📊 匯出 Excel"
}

CAT_COLOR_RGBA = {
    "美食": [255, 87, 51, 235],
    "購物": [155, 89, 182, 235],
    "景點與地標": [39, 174, 96, 235],
    "交通/住宿": [41, 128, 185, 235],
    "住宿飯店": [241, 196, 15, 245],
    "未分類": [127, 140, 141, 235]
}

PIN_SVG = """
<svg width="64" height="64" viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg">
  <path d="M32 4C19.85 4 10 13.85 10 26C10 42.5 32 60 32 60C32 60 54 42.5 54 26C54 13.85 44.15 4 32 4Z" fill="white" stroke="#2c3e50" stroke-width="3"/>
  <circle cx="32" cy="25" r="8" fill="#2c3e50"/>
</svg>
"""
PIN_ICON_URL = "data:image/svg+xml;charset=utf-8," + urllib.parse.quote(PIN_SVG.strip())

PIN_ICON_SPEC = {
    "url": PIN_ICON_URL,
    "width": 64,
    "height": 64,
    "anchorY": 60,
    "mask": True
}

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
# 自動智慧分區判定引擎
# ==========================================
def classify_kansai_region(lat=None, lon=None, name="", station_text="", note=""):
    combined_text = f"{name} {station_text} {note}"

    if lat is not None and lon is not None:
        if lat >= 34.85:
            return "京都/神戶/奈良"
        if lon <= 135.33:
            return "京都/神戶/奈良"
        if lon >= 135.68:
            return "京都/神戶/奈良"
        if (34.630 <= lat <= 34.685) and (135.380 <= lon <= 135.468):
            return "環球影城/港區"
        if (34.691 <= lat <= 34.718) and (135.475 <= lon <= 135.522):
            return "梅田"
        if (34.656 <= lat <= 34.684) and (135.485 <= lon <= 135.516):
            return "難波/心齋橋"
        if (34.638 <= lat <= 34.658) and (135.495 <= lon <= 135.528):
            return "天王寺/阿倍野"

    kyoto_kobe_nara_kw = [
        "京都", "嵐山", "伏見", "宇治", "祇園", "河原町", "清水寺", "金閣寺", "錦市場",
        "神戶", "神戸", "三宮", "元町", "有馬", "姬路", "姫路", "明石", "淡路",
        "奈良", "近鐵奈良", "東大寺", "春日大社", "生駒"
    ]
    if any(k in combined_text for k in kyoto_kobe_nara_kw):
        return "京都/神戶/奈良"

    usj_bay_kw = [
        "環球", "USJ", "ユニバーサル", "櫻島", "桜島", "天保山", "海遊館", "大阪港",
        "弁天町", "朝潮橋", "宇宙廣場", "コスモスクエア", "南港", "夢洲", "舞洲", "港區", "港区"
    ]
    if any(k in combined_text for k in usj_bay_kw):
        return "環球影城/港區"

    umeda_kw = [
        "梅田", "大阪站", "大阪駅", "北新地", "東梅田", "西梅田", "中崎町",
        "中津", "福島", "新福島", "大江橋", "渡邊橋", "渡辺橋", "茶屋町", "LUCUA", "GRAND FRONT", "HEP FIVE"
    ]
    if any(k in combined_text for k in umeda_kw):
        return "梅田"

    namba_shinsaibashi_kw = [
        "難波", "なんば", "心齋橋", "心斎橋", "道頓堀", "日本橋", "近鐵日本橋",
        "黑門", "黒門", "堀江", "四橋", "四ツ橋", "長堀橋", "美國村", "アメリカ村",
        "千日前", "宗右衛門町", "法善寺", "南海難波", "JR難波", "大阪難波", "湊町", "櫻川", "桜川", "大國町", "大国町"
    ]
    if any(k in combined_text for k in namba_shinsaibashi_kw):
        return "難波/心齋橋"

    tennoji_kw = [
        "天王寺", "阿倍野", "あべの", "新世界", "通天閣", "動物園前", "恵美須町",
        "新今宮", "寺田町", "四天王寺", "HARUKAS", "ハルカス", "Q's MALL"
    ]
    if any(k in combined_text for k in tennoji_kw):
        return "天王寺/阿倍野"

    return "其他/未指定"

def build_map_marker_layer(df_points, marker_style, is_day_route=False):
    if marker_style.startswith("📍"):
        df_icon = df_points.copy()
        df_icon["icon_data"] = [PIN_ICON_SPEC] * len(df_icon)
        return pdk.Layer(
            "IconLayer",
            data=df_icon,
            get_icon="icon_data",
            get_position="[lon, lat]",
            get_color="color",
            get_size=30 if is_day_route else 26,
            size_scale=1,
            size_min_pixels=20,
            size_max_pixels=36,
            pickable=True,
        )
    else:
        return pdk.Layer(
            "ScatterplotLayer",
            data=df_points,
            get_position="[lon, lat]",
            get_fill_color="color",
            get_line_color=[255, 255, 255, 255],
            line_width_min_pixels=1.5,
            stroked=True,
            get_radius=45 if is_day_route else 35,
            radius_min_pixels=6 if is_day_route else 5,
            radius_max_pixels=12 if is_day_route else 10,
            pickable=True,
        )

# ==========================================
# 2. 即時日幣匯率抓取工具（每 60 分鐘自動更新）
# ==========================================
@st.cache_data(ttl=3600, show_spinner=False)
def fetch_live_jpy_to_twd():
    tw_tz = timezone(timedelta(hours=8))
    now_str = datetime.datetime.now(tw_tz).strftime("%H:%M")

    try:
        res = requests.get("https://open.er-api.com/v6/latest/JPY", timeout=3)
        data = res.json()
        if data.get("result") == "success":
            twd_rate = float(data["rates"]["TWD"])
            if 0.10 <= twd_rate <= 0.50:
                return round(twd_rate, 4), f"🌐 即時市場匯率（{now_str} 更新）", True
    except Exception:
        pass

    try:
        res = requests.get(
            "https://cdn.jsdelivr.net/npm/@fawazahmed0/currency-api@latest/v1/currencies/jpy.json",
            timeout=3
        )
        data = res.json()
        twd_rate = float(data["jpy"]["twd"])
        if 0.10 <= twd_rate <= 0.50:
            return round(twd_rate, 4), f"🌐 即時市場匯率（{now_str} 更新）", True
    except Exception:
        pass

    return 0.2150, "⚠️ 離線預設參考匯率（可手動調整）", False

# ==========================================
# 3. 連接 Google 試算表 (含長效快取與自動重試保護)
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

def _retry_gspread(func, *args, **kwargs):
    for attempt in range(4):
        try:
            return func(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            if attempt < 3:
                time.sleep(1.5 * (attempt + 1))
            else:
                raise e

def get_or_create_ws(title, default_cols):
    def _op():
        sh = get_spreadsheet()
        try:
            ws = sh.worksheet(title)
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=title, rows=500, cols=len(default_cols) + 2)
            ws.append_row(default_cols)
        return ws
    return _retry_gspread(_op)

@st.cache_data(ttl=180, show_spinner=False)
def _cached_read_sheet(sheet_key, default_cols_tuple, defaults_items_tuple):
    default_cols = list(default_cols_tuple)
    defaults_map = dict(defaults_items_tuple)

    def _op():
        sh = get_spreadsheet()
        if sheet_key == "__MAIN__":
            ws = sh.sheet1
        else:
            try:
                ws = sh.worksheet(sheet_key)
            except gspread.WorksheetNotFound:
                ws = sh.add_worksheet(title=sheet_key, rows=500, cols=len(default_cols) + 2)
                ws.append_row(default_cols)

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
                df[col] = defaults_map.get(col, "")
                added = True

        df = df[default_cols].fillna("")
        if added:
            df_clean = df[default_cols].fillna("").astype(str)
            ws.clear()
            ws.update([df_clean.columns.values.tolist()] + df_clean.values.tolist())
        return df

    return _retry_gspread(_op)

def load_sheet_df(ws_or_title, default_cols, defaults_map=None):
    sheet_key = ws_or_title if isinstance(ws_or_title, str) else ws_or_title.title
    defaults_tuple = tuple(defaults_map.items()) if defaults_map else ()
    return _cached_read_sheet(sheet_key, tuple(default_cols), defaults_tuple).copy()

def save_sheet_df(ws_or_title, df, default_cols):
    def _op():
        if isinstance(ws_or_title, str):
            sh = get_spreadsheet()
            ws = sh.sheet1 if ws_or_title == "__MAIN__" else sh.worksheet(ws_or_title)
        else:
            ws = ws_or_title
        df_clean = df[default_cols].fillna("").astype(str)
        ws.clear()
        ws.update([df_clean.columns.values.tolist()] + df_clean.values.tolist())

    _retry_gspread(_op)
    _cached_read_sheet.clear()

def load_data():
    return _cached_read_sheet(
        "__MAIN__",
        tuple(DEFAULT_COLS),
        (("分類", "未分類"), ("區域", "其他/未指定"))
    ).copy()

def save_data(df):
    save_sheet_df("__MAIN__", df, DEFAULT_COLS)

def load_trip_dates():
    df_meta = load_sheet_df("旅遊設定", META_COLS)
    meta_dict = dict(zip(df_meta["設定鍵"], df_meta["設定值"]))
    start_str = meta_dict.get("start_date", "")
    end_str = meta_dict.get("end_date", "")
    try:
        s_date = datetime.date.fromisoformat(start_str)
        e_date = datetime.date.fromisoformat(end_str)
        if e_date >= s_date:
            return s_date, e_date
    except Exception:
        pass
    return None, None

def save_trip_dates(start_date, end_date):
    ws_meta = get_or_create_ws("旅遊設定", META_COLS)
    df_meta = load_sheet_df("旅遊設定", META_COLS)
    df_meta = df_meta[~df_meta["設定鍵"].isin(["start_date", "end_date"])].copy()
    new_rows = pd.DataFrame([
        {"設定鍵": "start_date", "設定值": start_date.isoformat()},
        {"設定鍵": "end_date", "設定值": end_date.isoformat()}
    ])
    df_meta = pd.concat([df_meta, new_rows], ignore_index=True)
    save_sheet_df(ws_meta, df_meta, META_COLS)

def load_user_tab_order(username):
    df_meta = load_sheet_df("旅遊設定", META_COLS)
    meta_dict = dict(zip(df_meta["設定鍵"], df_meta["設定值"]))
    raw_order = meta_dict.get(f"tab_order_{username}", "")
    if raw_order:
        saved_keys = []
        for raw_k in raw_order.split(","):
            k = "flight_hotel" if raw_k.strip() == "hotel" else raw_k.strip()
            if k in SHARED_TAB_KEYS and k not in saved_keys:
                saved_keys.append(k)
        for k in SHARED_TAB_KEYS:
            if k not in saved_keys:
                saved_keys.append(k)
        return saved_keys
    return list(SHARED_TAB_KEYS)

def save_user_tab_order(username, ordered_keys):
    ws_meta = get_or_create_ws("旅遊設定", META_COLS)
    df_meta = load_sheet_df("旅遊設定", META_COLS)
    key_name = f"tab_order_{username}"
    df_meta = df_meta[df_meta["設定鍵"] != key_name].copy()
    new_row = pd.DataFrame([{"設定鍵": key_name, "設定值": ",".join(ordered_keys)}])
    df_meta = pd.concat([df_meta, new_row], ignore_index=True)
    save_sheet_df(ws_meta, df_meta, META_COLS)

def get_hotel_events_for_date(df_hotels, target_date):
    checkouts = []
    checkins = []
    staying = []
    if df_hotels is None or df_hotels.empty:
        return checkouts, checkins, staying

    for _, h in df_hotels.iterrows():
        try:
            ci_d = datetime.date.fromisoformat(str(h["入住日期"]).strip())
            co_d = datetime.date.fromisoformat(str(h["退房日期"]).strip())
        except Exception:
            continue

        if target_date == co_d:
            checkouts.append(h)
        if target_date == ci_d:
            checkins.append(h)
        if ci_d <= target_date < co_d:
            staying.append(h)

    return checkouts, checkins, staying

def get_flight_events_for_date(df_flights, target_date):
    flights_on_day = []
    if df_flights is None or df_flights.empty:
        return flights_on_day

    for _, f in df_flights.iterrows():
        dep_str = str(f["起飛日期"]).strip()
        arr_str = str(f["抵達日期"]).strip()
        if target_date.isoformat() in [dep_str, arr_str]:
            flights_on_day.append(f)
    return flights_on_day

# ==========================================
# 4. 帳號驗證、分帳與 AI 助手工具
# ==========================================
def hash_pin(nickname, pin_str):
    raw = f"osaka_trip_{nickname.strip()}_{pin_str.strip()}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()

def get_tw_now_str():
    tw_tz = timezone(timedelta(hours=8))
    return datetime.datetime.now(tw_tz).strftime("%Y-%m-%d %H:%M:%S")

def _solve_debt_graph(paid_map, owed_map):
    all_people = sorted(list(set(list(paid_map.keys()) + list(owed_map.keys()))))
    summary_rows = []
    debtors = []
    creditors = []

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

def calculate_settlements(df_exp, currency):
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

    return _solve_debt_graph(paid_map, owed_map)

def calculate_combined_twd_settlements(df_exp, jpy_to_twd_rate):
    if df_exp.empty:
        return pd.DataFrame(), []

    paid_map = {}
    owed_map = {}

    for _, row in df_exp.iterrows():
        try:
            amt = float(row["金額"])
        except Exception:
            continue
        if amt <= 0:
            continue

        curr = str(row["幣別"]).strip()
        amt_twd = amt * jpy_to_twd_rate if curr == "JPY" else amt

        payer = str(row["付款人"]).strip()
        participants = [p.strip() for p in str(row["分攤成員"]).split(",") if p.strip()]
        if not participants:
            participants = [payer]

        paid_map[payer] = paid_map.get(payer, 0.0) + amt_twd
        share = amt_twd / len(participants)
        for p in participants:
            owed_map[p] = owed_map.get(p, 0.0) + share

    return _solve_debt_graph(paid_map, owed_map)

def ask_travel_ai_assistant(user_prompt, df_itin, df_links, df_hotels, df_flights, trip_days_info):
    gemini_key = st.secrets.get("gemini_api_key", "").strip()
    if not gemini_key:
        return "⚠️ 尚未設定 `gemini_api_key`！請在 Streamlit Secrets 貼上免費的 Google AI Studio API Key 即可啟用 AI 旅遊小助手。"

    try:
        from google import genai
        client = genai.Client(api_key=gemini_key)

        flight_text = "（尚未登錄航班機票資訊）"
        if df_flights is not None and not df_flights.empty:
            f_lines = []
            for _, f in df_flights.iterrows():
                f_lines.append(
                    f"- [{f['航段類型']}] {f['航空與班次']}｜出發：{f['出發機場航廈']} ({f['起飛日期']} {f['起飛時間']}) ➔ 抵達：{f['抵達機場航廈']} ({f['抵達日期']} {f['抵達時間']})｜備註：{f['訂位與行李備註']}"
                )
            flight_text = "\n".join(f_lines)

        hotel_text = "（尚未登錄預訂飯店）"
        if df_hotels is not None and not df_hotels.empty:
            h_lines = []
            for _, h in df_hotels.iterrows():
                h_lines.append(
                    f"- {h['飯店名稱']}｜入住：{h['入住日期']} {h['入住時間']} ～ 退房：{h['退房日期']} {h['退房時間']}｜最近車站：{h['最近車站']}｜備註：{h['訂房備註']}"
                )
            hotel_text = "\n".join(h_lines)

        itin_text = "（目前尚未排入任何每日行程）"
        if not df_itin.empty:
            lines = []
            for _, r in df_itin.iterrows():
                lines.append(
                    f"- {r['天數']} ({r['日期']}) {r['時間']} | {r['行程名稱']} [{r['分類']}] | 車站：{r['最近車站']} | 備註：{r['備註']}"
                )
            itin_text = "\n".join(lines)

        pool_text = "（目前清單無資料）"
        if not df_links.empty:
            p_lines = []
            for _, r in df_links.iterrows():
                name = str(r["自訂名稱"]).strip() or str(r["自動抓取標題"]).strip()
                p_lines.append(
                    f"- {name}（區域：{r['區域']}｜分類：{r['分類']}｜車站：{r['最近地鐵站']}｜備註：{r['備註']}）"
                )
            pool_text = "\n".join(p_lines)

        system_context = f"""
你是一位專業、貼心且熟悉日本大阪、京都、神戶、奈良自由行的「AI 旅遊規劃小助手」。
請根據以下這組旅客的真實航班起降時間、住宿飯店、行程與收藏清單，用繁體中文給出具體、實用、條理分明的建議（包含首日下飛機後往市區/飯店的交通銜接、末日前往機場的建議出發時間、每天從住宿飯店出發/返回是否順路、換飯店日的行李寄放與移動建議等）。

【旅程天數資訊】
{trip_days_info}

【已登錄的機票航班資訊】
{flight_text}

【已預訂的住宿飯店（含入住與退房時間）】
{hotel_text}

【目前已排入的每日行程表】
{itin_text}

【全團收藏的願望清單（含尚未排入的景點與美食）】
{pool_text}

【使用者提問】
{user_prompt}
"""
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=system_context
        )
        return response.text
    except Exception as e:
        return f"⚠️ AI 小助手連線發生錯誤：{e}"

# ==========================================
# 5. Google Places API (New) 座標與全鐵道解析（含快取）
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

        res = requests.get(url, headers=headers, cookies=cookies, timeout=4, allow_redirects=True)
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
        res = requests.get(url, headers=headers, timeout=3, allow_redirects=True)
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
    area_prefix = reg_first if reg_first else ""

    clean_note = ""
    if note and not any(k in note for k in ["很多分店", "多家分店", "燒肉店", "咖哩飯"]):
        clean_note = note.strip()[:25]

    text_query = f"{place_name} {area_prefix} {clean_note}".strip()
    if not any(k in text_query for k in ["大阪", "京都", "神戶", "神戸", "奈良", "Osaka", "Kyoto", "Kobe", "Nara", "関西", "關西"]):
        text_query = f"関西 {text_query}"

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
            timeout=3.5
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
            timeout=3
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

@st.cache_data(ttl=86400, show_spinner=False)
def cached_resolve_place_full(url="", custom_name="", region="", note=""):
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

def resolve_place_full(url="", custom_name="", region="", note=""):
    return cached_resolve_place_full(url.strip(), custom_name.strip(), region.strip(), note.strip())

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

def resolve_place_subway_and_region(url, custom_name="", note=""):
    title, subway_info, lat, lon, _, _ = resolve_place_full(url, custom_name, "", note)
    auto_reg = classify_kansai_region(lat=lat, lon=lon, name=title, station_text=subway_info, note=note)
    return title, subway_info, auto_reg

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
# 6. 產生美化 Excel 報表
# ==========================================
def generate_styled_excel(df_links, df_itin=None, df_hotels=None, df_flights=None, df_notes=None):
    all_list = []
    food_list, shop_list, spot_list, other_list = [], [], [], []
    sheet_updated = False

    progress_bar = st.progress(0, text="🔍 正在透過 Google API 解析座標、智慧分區與最近車站...")
    total_rows = max(1, len(df_links))

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
        if (not cached_title and not custom_name) or need_retry or (not region) or (region == "其他/未指定"):
            resolved_title, resolved_subway, auto_reg = resolve_place_subway_and_region(url, custom_name or cached_title, note)
            final_title = custom_name if custom_name else (cached_title if cached_title else resolved_title)
            if not cached_title:
                df_links.at[index, "自動抓取標題"] = final_title
                sheet_updated = True
            if resolved_subway != subway_info:
                subway_info = resolved_subway
                df_links.at[index, "最近地鐵站"] = subway_info
                sheet_updated = True
            if auto_reg != "其他/未指定" and auto_reg != region:
                region = auto_reg
                df_links.at[index, "區域"] = region
                sheet_updated = True
        else:
            final_title = custom_name if custom_name else cached_title

        title_search = final_title
        if not any(k in title_search.lower() for k in ["大阪", "京都", "神戶", "奈良", "osaka", "kyoto", "kobe", "nara"]):
            title_search = f"関西 {title_search}"

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
        if df_flights is not None and not df_flights.empty:
            df_flights_sorted = df_flights.sort_values(by=["起飛日期", "起飛時間"]).copy()
            df_flights_sorted.to_excel(writer, sheet_name="✈️ 機票航班表", index=False)
            sheets_written.append("✈️ 機票航班表")

        if df_hotels is not None and not df_hotels.empty:
            df_hotels_sorted = df_hotels.sort_values(by=["入住日期"]).copy()
            df_hotels_sorted.to_excel(writer, sheet_name="🏨 住宿飯店表", index=False)
            sheets_written.append("🏨 住宿飯店表")

        if df_itin is not None and not df_itin.empty:
            df_itin_sorted = df_itin.sort_values(by=["天數", "時間"]).copy()
            df_itin_sorted.to_excel(writer, sheet_name="📅 每日行程表", index=False)
            sheets_written.append("📅 每日行程表")

        if df_notes is not None and not df_notes.empty:
            df_notes_clean = df_notes[["標籤", "標題", "內文", "相關連結", "作者", "更新時間"]].copy()
            df_notes_clean.to_excel(writer, sheet_name="📝 共用筆記整理", index=False)
            sheets_written.append("📝 共用筆記整理")

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
                if col_name in ["網址", "Google地圖導航", "相關連結"]:
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
                        disp_text = "📍 開啟地圖" if col_name == "Google地圖導航" else "🔗 原始連結"
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
# 7. 頁面狀態與三層式導覽路由
# ==========================================
if "auth_user" not in st.session_state:
    st.session_state["auth_user"] = None
if "workspace_area" not in st.session_state:
    st.session_state["workspace_area"] = None

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
                            df_users = load_sheet_df("使用者帳號", USER_COLS)
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
            st.write("所有同行夥伴共同編輯與查詢的旅遊大本營（可自訂功能排列順序）：")
            st.markdown("""
            * 📅 **每日行程排程器 ＆ AI 顧問**：日曆排程、相鄰景點距離自動提示、AI 路線健檢
            * ✈️ **機票 ＆ 住宿整理**：並排憑證看板，整合去回程航班、航廈、訂位代號與住宿資訊
            * 🗺️ **全景點互動地圖總覽**：支援小圓點 / 地圖圖釘切換，檢視景點、飯店與單日動線
            * 📋 **共用旅遊清單**：自動智慧判定商圈分區（難波/梅田/港區/京阪神奈等）與最近車站
            * 🧭 **路線與距離試算**：一鍵套用已訂飯店雙向路線、任意兩定點對比
            * 📝 **共用筆記 ＆ 重點整理**：條理分明的筆記牆，可貼重點攻略、重要影片與置頂公告
            * 💰 **共同支出＆分帳計算機**：記錄代墊款項，即時匯率合併結算誰給誰錢
            * 📊 **匯出精美 Excel**：含機票航班、住宿飯店、每日行程、筆記與各分類上色分頁
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
            * 💱 **即時市場匯率換算**：自動抓取最新日幣匯率，勾選「已買到 ✅」即時統計花費
            * 🧳 **個人行李打包清單**：一鍵匯入日本自由行常用清單並勾選確認
            """)
            st.write("")
            st.write("")
            st.write("")
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
# 區塊 A：🌍 大家共用區（單頁按需載入架構，手機與電腦切換極速順暢）
# ==============================================================================
if st.session_state["workspace_area"] == "shared":
    df_links = load_data()
    df_flights = load_sheet_df("機票航班", FLIGHT_COLS)
    df_hotels = load_sheet_df("住宿飯店", HOTEL_COLS)

    has_gmap_key = bool(st.secrets.get("google_maps_api_key", "").strip())
    key_badge = "✅ Google 地圖 API 已連線" if has_gmap_key else "⚠️ 尚未偵測到 Google 地圖 API 金鑰"
    st.caption(f"📊 共收錄 **{len(df_links)}** 筆景點 ｜ ✈️ **{len(df_flights)}** 筆航班 ｜ 🏨 **{len(df_hotels)}** 間飯店 ｜ {key_badge} ｜ [📄 開啟 Google 試算表]({st.secrets['sheet_url']})")

    user_tab_keys = load_user_tab_order(current_user)

    with st.expander("⚙️ 自訂共用區功能排序（修改數字 1～10 即可把重要的功能排到最前面）"):
        st.caption("💡 直接點擊「顯示順序」欄位修改數字（數字越小越靠前），調整完畢後按下方 **「💾 套用並儲存我的排序」** 即可（設定僅影響你的帳號視角）：")

        order_df = pd.DataFrame([
            {
                "顯示順序": idx + 1,
                "功能名稱": SHARED_TAB_LABELS[k],
                "_key": k
            }
            for idx, k in enumerate(user_tab_keys)
        ])

        edited_order_df = st.data_editor(
            order_df[["顯示順序", "功能名稱"]],
            use_container_width=True,
            hide_index=True,
            disabled=["功能名稱"],
            column_config={
                "顯示順序": st.column_config.NumberColumn(
                    "顯示順序（1 = 最前面）",
                    min_value=1,
                    max_value=len(SHARED_TAB_KEYS),
                    step=1,
                    format="%d"
                )
            },
            key=f"tab_order_editor_{current_user}"
        )

        ord_c1, ord_c2, _ = st.columns([1.8, 1.8, 3.4])
        with ord_c1:
            if st.button("💾 套用並儲存我的排序", type="primary", use_container_width=True):
                merged_ord = order_df.copy()
                merged_ord["顯示順序"] = edited_order_df["顯示順序"].fillna(99).astype(int)
                merged_ord = merged_ord.sort_values(by=["顯示順序"], kind="stable")
                new_keys = merged_ord["_key"].tolist()
                save_user_tab_order(current_user, new_keys)
                st.success("✅ 已儲存你的專屬功能順序！")
                st.rerun()
        with ord_c2:
            if st.button("↩️ 恢復系統預設順序", use_container_width=True):
                save_user_tab_order(current_user, list(SHARED_TAB_KEYS))
                st.success("✅ 已恢復預設順序！")
                st.rerun()

    # 改用單頁即時切換選單（只執行當前選擇的頁面，避免手機背景同時載入 10 頁卡死）
    with st.container(border=True):
        active_page = st.radio(
            "🧭 選擇要開啟的功能頁面：",
            options=user_tab_keys,
            format_func=lambda k: SHARED_TAB_LABELS[k],
            horizontal=True,
            key="shared_active_page_nav"
        )

    # 按需載入旅程日期與行程資料
    saved_start, saved_end = (None, None)
    df_itin = pd.DataFrame(columns=ITINERARY_COLS)
    if active_page in ["itin", "flight_hotel", "map", "excel"]:
        saved_start, saved_end = load_trip_dates()
        df_itin = load_sheet_df("每日行程", ITINERARY_COLS)

    # ==========================================================
    # 頁面：📅 每日行程排程器 ＆ AI 旅遊顧問
    # ==========================================================
    if active_page == "itin":
        st.subheader("📅 直覺式每日行程排程器 ＆ 🤖 AI 路線顧問")

        with st.container(border=True):
            st.markdown("**步驟 1️⃣：在日曆選擇旅程的「出發日」與「回程日」**")
            d_col1, d_col2, d_col3 = st.columns([2, 2, 1])
            default_s = saved_start if saved_start else datetime.date.today() + timedelta(days=14)
            default_e = saved_end if saved_end else default_s + timedelta(days=4)

            with d_col1:
                pick_start = st.date_input("🛫 出發日期", value=default_s)
            with d_col2:
                pick_end = st.date_input("🛬 回程日期", value=max(default_e, pick_start))
            with d_col3:
                st.write("")
                st.write("")
                if st.button("💾 儲存旅程日期", type="primary", use_container_width=True):
                    if pick_end < pick_start:
                        st.error("❌ 回程日期不能早於出發日期喔！")
                    else:
                        save_trip_dates(pick_start, pick_end)
                        st.success("✅ 已同步儲存全團旅程日期！")
                        st.rerun()

            active_start = saved_start if saved_start else pick_start
            active_end = saved_end if saved_end else pick_end
            total_days = (active_end - active_start).days + 1
            st.caption(f"🗓️ 目前設定：**{active_start.strftime('%Y/%m/%d')} ({WEEKDAY_ZH[active_start.weekday()]})** ～ **{active_end.strftime('%Y/%m/%d')} ({WEEKDAY_ZH[active_end.weekday()]})**，共 **{total_days} 天 {max(0, total_days - 1)} 夜**")

        day_options = []
        day_date_map = {}
        day_date_obj_map = {}
        for d_idx in range(total_days):
            cur_d = active_start + timedelta(days=d_idx)
            w_str = WEEKDAY_ZH[cur_d.weekday()]
            day_tag = f"Day {d_idx + 1}"
            full_label = f"{day_tag}（{cur_d.strftime('%m/%d')} {w_str}）"
            day_options.append(full_label)
            day_date_map[full_label] = (day_tag, f"{cur_d.strftime('%Y-%m-%d')} ({w_str})")
            day_date_obj_map[day_tag] = cur_d

        with st.container(border=True):
            st.markdown("**步驟 2️⃣：將景點或活動排入行程**")
            opt_c1, opt_c2 = st.columns([3, 2])
            with opt_c1:
                add_source = st.radio(
                    "選擇加入方式：",
                    ["📋 從「共用旅遊清單」挑選景點", "✏️ 手動輸入自訂行程（如：搭飛機、逛街、自由時間）"],
                    horizontal=True
                )
            with opt_c2:
                only_unscheduled = st.checkbox("✨ 只顯示「尚未排進任何一天」的景點", value=True)

            scheduled_names_set = set(df_itin["行程名稱"].str.strip().tolist()) if not df_itin.empty else set()

            with st.form("add_itinerary_form", clear_on_submit=True):
                ic1, ic2, ic3 = st.columns([2, 1.2, 3.8])
                with ic1:
                    sel_day_label = st.selectbox("選擇第幾天 *", day_options)
                with ic2:
                    sel_time = st.time_input("預計時間 *", value=datetime.time(10, 0), step=900)

                if add_source.startswith("📋"):
                    with ic3:
                        if df_links.empty:
                            st.warning("共用清單目前是空的，請切換為「手動輸入自訂行程」或先去新增景點！")
                            sel_spot_idx = None
                        else:
                            spot_choices = {}
                            for idx, r in df_links.iterrows():
                                s_title = str(r["自訂名稱"]).strip() or str(r["自動抓取標題"]).strip() or str(r["網址"])[:30]
                                is_used = s_title in scheduled_names_set
                                if only_unscheduled and is_used:
                                    continue
                                tag_used = " ✅[已排]" if is_used else ""
                                s_lbl = f"[{idx}] {s_title}（{r['區域']}｜{r['分類']}）{tag_used}"
                                spot_choices[s_lbl] = idx

                            if not spot_choices:
                                st.success("🎉 共用清單裡的景點已經全部排進行程表囉！（取消勾選右上角篩選即可重複排入）")
                                sel_spot_idx = None
                            else:
                                chosen_spot_lbl = st.selectbox("選擇要排入的景點 *", list(spot_choices.keys()))
                                sel_spot_idx = spot_choices[chosen_spot_lbl]
                    in_note = st.text_input("行程備註（選填）", placeholder="例如：已訂位 18:30、記得先抽整理券")
                else:
                    with ic3:
                        custom_itin_name = st.text_input("行程 / 活動名稱 *", placeholder="例如：抵達關西機場、心齋橋自由購物")
                    ic4, ic5 = st.columns([2, 4])
                    with ic4:
                        custom_itin_cat = st.selectbox("分類", CAT_OPTIONS)
                    with ic5:
                        in_note = st.text_input("行程備註（選填）", placeholder="例如：搭乘南海電鐵特急 Rapi:t")

                if st.form_submit_button("➕ 加入該日行程表", type="primary"):
                    day_tag, date_str = day_date_map[sel_day_label]
                    time_str = sel_time.strftime("%H:%M")

                    if add_source.startswith("📋"):
                        if df_links.empty or sel_spot_idx is None:
                            st.warning("⚠️ 請先選擇有效的景點，或改用手動輸入！")
                        else:
                            r = df_links.loc[sel_spot_idx]
                            s_title = str(r["自訂名稱"]).strip() or str(r["自動抓取標題"]).strip()
                            orig_note = str(r["備註"]).strip()
                            merged_note = f"{in_note.strip()}（{orig_note}）" if (in_note.strip() and orig_note) else (in_note.strip() or orig_note)

                            new_it = pd.DataFrame([{
                                "天數": day_tag,
                                "日期": date_str,
                                "時間": time_str,
                                "行程名稱": s_title,
                                "分類": str(r["分類"]).strip(),
                                "最近車站": str(r["最近地鐵站"]).strip(),
                                "網址": str(r["網址"]).strip(),
                                "備註": merged_note,
                                "排程者": current_user
                            }])
                            df_itin = pd.concat([df_itin, new_it], ignore_index=True)
                            save_sheet_df("每日行程", df_itin, ITINERARY_COLS)
                            st.success(f"✅ 已將「{s_title}」排入 {sel_day_label} {time_str}！")
                            st.rerun()
                    else:
                        if not custom_itin_name.strip():
                            st.warning("⚠️ 請輸入行程名稱喔！")
                        else:
                            new_it = pd.DataFrame([{
                                "天數": day_tag,
                                "日期": date_str,
                                "時間": time_str,
                                "行程名稱": custom_itin_name.strip(),
                                "分類": custom_itin_cat,
                                "最近車站": "-",
                                "網址": "",
                                "備註": in_note.strip(),
                                "排程者": current_user
                            }])
                            df_itin = pd.concat([df_itin, new_it], ignore_index=True)
                            save_sheet_df("每日行程", df_itin, ITINERARY_COLS)
                            st.success(f"✅ 已將「{custom_itin_name.strip()}」排入 {sel_day_label} {time_str}！")
                            st.rerun()

        st.divider()

        view_c1, view_c2 = st.columns([2.2, 1.8])
        with view_c1:
            view_mode = st.radio(
                "👀 選擇行程呈現模式：",
                ["⏳ 時間軸卡片模式（含相鄰景點距離與航班/住宿提示）", "📊 總覽表格模式（一眼看全趟行程表）"],
                horizontal=True
            )
        with view_c2:
            filter_day_choice = st.selectbox("篩選檢視天數：", ["全部天數總覽"] + day_options)

        if df_itin.empty and df_hotels.empty and df_flights.empty:
            st.info("目前行程表還是空的，快從上方步驟 2 開始把景點排進每一天吧！")
        else:
            df_itin["_day_num"] = df_itin["天數"].apply(
                lambda x: int(re.search(r'\d+', str(x)).group(0)) if re.search(r'\d+', str(x)) else 999
            )
            df_itin_sorted = df_itin.sort_values(by=["_day_num", "時間"]).copy()

            if filter_day_choice != "全部天數總覽":
                target_d_tag = day_date_map[filter_day_choice][0]
                df_show_itin = df_itin_sorted[df_itin_sorted["天數"] == target_d_tag].copy()
            else:
                df_show_itin = df_itin_sorted.copy()

            if view_mode.startswith("⏳"):
                if df_show_itin.empty:
                    st.info(f"「{filter_day_choice}」目前還沒有安排任何行程喔！")
                else:
                    link_meta_map = {}
                    for _, lr in df_links.iterrows():
                        l_name = str(lr["自訂名稱"]).strip() or str(lr["自動抓取標題"]).strip()
                        if l_name:
                            link_meta_map[l_name] = (str(lr["網址"]).strip(), str(lr["區域"]).strip(), str(lr["備註"]).strip())

                    for day_tag, group in df_show_itin.groupby("天數", sort=False):
                        date_label = group.iloc[0]["日期"]
                        st.markdown(f"### 🗓️ {day_tag} ｜ {date_label}")

                        cur_date_obj = day_date_obj_map.get(day_tag)
                        if cur_date_obj:
                            day_flights = get_flight_events_for_date(df_flights, cur_date_obj)
                            for fl in day_flights:
                                st.warning(
                                    f"✈️ **當日航班（{fl['航段類型']}）**：**{fl['航空與班次']}** ｜ "
                                    f"`{fl['起飛時間']}` {fl['出發機場航廈']} ➔ `{fl['抵達時間']}` {fl['抵達機場航廈']}"
                                )

                            if not df_hotels.empty:
                                co_list, ci_list, stay_list = get_hotel_events_for_date(df_hotels, cur_date_obj)
                                hotel_msgs = []
                                for h in co_list:
                                    hotel_msgs.append(f"🧳 **上午退房**：{h['飯店名稱']}（退房時間 `{h['退房時間']}`）")
                                for h in ci_list:
                                    hotel_msgs.append(f"🔑 **下午入住**：{h['飯店名稱']}（入住時間 `{h['入住時間']}`｜{h['最近車站']}）")
                                if not ci_list and stay_list:
                                    h = stay_list[0]
                                    hotel_msgs.append(f"🏨 **今晚下榻飯店**：{h['飯店名稱']}（{h['最近車站']}）")

                                if hotel_msgs:
                                    h_col1, h_col2 = st.columns([4, 1.5])
                                    with h_col1:
                                        st.info(" ｜ ".join(hotel_msgs))
                                    with h_col2:
                                        if ci_list or co_list:
                                            if st.button(f"➕ 一鍵把入住/退房排入 {day_tag}", key=f"auto_add_hotel_{day_tag}", use_container_width=True):
                                                added_rows = []
                                                for h in co_list:
                                                    added_rows.append({
                                                        "天數": day_tag,
                                                        "日期": date_label,
                                                        "時間": str(h["退房時間"]).strip() or "11:00",
                                                        "行程名稱": f"🏨 飯店退房：{h['飯店名稱']}",
                                                        "分類": "交通/住宿",
                                                        "最近車站": str(h["最近車站"]).strip(),
                                                        "網址": str(h["網址"]).strip(),
                                                        "備註": str(h["訂房備註"]).strip(),
                                                        "排程者": current_user
                                                    })
                                                for h in ci_list:
                                                    added_rows.append({
                                                        "天數": day_tag,
                                                        "日期": date_label,
                                                        "時間": str(h["入住時間"]).strip() or "15:00",
                                                        "行程名稱": f"🔑 飯店入住：{h['飯店名稱']}",
                                                        "分類": "交通/住宿",
                                                        "最近車站": str(h["最近車站"]).strip(),
                                                        "網址": str(h["網址"]).strip(),
                                                        "備註": str(h["訂房備註"]).strip(),
                                                        "排程者": current_user
                                                    })
                                                df_itin_clean = df_itin.drop(columns=["_day_num"], errors="ignore")
                                                df_itin_clean = pd.concat([df_itin_clean, pd.DataFrame(added_rows)], ignore_index=True)
                                                save_sheet_df("每日行程", df_itin_clean, ITINERARY_COLS)
                                                st.rerun()

                        places_for_route = [
                            urllib.parse.quote(f"関西 {str(row['行程名稱']).replace('🏨 飯店退房：', '').replace('🔑 飯店入住：', '').strip()}")
                            for _, row in group.iterrows()
                        ]
                        if len(places_for_route) >= 2:
                            origin_p = places_for_route[0]
                            dest_p = places_for_route[-1]
                            waypoints = "|".join(places_for_route[1:-1])
                            multi_route_url = (
                                f"https://www.google.com/maps/dir/?api=1"
                                f"&origin={origin_p}&destination={dest_p}"
                                + (f"&waypoints={waypoints}" if waypoints else "")
                                + "&travelmode=transit"
                            )
                            st.markdown(f"👉 **[🗺️ 一鍵開啟 {day_tag} 完整順路多點導航（共 {len(places_for_route)} 站）]({multi_route_url})**")

                        group_items = list(group.iterrows())
                        resolved_coords_list = []
                        for _, r_item in group_items:
                            is_manual_custom = (str(r_item["最近車站"]).strip() == "-" and not str(r_item["網址"]).strip())
                            if is_manual_custom:
                                resolved_coords_list.append((None, None, []))
                            else:
                                s_n = str(r_item["行程名稱"]).replace("🏨 飯店退房：", "").replace("🔑 飯店入住：", "").strip()
                                u_str, reg_str, note_str = link_meta_map.get(s_n, (str(r_item["網址"]).strip(), "", ""))
                                _, _, lat_v, lon_v, _, st_lines_v = resolve_place_full(url=u_str, custom_name=s_n, region=reg_str, note=note_str)
                                resolved_coords_list.append((lat_v, lon_v, st_lines_v))

                        for pos_i, (orig_idx, row) in enumerate(group_items):
                            with st.container(border=True):
                                t_c1, t_c2, t_c3 = st.columns([1.5, 5.2, 1.8])
                                with t_c1:
                                    try:
                                        h_val, m_val = map(int, str(row["時間"]).split(":")[:2])
                                        cur_t_obj = datetime.time(h_val, m_val)
                                    except Exception:
                                        cur_t_obj = datetime.time(10, 0)

                                    new_t_obj = st.time_input(
                                        "⏰ 時間",
                                        value=cur_t_obj,
                                        step=900,
                                        key=f"edit_time_{orig_idx}",
                                        label_visibility="collapsed"
                                    )
                                    if new_t_obj.strftime("%H:%M") != str(row["時間"]).strip():
                                        if st.button("💾 更新時間", key=f"save_time_{orig_idx}", type="primary", use_container_width=True):
                                            df_itin.at[orig_idx, "時間"] = new_t_obj.strftime("%H:%M")
                                            df_itin_clean = df_itin.drop(columns=["_day_num"], errors="ignore")
                                            save_sheet_df("每日行程", df_itin_clean, ITINERARY_COLS)
                                            st.rerun()
                                    st.caption(f"🏷️ {row['分類']}")

                                with t_c2:
                                    st.markdown(f"**{row['行程名稱']}**")
                                    if str(row["最近車站"]).strip() and str(row["最近車站"]).strip() != "-":
                                        st.caption(f"🚉 {row['最近車站']}")
                                    if str(row["備註"]).strip():
                                        st.write(f"📝 {row['備註']}")

                                with t_c3:
                                    if str(row["網址"]).strip().startswith("http"):
                                        st.link_button("🔗 連結", str(row["網址"]).strip(), use_container_width=True)
                                    if st.button("🗑️ 移除", key=f"del_itin_{orig_idx}", use_container_width=True):
                                        df_itin = df_itin.drop(orig_idx).drop(columns=["_day_num"], errors="ignore").reset_index(drop=True)
                                        save_sheet_df("每日行程", df_itin, ITINERARY_COLS)
                                        st.rerun()

                            if pos_i < len(group_items) - 1:
                                lat1, lon1, lines1 = resolved_coords_list[pos_i]
                                lat2, lon2, lines2 = resolved_coords_list[pos_i + 1]
                                next_name = group_items[pos_i + 1][1]["行程名稱"]
                                if lat1 is not None and lat2 is not None:
                                    dkm = calc_haversine_km(lat1, lon1, lat2, lon2)
                                    is_dir = bool(set(lines1) & set(lines2))
                                    dur_txt = estimate_travel_duration(dkm, is_direct_line=is_dir)
                                    direct_badge = " ｜ ✨ 同一條鐵道線直達" if (is_dir and dkm > 1.2) else ""
                                    st.markdown(
                                        f"<div style='margin:-6px 0 8px 24px; padding:6px 14px; border-left:3px solid #2ecc71; background-color:rgba(46, 204, 113, 0.08); border-radius:4px; font-size:0.9rem;'>"
                                        f"⬇️ 前往下一站【<b>{next_name}</b>】：直線距離約 <b>{dkm:.2f} km</b> ｜ <b>{dur_txt}</b>{direct_badge}"
                                        f"</div>",
                                        unsafe_allow_html=True
                                    )
                                else:
                                    st.markdown(
                                        f"<div style='margin:-6px 0 8px 24px; padding:4px 14px; border-left:3px solid #bdc3c7; color:#7f8c8d; font-size:0.85rem;'>"
                                        f"⬇️ 前往下一站【{next_name}】"
                                        f"</div>",
                                        unsafe_allow_html=True
                                    )

            else:
                table_df = df_show_itin[["天數", "日期", "時間", "行程名稱", "分類", "最近車站", "備註", "網址", "排程者"]].copy()
                st.dataframe(
                    table_df,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "網址": st.column_config.LinkColumn("網址", display_text="🔗 開啟")
                    }
                )
                with st.expander("🗑️ 刪除特定行程項目"):
                    del_opts = [
                        f"[{idx}] {r['天數']} {r['時間']}｜{r['行程名稱']}"
                        for idx, r in df_itin_sorted.iterrows()
                    ]
                    sel_del_it = st.selectbox("選擇要移除的行程：", del_opts)
                    if st.button("確認移除此行程", type="secondary"):
                        del_idx = int(sel_del_it.split("]")[0].replace("[", ""))
                        df_itin = df_itin.drop(del_idx).drop(columns=["_day_num"], errors="ignore").reset_index(drop=True)
                        save_sheet_df("每日行程", df_itin, ITINERARY_COLS)
                        st.success("✅ 已從行程表移除！")
                        st.rerun()

        st.divider()

        with st.container(border=True):
            st.subheader("🤖 AI 旅遊規劃小助手（行程健檢 ＆ 交通顧問）")
            has_gemini = bool(st.secrets.get("gemini_api_key", "").strip())

            if not has_gemini:
                st.info("""
                💡 **尚未啟用 AI 小助手（設定只需 30 秒且完全免費）**：
                1. 前往 Google 官方的 **[Google AI Studio](https://aistudio.google.com/apikey)**，登入私人 Google 帳號並點 **「Create API key」** 複製金鑰。
                2. 打開 Streamlit 右下角 **Manage app → Settings → Secrets**，在最下方新增一行：
                   `gemini_api_key = "你的免費金鑰"` 並按 Save 即可開始跟 AI 對話！
                """)
            else:
                st.caption("AI 小助手已自動讀取你們的「機票航班時間」、「住宿飯店資訊」、「每日行程表」與「全團收藏清單」，你可以點下方快捷按鈕或直接發問：")
                q_b1, q_b2, q_b3 = st.columns(3)
                quick_ai_prompt = ""
                with q_b1:
                    if st.button("🔍 幫我健檢目前排程順不順？", use_container_width=True):
                        quick_ai_prompt = "請結合我們的航班起降時間、預訂的住宿飯店位置與入住退房時間，檢查目前已經排入的每日行程，分析首末日往返機場時間是否充裕、每天從飯店出發與返回的交通是否順路，並給出具體調整建議。"
                with q_b2:
                    if st.button("🧩 還沒排的景點怎麼塞進空檔？", use_container_width=True):
                        quick_ai_prompt = "請比對『已排入的每日行程』、『住宿飯店周邊』與『全團收藏清單中還沒排進去的景點/美食』，幫我推薦哪些未排入的店剛好順路可以塞進哪一天的空檔？"
                with q_b3:
                    if st.button("🌧️ 幫我規劃雨天備案與排隊提醒", use_container_width=True):
                        quick_ai_prompt = "針對我們目前的行程與收藏清單，請問有哪些餐廳需要提前預約或排隊很久？如果遇到下雨天，哪幾天的行程需要準備室內雨天備案？"

                user_ai_q = st.text_area(
                    "💬 或者直接輸入你想問 AI 小助手的問題：",
                    value=quick_ai_prompt,
                    height=90,
                    placeholder="例如：回程飛機起飛前幾點從飯店出發去關西機場最保險？"
                )

                if st.button("✨ 請 AI 小助手分析建議", type="primary"):
                    if not user_ai_q.strip():
                        st.warning("⚠️ 請先輸入問題或點擊上方快捷按鈕喔！")
                    else:
                        trip_info_str = f"{active_start.strftime('%Y/%m/%d')} 至 {active_end.strftime('%Y/%m/%d')}（共 {total_days} 天）"
                        with st.spinner("🤖 AI 小助手正在比對航班時間、住宿飯店、所有景點車站與行程動線中..."):
                            ai_reply = ask_travel_ai_assistant(
                                user_ai_q.strip(),
                                df_itin.drop(columns=["_day_num"], errors="ignore"),
                                df_links,
                                df_hotels,
                                df_flights,
                                trip_info_str
                            )
                        st.markdown("### 💡 AI 顧問分析報告")
                        st.markdown(ai_reply)

    # ==========================================================
    # 頁面：✈️ 機票 ＆ 住宿整理
    # ==========================================================
    elif active_page == "flight_hotel":
        st.subheader("✈️ 機票 ＆ 住宿確定資訊總覽")
        st.caption("把整趟旅程確定的機票航班、航廈、訂位代號與住宿飯店並排整理在這裡，出國查憑證一眼看完，並自動連動到每日行程表與地圖！")

        default_ci_date = saved_start if saved_start else datetime.date.today() + timedelta(days=14)
        default_co_date = saved_end if saved_end else default_ci_date + timedelta(days=4)

        view_fl_col, view_ht_col = st.columns(2)

        with view_fl_col:
            st.markdown("#### ✈️ 航班機票資訊")
            if df_flights.empty:
                with st.container(border=True):
                    st.info("尚未登錄機票資訊，請從下方「➕ 新增機票航班」填寫去回程班次！")
            else:
                df_flights_sorted = df_flights.sort_values(by=["起飛日期", "起飛時間"]).copy()
                for _, f_row in df_flights_sorted.iterrows():
                    try:
                        d_dep = datetime.date.fromisoformat(str(f_row["起飛日期"]).strip())
                        dep_d_disp = f"{d_dep.strftime('%m/%d')} ({WEEKDAY_ZH[d_dep.weekday()]})"
                    except Exception:
                        dep_d_disp = str(f_row["起飛日期"]).strip()

                    try:
                        d_arr = datetime.date.fromisoformat(str(f_row["抵達日期"]).strip())
                        arr_d_disp = f"{d_arr.strftime('%m/%d')} ({WEEKDAY_ZH[d_arr.weekday()]})"
                    except Exception:
                        arr_d_disp = str(f_row["抵達日期"]).strip()

                    with st.container(border=True):
                        st.markdown(f"`{f_row['航段類型']}` **{f_row['航空與班次']}**")
                        st.markdown(
                            f"🛫 **出發**：{f_row['出發機場航廈']} ｜ `{dep_d_disp} {f_row['起飛時間']}`\n\n"
                            f"🛬 **抵達**：{f_row['抵達機場航廈']} ｜ `{arr_d_disp} {f_row['抵達時間']}`"
                        )
                        if str(f_row["訂位與行李備註"]).strip():
                            st.caption(f"🎫 {f_row['訂位與行李備註']}")

        with view_ht_col:
            st.markdown("#### 🏨 住宿飯店資訊")
            if df_hotels.empty:
                with st.container(border=True):
                    st.info("尚未登錄住宿飯店，請從下方「➕ 新增住宿飯店」填寫訂房資訊！")
            else:
                df_hotels_sorted = df_hotels.sort_values(by=["入住日期"]).copy()
                for _, h_row in df_hotels_sorted.iterrows():
                    try:
                        d_in = datetime.date.fromisoformat(str(h_row["入住日期"]).strip())
                        d_out = datetime.date.fromisoformat(str(h_row["退房日期"]).strip())
                        nights = max(1, (d_out - d_in).days)
                        in_disp = f"{d_in.strftime('%m/%d')} ({WEEKDAY_ZH[d_in.weekday()]}) {h_row['入住時間']}"
                        out_disp = f"{d_out.strftime('%m/%d')} ({WEEKDAY_ZH[d_out.weekday()]}) {h_row['退房時間']}"
                        stay_badge = f"共 {nights} 晚"
                    except Exception:
                        in_disp = f"{h_row['入住日期']} {h_row['入住時間']}"
                        out_disp = f"{h_row['退房日期']} {h_row['退房時間']}"
                        stay_badge = ""

                    with st.container(border=True):
                        h_top1, h_top2 = st.columns([3.5, 1.3])
                        with h_top1:
                            st.markdown(f"**🏨 {h_row['飯店名稱']}** `{stay_badge}`")
                        with h_top2:
                            gmap_h_url = str(h_row["網址"]).strip()
                            if not gmap_h_url.startswith("http"):
                                q_h = urllib.parse.quote(f"関西 {h_row['飯店名稱']}")
                                gmap_h_url = f"https://www.google.com/maps/search/?api=1&query={q_h}"
                            st.link_button("📍 地圖", gmap_h_url, use_container_width=True)

                        st.markdown(f"🔑 **入住**：`{in_disp}` ➡️ 🧳 **退房**：`{out_disp}`")
                        if str(h_row["最近車站"]).strip():
                            st.caption(f"🚉 {h_row['最近車站']}")
                        if str(h_row["訂房備註"]).strip():
                            st.caption(f"📝 {h_row['訂房備註']}")

        st.divider()

        add_fl_col, add_ht_col = st.columns(2)

        with add_fl_col:
            with st.expander("➕ 新增機票航班（去程 / 回程 / 轉機）", expanded=df_flights.empty):
                with st.form("add_flight_form", clear_on_submit=True):
                    f_leg = st.selectbox("航段類型 *", FLIGHT_LEG_OPTIONS)
                    f_airline = st.text_input("航空公司 ＆ 航班編號 *", placeholder="例如：星宇航空 JX820 / 長榮 BR132")

                    fc_d1, fc_d2, fc_d3 = st.columns([2.2, 1.5, 1.3])
                    with fc_d1:
                        f_dep_port = st.text_input("🛫 出發機場/航廈 *", placeholder="桃園 TPE T1")
                    with fc_d2:
                        f_dep_date = st.date_input("起飛日期 *", value=default_ci_date)
                    with fc_d3:
                        f_dep_time = st.time_input("起飛時間 *", value=datetime.time(8, 30), step=300)

                    fc_a1, fc_a2, fc_a3 = st.columns([2.2, 1.5, 1.3])
                    with fc_a1:
                        f_arr_port = st.text_input("🛬 抵達機場/航廈 *", placeholder="關西 KIX T1")
                    with fc_a2:
                        f_arr_date = st.date_input("抵達日期 *", value=default_ci_date)
                    with fc_a3:
                        f_arr_time = st.time_input("抵達時間 *", value=datetime.time(12, 10), step=300)

                    f_note = st.text_input("📝 訂位代號 / 行李額度（選填）", placeholder="例如：訂位代號 ABC123、託運 23kg")

                    if st.form_submit_button("✅ 儲存航班資訊", type="primary", use_container_width=True):
                        if not f_airline.strip() or not f_dep_port.strip() or not f_arr_port.strip():
                            st.warning("⚠️ 請完整填寫「航空公司與班次」、「出發機場」與「抵達機場」喔！")
                        else:
                            new_f = pd.DataFrame([{
                                "航段類型": f_leg,
                                "航空與班次": f_airline.strip(),
                                "出發機場航廈": f_dep_port.strip(),
                                "起飛日期": f_dep_date.isoformat(),
                                "起飛時間": f_dep_time.strftime("%H:%M"),
                                "抵達機場航廈": f_arr_port.strip(),
                                "抵達日期": f_arr_date.isoformat(),
                                "抵達時間": f_arr_time.strftime("%H:%M"),
                                "訂位與行李備註": f_note.strip(),
                                "新增者": current_user
                            }])
                            df_flights = pd.concat([df_flights, new_f], ignore_index=True)
                            save_sheet_df("機票航班", df_flights, FLIGHT_COLS)
                            st.success(f"✅ 已新增航班「{f_airline.strip()}」！")
                            st.rerun()

        with add_ht_col:
            with st.expander("➕ 新增住宿飯店（支援單間或多間換飯店）", expanded=df_hotels.empty):
                with st.form("add_hotel_form", clear_on_submit=True):
                    h_name = st.text_input("🏨 飯店名稱 *", placeholder="例如：難波東方酒店 / 阪急龍仕柏飯店")
                    h_url = st.text_input("🔗 Google Maps 網址或訂房連結（選填）", placeholder="https://maps.app.goo.gl/...")

                    hc_i1, hc_i2 = st.columns([1.6, 1.4])
                    with hc_i1:
                        h_ci_date = st.date_input("📅 入住日期 *", value=default_ci_date)
                    with hc_i2:
                        h_ci_time = st.time_input("🕒 入住時間 *", value=datetime.time(15, 0), step=900)

                    hc_o1, hc_o2 = st.columns([1.6, 1.4])
                    with hc_o1:
                        h_co_date = st.date_input("📅 退房日期 *", value=default_co_date)
                    with hc_o2:
                        h_co_time = st.time_input("🕚 退房時間 *", value=datetime.time(11, 0), step=900)

                    h_note = st.text_input("📝 訂房代號 / 備註（選填）", placeholder="例如：Agoda #123456、可寄放行李")

                    if st.form_submit_button("✅ 儲存飯店資訊", type="primary", use_container_width=True):
                        if not h_name.strip():
                            st.warning("⚠️ 請輸入飯店名稱喔！")
                        elif h_co_date < h_ci_date:
                            st.error("❌ 退房日期不能早於入住日期喔！")
                        else:
                            with st.spinner("🔍 正在解析飯店座標與最近車站..."):
                                _, resolved_sub = resolve_place_and_subway(h_url.strip(), h_name.strip(), "", h_note.strip())

                            new_h = pd.DataFrame([{
                                "飯店名稱": h_name.strip(),
                                "網址": h_url.strip(),
                                "入住日期": h_ci_date.isoformat(),
                                "入住時間": h_ci_time.strftime("%H:%M"),
                                "退房日期": h_co_date.isoformat(),
                                "退房時間": h_co_time.strftime("%H:%M"),
                                "最近車站": resolved_sub,
                                "訂房備註": h_note.strip(),
                                "新增者": current_user
                            }])
                            df_hotels = pd.concat([df_hotels, new_h], ignore_index=True)
                            save_sheet_df("住宿飯店", df_hotels, HOTEL_COLS)
                            st.success(f"✅ 已新增飯店「{h_name.strip()}」！")
                            st.rerun()

        if not df_flights.empty or not df_hotels.empty:
            with st.expander("✏️ 修改或刪除已登錄的機票 / 住宿"):
                if not df_flights.empty:
                    st.markdown("**✈️ 管理機票航班**")
                    f_opts = [f"[{i}] {r['航段類型']}｜{r['航空與班次']}（{r['起飛日期']} {r['起飛時間']}）" for i, r in df_flights.iterrows()]
                    sel_f_lbl = st.selectbox("選擇要編輯或刪除的航班：", f_opts)
                    f_idx = int(sel_f_lbl.split("]")[0].replace("[", ""))
                    f_row = df_flights.loc[f_idx]

                    with st.form(f"edit_flight_form_{f_idx}"):
                        ef1, ef2 = st.columns([1.8, 3.2])
                        with ef1:
                            cur_leg_idx = FLIGHT_LEG_OPTIONS.index(f_row["航段類型"]) if f_row["航段類型"] in FLIGHT_LEG_OPTIONS else 0
                            ed_f_leg = st.selectbox("航段類型", FLIGHT_LEG_OPTIONS, index=cur_leg_idx)
                        with ef2:
                            ed_f_airline = st.text_input("航空公司 ＆ 航班編號", value=str(f_row["航空與班次"]))

                        ef3, ef4, ef5 = st.columns([2.5, 1.5, 1.2])
                        try:
                            cur_dep_d = datetime.date.fromisoformat(str(f_row["起飛日期"]).strip())
                        except Exception:
                            cur_dep_d = default_ci_date
                        try:
                            dh, dm = map(int, str(f_row["起飛時間"]).split(":")[:2])
                            cur_dep_t = datetime.time(dh, dm)
                        except Exception:
                            cur_dep_t = datetime.time(8, 30)

                        with ef3:
                            ed_f_dep_port = st.text_input("出發機場 / 航廈", value=str(f_row["出發機場航廈"]))
                        with ef4:
                            ed_f_dep_d = st.date_input("起飛日期", value=cur_dep_d)
                        with ef5:
                            ed_f_dep_t = st.time_input("起飛時間", value=cur_dep_t, step=300)

                        ef6, ef7, ef8 = st.columns([2.5, 1.5, 1.2])
                        try:
                            cur_arr_d = datetime.date.fromisoformat(str(f_row["抵達日期"]).strip())
                        except Exception:
                            cur_arr_d = default_ci_date
                        try:
                            ah, am = map(int, str(f_row["抵達時間"]).split(":")[:2])
                            cur_arr_t = datetime.time(ah, am)
                        except Exception:
                            cur_arr_t = datetime.time(12, 0)

                        with ef6:
                            ed_f_arr_port = st.text_input("抵達機場 / 航廈", value=str(f_row["抵達機場航廈"]))
                        with ef7:
                            ed_f_arr_d = st.date_input("抵達日期", value=cur_arr_d)
                        with ef8:
                            ed_f_arr_t = st.time_input("抵達時間", value=cur_arr_t, step=300)

                        ed_f_note = st.text_input("訂位與行李備註", value=str(f_row["訂位與行李備註"]))

                        fb1, fb2 = st.columns([2, 1])
                        with fb1:
                            save_f_btn = st.form_submit_button("💾 儲存航班修改", type="primary", use_container_width=True)
                        with fb2:
                            del_f_btn = st.form_submit_button("🗑️ 刪除此航班", type="secondary", use_container_width=True)

                        if save_f_btn:
                            df_flights.at[f_idx, "航段類型"] = ed_f_leg
                            df_flights.at[f_idx, "航空與班次"] = ed_f_airline.strip()
                            df_flights.at[f_idx, "出發機場航廈"] = ed_f_dep_port.strip()
                            df_flights.at[f_idx, "起飛日期"] = ed_f_dep_d.isoformat()
                            df_flights.at[f_idx, "起飛時間"] = ed_f_dep_t.strftime("%H:%M")
                            df_flights.at[f_idx, "抵達機場航廈"] = ed_f_arr_port.strip()
                            df_flights.at[f_idx, "抵達日期"] = ed_f_arr_d.isoformat()
                            df_flights.at[f_idx, "抵達時間"] = ed_f_arr_t.strftime("%H:%M")
                            df_flights.at[f_idx, "訂位與行李備註"] = ed_f_note.strip()
                            save_sheet_df("機票航班", df_flights, FLIGHT_COLS)
                            st.success("✅ 航班資訊已更新！")
                            st.rerun()

                        if del_f_btn:
                            df_flights = df_flights.drop(f_idx).reset_index(drop=True)
                            save_sheet_df("機票航班", df_flights, FLIGHT_COLS)
                            st.success("✅ 已刪除此筆航班！")
                            st.rerun()

                if not df_flights.empty and not df_hotels.empty:
                    st.divider()

                if not df_hotels.empty:
                    st.markdown("**🏨 管理住宿飯店**")
                    h_opts = [f"[{i}] {r['飯店名稱']}（{r['入住日期']} ～ {r['退房日期']}）" for i, r in df_hotels.iterrows()]
                    sel_h_lbl = st.selectbox("選擇要編輯或刪除的飯店：", h_opts)
                    h_idx = int(sel_h_lbl.split("]")[0].replace("[", ""))
                    h_row = df_hotels.loc[h_idx]

                    with st.form(f"edit_hotel_form_{h_idx}"):
                        eh1, eh2 = st.columns([3, 3])
                        with eh1:
                            e_h_name = st.text_input("飯店名稱", value=str(h_row["飯店名稱"]))
                        with eh2:
                            e_h_url = st.text_input("網址", value=str(h_row["網址"]))

                        eh3, eh4, eh5, eh6 = st.columns([2, 1.5, 2, 1.5])
                        try:
                            cur_ci_d = datetime.date.fromisoformat(str(h_row["入住日期"]).strip())
                            cur_co_d = datetime.date.fromisoformat(str(h_row["退房日期"]).strip())
                        except Exception:
                            cur_ci_d = default_ci_date
                            cur_co_d = default_co_date

                        try:
                            ci_h, ci_m = map(int, str(h_row["入住時間"]).split(":")[:2])
                            cur_ci_t = datetime.time(ci_h, ci_m)
                        except Exception:
                            cur_ci_t = datetime.time(15, 0)

                        try:
                            co_h, co_m = map(int, str(h_row["退房時間"]).split(":")[:2])
                            cur_co_t = datetime.time(co_h, co_m)
                        except Exception:
                            cur_co_t = datetime.time(11, 0)

                        with eh3:
                            e_ci_d = st.date_input("入住日期", value=cur_ci_d)
                        with eh4:
                            e_ci_t = st.time_input("入住時間", value=cur_ci_t, step=900)
                        with eh5:
                            e_co_d = st.date_input("退房日期", value=cur_co_d)
                        with eh6:
                            e_co_t = st.time_input("退房時間", value=cur_co_t, step=900)

                        e_h_note = st.text_input("訂房備註", value=str(h_row["訂房備註"]))
                        recheck_h_sub = st.checkbox("🔄 儲存時重新反查最近車站")

                        hb1, hb2 = st.columns([2, 1])
                        with hb1:
                            save_h_btn = st.form_submit_button("💾 儲存飯店修改", type="primary", use_container_width=True)
                        with hb2:
                            del_h_btn = st.form_submit_button("🗑️ 刪除此飯店", type="secondary", use_container_width=True)

                        if save_h_btn:
                            if e_co_d < e_ci_d:
                                st.error("❌ 退房日期不能早於入住日期喔！")
                            else:
                                final_h_sub = str(h_row["最近車站"]).strip()
                                if recheck_h_sub or e_h_name.strip() != str(h_row["飯店名稱"]).strip():
                                    _, final_h_sub = resolve_place_and_subway(e_h_url.strip(), e_h_name.strip(), "", e_h_note.strip())

                                df_hotels.at[h_idx, "飯店名稱"] = e_h_name.strip()
                                df_hotels.at[h_idx, "網址"] = e_h_url.strip()
                                df_hotels.at[h_idx, "入住日期"] = e_ci_d.isoformat()
                                df_hotels.at[h_idx, "入住時間"] = e_ci_t.strftime("%H:%M")
                                df_hotels.at[h_idx, "退房日期"] = e_co_d.isoformat()
                                df_hotels.at[h_idx, "退房時間"] = e_co_t.strftime("%H:%M")
                                df_hotels.at[h_idx, "最近車站"] = final_h_sub
                                df_hotels.at[h_idx, "訂房備註"] = e_h_note.strip()
                                save_sheet_df("住宿飯店", df_hotels, HOTEL_COLS)
                                st.success("✅ 飯店資訊已更新！")
                                st.rerun()

                        if del_h_btn:
                            df_hotels = df_hotels.drop(h_idx).reset_index(drop=True)
                            save_sheet_df("住宿飯店", df_hotels, HOTEL_COLS)
                            st.success("✅ 已刪除此間飯店！")
                            st.rerun()

    # ==========================================================
    # 頁面：🗺️ 互動地圖總覽
    # ==========================================================
    elif active_page == "map":
        st.subheader("🗺️ 全景點 ＆ 每日動線互動地圖總覽")
        st.caption("🟡 住宿飯店 ｜ 🔴 美食 ｜ 🟣 購物 ｜ 🟢 景點與地標 ｜ 🔵 交通/住宿（滑鼠移到標記上或手機點擊可查看店名與車站資訊）")

        ctrl_m1, ctrl_m2 = st.columns([3, 2])
        with ctrl_m1:
            map_mode = st.radio(
                "選擇地圖顯示模式：",
                [
                    "🌍 全部收藏景點 ＆ 住宿飯店總覽",
                    "🗓️ 單日行程動線連線模式（用路線連起當天每一站）"
                ],
                horizontal=True
            )
        with ctrl_m2:
            marker_style = st.radio(
                "選擇標記樣式：",
                [
                    "● 精緻小圓點（清爽不遮路）",
                    "📍 經典地圖圖釘"
                ],
                horizontal=True
            )

        if map_mode.startswith("🌍"):
            if df_links.empty and df_hotels.empty:
                st.info("目前共用清單與住宿飯店還沒有資料喔！")
            else:
                mf1, mf2, mf3 = st.columns([2, 2, 1.5])
                with mf1:
                    map_cats = st.multiselect("篩選顯示分類：", CAT_OPTIONS, default=CAT_OPTIONS, key="map_cat_filter")
                with mf2:
                    map_regs = st.multiselect("篩選顯示區域：", REGION_OPTIONS, default=REGION_OPTIONS, key="map_reg_filter")
                with mf3:
                    st.write("")
                    show_hotels_on_map = st.checkbox("🏨 同時標示預訂飯店", value=True)

                sub_map_df = df_links[df_links["分類"].isin(map_cats) & df_links["區域"].isin(map_regs)].copy()

                map_points = []
                unmapped_names = []
                with st.spinner("🗺️ 正在定位地圖座標..."):
                    if show_hotels_on_map and not df_hotels.empty:
                        for _, hr in df_hotels.iterrows():
                            h_n = str(hr["飯店名稱"]).strip()
                            h_u = str(hr["網址"]).strip()
                            h_sub = str(hr["最近車站"]).strip()
                            _, _, h_lat, h_lon, _, _ = resolve_place_full(url=h_u, custom_name=h_n)
                            if h_lat is not None and h_lon is not None:
                                map_points.append({
                                    "name": f"🏨 {h_n}",
                                    "category": "住宿飯店",
                                    "region": f"{hr['入住日期']} ~ {hr['退房日期']}",
                                    "station": h_sub,
                                    "note": str(hr["訂房備註"]).strip() or "已預訂住宿飯店",
                                    "lat": h_lat,
                                    "lon": h_lon,
                                    "color": CAT_COLOR_RGBA["住宿飯店"]
                                })

                    for _, r in sub_map_df.iterrows():
                        s_name = str(r["自訂名稱"]).strip() or str(r["自動抓取標題"]).strip()
                        s_url = str(r["網址"]).strip()
                        s_cat = str(r["分類"]).strip()
                        s_reg = str(r["區域"]).strip()
                        s_note = str(r["備註"]).strip()

                        _, s_sub, lat_v, lon_v, _, _ = resolve_place_full(url=s_url, custom_name=s_name, region=s_reg, note=s_note)
                        if lat_v is not None and lon_v is not None:
                            map_points.append({
                                "name": s_name,
                                "category": s_cat,
                                "region": s_reg,
                                "station": s_sub,
                                "note": s_note if s_note else "無備註",
                                "lat": lat_v,
                                "lon": lon_v,
                                "color": CAT_COLOR_RGBA.get(s_cat, [127, 140, 141, 235])
                            })
                        else:
                            unmapped_names.append(s_name)

                if not map_points:
                    st.warning("⚠️ 目前篩選的地點尚未取得精準座標（可能是未指定分店的連鎖店）。")
                else:
                    df_pts = pd.DataFrame(map_points)
                    center_lat = float(df_pts["lat"].mean())
                    center_lon = float(df_pts["lon"].mean())

                    marker_layer = build_map_marker_layer(df_pts, marker_style, is_day_route=False)

                    view_state = pdk.ViewState(
                        latitude=center_lat,
                        longitude=center_lon,
                        zoom=12.4,
                        pitch=0
                    )

                    tooltip = {
                        "html": "<b>{name}</b><br/>🏷️ {category} ｜ 🗺️ {region}<br/>🚉 {station}<br/>📝 {note}",
                        "style": {"backgroundColor": "#1e272e", "color": "white", "fontSize": "13px", "borderRadius": "8px", "padding": "8px"}
                    }

                    st.pydeck_chart(pdk.Deck(
                        layers=[marker_layer],
                        initial_view_state=view_state,
                        tooltip=tooltip,
                        map_style="https://basemaps.cartocdn.com/gl/voyager-gl-style/style.json"
                    ))
                    st.success(f"📍 地圖上已成功標示 **{len(df_pts)}** 個地點（含已預訂飯店）！")
                    if unmapped_names:
                        st.caption(f"ℹ️ 以下地點因未指定分店或尚無座標，未顯示於地圖上：{', '.join(unmapped_names)}")

        else:
            if df_itin.empty:
                st.info("目前「每日行程排程」還沒有資料，請先去排入行程喔！")
            else:
                df_itin_m = df_itin.copy()
                df_itin_m["_d_num"] = df_itin_m["天數"].apply(
                    lambda x: int(re.search(r'\d+', str(x)).group(0)) if re.search(r'\d+', str(x)) else 999
                )
                df_itin_m = df_itin_m.sort_values(by=["_d_num", "時間"])
                avail_days = df_itin_m["天數"].unique().tolist()
                chosen_map_day = st.selectbox("選擇要檢視動線的天數：", avail_days)

                day_sub = df_itin_m[df_itin_m["天數"] == chosen_map_day].copy()
                link_meta_map = {}
                for _, lr in df_links.iterrows():
                    l_name = str(lr["自訂名稱"]).strip() or str(lr["自動抓取標題"]).strip()
                    if l_name:
                        link_meta_map[l_name] = (str(lr["網址"]).strip(), str(lr["區域"]).strip(), str(lr["備註"]).strip())

                day_pts = []
                with st.spinner(f"🧭 正在繪製 {chosen_map_day} 動線地圖..."):
                    for seq_idx, (_, r) in enumerate(day_sub.iterrows(), start=1):
                        s_name = str(r["行程名稱"]).replace("🏨 飯店退房：", "").replace("🔑 飯店入住：", "").strip()
                        u_str, reg_str, note_str = link_meta_map.get(s_name, (str(r["網址"]).strip(), "", ""))
                        _, s_sub, lat_v, lon_v, _, _ = resolve_place_full(url=u_str, custom_name=s_name, region=reg_str, note=note_str)
                        if lat_v is not None and lon_v is not None:
                            day_pts.append({
                                "seq_label": f"第 {seq_idx} 站（{r['時間']}）",
                                "name": str(r["行程名稱"]).strip(),
                                "category": str(r["分類"]).strip(),
                                "station": s_sub,
                                "note": str(r["備註"]).strip() or "無備註",
                                "lat": lat_v,
                                "lon": lon_v,
                                "color": CAT_COLOR_RGBA.get(str(r["分類"]).strip(), [230, 126, 34, 235])
                            })

                if not day_pts:
                    st.warning(f"⚠️ {chosen_map_day} 的行程項目目前無法解析座標，請確認是否為自訂文字或補齊景點名稱。")
                else:
                    df_day_pts = pd.DataFrame(day_pts)
                    layers = []

                    if len(day_pts) >= 2:
                        path_coords = [[p["lon"], p["lat"]] for p in day_pts]
                        path_df = pd.DataFrame([{"path": path_coords}])
                        path_layer = pdk.Layer(
                            "PathLayer",
                            data=path_df,
                            get_path="path",
                            get_color=[230, 126, 34, 210],
                            width_scale=1,
                            width_min_pixels=3,
                        )
                        layers.append(path_layer)

                    pts_layer = build_map_marker_layer(df_day_pts, marker_style, is_day_route=True)
                    layers.append(pts_layer)

                    view_state_day = pdk.ViewState(
                        latitude=float(df_day_pts["lat"].mean()),
                        longitude=float(df_day_pts["lon"].mean()),
                        zoom=12.7,
                        pitch=0
                    )

                    tooltip_day = {
                        "html": "<b>{seq_label}：{name}</b><br/>🏷️ {category}<br/>🚉 {station}<br/>📝 {note}",
                        "style": {"backgroundColor": "#1e272e", "color": "white", "fontSize": "13px", "borderRadius": "8px", "padding": "8px"}
                    }

                    st.pydeck_chart(pdk.Deck(
                        layers=layers,
                        initial_view_state=view_state_day,
                        tooltip=tooltip_day,
                        map_style="https://basemaps.cartocdn.com/gl/voyager-gl-style/style.json"
                    ))

                    if len(day_pts) >= 2:
                        st.markdown(f"#### 📏 {chosen_map_day} 各站移動距離一覽")
                        seg_rows = []
                        total_km = 0.0
                        for k in range(len(day_pts) - 1):
                            p1, p2 = day_pts[k], day_pts[k + 1]
                            dkm = calc_haversine_km(p1["lat"], p1["lon"], p2["lat"], p2["lon"])
                            total_km += dkm
                            seg_rows.append({
                                "路段": f"{p1['seq_label']} {p1['name']} ➔ {p2['seq_label']} {p2['name']}",
                                "直線距離": f"{dkm:.2f} km",
                                "預估移動時間": estimate_travel_duration(dkm)
                            })
                        st.dataframe(pd.DataFrame(seg_rows), use_container_width=True, hide_index=True)
                        st.caption(f"🏁 {chosen_map_day} 全日累計直線移動距離：約 **{total_km:.2f} km**")

    # ==========================================================
    # 頁面：📋 共用旅遊清單
    # ==========================================================
    elif active_page == "list":
        if df_links.empty:
            st.info("目前清單空空如也，快切換到「➕ 新增連結」加入第一個大阪景點吧！")
        else:
            btn_col1, btn_col2, btn_col3 = st.columns(3)
            with btn_col1:
                if st.button("🧭 一鍵智慧自動分區（重新整理所有景點分區）", type="primary", use_container_width=True):
                    prog = st.progress(0, text="🧭 正在根據經緯度座標與最近車站重新判定分區...")
                    changed_count = 0
                    for i, idx in enumerate(df_links.index.tolist()):
                        row = df_links.loc[idx]
                        c_name = str(row["自訂名稱"]).strip() or str(row["自動抓取標題"]).strip()
                        u_str = str(row["網址"]).strip()
                        n_str = str(row["備註"]).strip()
                        old_reg = str(row["區域"]).strip()

                        res_title, res_sub, lat_v, lon_v, _, _ = resolve_place_full(
                            url=u_str, custom_name=c_name, region="", note=n_str
                        )
                        final_t = c_name if c_name else res_title
                        new_reg = classify_kansai_region(
                            lat=lat_v, lon=lon_v, name=final_t, station_text=res_sub, note=n_str
                        )

                        if not str(row["自動抓取標題"]).strip():
                            df_links.at[idx, "自動抓取標題"] = final_t
                        if not str(row["最近地鐵站"]).strip() or str(row["最近地鐵站"]).startswith(("查無", "未取得", "⚠️")):
                            df_links.at[idx, "最近地鐵站"] = res_sub
                        if new_reg != old_reg:
                            df_links.at[idx, "區域"] = new_reg
                            changed_count += 1
                        prog.progress((i + 1) / len(df_links))
                    prog.empty()
                    save_data(df_links)
                    st.success(f"✅ 智慧分區完成！共自動更新了 **{changed_count}** 筆景點的區域分類。")
                    st.rerun()

            with btn_col2:
                def needs_subway_lookup(val):
                    s = str(val).strip()
                    return (not s) or s.startswith(("查無", "未取得", "地鐵查詢逾時", "車站查詢逾時", "非地圖連結", "⏭️", "⚠️")) or ("JROsaka" in s)

                retry_mask = df_links["最近地鐵站"].apply(needs_subway_lookup)
                retry_count = int(retry_mask.sum())
                if retry_count > 0:
                    if st.button(f"🚇 補查 {retry_count} 筆缺漏的最近車站", type="secondary", use_container_width=True):
                        prog = st.progress(0, text="🚇 正在透過 Google API 精準反查最近車站...")
                        retry_indices = df_links[retry_mask].index.tolist()
                        for i, idx in enumerate(retry_indices):
                            row = df_links.loc[idx]
                            c_name = str(row["自訂名稱"]).strip() or str(row["自動抓取標題"]).strip()
                            resolved_title, resolved_subway, auto_reg = resolve_place_subway_and_region(
                                str(row["網址"]).strip(),
                                c_name,
                                str(row["備註"]).strip()
                            )
                            if not str(row["自動抓取標題"]).strip():
                                df_links.at[idx, "自動抓取標題"] = c_name if c_name else resolved_title
                            df_links.at[idx, "最近地鐵站"] = resolved_subway
                            if auto_reg != "其他/未指定":
                                df_links.at[idx, "區域"] = auto_reg
                            prog.progress((i + 1) / len(retry_indices))
                        prog.empty()
                        save_data(df_links)
                        st.success("✅ 已完成車站與分區更新！")
                        st.rerun()
                else:
                    st.button("✅ 所有景點車站皆已解析", disabled=True, use_container_width=True)

            with btn_col3:
                if st.button("🔄 全部重新查詢車站與分區", type="secondary", use_container_width=True):
                    prog = st.progress(0, text="🚇 正在全部重新解析座標、車站與分區...")
                    for i, idx in enumerate(df_links.index.tolist()):
                        row = df_links.loc[idx]
                        c_name = str(row["自訂名稱"]).strip() or str(row["自動抓取標題"]).strip()
                        resolved_title, resolved_subway, auto_reg = resolve_place_subway_and_region(
                            str(row["網址"]).strip(),
                            c_name,
                            str(row["備註"]).strip()
                        )
                        if not str(row["自動抓取標題"]).strip():
                            df_links.at[idx, "自動抓取標題"] = c_name if c_name else resolved_title
                        df_links.at[idx, "最近地鐵站"] = resolved_subway
                        df_links.at[idx, "區域"] = auto_reg
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

    # ==========================================================
    # 頁面：🧭 路線與距離試算
    # ==========================================================
    elif active_page == "route":
        st.subheader("🧭 路線距離與電車轉乘試算")
        calc_mode = st.radio(
            "請選擇試算模式：",
            [
                "🏨 基準定點（已訂飯店/機場）↔ 📋 清單景點（批量雙向比對）",
                "📍 任意兩定點直接對比（例如：比對機場與候選飯店距離）"
            ],
            horizontal=True
        )

        if calc_mode.startswith("🏨"):
            with st.container(border=True):
                st.markdown("**步驟 1️⃣：設定基準定點（可直接點選已預訂的飯店，或輸入機場/車站）**")

                if not df_hotels.empty:
                    h_btn_cols = st.columns(min(4, len(df_hotels)))
                    for idx_h, (_, hr) in enumerate(df_hotels.iterrows()):
                        with h_btn_cols[idx_h % 4]:
                            h_label_name = str(hr["飯店名稱"]).strip()
                            if st.button(f"🏨 帶入：{h_label_name}", key=f"quick_base_h_{idx_h}", use_container_width=True):
                                raw_val = f"{h_label_name} {str(hr['網址']).strip()}".strip()
                                parsed = resolve_free_text_or_url(raw_val)
                                if parsed["lat"] is not None:
                                    st.session_state["base_raw_input"] = raw_val
                                    st.session_state["base_info"] = parsed
                                    st.rerun()

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

    # ==========================================================
    # 頁面：📝 共用筆記 ＆ 重點整理
    # ==========================================================
    elif active_page == "notes":
        df_notes = load_sheet_df("共用筆記", NOTE_COLS)
        st.subheader("📝 共用筆記 ＆ 重點整理牆")
        st.caption("把全團需要一起看的攻略重點、必看 YouTube 影片、交通票券提醒或行前備忘整理在這裡，條理分明不洗版！")

        with st.expander("➕ 新增一則共用筆記 / 攻略重點", expanded=df_notes.empty):
            with st.form("add_shared_note_form", clear_on_submit=True):
                nc1, nc2, nc3 = st.columns([1.8, 3.2, 1])
                with nc1:
                    n_tag = st.selectbox("🏷️ 筆記分類標籤", NOTE_TAG_OPTIONS)
                with nc2:
                    n_title = st.text_input("📌 筆記標題 *", placeholder="例如：關西機場搭南海電鐵到難波教學、環球影城快速通關攻略")
                with nc3:
                    st.write("")
                    st.write("")
                    n_pin = st.checkbox("📌 置頂顯示", value=False)

                n_content = st.text_area(
                    "📄 筆記內文（支援多行條列整理重點）*",
                    height=140,
                    placeholder="例如：\n1. 下飛機領完行李往 2 樓走過空橋到南海電鐵櫃檯\n2. 建議事先在 Klook 買好 Rapi:t 特急券並劃位\n3. 記得隨身攜帶護照辦退稅"
                )
                n_link = st.text_input("🔗 相關影片 / 參考文章連結（選填）", placeholder="https://www.youtube.com/watch?v=...")

                if st.form_submit_button("✅ 發布共用筆記", type="primary"):
                    if not n_title.strip():
                        st.warning("⚠️ 請輸入筆記標題喔！")
                    elif not n_content.strip():
                        st.warning("⚠️ 請輸入筆記內文喔！")
                    else:
                        new_note_row = pd.DataFrame([{
                            "置頂": "TRUE" if n_pin else "FALSE",
                            "標籤": n_tag,
                            "標題": n_title.strip(),
                            "內文": n_content.strip(),
                            "相關連結": n_link.strip(),
                            "作者": current_user,
                            "更新時間": get_tw_now_str()
                        }])
                        df_notes = pd.concat([df_notes, new_note_row], ignore_index=True)
                        save_sheet_df("共用筆記", df_notes, NOTE_COLS)
                        st.success(f"✅ 已發布筆記「{n_title.strip()}」！")
                        st.rerun()

        st.divider()

        if df_notes.empty:
            st.info("目前還沒有任何共用筆記，點擊上方「➕ 新增一則共用筆記 / 攻略重點」開始整理吧！")
        else:
            nf1, nf2 = st.columns([2.5, 2.5])
            with nf1:
                filter_note_tags = st.multiselect("🏷️ 依分類標籤篩選：", NOTE_TAG_OPTIONS, default=NOTE_TAG_OPTIONS)
            with nf2:
                search_note_kw = st.text_input("🔍 搜尋筆記標題或內文關鍵字：", placeholder="輸入關鍵字篩選...")

            df_notes_disp = df_notes.copy()
            df_notes_disp["_is_pinned"] = df_notes_disp["置頂"].apply(lambda x: str(x).strip().upper() == "TRUE")
            df_notes_disp = df_notes_disp[df_notes_disp["標籤"].isin(filter_note_tags)]

            if search_note_kw.strip():
                kw = search_note_kw.strip().lower()
                mask_kw = (
                    df_notes_disp["標題"].str.lower().str.contains(kw, na=False)
                    | df_notes_disp["內文"].str.lower().str.contains(kw, na=False)
                )
                df_notes_disp = df_notes_disp[mask_kw]

            df_notes_disp = df_notes_disp.sort_values(by=["_is_pinned", "更新時間"], ascending=[False, False])

            if df_notes_disp.empty:
                st.info("找不到符合條件的筆記喔！")
            else:
                for orig_n_idx, nrow in df_notes_disp.iterrows():
                    is_pinned = nrow["_is_pinned"]
                    pin_badge = "📌 【置頂】" if is_pinned else ""
                    with st.container(border=True):
                        nh1, nh2 = st.columns([5.5, 1.5])
                        with nh1:
                            st.markdown(f"#### {pin_badge}`{nrow['標籤']}` {nrow['標題']}")
                            st.caption(f"✍️ 整理者：**{nrow['作者']}** ｜ 🕒 更新於：{nrow['更新時間']}")
                        with nh2:
                            if str(nrow["相關連結"]).strip().startswith("http"):
                                st.link_button("🎬 開啟相關連結", str(nrow["相關連結"]).strip(), use_container_width=True)

                        st.markdown(nrow["內文"])

                        with st.expander("✏️ 編輯 / 置頂 / 刪除此則筆記"):
                            with st.form(f"edit_note_form_{orig_n_idx}"):
                                ec_n1, ec_n2, ec_n3 = st.columns([1.8, 3.2, 1])
                                with ec_n1:
                                    cur_tag_idx = NOTE_TAG_OPTIONS.index(nrow["標籤"]) if nrow["標籤"] in NOTE_TAG_OPTIONS else 0
                                    ed_n_tag = st.selectbox("分類標籤", NOTE_TAG_OPTIONS, index=cur_tag_idx)
                                with ec_n2:
                                    ed_n_title = st.text_input("筆記標題", value=str(nrow["標題"]))
                                with ec_n3:
                                    st.write("")
                                    st.write("")
                                    ed_n_pin = st.checkbox("📌 置頂", value=is_pinned)

                                ed_n_content = st.text_area("筆記內文", value=str(nrow["內文"]), height=130)
                                ed_n_link = st.text_input("相關連結", value=str(nrow["相關連結"]))

                                btn_n1, btn_n2 = st.columns([2, 1])
                                with btn_n1:
                                    save_n_btn = st.form_submit_button("💾 儲存筆記修改", type="primary", use_container_width=True)
                                with btn_n2:
                                    del_n_btn = st.form_submit_button("🗑️ 刪除此筆記", type="secondary", use_container_width=True)

                                if save_n_btn:
                                    if not ed_n_title.strip() or not ed_n_content.strip():
                                        st.warning("⚠️ 標題與內文不能為空喔！")
                                    else:
                                        df_notes.at[orig_n_idx, "置頂"] = "TRUE" if ed_n_pin else "FALSE"
                                        df_notes.at[orig_n_idx, "標籤"] = ed_n_tag
                                        df_notes.at[orig_n_idx, "標題"] = ed_n_title.strip()
                                        df_notes.at[orig_n_idx, "內文"] = ed_n_content.strip()
                                        df_notes.at[orig_n_idx, "相關連結"] = ed_n_link.strip()
                                        df_notes.at[orig_n_idx, "作者"] = current_user
                                        df_notes.at[orig_n_idx, "更新時間"] = get_tw_now_str()
                                        save_sheet_df("共用筆記", df_notes, NOTE_COLS)
                                        st.success("✅ 筆記已更新！")
                                        st.rerun()

                                if del_n_btn:
                                    df_notes = df_notes.drop(orig_n_idx).reset_index(drop=True)
                                    save_sheet_df("共用筆記", df_notes, NOTE_COLS)
                                    st.success("✅ 已刪除此則筆記！")
                                    st.rerun()

    # ==========================================================
    # 頁面：➕ 新增連結
    # ==========================================================
    elif active_page == "add":
        with st.form("add_link_form", clear_on_submit=True):
            st.subheader("新增景點 / 餐廳 / 購物連結")
            st.caption("✨ 貼上網址或輸入店名後，系統會透過 Google API 自動定位座標、反查最近車站，並**自動判斷所屬商圈分區**！")
            raw_input = st.text_input("🔗 貼上網址（支援 Google Maps、IG、食記或一般網頁）*", placeholder="https://...")
            c_name = st.text_input("✏️ 自訂中文/日文名稱（選填；若貼 IG 或食記建議填寫店名，系統會自動搜尋座標與分區）", placeholder="例如：一蘭拉麵 道頓堀店")

            c1, c2 = st.columns([2, 3])
            with c1:
                chosen_cat = st.selectbox("🏷️ 主分類", CAT_OPTIONS)
            with c2:
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
                            with st.spinner("🔍 正在透過 Google API 解析地點座標、最近車站並自動判定商圈分區..."):
                                fetched_title, subway_result, auto_region = resolve_place_subway_and_region(
                                    extracted_url,
                                    c_name.strip(),
                                    note_text.strip()
                                )

                            new_row = pd.DataFrame([{
                                "網址": extracted_url,
                                "分類": chosen_cat,
                                "區域": auto_region,
                                "自訂名稱": c_name.strip(),
                                "自動抓取標題": fetched_title,
                                "最近地鐵站": subway_result,
                                "備註": note_text.strip(),
                                "新增者": current_user,
                                "新增時間": get_tw_now_str()
                            }])
                            df_updated = pd.concat([df_links, new_row], ignore_index=True)
                            save_data(df_updated)
                            st.success(f"✅ 新增成功！「{fetched_title}」｜自動分區：【{auto_region}】｜{subway_result}")
                            st.rerun()

    # ==========================================================
    # 頁面：✏️ 修改 / 刪除
    # ==========================================================
    elif active_page == "edit":
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
                    edit_region = st.selectbox("🗺️ 區域 / 商圈（若勾選下方重新反查將自動更新）", REGION_OPTIONS, index=reg_idx)

                edit_subway = st.text_input("🚇 最近車站（可手動填寫，或勾選下方重新反查）", value=str(target_row["最近地鐵站"]))
                recheck_subway = st.checkbox("🔄 儲存時重新反查最近車站與自動分區")
                edit_note = st.text_input("📝 備註", value=str(target_row["備註"]))

                save_edit = st.form_submit_button("💾 儲存修改", type="primary")
                if save_edit:
                    clean_url = edit_url.strip()
                    new_name = edit_name.strip()

                    if recheck_subway:
                        resolved_title, resolved_subway, auto_reg = resolve_place_subway_and_region(
                            clean_url,
                            new_name,
                            edit_note.strip()
                        )
                        df_links.at[target_idx, "自動抓取標題"] = new_name if new_name else resolved_title
                        final_subway = resolved_subway
                        final_region = auto_reg
                    else:
                        if new_name != str(target_row["自訂名稱"]).strip():
                            df_links.at[target_idx, "自動抓取標題"] = new_name if new_name else str(target_row["自動抓取標題"])
                        final_subway = edit_subway.strip()
                        final_region = edit_region

                    df_links.at[target_idx, "網址"] = clean_url
                    df_links.at[target_idx, "自訂名稱"] = new_name
                    df_links.at[target_idx, "分類"] = edit_cat
                    df_links.at[target_idx, "區域"] = final_region
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

    # ==========================================================
    # 頁面：💰 共同支出記帳＆自動分帳計算機
    # ==========================================================
    elif active_page == "expense":
        st.subheader("💰 共同支出記帳區 ＆ 自動分帳計算機")
        live_rate, rate_status, _ = fetch_live_jpy_to_twd()

        rate_bar1, rate_bar2 = st.columns([4, 1])
        with rate_bar1:
            st.caption(f"記錄誰先代墊了機票、住宿、門票或餐費 ｜ **{rate_status}：`1 JPY = {live_rate:.4f} TWD`**")
        with rate_bar2:
            if st.button("🔄 刷新即時匯率", key="refresh_rate_shared", use_container_width=True):
                fetch_live_jpy_to_twd.clear()
                st.rerun()

        df_users = load_sheet_df("使用者帳號", USER_COLS)
        registered_members = [u.strip() for u in df_users["暱稱"].tolist() if u.strip()]
        if current_user not in registered_members:
            registered_members.append(current_user)

        df_exp = load_sheet_df("共同支出", EXPENSE_COLS)

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
                    exp_note = st.text_input("備註（選填）", placeholder="例如：含飲料、刷卡")

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
                        save_sheet_df("共同支出", df_exp, EXPENSE_COLS)
                        st.success(f"✅ 已新增支出：{exp_title.strip()}（{curr_code} {int(exp_amt):,}）")
                        st.rerun()

        if df_exp.empty:
            st.info("目前還沒有任何共同支出紀錄，第一筆代墊由你開始記錄吧！")
        else:
            st.markdown("### 📜 所有共同支出明細")
            disp_exp = df_exp.copy()
            def calc_twd_preview(r):
                try:
                    a = float(r["金額"])
                    if str(r["幣別"]).strip() == "JPY":
                        return f"約 NT$ {round(a * live_rate):,}"
                    return f"NT$ {round(a):,}"
                except Exception:
                    return "-"
            disp_exp.insert(4, "約合台幣(即時匯率)", disp_exp.apply(calc_twd_preview, axis=1))
            st.dataframe(disp_exp, use_container_width=True, hide_index=True)

            with st.expander("🗑️ 刪除錯誤的帳目紀錄"):
                del_exp_opts = [
                    f"[{i}] {r['日期時間']}｜{r['項目名稱']}（{r['幣別']} {r['金額']}，付款：{r['付款人']}）"
                    for i, r in df_exp.iterrows()
                ]
                sel_del_exp = st.selectbox("選擇要刪除的帳目：", del_exp_opts)
                if st.button("確認刪除此筆帳目", type="secondary"):
                    del_i = int(sel_del_exp.split("]")[0].replace("[", ""))
                    df_exp = df_exp.drop(del_i).reset_index(drop=True)
                    save_sheet_df("共同支出", df_exp, EXPENSE_COLS)
                    st.success("✅ 已刪除該筆帳目！")
                    st.rerun()

            st.divider()

            st.markdown("### 🧮 自動分帳結算結果（誰該給誰多少錢）")
            settle_mode = st.radio(
                "選擇結算方式：",
                [
                    "💱 方式一：依即時匯率合併成【台幣 TWD】一次總結清（最推薦，不用換兩種幣轉帳）",
                    "💴 方式二：日幣（JPY）與台幣（TWD）分開各自結算"
                ],
                horizontal=True
            )

            if settle_mode.startswith("💱"):
                custom_settle_rate = st.number_input(
                    "結算採用匯率（預設已自動抓最新市場即時匯率，也可微調成實際換匯/刷卡匯率）：",
                    min_value=0.1000,
                    max_value=0.5000,
                    value=float(live_rate),
                    step=0.0005,
                    format="%.4f"
                )
                comb_sum_df, comb_transfers = calculate_combined_twd_settlements(df_exp, custom_settle_rate)
                if not comb_sum_df.empty:
                    st.dataframe(comb_sum_df, use_container_width=True, hide_index=True)
                    if not comb_transfers:
                        st.success("🎉 所有帳目已完全打平，無需互相轉帳！")
                    else:
                        st.markdown("**👉 最終台幣轉帳建議（最少轉帳次數，直接轉台幣結清即可）：**")
                        for debtor, creditor, amt in comb_transfers:
                            st.warning(f"💸 **{debtor}** ➔ 轉給 **{creditor}**：**NT$ {amt:,} TWD**")
            else:
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
                            st.markdown("**👉 日幣轉帳建議：**")
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
                            st.markdown("**👉 台幣轉帳建議：**")
                            for debtor, creditor, amt in twd_transfers:
                                st.warning(f"💸 **{debtor}** ➔ 轉給 **{creditor}**：**NT$ {amt:,} TWD**")

    # ==========================================================
    # 頁面：📊 匯出精美 Excel
    # ==========================================================
    elif active_page == "excel":
        df_notes = load_sheet_df("共用筆記", NOTE_COLS)
        st.write("點擊下方按鈕，系統會自動補查尚未解析的最近車站（地鐵/電鐵）、智慧判定商圈分區並依區域上色分類，產生包含「機票航班表」、「住宿飯店表」、「每日行程表」、「共用筆記整理」、「總表」與各分類分頁的 Excel 檔案。")
        if df_links.empty and df_itin.empty and df_hotels.empty and df_flights.empty and df_notes.empty:
            st.info("目前沒有資料可以匯出喔！")
        else:
            if st.button("🚀 開始統整並產生 Excel", type="primary"):
                excel_data = generate_styled_excel(df_links.copy(), df_itin.copy(), df_hotels.copy(), df_flights.copy(), df_notes.copy())
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

    priv_page = st.radio(
        "選擇私人功能：",
        ["🛍️ 私人購物 ＆ 代購清單", "🧳 個人行李打包清單"],
        horizontal=True,
        key="private_active_page_nav"
    )

    if priv_page.startswith("🛍️"):
        df_shop_all = load_sheet_df("私人購物清單", SHOPPING_COLS)
        my_shop = df_shop_all[df_shop_all["擁有者"].str.strip() == current_user].copy()

        live_rate, rate_status, _ = fetch_live_jpy_to_twd()

        rate_col, m1, m2, m3 = st.columns([1.4, 1, 1, 1])
        with rate_col:
            jpy_rate = st.number_input(
                "💱 即時匯率 (1 日幣 = ? 台幣)",
                min_value=0.1000,
                max_value=0.5000,
                value=float(live_rate),
                step=0.0005,
                format="%.4f",
                help="已自動抓取網路最新即時匯率，你也可以手動輸入自己換匯的匯率"
            )
            st.caption(f"{rate_status}")
            if st.button("🔄 更新即時匯率", key="refresh_rate_private"):
                fetch_live_jpy_to_twd.clear()
                st.rerun()

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
                        save_sheet_df("私人購物清單", df_shop_all, SHOPPING_COLS)
                        st.success(f"✅ 已加入「{s_item.strip()}」！")
                        st.rerun()

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
                save_sheet_df("私人購物清單", merged_shop, SHOPPING_COLS)
                st.success("✅ 私人購物清單已同步儲存！")
                st.rerun()

    else:
        df_pack_all = load_sheet_df("個人行李清單", PACKING_COLS)
        my_pack = df_pack_all[df_pack_all["擁有者"].str.strip() == current_user].copy()

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
                save_sheet_df("個人行李清單", df_pack_all, PACKING_COLS)
                st.success("🎉 已匯入 10 項日本旅遊必備物品！")
                st.rerun()
        else:
            total_items = len(my_pack)
            packed_items = sum(1 for v in my_pack["已打包"] if str(v).strip().upper() in ["TRUE", "1", "是", "✅"])
            pct = packed_items / total_items if total_items > 0 else 0
            st.progress(pct, text=f"🧳 行李打包進度：已完成 **{packed_items} / {total_items}** 項（{int(pct * 100)}%）")

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
                        save_sheet_df("個人行李清單", df_pack_all, PACKING_COLS)
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
                save_sheet_df("個人行李清單", merged_pack, PACKING_COLS)
                st.success("✅ 個人行李清單已同步儲存！")
                st.rerun()
