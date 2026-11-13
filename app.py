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

# 大阪地鐵 (Osaka Metro) 路線關鍵字
OSAKA_METRO_LINES = [
    "御堂筋線", "谷町線", "四つ橋線", "中央線",
    "千日前線", "堺筋線", "長堀鶴見緑地線", "今里筋線", "南港ポートタウン線"
]

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
# 3. Google 地圖反查座標與最近地鐵站邏輯
# ==========================================
def is_google_maps_url(url):
    return any(domain in url.lower() for domain in [
        'maps.app.goo.gl', 'goo.gl/maps', 'google.com/maps', 'maps.google.com'
    ])

def is_valid_kansai_coords(lat, lon):
    """確認座標位於關西圈合理範圍內，避免誤抓網頁雜訊數字"""
    return (33.0 <= lat <= 36.2) and (134.0 <= lon <= 137.0)

def extract_coords_from_text(text):
    """從網址或網頁原始碼中提取精準經緯度"""
    patterns = [
        r'!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)',
        r'@(-?\d+\.\d+),(-?\d+\.\d+)',
        r'\[null,null,(-?\d+\.\d+),(-?\d+\.\d+)\]',
        r'center=(-?\d+\.\d+)%2C(-?\d+\.\d+)',
        r'"latitude"\s*:\s*"?(-?\d+\.\d+)"?\s*,\s*"longitude"\s*:\s*"?(-?\d+\.\d+)"?'
    ]
    for pat in patterns:
        for m in re.finditer(pat, text):
            lat, lon = float(m.group(1)), float(m.group(2))
            if is_valid_kansai_coords(lat, lon):
                return lat, lon
    return None, None

def has_specific_location_clue(name, region="", note=""):
    """判斷是否已具備單一分店或具體位置線索"""
    combined = f"{name} {note}".lower()
    # 若備註明確寫出有很多分店，直接判定為未指定單一分店
    if any(k in note for k in ["很多分店", "多家分店", "分店很多", "連鎖"]):
        return False
    # 若已選擇具體區域（非「其他/未指定」），代表已縮小到該商圈分店
    if region and region != "其他/未指定":
        return True
    # 若名稱或備註包含分店、本店、百貨、地標字眼
    if any(ind.lower() in combined for ind in BRANCH_INDICATORS):
        return True
    return False

def is_ambiguous_chain(name, region="", note=""):
    """判斷是否為多家分店且完全未指定區域或分店"""
    if any(k in note for k in ["很多分店", "多家分店", "分店很多"]):
        return True
    if has_specific_location_clue(name, region, note):
        return False
    return any(chain.lower() in name.lower() for chain in COMMON_CHAINS)

def resolve_google_maps_url(url):
    """解析直接貼上的 Google 地圖連結，回傳 (店家名稱, 緯度, 經度)"""
    title = ""
    lat, lon = None, None
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
            'Accept-Language': 'zh-TW,zh;q=0.9,ja;q=0.8'
        }
        res = requests.get(url, headers=headers, timeout=5, allow_redirects=True)
        resolved_url = res.url

        place_match = re.search(r'/place/([^/]+)', resolved_url)
        if place_match:
            title = urllib.parse.unquote(place_match.group(1)).replace('+', ' ')
        else:
            soup = BeautifulSoup(res.text, 'html.parser')
            if soup.title and soup.title.string:
                t = soup.title.string.strip()
                for suffix in ['- Google Maps', '- Google 地圖', 'Google Maps', 'Google 地圖']:
                    t = t.replace(suffix, '')
                title = t.strip()

        lat, lon = extract_coords_from_text(resolved_url)
        if lat is None:
            lat, lon = extract_coords_from_text(res.text)
    except Exception:
        pass
    return title, lat, lon

def scrape_webpage_title_and_coords(url):
    """從一般網頁（食記、IG、官網）抓取標題與內嵌地圖座標"""
    title = ""
    lat, lon = None, None
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36'
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

def search_google_maps_by_clues(place_name, region="", note=""):
    """
    利用「店家名稱 + 區域 + 備註線索」直接向 Google 地圖反查精準座標。
    """
    if not place_name or place_name in ["未知地點 (需手動確認)", "YouTube 推薦影片"]:
        return None, None

    # 組合搜尋關鍵字
    region_hint = ""
    if region and region != "其他/未指定":
        region_hint = region.split("/")[0]  # 例如 "難波/心齋橋" -> "難波"

    # 從備註提取地點/百貨線索（排除純餐點描述）
    note_hint = ""
    if note and not any(k in note for k in ["很多分店", "多家分店"]):
        note_hint = note.strip()[:20]

    queries_to_try = []
    base_prefix = "" if any(k in place_name for k in ["大阪", "京都", "神戶", "奈良"]) else "大阪 "
    if region_hint or note_hint:
        queries_to_try.append(f"{base_prefix}{region_hint} {place_name} {note_hint}".strip())
    queries_to_try.append(f"{base_prefix}{place_name}".strip())

    # 若備註提到明確建築物（如 OPA百貨 五樓），加入建築物備援搜尋
    if note_hint and any(k in note_hint for k in ["百貨", "OPA", "PARCO", "大丸", "高島屋", "阪急", "阪神", "LUCUA", "HEP", "MIO", "Q's"]):
        queries_to_try.append(f"大阪 {region_hint} {note_hint}".strip())

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
        'Accept-Language': 'zh-TW,zh;q=0.9,ja;q=0.8'
    }

    # 1. 優先向 Google Maps 搜尋端點抓取座標
    for q in queries_to_try:
        try:
            gmap_search_url = f"https://www.google.com/maps/search/{urllib.parse.quote(q)}?hl=zh-TW"
            res = requests.get(gmap_search_url, headers=headers, timeout=5, allow_redirects=True)
            lat, lon = extract_coords_from_text(res.url)
            if lat is None:
                lat, lon = extract_coords_from_text(res.text)
            if lat is not None and lon is not None:
                return lat, lon
        except Exception:
            continue

    # 2. 備援：透過 OpenStreetMap Nominatim 查詢
    for q in queries_to_try:
        try:
            res = requests.get(
                "https://nominatim.openstreetmap.org/search",
                params={"q": q, "format": "json", "limit": 1},
                headers={"User-Agent": "OsakaTripPlannerApp/2.0"},
                timeout=3
            )
            data = res.json()
            if data:
                lat, lon = float(data[0]["lat"]), float(data[0]["lon"])
                if is_valid_kansai_coords(lat, lon):
                    return lat, lon
        except Exception:
            continue

    return None, None

def format_walking_time(dist_str):
    """將車站直線距離換算為預估實際步行分鐘數（日本標準 80m/分 × 1.25 街區係數）"""
    try:
        num_match = re.search(r'(\d+(?:\.\d+)?)', str(dist_str))
        if not num_match:
            return f"約 {dist_str}"
        val = float(num_match.group(1))
        meters = val * 1000 if "km" in str(dist_str).lower() else val
        walk_mins = max(1, math.ceil((meters * 1.25) / 80))
        return f"步行約 {walk_mins} 分鐘 / {dist_str}"
    except Exception:
        return f"約 {dist_str}"

def query_heartrails_metro(lat, lon):
    """根據經緯度查詢最近的大阪地鐵站名、路線與步行時間"""
    try:
        res = requests.get(
            "https://express.heartrails.com/api/json",
            params={"method": "getStations", "x": lon, "y": lat},
            timeout=4
        )
        stations = res.json().get("response", {}).get("station", [])
        if not stations:
            return "查無鄰近車站"

        # 篩選大阪地鐵 (Osaka Metro)
        metro_stations = [
            s for s in stations
            if any(line_kw in s.get("line", "") for line_kw in OSAKA_METRO_LINES)
            and ("大阪" in s.get("line", "") or "Osaka" in s.get("line", ""))
        ]

        if metro_stations:
            nearest = metro_stations[0]
            st_name = nearest.get("name", "")
            walk_text = format_walking_time(nearest.get("distance", ""))

            same_st_lines = []
            for s in metro_stations:
                if s.get("name") == st_name:
                    clean_line = s.get("line", "").replace("大阪", "Osaka Metro ")
                    if clean_line not in same_st_lines:
                        same_st_lines.append(clean_line)
            lines_str = " / ".join(same_st_lines)
            return f"🚇 {st_name}站（{lines_str}，{walk_text}）"

        # 若附近無大阪地鐵（如環球影城、豐中、京都、奈良），顯示最近的一般車站
        fallback = stations[0]
        walk_text = format_walking_time(fallback.get("distance", ""))
        return f"🚃 {fallback.get('name', '')}站（{fallback.get('line', '')}・非地鐵，{walk_text}）"
    except Exception:
        return "地鐵查詢逾時（可稍後補抓）"

def resolve_place_and_subway(url, custom_name="", region="", note=""):
    """
    完整兩段式解析：
    1. 先取得店家名稱與 Google 地圖座標（支援直接地圖連結、網頁內嵌座標、或店名+區域+備註反查 Google 地圖）
    2. 確認非未指定分店的連鎖店後，計算最近大阪地鐵站、路線與步行分鐘數
    """
    final_title = custom_name.strip()
    lat, lon = None, None

    # 1. 若為 Google 地圖網址，直接解析
    if is_google_maps_url(url):
        gmap_title, g_lat, g_lon = resolve_google_maps_url(url)
        if not final_title:
            final_title = gmap_title if gmap_title else "未知地點 (需手動確認)"
        lat, lon = g_lat, g_lon
    else:
        # 若為一般網頁/社群連結，先抓標題與網頁內嵌座標
        web_title, w_lat, w_lon = scrape_webpage_title_and_coords(url)
        if not final_title:
            final_title = web_title if web_title else ("YouTube 推薦影片" if ('youtube.com' in url or 'youtu.be' in url) else "未知地點 (需手動確認)")
        lat, lon = w_lat, w_lon

    # 2. 檢查是否為「未指定區域與分店的多家分店連鎖店」
    if lat is None and is_ambiguous_chain(final_title, region, note):
        return final_title, "⏭️ 疑有多家分店（未指定分店或區域，不自動匹配地鐵）"

    # 3. 若網址本身沒帶座標，利用「店名 + 區域 + 備註」向 Google 地圖反查座標
    if lat is None or lon is None:
        lat, lon = search_google_maps_by_clues(final_title, region, note)

    if lat is not None and lon is not None:
        subway_info = query_heartrails_metro(lat, lon)
    else:
        subway_info = "查無單一店面座標（建議補貼 Google 地圖連結）"

    return final_title, subway_info

# ==========================================
# 4. 產生美化 Excel 報表 (記憶體串流版)
# ==========================================
def generate_styled_excel(df_links):
    all_list = []
    food_list, shop_list, spot_list, other_list = [], [], [], []
    sheet_updated = False

    progress_bar = st.progress(0, text="🔍 正在反查 Google 地圖座標、最近地鐵站與步行時間...")
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

        need_retry = (not subway_info) or subway_info.startswith(("查無", "未取得", "地鐵查詢逾時", "非地圖連結"))
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
            "最近地鐵站": subway_info,
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
            "最近地鐵站": subway_info,
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
                adjusted_width = min(max(max_len + 4, 12), 48)
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
        # 找出尚未查詢或先前查無座標的項目，提供一鍵用新版 Google 地圖引擎重查
        def needs_subway_lookup(val):
            s = str(val).strip()
            return (not s) or s.startswith(("查無", "未取得", "地鐵查詢逾時", "非地圖連結", "⏭️"))

        retry_mask = df_links["最近地鐵站"].apply(needs_subway_lookup)
        retry_count = int(retry_mask.sum())
        if retry_count > 0:
            if st.button(f"🚇 一鍵為 {retry_count} 筆資料反查 Google 地圖與最近地鐵站", type="secondary"):
                prog = st.progress(0, text="🚇 正在透過 Google 地圖反查店家位置與步行分鐘數...")
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
                st.success("✅ 已完成 Google 地圖位置反查與地鐵站更新！")
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
        filtered_df["最近地鐵站"] = filtered_df["最近地鐵站"].apply(
            lambda x: x if str(x).strip() else "(尚未查詢)"
        )

        display_df = filtered_df[["分類", "區域", "顯示名稱", "最近地鐵站", "備註", "網址", "新增者"]]
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
        c_name = st.text_input("✏️ 自訂中文名稱（選填；系統會結合區域與備註自動反查 Google 地圖位置）", placeholder="例如：一蘭拉麵 道頓堀店")

        c1, c2 = st.columns(2)
        with c1:
            chosen_cat = st.selectbox("🏷️ 主分類", CAT_OPTIONS)
        with c2:
            chosen_region = st.selectbox("🗺️ 區域 / 商圈", REGION_OPTIONS)

        note_text = st.text_input("📝 備註（選填；填入百貨或建築物名稱如「OPA百貨 五樓」可大幅提升定位準確度）", placeholder="例如：OPA百貨 五樓、必點牛舌")
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
                        with st.spinner("🔍 正在反查 Google 地圖位置並計算最近地鐵站與步行時間..."):
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

            edit_subway = st.text_input("🚇 最近地鐵站（可手動填寫，或勾選下方重新反查）", value=str(target_row["最近地鐵站"]))
            recheck_subway = st.checkbox("🔄 儲存時重新反查 Google 地圖與最近大阪地鐵站")
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
    st.write("點擊下方按鈕，系統會自動補查尚未解析的 Google 地圖位置與最近地鐵站、依區域上色分類，並產生包含「總表」與各分類分頁的 Excel 檔案。")
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
