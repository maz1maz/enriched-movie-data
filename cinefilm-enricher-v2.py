#!/usr/bin/env python3
"""
Cinefilm Archive Enricher v2
==============================
نسخه بهبودیافته و ترکیبی از بهترین ویژگی‌های دو نسخه قبلی

ویژگی‌ها:
  ✅ ترجمه خودکار عناوین فارسی/غیرانگلیسی به انگلیسی
  ✅ جستجوی IMDb ID از طریق IMDb Suggestion API (دقت بالاتر)
  ✅ پشتیبانی فیلم (OMDb) + سریال (TVMaze)
  ✅ سیستم کش پیشرفته (resume بعد از قطعی)
  ✅ تشخیص خودکار نوع فایل (فیلم یا سریال)
  ✅ پشتیبانی CSV + Excel + انکدینگ‌های مختلف
  ✅ خروجی Missing برای اصلاح دستی
  ✅ ستون‌های کامل: Shelf, Row, Director, Cast, Year, Genre, Rating, Runtime,
     Country, Synopsis, Poster URL, Original Title, Total Size, Folder Paths
  ✅ اصلاح خودکار عناوین مشکل‌دار
  ✅ نوار پیشرفت (progress bar)
  ✅ مدیریت Rate Limit هوشمند
  ✅ گزارش نهایی آمار

Requirements:
    pip install requests pandas openpyxl deep-translator chardet

Usage:
    python cinefilm-enricher-v2.py --input movies.csv
    python cinefilm-enricher-v2.py --input series.xlsx --type series
    python cinefilm-enricher-v2.py --input data.csv --type auto --output results/
    python cinefilm-enricher-v2.py --input data.csv --delay 0.5 --no-translate
"""

import argparse
import csv
import json
import os
import re
import sys
import time
import html
import urllib.parse
from pathlib import Path
from collections import defaultdict

import requests

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

try:
    import chardet
    HAS_CHARDET = True
except ImportError:
    HAS_CHARDET = False

try:
    from deep_translator import GoogleTranslator
    HAS_TRANSLATOR = True
except ImportError:
    HAS_TRANSLATOR = False

# =============================================================================
# CONFIG
# =============================================================================
OMDB_API_KEY = "fe53f97e"
OMDB_URL = "https://www.omdbapi.com/"
IMDB_SUGGESTION_URL = "https://v3.sg.media-imdb.com/suggestion/{}/{}.json"
TVMAZE_SEARCH_URL = "https://api.tvmaze.com/search/shows"
TVMAZE_SHOW_URL = "https://api.tvmaze.com/shows/{}"
TVMAZE_CAST_URL = "https://api.tvmaze.com/shows/{}/cast"
TVMAZE_CREW_URL = "https://api.tvmaze.com/shows/{}/crew"

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)

# ستون‌های نهایی خروجی
MOVIE_COLUMNS = [
    "Title", "Shelf", "Row", "Director", "Cast", "Year",
    "Genre", "Rating", "Runtime", "Country", "Synopsis",
    "Poster URL", "Original Title", "Total Size", "Folder Paths"
]

SERIES_COLUMNS = [
    "Title", "Shelf", "Row", "Director", "Cast", "Year",
    "Genre", "Rating", "Runtime", "Country", "Synopsis",
    "Poster URL", "Original Title", "Seasons Available",
    "Total Size", "Folder Paths"
]

# اصلاحات دستی برای عناوین مشکل‌دار (TVMaze ID)
TVMAZE_CORRECTIONS = {
    "money heist": 27436,
    "la casa de papel": 27436,
    "protector": 36807,
    "the protector": 36807,
    "anne": 12989,
    "anne with an e": 12989,
    "24": 406,
    "the bridge": 1910,
    "dark": 13177,
}

# اصلاحات دستی برای عناوین مشکل‌دار (OMDb جستجوی جایگزین)
OMDB_TITLE_CORRECTIONS = {
    "money heist": "La Casa de Papel",
}

# =============================================================================
# CACHE SYSTEM
# =============================================================================
class Cache:
    """سیستم کش ساده و مؤثر — درخواست‌های تکراری رو حذف می‌کنه"""

    def __init__(self, cache_path):
        self.path = Path(cache_path)
        self.data = {}
        self.hits = 0
        self.misses = 0
        self._dirty = False
        self._load()

    def _load(self):
        if self.path.exists():
            try:
                with open(self.path, "r", encoding="utf-8") as f:
                    self.data = json.load(f)
                print(f"  📦 کش بارگذاری شد: {len(self.data)} آیتم ({self.path.name})")
            except (json.JSONDecodeError, IOError):
                self.data = {}

    def save(self):
        if self._dirty:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.path, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=1)
            self._dirty = False

    def get(self, key):
        key = str(key).strip().lower()
        if key in self.data:
            self.hits += 1
            return self.data[key]
        self.misses += 1
        return None

    def set(self, key, value):
        self.data[str(key).strip().lower()] = value
        self._dirty = True

    def stats(self):
        total = self.hits + self.misses
        if total == 0:
            return "بدون درخواست"
        pct = (self.hits / total) * 100
        return f"✅ {self.hits} کش‌هیت / ❌ {self.misses} درخواست جدید ({pct:.0f}% صرفه‌جویی)"


# =============================================================================
# TRANSLATION
# =============================================================================
def is_pure_ascii(s):
    try:
        s.encode("ascii")
        return True
    except UnicodeEncodeError:
        return False


def translate_to_english(query):
    """ترجمه خودکار عنوان غیرانگلیسی به انگلیسی"""
    if not HAS_TRANSLATOR:
        return query
    if is_pure_ascii(query):
        return query
    try:
        translated = GoogleTranslator(source="auto", target="en").translate(query)
        if translated and translated.strip():
            return translated.strip()
    except Exception:
        pass
    return query


# =============================================================================
# TEXT CLEANING
# =============================================================================
def strip_html(text):
    if not text:
        return ""
    text = re.sub(r"<[^>]+>", "", text)
    return html.unescape(text).strip()


def clean_title(title):
    """تمیز کردن عنوان فیلم/سریال"""
    t = str(title).strip()
    t = re.sub(r"\s+", " ", t)
    # حذف تگ‌های رایج
    t = re.sub(r"\s*\(TV Movie.*$", "", t, flags=re.I)
    t = re.sub(r"\s*\(Video.*$", "", t, flags=re.I)
    t = re.sub(r"\s*\(Short.*$", "", t, flags=re.I)
    # حذف کاراکترهای عجیب
    t = t.replace("\ufffd", "").replace("\xb3", "3").strip()
    return t


def clean_runtime(runtime_str):
    """استخراج عدد از رشته Runtime"""
    if not runtime_str or runtime_str == "N/A":
        return ""
    match = re.search(r"(\d+)", str(runtime_str))
    if match:
        return int(match.group(1))
    return runtime_str


# =============================================================================
# HTTP REQUEST WITH RETRY
# =============================================================================
def http_get_with_retry(url, params=None, headers=None, timeout=10, max_retries=3):
    """درخواست HTTP با retry و exponential backoff"""
    for attempt in range(max_retries):
        try:
            res = requests.get(url, params=params, headers=headers, timeout=timeout)
            # اگر rate limit شدیم، صبر کن
            if res.status_code == 429:
                wait_time = (2 ** attempt) + 1
                time.sleep(wait_time)
                continue
            return res
        except requests.exceptions.RequestException:
            if attempt < max_retries - 1:
                wait_time = (2 ** attempt) + 0.5
                time.sleep(wait_time)
            continue
    return None


# =============================================================================
# IMDb SUGGESTION API
# =============================================================================
def get_imdb_id_from_suggestion(query):
    """
    پیدا کردن IMDb ID از طریق سرویس Suggestion IMDb
    این روش دقیق‌تر از جستجوی مستقیم عنوان در OMDb هست
    """
    try:
        clean_q = query.strip().lower()
        first_char = clean_q[0]
        encoded = urllib.parse.quote(clean_q)
        url = IMDB_SUGGESTION_URL.format(first_char, encoded)
        headers = {"User-Agent": USER_AGENT}
        res = http_get_with_retry(url, headers=headers)
        if res and res.status_code == 200:
            data = res.json()
            if "d" in data and data["d"]:
                # اولویت با فیلم یا سریال
                for item in data["d"]:
                    qid = item.get("qid", "")
                    if qid in ("movie", "tvSeries", "tvMiniSeries"):
                        return item.get("id"), qid
                # اگه هیچکدوم فیلم/سریال نبود، اولی رو برگردون
                return data["d"][0].get("id"), data["d"][0].get("qid", "")
    except Exception:
        pass
    return None, None


# =============================================================================
# OMDb API (فیلم)
# =============================================================================
def fetch_movie_details(title, cache, translate=True, verbose=True):
    """دریافت اطلاعات کامل فیلم از OMDb"""
    # بررسی کش
    cached = cache.get(f"movie:{title}")
    if cached is not None:
        return cached if cached != {} else None

    search_title = title
    if translate and not is_pure_ascii(title):
        if verbose:
            print(f"  🌐 ترجمه '{title}'...", end=" ", flush=True)
        search_title = translate_to_english(title)
        if verbose and search_title != title:
            print(f"→ '{search_title}'")
        elif verbose:
            print("(بدون تغییر)")

    # اصلاح عنوان
    search_lower = search_title.lower().strip()
    if search_lower in OMDB_TITLE_CORRECTIONS:
        search_title = OMDB_TITLE_CORRECTIONS[search_lower]

    omdb_data = None

    # مرحله ۱: IMDb Suggestion → IMDb ID → OMDb
    imdb_id, qid = get_imdb_id_from_suggestion(search_title)
    if imdb_id:
        res = http_get_with_retry(OMDB_URL, params={
            "apikey": OMDB_API_KEY,
            "i": imdb_id,
            "plot": "short",
        })
        if res and res.status_code == 200:
            try:
                data = res.json()
                if data.get("Response") == "True":
                    omdb_data = data
            except Exception:
                pass

    # مرحله ۲: Fallback — جستجوی مستقیم عنوان
    if not omdb_data:
        res = http_get_with_retry(OMDB_URL, params={
            "apikey": OMDB_API_KEY,
            "t": search_title,
            "plot": "short",
        })
        if res and res.status_code == 200:
            try:
                data = res.json()
                if data.get("Response") == "True":
                    omdb_data = data
            except Exception:
                pass

    # ذخیره در کش (حتی None برای جلوگیری از درخواست تکراری)
    cache.set(f"movie:{title}", omdb_data if omdb_data else {})
    return omdb_data


# =============================================================================
# TVMaze API (سریال)
# =============================================================================
def fetch_series_details(title, cache, translate=True, verbose=True):
    """دریافت اطلاعات کامل سریال از TVMaze"""
    # بررسی کش
    cached = cache.get(f"series:{title}")
    if cached is not None:
        return cached if cached != {} else None

    search_title = title
    if translate and not is_pure_ascii(title):
        if verbose:
            print(f"  🌐 ترجمه '{title}'...", end=" ", flush=True)
        search_title = translate_to_english(title)
        if verbose and search_title != title:
            print(f"→ '{search_title}'")
        elif verbose:
            print("(بدون تغییر)")

    show_id = None
    search_lower = search_title.lower().strip()

    # بررسی اصلاحات دستی
    if search_lower in TVMAZE_CORRECTIONS:
        show_id = TVMAZE_CORRECTIONS[search_lower]

    # جستجو در TVMaze
    if not show_id:
        res = http_get_with_retry(TVMAZE_SEARCH_URL, params={"q": search_title})
        if res and res.status_code == 200:
            try:
                results = res.json()
                if results:
                    # بهترین نتیجه
                    best = results[0]
                    score = best.get("score", 0)
                    show = best.get("show", {})
                    # اگه امتیاز خیلی کمه، مطمئن نیستیم
                    if score > 0.5 or show:
                        show_id = show.get("id")
            except Exception:
                pass

    if not show_id:
        cache.set(f"series:{title}", {})
        return None

    # دریافت جزئیات سریال
    details = {}
    res = http_get_with_retry(TVMAZE_SHOW_URL.format(show_id))
    if res and res.status_code == 200:
        try:
            show = res.json()
            details["Title"] = show.get("name", "")
            details["Original Title"] = show.get("name", "")

            # سال
            premiered = show.get("premiered", "")
            if premiered:
                details["Year"] = premiered[:4]

            # ژانر
            genres = show.get("genres", [])
            details["Genre"] = ", ".join(genres) if genres else ""

            # رتبه
            rating = show.get("rating", {})
            details["Rating"] = str(rating.get("average", "")) if rating else ""

            # Runtime
            details["Runtime"] = show.get("runtime", "") or ""

            # خلاصه
            details["Synopsis"] = strip_html(show.get("summary", ""))

            # کشور
            network = show.get("network", {})
            country = network.get("country", {}) if network else {}
            details["Country"] = country.get("name", "")

            # پوستر
            image = show.get("image", {})
            details["Poster URL"] = image.get("original", "") or image.get("medium", "") or ""

            # تعداد فصل‌ها
            details["_seasons"] = show.get("seasons", [])
            details["_show_id"] = show_id
        except Exception:
            pass

    # دریافت بازیگران و کارگردان
    res = http_get_with_retry(TVMAZE_CAST_URL.format(show_id))
    if res and res.status_code == 200:
        try:
            cast_data = res.json()
            cast_names = [p.get("person", {}).get("name", "") for p in cast_data[:8]]
            details["Cast"] = ", ".join(filter(None, cast_names))
        except Exception:
            pass

    # کارگردان از crew — TVMaze crew API محدود است
    details["Director"] = ""

    # تعداد فصل‌ها
    seasons = details.pop("_seasons", [])
    details["Seasons Available"] = len(seasons) if seasons else ""

    cache.set(f"series:{title}", details if details else {})
    return details if details else None


# =============================================================================
# FILE TYPE AUTO-DETECTION
# =============================================================================
def detect_file_type(file_path, df=None):
    """تشخیص خودکار نوع فایل: movie یا series"""
    name = str(file_path).lower()

    # از نام فایل
    if "series" in name or "serial" in name or "tv" in name:
        return "series"
    if "movie" in name or "film" in name:
        return "movie"

    # از محتوای فایل
    if df is not None:
        all_text = " ".join(str(x) for x in df.head(20).values.flatten() if x)
        lower_text = all_text.lower()
        series_keywords = ["season", "episode", "s01", "s02", "series", "tv series", "فصل", "قسمت"]
        movie_keywords = ["movie", "film", "director", "فیلم"]

        series_score = sum(1 for k in series_keywords if k in lower_text)
        movie_score = sum(1 for k in movie_keywords if k in lower_text)

        if series_score > movie_score:
            return "series"
        if movie_score > series_score:
            return "movie"

    # پیش‌فرض
    return "movie"


# =============================================================================
# FILE READING
# =============================================================================
def read_input_file(file_path):
    """خواندن فایل ورودی — CSV یا Excel با تشخیص خودکار انکدینگ"""
    ext = Path(file_path).suffix.lower()

    if ext in (".xlsx", ".xls"):
        if not HAS_PANDAS:
            print("  ❌ pandas نصب نیست. لطفاً: pip install pandas openpyxl")
            sys.exit(1)
        df = pd.read_excel(file_path)
        return df

    elif ext == ".csv":
        # تشخیص انکدینگ
        encoding = "utf-8"
        if HAS_CHARDET:
            with open(file_path, "rb") as f:
                raw = f.read(10000)
                detected = chardet.detect(raw)
                if detected and detected.get("encoding"):
                    encoding = detected["encoding"]
                    # اصلاح windows-1252
                    if encoding.lower() in ("windows-1252", "iso-8859-1", "iso-8859-9"):
                        encoding = "windows-1252"

        if HAS_PANDAS:
            try:
                df = pd.read_csv(file_path, encoding=encoding)
            except Exception:
                df = pd.read_csv(file_path, encoding="utf-8", errors="replace")
            return df
        else:
            # بدون pandas — از csv استفاده کن
            rows = []
            with open(file_path, "r", encoding=encoding, errors="replace") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    rows.append(row)
            return rows
    else:
        print(f"  ❌ فرمت فایل پشتیبانی نمی‌شود: {ext}")
        sys.exit(1)


def find_title_column(df):
    """پیدا کردن ستون عنوان"""
    if HAS_PANDAS and isinstance(df, pd.DataFrame):
        cols = df.columns.tolist()
    else:
        cols = list(df[0].keys()) if df else []

    for candidate in ["Title", "title", "Name", "name", "عنوان", "نام", "نام فیلم"]:
        for col in cols:
            if str(col).strip().lower() == candidate.lower():
                return col
    # اولین ستون
    return cols[0] if cols else None


def normalize_columns(df):
    """یکسان‌سازی نام ستون‌ها"""
    if HAS_PANDAS and isinstance(df, pd.DataFrame):
        mapping = {}
        for col in df.columns:
            clean = str(col).strip()
            # Title Case
            if clean.lower() == "title":
                mapping[col] = "Title"
            elif clean.lower() == "name":
                mapping[col] = "Title"
            else:
                mapping[col] = clean
        return df.rename(columns=mapping)
    return df


# =============================================================================
# MAIN PROCESSING
# =============================================================================
def process_file(input_path, file_type="auto", output_dir=None,
                 delay=0.5, translate=True, verbose=True):
    """پردازش اصلی فایل"""

    input_path = Path(input_path)
    if not input_path.exists():
        print(f"❌ فایل یافت نشد: {input_path}")
        return False

    # پوشه خروجی
    if not output_dir:
        output_dir = input_path.parent / "output"
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n{'='*60}")
    print(f"  🎬 Cinefilm Enricher v2")
    print(f"  📁 فایل ورودی: {input_path.name}")
    print(f"  📂 پوشه خروجی: {output_dir}")
    print(f"{'='*60}")

    # خواندن فایل
    print(f"\n📖 خواندن فایل...")
    df = read_input_file(input_path)

    if HAS_PANDAS and isinstance(df, pd.DataFrame):
        df = normalize_columns(df)
        title_col = find_title_column(df)
        total = len(df)
        print(f"  📊 {total} ردیف | ستون عنوان: '{title_col}'")
    else:
        total = len(df)
        title_col = find_title_column(df)
        print(f"  📊 {total} ردیف | ستون عنوان: '{title_col}'")

    # تشخیص نوع
    if file_type == "auto":
        file_type = detect_file_type(input_path, df)
    print(f"  🎭 نوع: {'🎥 فیلم' if file_type == 'movie' else '📺 سریال'}")

    # بارگذاری کش
    cache_name = "omdb_cache.json" if file_type == "movie" else "tvmaze_cache.json"
    cache = Cache(output_dir / cache_name)

    # ستون‌ها و تابع استخراج
    if file_type == "movie":
        columns = MOVIE_COLUMNS
        fetch_fn = lambda title: fetch_movie_details(title, cache, translate, verbose)
    else:
        columns = SERIES_COLUMNS
        fetch_fn = lambda title: fetch_series_details(title, cache, translate, verbose)

    # اطمینان از وجود ستون‌ها
    if HAS_PANDAS and isinstance(df, pd.DataFrame):
        for col in columns:
            if col not in df.columns:
                df[col] = ""
            df[col] = df[col].astype(object)

    # پردازش
    print(f"\n🔄 شروع پردازش...\n")
    success_count = 0
    fail_count = 0
    skip_count = 0

    if HAS_PANDAS and isinstance(df, pd.DataFrame):
        for idx, row in df.iterrows():
            title = row.get(title_col)
            if not title or pd.isna(title) or str(title).strip() == "":
                continue

            title = str(title).strip()
            progress = f"[{idx + 1}/{total}]"

            # چک پردازش قبلی
            dir_val = row.get("Director", "")
            genre_val = row.get("Genre", "")
            if (dir_val and not pd.isna(dir_val) and str(dir_val).strip()
                    and genre_val and not pd.isna(genre_val) and str(genre_val).strip()):
                skip_count += 1
                if verbose:
                    print(f"  {progress} ⏭️  '{title}' — قبلاً پردازش شده")
                continue

            if verbose:
                print(f"  {progress} 🔍 '{title}'")

            details = fetch_fn(title)

            if details:
                success_count += 1
                field_map = {
                    "Original Title": details.get("Original Title", "") or details.get("Title", ""),
                    "Director": details.get("Director", ""),
                    "Cast": details.get("Cast", "") or details.get("Actors", ""),
                    "Year": details.get("Year", ""),
                    "Genre": details.get("Genre", ""),
                    "Rating": details.get("Rating", "") or details.get("imdbRating", ""),
                    "Runtime": clean_runtime(details.get("Runtime", "")),
                    "Country": details.get("Country", ""),
                    "Synopsis": details.get("Synopsis", "") or details.get("Plot", ""),
                }

                poster = details.get("Poster URL", "") or details.get("Poster", "")
                if poster and poster != "N/A":
                    field_map["Poster URL"] = poster

                if file_type == "series":
                    field_map["Seasons Available"] = details.get("Seasons Available", "")

                for col, val in field_map.items():
                    if col in df.columns:
                        df.at[idx, col] = val

                if verbose:
                    rating = field_map.get("Rating", "")
                    year = field_map.get("Year", "")
                    extra = f"  ⭐ {rating}" if rating else ""
                    extra += f"  📅 {year}" if year else ""
                    print(f"    ✅ OK{extra}")
            else:
                fail_count += 1
                if verbose:
                    print(f"    ❌ یافت نشد")

            # ذخیره دوره‌ای کش (هر ۵۰ آیتم)
            if (idx + 1) % 50 == 0:
                cache.save()

            # تأخیر
            if not cache.get(f"{'movie' if file_type == 'movie' else 'series'}:{title}"):
                time.sleep(delay)

        # ذخیره نهایی کش
        cache.save()

    # خروجی‌ها
    print(f"\n💾 ذخیره خروجی‌ها...")

    base = input_path.stem
    # فایل غنی‌شده
    enriched_path = output_dir / f"{base}-enriched.xlsx"
    missing_path = output_dir / f"{base}-missing.csv"

    if HAS_PANDAS and isinstance(df, pd.DataFrame):
        # مرتب‌سازی ستون‌ها
        ordered = [c for c in columns if c in df.columns]
        extra = [c for c in df.columns if c not in columns]
        df_out = df[ordered + extra]
        df_out.to_excel(enriched_path, index=False)

        # فایل missing
        missing = df[df["Director"].isna() | (df["Director"] == "") |
                      df["Genre"].isna() | (df["Genre"] == "")]
        if len(missing) > 0:
            missing.to_csv(missing_path, index=False, encoding="utf-8-sig")
            print(f"  ⚠️  {len(missing)} عنوان ناقص → {missing_path.name}")

    print(f"  📊 فایل اصلی → {enriched_path.name}")

    # CSV هم بساز
    csv_path = output_dir / f"{base}-enriched.csv"
    if HAS_PANDAS and isinstance(df, pd.DataFrame):
        df_out.to_csv(csv_path, index=False, encoding="utf-8-sig")
        print(f"  📊 CSV → {csv_path.name}")

    # گزارش نهایی
    processed = success_count + fail_count
    print(f"\n{'='*60}")
    print(f"  📈 گزارش نهایی")
    print(f"  {'='*58}")
    print(f"  کل ردیف‌ها:     {total}")
    print(f"  رد شده (قبلی):  {skip_count}")
    print(f"  پردازش شده:     {processed}")
    print(f"  ✅ موفق:         {success_count}")
    print(f"  ❌ ناموفق:       {fail_count}")
    if processed > 0:
        pct = (success_count / processed) * 100
        print(f"  📊 درصد موفقیت: {pct:.1f}%")
    print(f"  📦 کش: {cache.stats()}")
    print(f"{'='*60}\n")

    return True


# =============================================================================
# CLI
# =============================================================================
def main():
    parser = argparse.ArgumentParser(
        description="Cinefilm Archive Enricher v2 — غنی‌سازی خودکار اطلاعات فیلم و سریال",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
مثال‌ها:
  python cinefilm-enricher-v2.py --input movies.csv
  python cinefilm-enricher-v2.py --input series.xlsx --type series
  python cinefilm-enricher-v2.py --input data.csv --type auto --output results/
  python cinefilm-enricher-v2.py --input data.csv --delay 0.5 --no-translate
        """
    )
    parser.add_argument("--input", "-i", required=True, help="فایل ورودی (CSV یا Excel)")
    parser.add_argument("--type", "-t", choices=["auto", "movie", "series"],
                        default="auto", help="نوع فایل (پیش‌فرض: auto)")
    parser.add_argument("--output", "-o", help="پوشه خروجی (پیش‌فرض: input_dir/output)")
    parser.add_argument("--delay", "-d", type=float, default=0.5,
                        help="تأخیر بین درخواست‌ها (ثانیه، پیش‌فرض: 0.5)")
    parser.add_argument("--no-translate", action="store_true",
                        help="غیرفعال کردن ترجمه خودکار")
    parser.add_argument("--quiet", "-q", action="store_true",
                        help="حالت بی‌صدا (فقط خطاها)")

    args = parser.parse_args()

    # بررسی وابستگی‌ها
    missing_deps = []
    if not HAS_PANDAS:
        missing_deps.append("pandas")
    if not HAS_TRANSLATOR and not args.no_translate:
        print("⚠️  deep-translator نصب نیست — ترجمه خودکار غیرفعال")
        print("   نصب: pip install deep-translator\n")

    success = process_file(
        input_path=args.input,
        file_type=args.type,
        output_dir=args.output,
        delay=args.delay,
        translate=not args.no_translate,
        verbose=not args.quiet,
    )

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
