# 🎬 Cinefilm Archive Enricher v2

نسخه بهبودیافته و ترکیبی از بهترین ویژگی‌های دو نسخه قبلی برای غنی‌سازی خودکار آرشیو فیلم و سریال.

## ✨ ویژگی‌های جدید و بهبودیافته

### 🆕 ویژگی‌های اضافه شده از کد جدید
- ✅ **ترجمه خودکار فارسی/عربی به انگلیسی** — با `deep-translator` عناوین غیرانگلیسی رو ترجمه می‌کنه
- ✅ **جستجوی IMDb ID قبل از OMDb** — از IMDb Suggestion API استفاده می‌کنه که دقیق‌تره
- ✅ **API Key بهتر** — کلید `fe53f97e` به جای کلید عمومی `thewdb`
- ✅ **Fallback هوشمند** — اگه با IMDb ID نتیجه نگرفت، مستقیم با عنوان جستجو می‌کنه

### 🔄 ویژگی‌های حفظ شده از کد ریپازیتوری
- ✅ **پشتیبانی فیلم + سریال** — OMDb برای فیلم، TVMaze برای سریال
- ✅ **سیستم کش پیشرفته** — `omdb_cache.json` و `tvmaze_cache.json` برای resume بعد از قطعی
- ✅ **تشخیص خودکار نوع فایل** — movie یا series رو خودش تشخیص میده
- ✅ **پشتیبانی CSV + Excel** — هر دو فرمت ورودی
- ✅ **مدیریت انکدینگ** — تشخیص خودکار `windows-1252`، `UTF-8` و سایر انکدینگ‌ها
- ✅ **خروجی Missing** — لیست عناوین ناقص برای اصلاح دستی
- ✅ **ستون‌های کامل** — Shelf, Row, Director, Cast, Year, Genre, Rating, Runtime, Country, Synopsis, Poster URL, Original Title, Total Size, Folder Paths
- ✅ **اصلاح عناوین مشکل‌دار** — Money Heist → La Casa de Papel و مشابه‌ها

### 🚀 ویژگی‌های جدید v2
- ✅ **HTTP Retry با Exponential Backoff** — مدیریت خودکار خطاهای شبکه و Rate Limit
- ✅ **نوار پیشرفت (Progress)** — نمایش `[1/100]` برای هر عنوان
- ✅ **گزارش نهایی آمار** — درصد موفقیت، تعداد کش‌هیت، صرفه‌جویی
- ✅ **ذخیره دوره‌ای کش** — هر ۵۰ آیتم کش ذخیره میشه
- ✅ **Skip هوشمند** — ردیف‌هایی که Director + Genre پر باشه رد میشن
- ✅ **حالت Quiet** — `--quiet` برای کاهش لاگ
- ✅ **Delay قابل تنظیم** — `--delay 0.5` برای کنترل سرعت درخواست‌ها
- ✅ **غیرفعال کردن ترجمه** — `--no-translate` اگه نمی‌خوای ترجمه کنه

## 📦 نصب

```bash
pip install requests pandas openpyxl deep-translator chardet
```

یا:

```bash
pip install -r requirements-v2.txt
```

## 🎯 استفاده

### ساده‌ترین حالت
```bash
python cinefilm-enricher-v2.py --input movies.csv
```

### مشخص کردن نوع فایل
```bash
python cinefilm-enricher-v2.py --input series.xlsx --type series
```

### تشخیص خودکار (پیش‌فرض)
```bash
python cinefilm-enricher-v2.py --input data.csv --type auto
```

### پوشه خروجی سفارشی
```bash
python cinefilm-enricher-v2.py --input movies.csv --output results/
```

### تنظیم تأخیر بین درخواست‌ها
```bash
python cinefilm-enricher-v2.py --input movies.csv --delay 1.0
```

### بدون ترجمه (سریع‌تر)
```bash
python cinefilm-enricher-v2.py --input movies.csv --no-translate
```

### حالت بی‌صدا
```bash
python cinefilm-enricher-v2.py --input movies.csv --quiet
```

## 📊 خروجی‌ها

### فایل غنی‌شده
- `output/{filename}-enriched.xlsx` — فایل اصلی با اطلاعات کامل
- `output/{filename}-enriched.csv` — نسخه CSV

### فایل Missing
- `output/{filename}-missing.csv` — عناوینی که اطلاعاتشون ناقص مونده

### کش
- `output/omdb_cache.json` — کش فیلم‌ها
- `output/tvmaze_cache.json` — کش سریال‌ها

## 🔧 API ها

### فیلم (OMDb)
- **API**: OMDb API
- **Key**: `fe53f97e` (می‌تونی با کلید خودت عوض کنی)
- **Limit**: بستگی به نوع کلید داره

### سریال (TVMaze)
- **API**: TVMaze API
- **Key**: نیاز نیست (کاملاً رایگان)
- **Limit**: نامحدود

### ترجمه (Google Translate)
- **API**: Google Translate (از طریق deep-translator)
- **Limit**: نامحدود (ولی ممکنه rate limit بشه)

## 📈 آمار و گزارش

بعد از اجرا، گزارش نهایی شامل:
- کل ردیف‌ها
- رد شده (قبلاً پردازش شده)
- پردازش شده
- موفق / ناموفق
- درصد موفقیت
- آمار کش (تعداد کش‌هیت و صرفه‌جویی)

مثال:
```
============================================================
  📈 گزارش نهایی
  ==========================================================
  کل ردیف‌ها:     1000
  رد شده (قبلی):  150
  پردازش شده:     850
  ✅ موفق:         780
  ❌ ناموفق:       70
  📊 درصد موفقیت: 91.8%
  📦 کش: ✅ 150 کش‌هیت / ❌ 850 درخواست جدید (15% صرفه‌جویی)
============================================================
```

## 🎭 تشخیص خودکار نوع فایل

برنامه از این روش‌ها برای تشخیص movie/series استفاده می‌کنه:

1. **نام فایل**: `series`, `serial`, `tv` → سریال | `movie`, `film` → فیلم
2. **محتوای فایل**: کلمات کلیدی مثل `season`, `episode`, `s01` → سریال
3. **پیش‌فرض**: اگه هیچکدوم نبود → فیلم

## 🛠️ پارامترها

```
--input, -i        فایل ورودی (CSV یا Excel) [اجباری]
--type, -t         نوع فایل: auto, movie, series [پیش‌فرض: auto]
--output, -o       پوشه خروجی [پیش‌فرض: input_dir/output]
--delay, -d        تأخیر بین درخواست‌ها (ثانیه) [پیش‌فرض: 0.5]
--no-translate     غیرفعال کردن ترجمه خودکار
--quiet, -q        حالت بی‌صدا (فقط خطاها)
```

## 🔄 Resume بعد از قطعی

اگه اینترنت قطع شد یا برنامه متوقف شد:
1. دوباره همون دستور رو اجرا کن
2. برنامه از کش استفاده می‌کنه و درخواست‌های تکراری نمی‌زنه
3. فقط عناوین جدید پردازش میشن

## 📝 نکات مهم

### انکدینگ فایل‌ها
- برنامه خودکار انکدینگ رو تشخیص میده
- اگه مشکل داشتی، فایل رو با UTF-8 ذخیره کن

### عناوین مشکل‌دار
- بعضی عناوین مثل "Money Heist" ممکنه نتیجه اشتباه بدن
- برنامه اصلاحات دستی داره (فایل رو ببین: `TVMAZE_CORRECTIONS`)
- می‌تونی خودت اضافه کنی

### Rate Limit
- OMDb: بستگی به کلیدت داره
- TVMaze: 20 درخواست در 10 ثانیه
- برنامه خودش مدیریت می‌کنه (retry + backoff)

### ترجمه
- اگه `deep-translator` نصب نباشه، ترجمه خودکار غیرفعال میشه
- می‌تونی با `--no-translate` دستی غیرفعالش کنی

## 🆚 مقایسه با نسخه‌های قبلی

| ویژگی | کد ریپازیتوری | کد جدید | **v2** |
|---|---|---|---|
| ترجمه فارسی | ❌ | ✅ | ✅ |
| IMDb Suggestion API | ❌ | ✅ | ✅ |
| سریال | ✅ | ❌ | ✅ |
| کش | ✅ | ❌ | ✅ |
| Retry/Backoff | ❌ | ❌ | ✅ |
| CSV + Excel | ✅ | فقط Excel | ✅ |
| تشخیص خودکار نوع | ✅ | ❌ | ✅ |
| Progress bar | ❌ | ❌ | ✅ |
| گزارش آمار | ❌ | ❌ | ✅ |
| مدیریت انکدینگ | ✅ | ❌ | ✅ |
| خروجی Missing | ✅ | ❌ | ✅ |

## 🐛 عیب‌یابی

### خطای SSL
```bash
# آپدیت certificates
sudo apt-get install ca-certificates
```

### ترجمه کار نمی‌کنه
```bash
# نصب deep-translator
pip install deep-translator
```

### Rate Limit شدید
```bash
# افزایش delay
python cinefilm-enricher-v2.py --input movies.csv --delay 2.0
```

### فایل خوانده نمیشه
```bash
# بررسی انکدینگ
file -i movies.csv
# اگه مشکل داشت، با UTF-8 ذخیره کن
```

## 📄 لایسنس

آزاد برای استفاده شخصی و تجاری.

## 🙏 تقدیر

- OMDb API برای اطلاعات فیلم
- TVMaze API برای اطلاعات سریال
- Google Translate برای ترجمه
- IMDb برای Suggestion API

---

**ساخته شده با ❤️ برای آرشیو Cinefilm**
