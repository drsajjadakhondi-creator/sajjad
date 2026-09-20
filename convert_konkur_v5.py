# -*- coding: utf-8 -*-
"""
تبدیل PDF انتخاب رشته کنکور به اکسل (نسخه ۵)

ساختار خروجی:
  1. عنوان رشته
  2. دوره تحصیلی
  3. دانشگاه
  4. نیم‌سال (اول / دوم)
  5. کدرشته محل
  6. توضیحات
  7. نحوه پذیرش

- ردیف‌های هدر جدول وارد اکسل نمی‌شوند
- اگر جدولی تیتر مستقیم بالای خودش نداشت، از تیتر جدول قبلی استفاده می‌شود
- متن‌های چندخطی عنوان رشته و توضیحات جمع‌آوری می‌شوند
- اصلاح RTL + اصلاحات رایج املایی
- نسخه ۴: رفع اشکال «توضیحات» چندخطی که بین دو ردیف کدرشته‌ی مجاور
  پاره می‌شد (به‌خاطر همپوشانی عمودی متن توضیحات با ردیف بالایی‌اش در PDF)
- نسخه ۵: رفع اشکال «دانشگاه/استان» که گاهی نام قبلی (مثلا یزد) به‌جای
  نام واقعی جدول جاری نوشته می‌شد. به‌جای حدس‌زدن از روی کلیدواژه‌هایی مثل
  «دانشگاه/دانشکده»، حالا دقیقا همان یک یا دو خطی که بلافاصله بالای هر
  جدول قرار دارد (همان خطی که در PDF عملا اسم استان/دانشگاه است) عینا
  به‌عنوان تیتر برداشته می‌شود؛ فقط پیشوند «ادامه» از آن حذف می‌شود.

نحوه اجرا:
    python convert_konkur_v5.py konkur.pdf konkur_v5.xlsx
"""

import sys
import re
import pdfplumber
import pandas as pd
from collections import defaultdict

DEFAULT_PDF_PATH = "konkur.pdf"
DEFAULT_OUTPUT_PATH = "konkur_v5.xlsx"

# پیشوندهای «ادامه» که باید از ابتدای تیتر حذف شوند (ادامه، ادامه‌ی، ادامهٔ، ...)
CONTINUATION_PREFIX = re.compile(r"^ادامه[\s\u200c]*(ی|ء|ٔ)?[\s\u200c]*[-–—:]*\s*")

# اگر فاصله‌ی عمودی بین دو خط کمتر از این مقدار (بر حسب واحد PDF) باشد،
# آن دو خط را بخشی از یک تیتر واحدِ چندخطی در نظر می‌گیریم
LINE_MERGE_GAP = 6.0


PERSIAN_RANGE = re.compile(r"[\u0600-\u06FF]")
CODE_RE = re.compile(r"^\d{4,6}$")

# اصلاحات رایج استخراج اشتباه از این نوع PDF
COMMON_FIXES = [
    ("اسالمي", "اسلامي"),
    ("اسالمی", "اسلامی"),
    ("ايالم", "ايلام"),
    ("ایالم", "ایلام"),
    ("اطالعات", "اطلاعات"),
    ("سالمت", "سلامت"),
    ("گيالن", "گيلان"),
    ("گیالن", "گیلان"),
    ("خلخلا", "خلخال"),
    ("لابرز", "البرز"),
    ("بقيهلاه", "بقيه الله"),
    ("بقیهلاه", "بقیه الله"),
    ("شملاي", "شمالی"),
    ("شملای", "شمالی"),
    ("تربتجام", "تربت جام"),
    ("پيوستها", "پیوست‌ها"),
    ("پيوست‌ها", "پیوست‌ها"),
    ("بهياران", "بهیاران"),
    ("ويژه", "ویژه"),
    ("خوي", "خوی"),
    ("گرمي", "گرمی"),
    ("مشگين", "مشگین"),
    ("خميني", "خمینی"),
    ("آذربايجان", "آذربایجان"),
    ("غربي", "غربی"),
    ("پزشکي", "پزشکی"),
    ("داروسازي", "داروسازی"),
    ("دندانپزشکي", "دندانپزشکی"),
    ("دامپزشکي", "دامپزشکی"),
    ("پرستاري", "پرستاری"),
    ("مامايي", "مامایی"),
    ("هوشبري", "هوشبری"),
    ("فيزيوتراپي", "فیزیوتراپی"),
    ("کاردرماني", "کاردرمانی"),
    ("گفتار درماني", "گفتار درمانی"),
    ("شنوايي شناسي", "شنوایی شناسی"),
    ("علوم آزمايشگاهي", "علوم آزمایشگاهی"),
    ("علوم تغذيه", "علوم تغذیه"),
    ("فناوري", "فناوری"),
    ("فوريت هاي", "فوریت‌های"),
    ("فوريت‌هاي", "فوریت‌های"),
    ("پيش بيمارستاني", "پیش‌بیمارستانی"),
    ("پيش‌بيمارستاني", "پیش‌بیمارستانی"),
    ("تکنولوژي", "تکنولوژی"),
    ("پرتوشناسي", "پرتوشناسی"),
    ("مهندسي", "مهندسی"),
    ("بهداشت محيط", "بهداشت محیط"),
    ("حرفه اي", "حرفه‌ای"),
    ("ايمني كار", "ایمنی کار"),
    ("کتابداري", "کتابداری"),
    ("اطالع رساني", "اطلاع‌رسانی"),
    ("اطلاع رساني", "اطلاع‌رسانی"),
    ("مديريت", "مدیریت"),
    ("بهداشتي درماني", "بهداشتی درمانی"),
    ("پروتزهاي دنداني", "پروتزهای دندانی"),
    ("اعضاي مصنوعي", "اعضای مصنوعی"),
    ("وسايل کمکي", "وسایل کمکی"),
    ("بيولوژي", "بیولوژی"),
    ("ناقلين", "ناقلین"),
    ("شهريه پرداز", "شهریه پرداز"),
    ("واگذاري", "واگذاری"),
    ("محدوديت", "محدودیت"),
    ("معرفي", "معرفی"),
    ("خوابگاههاي", "خوابگاه‌های"),
    ("ملکي", "ملکی"),
    ("براي", "برای"),
    ("داراي", "دارای"),
    ("شرايط", "شرایط"),
    ("جمهوري", "جمهوری"),
    ("اسلامي", "اسلامی"),
    ("فرماندهي", "فرماندهی"),
    ("انتظامي", "انتظامی"),
    ("فراجا", "فراجا"),
    ("ضوابط", "ضوابط"),
    ("گزینش", "گزینش"),
    ("كارداني", "کاردانی"),
    ("دكتري", "دکتری"),
    ("عمومي", "عمومی"),
    ("تحصيل", "تحصیل"),
    ("عالي", "عالی"),
    ("نجفآباد", "نجف‌آباد"),
    ("بيدگل", "بیدگل"),
]


def apply_common_fixes(text: str) -> str:
    if not text:
        return text
    for wrong, correct in COMMON_FIXES:
        text = text.replace(wrong, correct)
    return text


def fix_rtl_word(word: str) -> str:
    if not word:
        return word
    if PERSIAN_RANGE.search(word):
        return word[::-1]
    return word


def fix_rtl_text(text) -> str:
    if text is None:
        return ""
    text = str(text).strip()
    if not text:
        return ""
    words = text.split()
    fixed = [fix_rtl_word(w) for w in reversed(words)]
    return apply_common_fixes(" ".join(fixed))


def fix_rtl_cell(cell) -> str:
    return fix_rtl_text(cell)


def is_number_like(val: str) -> bool:
    val = (val or "").strip().replace(",", "").replace("٬", "")
    if not val or val in ("-", "–", "—", "ـ"):
        return False
    try:
        float(val)
        return True
    except ValueError:
        return False


def strip_continuation_prefix(text: str) -> str:
    """پیشوند «ادامه / ادامه‌ی / ادامهٔ ...» را از ابتدای تیتر حذف می‌کند."""
    if not text:
        return text
    return CONTINUATION_PREFIX.sub("", text).strip()


def get_lines_in_region(page, top_bound, bottom_bound):
    """خط‌های موجود در یک محدوده عمودی از صفحه را برمی‌گرداند، هرکدام همراه
    با موقعیت عمودی‌شان (top, bottom) تا بشود فاصله‌ی بین خط‌ها را سنجید."""
    words = page.extract_words(use_text_flow=False)
    region = [w for w in words if top_bound <= w["top"] < bottom_bound]
    region.sort(key=lambda w: (round(w["top"] / 3), w["x0"]))
    lines = defaultdict(list)
    for w in region:
        lines[round(w["top"] / 3)].append(w)
    result = []
    for key in sorted(lines.keys()):
        line_words = sorted(lines[key], key=lambda w: w["x0"])
        raw = " ".join(w["text"] for w in line_words)
        result.append({
            "top": min(w["top"] for w in line_words),
            "bottom": max(w["bottom"] for w in line_words),
            "text": fix_rtl_text(raw),
        })
    return result


def extract_heading_from_lines(lines, max_lines=2):
    """دقیقا همان یک یا دو خطی که بلافصل بالای جدول قرار دارد را به‌عنوان
    تیتر (اسم استان/دانشگاه) برمی‌گرداند. اگر دو خط آخر از نظر عمودی خیلی
    به هم نزدیک باشند (یعنی عملا یک تیتر است که در PDF شکسته شده)، هر دو
    را به هم می‌چسباند؛ در غیر این صورت فقط آخرین خط را برمی‌دارد."""
    if not lines:
        return None
    chosen = [lines[-1]]
    idx = len(lines) - 1
    while len(chosen) < max_lines and idx > 0:
        gap = lines[idx]["top"] - lines[idx - 1]["bottom"]
        if gap <= LINE_MERGE_GAP:
            chosen.insert(0, lines[idx - 1])
            idx -= 1
        else:
            break
    heading = " ".join(ln["text"] for ln in chosen)
    heading = apply_common_fixes(heading)
    heading = strip_continuation_prefix(heading)
    heading = re.sub(r"\s+", " ", heading).strip()
    return heading or None


def is_header_row(row):
    text = " ".join(c for c in row if c)
    indicators = [
        "نحوه پذیرش", "دوره تحصیلی", "کد رشته", "عنوان رشته", "توضیحات",
        "ظرفیت", "جنس پذیرش", "نیمسال", "نیم‌سال", "کدرشته",
        "پذيرش", "توضيحات", "هتشردک",
    ]
    return sum(1 for ind in indicators if ind in text) >= 2


def has_code(row):
    for c in row:
        if CODE_RE.match((c or "").strip()):
            return True
    return False


def find_code(row):
    for c in row:
        s = (c or "").strip()
        if CODE_RE.match(s):
            return s
    return ""


# ایندکس‌های پیش‌فرض برای ساختار رایج ۹ ستونه (استخراج LTR)
# [0:توضیحات, 1:مرد, 2:زن, 3:دوم, 4:اول, 5:عنوان رشته, 6:کد, 7:دوره, 8:نحوه پذیرش]
DEFAULT_COL_MAP = {
    "desc": 0,
    "gender_m": 1,
    "gender_f": 2,
    "cap_second": 3,
    "cap_first": 4,
    "title": 5,
    "code": 6,
    "period": 7,
    "admission": 8,
}



# هر نقش با چند کلیدواژه که فقط مخصوص خودش است شناسایی می‌شود.
# نکته‌ی مهم: چون عبارت «کد رشته» خودش شامل کلمه‌ی «رشته» است، باید
# قبل از «عنوان رشته» بررسی و «تصاحب» شود، وگرنه ستون کد به اشتباه
# به عنوان ستون «عنوان رشته» شناسایی می‌شود (باگ اصلی نسخه‌ی قبلی).
COLUMN_ROLE_KEYWORDS = [
    ("desc", ("توضیحات", "توضيحات")),
    ("code", ("کدرشته", "کد رشته", "هتشردک", "هتشر دک")),
    ("admission", ("نحوه پذیرش", "نحوه", "شريذپ هوحن")),
    ("period", ("دوره تحصیلی", "دوره", "يليصحت")),
    ("title", ("عنوان رشته", "عنوان")),
    ("cap_first", ("اول", "لوا")),
    ("cap_second", ("دوم", "مود")),
]


def detect_col_map(header_rows, n_cols):
    """از روی متن هدر جدول، ایندکس هر ستون را تشخیص می‌دهد.
    هر ستون فقط یک‌بار «تصاحب» می‌شود تا دو نقش روی یک ستون قرار نگیرند."""
    col_map = dict(DEFAULT_COL_MAP)
    if not header_rows:
        return col_map

    joined_per_col = [[] for _ in range(n_cols)]
    for row in header_rows:
        for i, c in enumerate(row):
            if i < n_cols and c:
                joined_per_col[i].append(c)
    combined = [" ".join(cells) for cells in joined_per_col]

    assigned = set()
    for role, keywords in COLUMN_ROLE_KEYWORDS:
        for i, text in enumerate(combined):
            if i in assigned:
                continue
            if any(kw in text for kw in keywords):
                col_map[role] = i
                assigned.add(i)
                break
    return col_map


def process_table(raw_rows, university, page_number, last_col_map=None):
    if not raw_rows:
        return [], last_col_map

    fixed_rows = []
    for row in raw_rows:
        if row is None:
            continue
        fixed_rows.append([fix_rtl_cell(c) for c in row])

    if not fixed_rows:
        return [], last_col_map

    n_cols = max(len(r) for r in fixed_rows)
    # pad
    fixed_rows = [r + [""] * (n_cols - len(r)) for r in fixed_rows]

    # جدا کردن هدر
    header_rows = []
    data_start = 0
    for i, row in enumerate(fixed_rows[:8]):
        if is_header_row(row):
            header_rows.append(row)
            data_start = i + 1
        elif header_rows:
            break

    if header_rows:
        col_map = detect_col_map(header_rows, n_cols)
    elif last_col_map and all(v < n_cols for v in last_col_map.values()):
        # فقط وقتی نقشه‌ی قبلی هنوز با تعداد ستون فعلی سازگار است از آن استفاده کن؛
        # اگر تعداد ستون‌ها عوض شده باشد، نقشه‌ی قدیمی می‌تواند اشتباه باشد.
        col_map = last_col_map
    else:
        col_map = detect_col_map([], n_cols)

    records = []
    consumed = set()  # ردیف‌هایی که به عنوان ادامه رکورد قبلی مصرف شده‌اند
    i = data_start
    while i < len(fixed_rows):
        if i in consumed:
            i += 1
            continue
        row = fixed_rows[i]
        if not has_code(row):
            i += 1
            continue

        code = find_code(row)
        title = row[col_map["title"]] if col_map["title"] < len(row) else ""
        period = row[col_map["period"]] if col_map["period"] < len(row) else ""
        admission = row[col_map["admission"]] if col_map["admission"] < len(row) else ""
        desc = row[col_map["desc"]] if col_map["desc"] < len(row) else ""
        cap_first = row[col_map["cap_first"]] if col_map["cap_first"] < len(row) else ""
        cap_second = row[col_map["cap_second"]] if col_map["cap_second"] < len(row) else ""

        HEADER_LEAK = {
            "محل", "اول", "دوم", "مرد", "زن", "ظرفیت", "جنس", "پذیرش",
            "نیمسال", "نیم‌سال", "کدرشته", "عنوان", "رشته", "دوره", "نحوه",
        }

        def is_leak(val: str) -> bool:
            v = val.strip()
            if not v or is_number_like(v):
                return True
            if v in HEADER_LEAK:
                return True
            if len(v) <= 2:
                return True
            return False

        # فقط ردیف‌های بلافاصله قبل که هنوز مصرف نشده‌اند
        pre_title, pre_desc = [], []
        k = i - 1
        while k >= data_start and k not in consumed:
            prev = fixed_rows[k]
            if has_code(prev) or is_header_row(prev):
                break
            non_empty = [(idx, c.strip()) for idx, c in enumerate(prev) if c and c.strip()]
            if not non_empty:
                k -= 1
                continue
            accepted = False
            for idx, val in non_empty:
                if is_leak(val):
                    continue
                if idx == col_map["title"]:
                    pre_title.insert(0, val)
                    accepted = True
                elif idx == col_map["desc"]:
                    pre_desc.insert(0, val)
                    accepted = True
            if not accepted:
                break
            consumed.add(k)
            k -= 1

        post_title, post_desc = [], []
        j = i + 1
        while j < len(fixed_rows):
            nxt = fixed_rows[j]
            if has_code(nxt) or is_header_row(nxt):
                break
            non_empty = [(idx, c.strip()) for idx, c in enumerate(nxt) if c and c.strip()]
            if not non_empty:
                consumed.add(j)
                j += 1
                continue
            accepted = False
            for idx, val in non_empty:
                if is_leak(val):
                    continue
                if idx == col_map["title"]:
                    post_title.append(val)
                    accepted = True
                elif idx == col_map["desc"]:
                    post_desc.append(val)
                    accepted = True
            if not accepted:
                break
            consumed.add(j)
            j += 1

        full_title = " ".join(pre_title + ([title] if title else []) + post_title).strip()
        full_desc = " ".join(pre_desc + ([desc] if desc else []) + post_desc).strip()
        for leak in ("محل اول دوم", "اول دوم", "محل اول", "محل دوم"):
            full_desc = full_desc.replace(leak, "").strip()
        full_desc = re.sub(r"\s+", " ", full_desc).strip(" -–—")

        # نیم‌سال
        semester = ""
        if is_number_like(cap_first) and not is_number_like(cap_second):
            semester = "اول"
        elif is_number_like(cap_second) and not is_number_like(cap_first):
            semester = "دوم"
        elif is_number_like(cap_first):
            semester = "اول"

        full_title = apply_common_fixes(full_title)
        full_desc = apply_common_fixes(full_desc)
        period = apply_common_fixes(period)
        admission = apply_common_fixes(admission)
        uni = apply_common_fixes(university or "")

        records.append({
            "عنوان رشته": full_title,
            "دوره تحصیلی": period,
            "دانشگاه": uni,
            "نیم‌سال": semester,
            "کدرشته محل": code,
            "توضیحات": full_desc,
            "نحوه پذیرش": admission,
        })
        i = j

    return records, col_map


# ---------------------------------------------------------------------------
# رفع اشکال «توضیحات» چندخطی که بین دو (یا چند) ردیف کدرشته پاره می‌شود
# ---------------------------------------------------------------------------
# علت اشکال: وقتی متنِ ستون «توضیحات» بیش از یک خط باشد ولی ارتفاع ردیفِ خودش
# در جدول PDF فقط برای یک خط جا دارد، خطوط اضافه به‌صورت بصری به سمت بالا
# سرریز می‌کند و داخل محدوده‌ی ردیفِ (کدرشته‌ی) قبلی افتاده و توسط pdfplumber
# به آن ردیف اشتباه نسبت داده می‌شود. متن واقعاً متعلق به ردیفی است که در
# پایین‌ترین نقطه قرار دارد نه ردیف بالایی.
#
# راه‌حل: یک «دیکشنری مطمئن» از توضیحات کامل و مستقل می‌سازیم (توضیحاتی که
# در فایل، ردیف قبل و بعدشان -در همان دانشگاه- خالی است، پس مطمئنیم که کامل
# و بدون پارگی استخراج شده‌اند). سپس هر توضیحِ ناقص را با استفاده از این
# دیکشنری با ردیف(های) بعدی خودش ترکیب می‌کنیم تا به یک عبارتِ کاملِ شناخته‌شده
# برسیم؛ اگر به تطبیق دقیقی نرسیدیم، برای جلوگیری از قاطی‌شدن اطلاعات، متن را
# دست‌نخورده رها می‌کنیم (محافظه‌کارانه عمل می‌کنیم).

MAX_WORDS_GROW = 40
DANGLING_WORDS = {
    "در", "و", "با", "به", "از", "برای", "که", "را", "تا", "یا", "شرایط",
    "خوابگاه‌های", "معرفی",
}
SUSPICIOUS_START_CHARS = ("»", "«", ")", "”", "’")
# کلماتی که در این نوع سند تقریبا هیچ‌وقت ابتدای یک یادداشتِ مستقل نیستند
# (حرف اضافه‌اند، یا وسطِ یک عبارتِ ثابت مثل «شرایط در دفترچه» قرار می‌گیرند)
SUSPICIOUS_START_WORDS = {
    "در", "دفترچه", "به", "با", "از", "برای", "که", "را", "تا", "یا", "و", "شرایط",
    "خودگردان",
}


def _build_trusted_descriptions(records):
    n = len(records)
    trusted = set()
    for i in range(n):
        d = (records[i].get("توضیحات") or "").strip()
        if not d:
            continue
        prev_empty = (
            i == 0
            or not (records[i - 1].get("توضیحات") or "").strip()
            or records[i - 1].get("دانشگاه") != records[i].get("دانشگاه")
        )
        next_empty = (
            i == n - 1
            or not (records[i + 1].get("توضیحات") or "").strip()
            or records[i + 1].get("دانشگاه") != records[i].get("دانشگاه")
        )
        if prev_empty and next_empty:
            trusted.add(d)
    return trusted


def _peel_prefix_notes(text, trusted):
    """اگر ابتدای متن، یکی از یادداشت‌های کامل و شناخته‌شده باشد، آن را جدا می‌کند
    (برای حالتی که دو یادداشت مستقل بدون فاصله‌گذار مشخص به هم چسبیده‌اند،
    مثل «ویژه بهیاران دارای مصاحبه - ...»)."""
    notes = []
    remaining = (text or "").strip(" -")
    changed = True
    while changed and remaining:
        changed = False
        for t in trusted:
            if t != remaining and remaining.startswith(t + " "):
                notes.append(t)
                remaining = remaining[len(t):].strip(" -")
                changed = True
                break
    return notes, remaining


def reconstruct_descriptions(records):
    """توضیحات پاره‌شده‌ی بین ردیف‌های مجاور (در همان دانشگاه) را بازسازی می‌کند."""
    n = len(records)
    descs = [(r.get("توضیحات") or "").strip() for r in records]
    unis = [r.get("دانشگاه") for r in records]
    trusted = _build_trusted_descriptions(records)

    def is_prefix_of_trusted(txt):
        return any(len(t) > len(txt) and t.startswith(txt + " ") for t in trusted)

    def try_grow_to_trusted(pending, start_row):
        buf = pending
        words_added = 0
        row = start_row
        while row + 1 < n and unis[row + 1] == unis[start_row] and words_added < MAX_WORDS_GROW:
            row += 1
            next_words = descs[row].split(" ") if descs[row] else []
            if not next_words or next_words == [""]:
                return None
            for k, w in enumerate(next_words):
                cand = (buf + " " + w).strip()
                words_added += 1
                if cand in trusted:
                    leftover = " ".join(next_words[k + 1:]).strip(" -")
                    return cand, row, leftover
                if is_prefix_of_trusted(cand):
                    buf = cand
                    continue
                return None
        return None

    result = [[] for _ in range(n)]
    consumed = [False] * n
    i = 0
    while i < n:
        if consumed[i]:
            i += 1
            continue
        if not descs[i]:
            i += 1
            continue
        notes, pending = _peel_prefix_notes(descs[i], trusted)
        result[i].extend(notes)
        if not pending:
            i += 1
            continue
        if pending in trusted or not is_prefix_of_trusted(pending):
            # کامل است یا این‌که هیچ نشانه‌ای از پارگی نداریم؛ دست‌نخورده می‌ماند
            result[i].append(pending)
            i += 1
            continue
        grown = try_grow_to_trusted(pending, i)
        if grown is None:
            # نتوانستیم با اطمینان کامل کنیم؛ برای جلوگیری از قاطی‌شدن اطلاعات دست‌نخورده می‌ماند
            result[i].append(pending)
            i += 1
            continue
        final_text, last_row, leftover = grown
        result[last_row].append(final_text)
        for r_idx in range(i + 1, last_row):
            consumed[r_idx] = True
        if leftover:
            more_notes, more_pending = _peel_prefix_notes(leftover, trusted)
            result[last_row].extend(more_notes)
            if more_pending:
                result[last_row].append(more_pending)
        i = last_row + 1

    for idx, r in enumerate(records):
        if consumed[idx]:
            r["توضیحات"] = ""
        else:
            r["توضیحات"] = " - ".join(dict.fromkeys(result[idx])) if result[idx] else ""

    # --- پاس دوم (شبکهٔ ایمنی) -------------------------------------------
    # بعضی بلوک‌ها بیش از دو یادداشت را هم‌زمان روی هم می‌ریزند (مثلا وقتی
    # سه یا چهار یادداشتِ جداگانه پشت سر هم سرریز می‌کنند)، و پاس اول بالا
    # (که فقط دو یادداشتِ متوالی را با دیکشنری تطبیق می‌دهد) نمی‌تواند آن‌ها
    # را با اطمینان کامل بازسازی کند. برای این بلوک‌های به‌شدت درهم، به‌جای
    # نوشتن متنِ ناقص/بریده روی یک ردیف اشتباه، کل متنِ خام همان بلوک را
    # (بدون حدس زدن مرز دقیق هر یادداشت) در تمام ردیف‌های همان بلوک تکرار
    # می‌کنیم؛ این تضمین می‌کند حداقل هیچ توضیحی نصفه/بی‌معنی باقی نماند،
    # هرچند ممکن است در این موارد نادر یک یادداشت به‌صورت اضافه هم روی
    # ردیف‌های مجاور دیده شود.
    #
    # عمداً اینجا از تطبیق «آیا این تکه پیشوند/پسوند یک عبارت طولانی‌تر است»
    # استفاده نمی‌کنیم: خیلی از یادداشت‌های این سند (مثلا تعهد خدمت برای
    # ده‌ها دانشگاه مختلف) عبارت‌های تقریبا مشابه و هم‌پوشان زیادی دارند و
    # چنین تطبیقی باعث می‌شد ردیف‌های کاملا سالمِ نامرتبط هم اشتباها به هم
    # بچسبند. به همین دلیل فقط از نشانه‌های محدود و مطمئنِ «بریدگی» (پایان‌
    # یافتن با یک حرف اضافه‌ی معلق، یا شروع‌شدن از وسط یک گیومه/پرانتز)
    # استفاده می‌کنیم.
    def looks_incomplete(text):
        text = (text or "").strip()
        if not text:
            return False
        if text in trusted:
            return False
        if text.startswith(SUSPICIOUS_START_CHARS):
            return True
        words = text.split()
        if words and words[0] in SUSPICIOUS_START_WORDS:
            return True
        last_clause = text.split(" - ")[-1].strip()
        if not last_clause:
            return True
        if last_clause.split()[-1] in DANGLING_WORDS:
            return True
        return False

    i = 0
    while i < n:
        if not descs[i]:
            i += 1
            continue
        j = i
        while j + 1 < n and descs[j + 1] and unis[j + 1] == unis[i]:
            j += 1
        run = list(range(i, j + 1))
        # فقط زیربازه‌های به‌هم‌پیوسته‌ی مشکل‌دار را اصلاح می‌کنیم، نه کل run را،
        # تا ردیف‌های سالم و مستقلِ کناری (که هیچ نشانه‌ی پارگی ندارند) دست‌نخورده بمانند
        trouble = [looks_incomplete(records[r_idx]["توضیحات"]) or consumed[r_idx] for r_idx in run]
        seg_start = None
        for k, r_idx in enumerate(run):
            if trouble[k]:
                if seg_start is None:
                    seg_start = k
            else:
                if seg_start is not None:
                    _apply_fallback(records, descs, run[seg_start:k])
                    seg_start = None
        if seg_start is not None:
            _apply_fallback(records, descs, run[seg_start:])
        i = j + 1

    return records


def _apply_fallback(records, descs, seg):
    """کل متن خام همان زیربازهٔ مشکل‌دار را (بدون حدس مرز دقیق) در تمام
    ردیف‌های همان زیربازه تکرار می‌کند تا اطلاعاتی نصفه/بی‌معنی باقی نماند.
    فقط برای بازه‌های کوتاه (حداکثر ۴ ردیف) و با تکه‌های نسبتاً کوتاه اعمال
    می‌شود؛ اگر خودِ یکی از تکه‌های خام از قبل طولانی باشد (مثلا یادداشت‌های
    «تعهد خدمت» که ذاتاً طولانی و تکراری‌اند)، به‌جای ادغام و طولانی‌ترکردنِ
    بیشتر، دست‌نخورده رها می‌شود تا خطر قاطی‌شدنِ اطلاعاتِ نامرتبط پیش نیاید."""
    if len(seg) < 2 or len(seg) > 4:
        return
    if any(len(descs[r_idx]) > 120 for r_idx in seg):
        return
    combined = re.sub(
        r"\s+", " ",
        " ".join(descs[r_idx] for r_idx in seg if descs[r_idx]),
    ).strip()
    for r_idx in seg:
        records[r_idx]["توضیحات"] = combined if descs[r_idx] else ""


def extract_all(pdf_path):
    all_records = []
    current_university = ""
    last_col_map = None

    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            tables = page.find_tables()
            if not tables:
                continue
            tables_sorted = sorted(tables, key=lambda t: t.bbox[1])
            prev_bottom = 0.0

            for table in tables_sorted:
                table_top = table.bbox[1]
                lines_above = get_lines_in_region(page, prev_bottom, table_top)
                found = extract_heading_from_lines(lines_above)
                if found:
                    # همان یک/دو خط بلافصل بالای جدول عینا به‌عنوان تیتر
                    # (استان/دانشگاه) گرفته می‌شود؛ اگر خطی بالای جدول نبود
                    # (found=None)، current_university قبلی دست‌نخورده
                    # می‌ماند و همان تیتر جدول قبلی ادامه پیدا می‌کند.
                    current_university = found

                raw = table.extract()
                records, last_col_map = process_table(
                    raw, current_university, page_number, last_col_map
                )
                all_records.extend(records)
                prev_bottom = table.bbox[3]

    return all_records


def main(pdf_path, output_path):
    print("در حال استخراج از PDF ... (ممکن است کمی طول بکشد)")
    records = extract_all(pdf_path)
    print(f"تعداد رکورد: {len(records)}")

    if not records:
        print("هیچ رکوردی پیدا نشد.")
        return

    print("در حال بازسازی توضیحاتِ چندخطی پاره‌شده ...")
    records = reconstruct_descriptions(records)

    columns = [
        "عنوان رشته",
        "دوره تحصیلی",
        "دانشگاه",
        "نیم‌سال",
        "کدرشته محل",
        "توضیحات",
        "نحوه پذیرش",
    ]
    df = pd.DataFrame(records)[columns]
    df.to_excel(output_path, index=False)
    print(f"ذخیره شد: {output_path}")


if __name__ == "__main__":
    pdf_path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_PDF_PATH
    output_path = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_OUTPUT_PATH
    main(pdf_path, output_path)
