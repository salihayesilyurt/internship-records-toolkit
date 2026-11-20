# -*- coding: utf-8 -*-
"""
Internship Excel Checker / Staj Excel Kontrol Scripti
=====================================================
Usage (computer):  python internship_checker.py STAJ_2026.xlsx
Colab:             open internship_checker.ipynb and choose "Runtime -> Run all".

What it checks:
  1. Student number: correct number of digits; same number with different names,
     same name with different numbers; unusual department code
  2. Dates: readable, start <= end, no overlapping internships for the same student
  3. Workdays: weekdays between start and end minus public holidays
     (Saturdays also count when the note mentions "cumartesi") vs. "Yapılan Gün"
  4. Acceptance rules: signature "Hayır" / attendance "Yok" / evaluation form "Yok"
     => accepted days must be 0; accepted days <= worked days
  5. Notes: rewrites notes with the same meaning in one standard wording and removes
     voice-entry leftovers such as "yeni satış"
  6. Reports different spellings of the same company / supervisor

Output: <file>_checked.xlsx
  - Sheet1: corrected data (yellow = fixed automatically, orange = check manually)
  - Kontrol Raporu: every finding with current value and suggestion
  - İş Günü Hesabı: workday calculation for each row
  - Yazım Farkları: possible duplicate spellings

The report texts written into the Excel file are Turkish, because the data is Turkish.
"""
import re
import sys
from collections import defaultdict
from datetime import date, datetime, timedelta

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# 1) PUBLIC HOLIDAYS  (add the new year's holidays here every year)
#    Holiday eves ("arefe") are half days and count as workdays by default.
#    Set EVE_IS_HOLIDAY = True to treat them as holidays.
# --------------------------------------------------------------------------
HOLIDAYS = {
    # 2025
    date(2025, 1, 1): "Yılbaşı",
    date(2025, 3, 30): "Ramazan Bayramı 1", date(2025, 3, 31): "Ramazan Bayramı 2",
    date(2025, 4, 1): "Ramazan Bayramı 3",
    date(2025, 4, 23): "23 Nisan", date(2025, 5, 1): "1 Mayıs", date(2025, 5, 19): "19 Mayıs",
    date(2025, 6, 6): "Kurban Bayramı 1", date(2025, 6, 7): "Kurban Bayramı 2",
    date(2025, 6, 8): "Kurban Bayramı 3", date(2025, 6, 9): "Kurban Bayramı 4",
    date(2025, 7, 15): "15 Temmuz", date(2025, 8, 30): "30 Ağustos",
    date(2025, 10, 29): "29 Ekim",
    # 2026
    date(2026, 1, 1): "Yılbaşı",
    date(2026, 3, 20): "Ramazan Bayramı 1", date(2026, 3, 21): "Ramazan Bayramı 2",
    date(2026, 3, 22): "Ramazan Bayramı 3",
    date(2026, 4, 23): "23 Nisan", date(2026, 5, 1): "1 Mayıs", date(2026, 5, 19): "19 Mayıs",
    date(2026, 5, 27): "Kurban Bayramı 1", date(2026, 5, 28): "Kurban Bayramı 2",
    date(2026, 5, 29): "Kurban Bayramı 3", date(2026, 5, 30): "Kurban Bayramı 4",
    date(2026, 7, 15): "15 Temmuz", date(2026, 8, 30): "30 Ağustos",
    date(2026, 10, 29): "29 Ekim",
}
HOLIDAY_EVES = {
    date(2025, 3, 29): "Ramazan arefesi", date(2025, 6, 5): "Kurban arefesi",
    date(2025, 10, 28): "29 Ekim arefesi",
    date(2026, 3, 19): "Ramazan arefesi", date(2026, 5, 26): "Kurban arefesi",
    date(2026, 10, 28): "29 Ekim arefesi",
}
EVE_IS_HOLIDAY = False

STUDENT_NO_DIGITS = 10  # number of digits in a student number (e.g. 2023141043)

# Column headers - must match the Excel headers exactly
COLUMNS = dict(
    row_no="SN", student_no="Öğrenci No", name="Ad-Soyad", company="Staj Yeri",
    supervisor="Sorumlu Kişi", start="Başlama Tarihi", end="Bitiş Tarihi",
    eval_form="Gizli Sicil Fişi", signed="Her Sayfa İmzalı mı", attendance="Puantaj Cetveli",
    days_worked="Yapılan Gün", days_accepted="Kabul Edilen Gün", notes="Açıklama",
)

YELLOW = PatternFill("solid", fgColor="FFF2CC")  # fixed automatically
ORANGE = PatternFill("solid", fgColor="F8CBAD")  # needs a manual check

STATUS_FIXED = "Düzeltildi"
STATUS_CHECK = "Kontrol edin"


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
def tr_lower(text):
    """Lower-case with Turkish I/İ rules."""
    return text.replace("I", "ı").replace("İ", "i").lower()


def tr_capitalize(text):
    """Capitalise the first letter with Turkish i/ı rules."""
    text = text.strip()
    if not text:
        return text
    first = {"i": "İ", "ı": "I"}.get(text[0], text[0].upper())
    return first + text[1:]


def to_date(value):
    """Convert a cell value to a date and try to repair broken text dates (03.08.206).
    Returns (date or None, was_repaired)."""
    if value is None or value == "":
        return None, False
    if isinstance(value, datetime):
        return value.date(), False
    if isinstance(value, date):
        return value, False
    match = re.match(r"^\s*(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})\s*$", str(value))
    if not match:
        return None, False
    day, month, year = int(match.group(1)), int(match.group(2)), match.group(3)
    if len(year) == 2:
        year = 2000 + int(year)
    elif len(year) == 3 and year.startswith("20"):  # 206 -> 2026 (a digit was dropped)
        year = 2020 + int(year[2])
    else:
        year = int(year)
    try:
        return date(year, month, day), True  # it was text, so it is converted
    except ValueError:
        return None, False


def is_workday(day, saturday_counts):
    if day in HOLIDAYS:
        return False
    if EVE_IS_HOLIDAY and day in HOLIDAY_EVES:
        return False
    weekday = day.weekday()
    if weekday == 6:
        return False
    if weekday == 5 and not saturday_counts:
        return False
    return True


def count_workdays(start, end, saturday_counts=False):
    """Returns (number of workdays, list of holidays that fall on a working day)."""
    if not start or not end or end < start:
        return None, []
    count, holidays_hit, day = 0, [], start
    while day <= end:
        if is_workday(day, saturday_counts):
            count += 1
        elif day in HOLIDAYS and day.weekday() < (6 if saturday_counts else 5):
            holidays_hit.append(f"{day:%d.%m} {HOLIDAYS[day]}")
        day += timedelta(days=1)
    return count, holidays_hit


def nth_workday(from_day, n, saturday_counts, step):
    """The n-th workday counting from from_day (step=+1 forward, -1 backward)."""
    day, found = from_day, 0
    while True:
        if is_workday(day, saturday_counts):
            found += 1
            if found == n:
                return day
        day += timedelta(days=step)


# --------------------------------------------------------------------------
# Note normalisation
# --------------------------------------------------------------------------
VOICE_ARTIFACT = re.compile(r"\byeni\s+sat[ıi][şsr]\w*(\s*\d+)?", re.I)  # "yeni satış 2021" etc.

# (pattern on the lower-cased note, function returning the standard wording)
NOTE_RULES = [
    (r"gizli\s*s[iı]\w*'?\s*(?:de|da|formunda|fişinde)?\s*fotoğraf\w*\s*(?:yok|eksik)\w*|fotoğraf\w*\s*(?:yok|eksik)",
     lambda m, t: "Gizli sicilde fotoğraf yok"),
    (r"(?:gizli\s*s[iı]\w*'?\s*(?:de|da)?\s*)?çalışmadığı\s+gün\s+sayısı\w*\s*(\d+)?[^.,;]*",
     lambda m, t: _days_not_worked_note(m, t)),
    (r"(?:hiçbir\s+sayfa|sayfaların\s+hiçbiri)\s+imza\w*\s*(?:değil|yok)?",
     lambda m, t: "Hiçbir sayfa imzalı değil"),
    (r"(?:staj\s+defteri\s+)?her\s+sayfa\w*\s+(?:imzalı\s+değil|imza\s+yok)",
     lambda m, t: "Her sayfa imzalı değil"),
    (r"puantaj\s*(?:cetveli)?\s*(?:yok|eksik)", lambda m, t: "Puantaj cetveli yok"),
    (r"cumartesi\w*\s*(?:günleri\s*)?(?:çalışma\s*)?d[iı]l\w*\s*(?:var)?",
     lambda m, t: "Cumartesi çalışma dilekçesi var"),
    (r"tarihlerin\s+hepsi\s+(20\d\d)[^.;]*",
     lambda m, t: f"Belgelerdeki tarihler {m.group(1)} yazılmış"),
    (r"gizli\s*s[iı]\w*'?\s*(?:de|da)?\s*15\s*temmuz\s*izni\s*yazılmış",
     lambda m, t: "Gizli sicilde 15 Temmuz izin olarak yazılmış"),
]


def _days_not_worked_note(match, whole_note):
    number = match.group(1)
    if number is None:
        found = re.search(r"(\d+)\s*gün", whole_note)
        number = found.group(1) if found else None
    reason = ""
    if "resmi tatil" in whole_note:
        reason = " (resmi tatil)"
    elif "izin" in whole_note and "15 temmuz" not in whole_note:
        reason = " (izin)"
    if number:
        return f"Gizli sicilde çalışmadığı gün sayısı {number}{reason}"
    if reason:
        return "Gizli sicilde çalışmadığı gün olarak resmi tatil yazılmış"
    return "Gizli sicilde çalışmadığı gün belirtilmiş"


def normalize_note(note):
    """Rewrite a note in standard wording; unknown parts are kept but tidied."""
    if note is None or str(note).strip() == "":
        return None
    text = VOICE_ARTIFACT.sub("", str(note))
    text = re.sub(r"\s+", " ", text).strip()
    lower = tr_lower(text)
    parts = []            # (position, sentence)
    mask = list(lower)    # matched characters are replaced with "\0"
    for pattern, make_sentence in NOTE_RULES:
        for match in re.finditer(pattern, lower):
            if all(ch == "\0" for ch in mask[match.start():match.end()]):
                continue
            sentence = make_sentence(match, lower)
            if sentence not in [p[1] for p in parts]:
                parts.append((match.start(), sentence))
            for i in range(match.start(), match.end()):
                mask[i] = "\0"
    # parts that matched no rule, in their original spelling
    leftover = "".join(text[i] if mask[i] != "\0" else "\0" for i in range(len(text)))
    for chunk in re.finditer(r"[^\0]+", leftover):
        for piece in re.split(r"[.;]", chunk.group()):
            p = piece.strip(" ,.;-")
            p = re.sub(r"^(ve|ayrıca|ama)\s+", "", p, flags=re.I).strip()
            p = re.sub(r"\s+(ve|ayrıca)$", "", p, flags=re.I).strip()
            p = re.sub(r"^\d+\s*gün\s*izin$", "", p, flags=re.I)  # already merged into days-not-worked
            if len(p) < 3 or tr_lower(p) in ("ayrıca", "ve", "ama"):
                continue
            p = re.sub(r"gizli\s*s[iı]\w*'?(de|da)?\b", "gizli sicilde", p, flags=re.I)
            p = re.sub(r"\s{2,}", " ", p)
            parts.append((chunk.start(), tr_capitalize(p)))
    parts.sort(key=lambda x: x[0])
    result = ". ".join(p[1] for p in parts)
    return (result + ".") if result else None


def normalize_company(name):
    if not isinstance(name, str):
        return name
    text = re.sub(r"\s+", " ", name).strip()
    text = re.sub(r"\bA\.?\s?ş\.?(?=\s|$)", "A.Ş.", text)  # A.ş -> A.Ş.
    text = re.sub(r"\bA\.Ş(?=\s|$)", "A.Ş.", text)
    text = re.sub(r"\bLtd\.?\s*Şti\.?", "Ltd. Şti.", text)
    text = re.sub(r"\bSan\.?\s*Ve\s*Tic\b\.?", "San. Ve Tic.", text)
    return text


def similarity_key(value):
    """Rough key for spotting spelling variants of the same company / person."""
    text = tr_lower(str(value))
    text = text.translate(str.maketrans("çğıöşü", "cgiosu"))
    text = re.sub(r"\b(a\.?s\.?|ltd\.?|sti\.?|san\.?|tic\.?|ve|sanayi|ticaret|anonim|sirketi|limited|"
                  r"yazilim|teknolojileri|teknoloji|muhendislik|danismanlik|his\.?|hizmetleri)\b", " ", text)
    text = re.sub(r"[^a-z0-9 ]", " ", text)
    return " ".join(text.split()[:2])


def is_number(value):
    return isinstance(value, (int, float))


# --------------------------------------------------------------------------
# Main check
# --------------------------------------------------------------------------
def check_workbook(input_path, output_path=None):
    workbook = openpyxl.load_workbook(input_path)
    sheet = workbook.worksheets[0]
    headers = [c.value.strip() if isinstance(c.value, str) else c.value for c in sheet[1]]
    col = {key: headers.index(title) + 1 for key, title in COLUMNS.items()}

    findings = []      # (excel row, SN, student no, name, column, problem, current, suggestion, status)
    workday_rows = []
    data_rows = []

    def cell(row, key):
        return sheet.cell(row, col[key])

    def report(row, key, problem, current, suggestion, auto_fixed):
        c = cell(row, key)
        c.fill = YELLOW if auto_fixed else ORANGE
        previous = c.comment.text + "\n" if c.comment else ""
        detail = f"\nÖnceki: {current}" if auto_fixed else (f"\nÖneri: {suggestion}" if suggestion else "")
        c.comment = Comment(previous + problem + detail, "Kontrol")
        findings.append((row, cell(row, "row_no").value, cell(row, "student_no").value, cell(row, "name").value,
                         COLUMNS[key], problem, current, suggestion,
                         STATUS_FIXED if auto_fixed else STATUS_CHECK))

    for row in range(2, sheet.max_row + 1):
        if all(sheet.cell(row, col[k]).value in (None, "") for k in ("student_no", "name", "company")):
            continue
        data_rows.append(row)

    # --- names replaced by a voice command ("Aynen", "aynı") -> take the name from the same number
    rows_by_number = defaultdict(list)
    for row in data_rows:
        rows_by_number[str(cell(row, "student_no").value).strip()].append(row)
    voice_words = ("aynen", "aynı", "aynısı", "ayni")
    for row in data_rows:
        name = str(cell(row, "name").value or "").strip()
        if tr_lower(name) in voice_words:
            others = [str(cell(x, "name").value).strip()
                      for x in rows_by_number[str(cell(row, "student_no").value).strip()]
                      if x != row and tr_lower(str(cell(x, "name").value)) not in voice_words]
            if others:
                report(row, "name", "Ad yerine sesli komut yazılmış, aynı numaralı satırdan alındı",
                       name, others[0], True)
                cell(row, "name").value = others[0]

    # --- per row: number, dates, workdays, acceptance rules, notes, spelling
    for row in data_rows:
        # student number
        number = cell(row, "student_no").value
        digits = re.sub(r"\D", "", str(number or ""))
        if len(digits) != STUDENT_NO_DIGITS:
            report(row, "student_no", f"Öğrenci numarası {len(digits)} haneli ({STUDENT_NO_DIGITS} olmalı)",
                   number, "", False)
        elif not isinstance(number, int):
            cell(row, "student_no").value = int(digits)

        # dates
        start_raw, end_raw = cell(row, "start").value, cell(row, "end").value
        start, start_fixed = to_date(start_raw)
        end, end_fixed = to_date(end_raw)
        for key, raw, parsed, fixed in (("start", start_raw, start, start_fixed), ("end", end_raw, end, end_fixed)):
            if parsed is None:
                report(row, key, "Tarih okunamadı", raw, "", False)
            elif fixed:
                report(row, key, "Tarih metin/bozuk yazılmış, düzeltildi", raw, f"{parsed:%d.%m.%Y}", True)
                cell(row, key).value = datetime(parsed.year, parsed.month, parsed.day)
                cell(row, key).number_format = "dd.mm.yyyy"

        old_note = cell(row, "notes").value
        saturday_counts = "cumartesi" in tr_lower(str(old_note or ""))
        worked = cell(row, "days_worked").value
        accepted = cell(row, "days_accepted").value

        calculated, holidays_hit = count_workdays(start, end, saturday_counts)
        if start and end and end < start:
            suggestion = ""
            if is_number(worked):
                expected_start = nth_workday(end, worked, saturday_counts, -1)
                suggestion = f"Bitiş {end:%d.%m.%Y} ve {worked} gün ise başlama {expected_start:%d.%m.%Y} olmalı"
            report(row, "start", "Başlama tarihi bitiş tarihinden sonra",
                   f"{start:%d.%m.%Y} > {end:%d.%m.%Y}", suggestion, False)
        if calculated is not None and is_number(worked) and calculated != worked:
            with_saturdays, _ = count_workdays(start, end, True)
            if not saturday_counts and with_saturdays == worked:
                report(row, "days_worked",
                       f"Hafta içi {calculated} iş günü; cumartesiler sayılınca {with_saturdays} ediyor "
                       f"ve Yapılan Gün ile tutuyor", worked,
                       "Cumartesi çalışmış görünüyor: dilekçeyi kontrol edip nota "
                       "'Cumartesi çalışma dilekçesi var.' ekleyin", False)
            else:
                expected_end = nth_workday(start, worked, saturday_counts, +1)
                report(row, "days_worked",
                       f"Tarihlere göre {calculated} iş günü, Yapılan Gün {worked}"
                       + (f" (aradaki tatil: {', '.join(holidays_hit)})" if holidays_hit else ""),
                       worked, f"Yapılan Gün {worked} ise bitiş {expected_end:%d.%m.%Y} olmalı", False)
        workday_rows.append((cell(row, "row_no").value, number, cell(row, "name").value,
                             start, end, "Evet" if saturday_counts else "Hayır",
                             ", ".join(holidays_hit), calculated, worked,
                             (worked - calculated) if (calculated is not None and is_number(worked)) else None))

        # acceptance rules
        signed, attendance, eval_form = (str(cell(row, k).value or "").strip()
                                         for k in ("signed", "attendance", "eval_form"))
        reasons = [text for text, condition in (("imza Hayır", signed == "Hayır"),
                                                ("puantaj Yok", attendance == "Yok"),
                                                ("gizli sicil Yok", eval_form == "Yok")) if condition]
        if reasons and accepted not in (0, None):
            report(row, "days_accepted", f"{', '.join(reasons)} ama Kabul {accepted}", accepted, 0, False)
        if not reasons and accepted == 0:
            report(row, "days_accepted", "Belgeler tam görünüyor ama Kabul 0", accepted, worked, False)
        if is_number(worked) and is_number(accepted) and accepted > worked:
            report(row, "days_accepted", "Kabul edilen gün yapılan günden fazla", accepted, worked, False)
        for key, allowed in (("eval_form", ("Var", "Yok")), ("attendance", ("Var", "Yok")),
                             ("signed", ("Evet", "Hayır"))):
            value = str(cell(row, key).value or "").strip()
            if value not in allowed:
                report(row, key, f"Geçersiz değer (beklenen {'/'.join(allowed)})", value, "", False)

        # note
        new_note = normalize_note(old_note)
        if (old_note or None) != new_note:
            old_plain = str(old_note).strip().rstrip(". ") if old_note else None
            new_plain = new_note.rstrip(".") if new_note else None
            if old_plain != new_plain:  # do not report a change that only adds a full stop
                report(row, "notes", "Not standart biçime getirildi", old_note, new_note, True)
            cell(row, "notes").value = new_note

        # company / supervisor spelling
        for key in ("company", "supervisor"):
            value = cell(row, key).value
            if key == "company":
                new_value = normalize_company(value)
            else:
                new_value = re.sub(r"\s+", " ", value).strip() if isinstance(value, str) else value
            if new_value != value:
                if re.sub(r"\s+", " ", str(value)).strip() != new_value:  # report more than whitespace changes
                    report(row, key, "Yazım standartlaştırıldı (A.Ş., Ltd. Şti. vb.)", value, new_value, True)
                cell(row, key).value = new_value

    # --- unusual department code (digits 5-7 of the student number)
    code_counts = defaultdict(int)
    for row in data_rows:
        number = str(cell(row, "student_no").value)
        if len(number) == STUDENT_NO_DIGITS:
            code_counts[number[4:7]] += 1
    common_codes = {code for code, count in code_counts.items() if count >= 3}
    for row in data_rows:
        number = str(cell(row, "student_no").value)
        if len(number) == STUDENT_NO_DIGITS and number[4:7] not in common_codes:
            report(row, "student_no",
                   f"Numaradaki bölüm kodu {number[4:7]}, diğerlerinden farklı ({'/'.join(sorted(common_codes))})",
                   cell(row, "student_no").value, "Yatay/dikey geçiş olabilir; belgeden doğrulayın", False)

    # --- same number with different names / same name with different numbers / overlaps
    by_number, by_name = defaultdict(list), defaultdict(list)
    for row in data_rows:
        by_number[str(cell(row, "student_no").value)].append(row)
        by_name[tr_lower(str(cell(row, "name").value).strip())].append(row)
    for number, rows in by_number.items():
        names = {str(cell(row, "name").value).strip() for row in rows}
        if len(names) > 1:
            for row in rows:
                report(row, "name", f"Aynı numarada farklı ad yazımı: {' / '.join(sorted(names))}",
                       cell(row, "name").value, "Belgeden doğru yazımı kontrol edin", False)
        periods = []
        for row in rows:
            s, _ = to_date(cell(row, "start").value)
            e, _ = to_date(cell(row, "end").value)
            if s and e and s <= e:
                periods.append((s, e, row))
        periods.sort()
        for (s1, e1, r1), (s2, e2, r2) in zip(periods, periods[1:]):
            if s2 <= e1:
                report(r2, "start", f"SN {cell(r1, 'row_no').value} ile tarihler çakışıyor",
                       f"{s2:%d.%m.%Y}", "", False)
    for name, rows in by_name.items():
        numbers = {str(cell(row, "student_no").value) for row in rows}
        if len(numbers) > 1:
            for row in rows:
                report(row, "student_no", f"Aynı ad farklı numaralarla: {', '.join(sorted(numbers))}",
                       cell(row, "student_no").value, "", False)

    # --- spelling variants of the same company / supervisor (information only)
    spelling_variants = []
    for key in ("company", "supervisor"):
        groups = defaultdict(set)
        for row in data_rows:
            value = cell(row, key).value
            if value:
                groups[similarity_key(value)].add(str(value).strip())
        for variants in groups.values():
            if len(variants) > 1:
                spelling_variants.append((COLUMNS[key], " | ".join(sorted(variants))))

    # --- report sheets (Turkish, like the data)
    for title in ("Kontrol Raporu", "İş Günü Hesabı", "Yazım Farkları"):
        if title in workbook.sheetnames:
            del workbook[title]
    report_sheet = workbook.create_sheet("Kontrol Raporu")
    report_sheet.append(["Excel Satırı", "SN", "Öğrenci No", "Ad-Soyad", "Alan", "Sorun",
                         "Mevcut / Önceki", "Öneri / Yeni", "Durum"])
    for finding in sorted(findings, key=lambda x: (x[8] != STATUS_CHECK, x[0])):
        report_sheet.append([v.strftime("%d.%m.%Y") if isinstance(v, (date, datetime)) else v for v in finding])
    workday_sheet = workbook.create_sheet("İş Günü Hesabı")
    workday_sheet.append(["SN", "Öğrenci No", "Ad-Soyad", "Başlama", "Bitiş", "Cumartesi sayıldı",
                          "Aradaki resmi tatiller", "Hesaplanan İş Günü", "Yapılan Gün", "Fark"])
    for values in workday_rows:
        workday_sheet.append(list(values))
        for c in (4, 5):
            workday_sheet.cell(workday_sheet.max_row, c).number_format = "dd.mm.yyyy"
        if values[9] not in (0, None):
            for c in range(1, 11):
                workday_sheet.cell(workday_sheet.max_row, c).fill = ORANGE
    spelling_sheet = workbook.create_sheet("Yazım Farkları")
    spelling_sheet.append(["Alan", "Farklı yazımlar (aynı kurum/kişi olabilir)"])
    for values in spelling_variants:
        spelling_sheet.append(list(values))

    widths = {report_sheet: [12, 8, 13, 26, 18, 60, 40, 45, 13],
              workday_sheet: [6, 13, 26, 12, 12, 10, 30, 12, 11, 7],
              spelling_sheet: [14, 120]}
    for sh, column_widths in widths.items():
        for c in sh[1]:
            c.font = Font(bold=True)
        sh.freeze_panes = "A2"
        for i, width in enumerate(column_widths, 1):
            sh.column_dimensions[get_column_letter(i)].width = width
        for cells in sh.iter_rows(min_row=2):
            for c in cells:
                c.alignment = Alignment(wrap_text=True, vertical="top")

    output_path = output_path or re.sub(r"\.xlsx$", "", input_path, flags=re.I) + "_checked.xlsx"
    workbook.save(output_path)
    open_items = sum(1 for f in findings if f[8] == STATUS_CHECK)
    print(f"{len(data_rows)} rows checked: {len(findings) - open_items} fixed automatically, "
          f"{open_items} need a manual check. Output: {output_path}")
    print(f"{len(data_rows)} satır kontrol edildi: {len(findings) - open_items} otomatik düzeltme, "
          f"{open_items} elle kontrol gerektiren madde. Çıktı: {output_path}")
    return output_path


if __name__ == "__main__":
    try:
        from google.colab import files  # running in Colab
        uploaded = files.upload()
        result = check_workbook(list(uploaded.keys())[0])
        files.download(result)
    except ImportError:
        if len(sys.argv) < 2:
            print("Usage / Kullanım: python internship_checker.py STAJ_2026.xlsx [output.xlsx]")
            sys.exit(1)
        check_workbook(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
