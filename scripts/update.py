#!/usr/bin/env python3
"""
עדכון נתוני התחנות ובניית index.html.

מה הסקריפט עושה:
1. מחפש קובץ Excel חדש בתיקייה upload/ (דוח פזומט מאתר פז).
2. מנקה את הנתונים ומשווה לגרסה הקודמת (data/stations.json).
3. מאתר קואורדינטות (Geocoding) רק לתחנות חדשות או שהכתובת שלהן השתנתה,
   בעזרת Nominatim (OpenStreetMap), לפי מדיניות השימוש: בקשה אחת לשנייה לכל היותר.
4. מחיל עריכות ידניות מ-data/overrides.jsonl (מיקום ושעות) – הן לא נמחקות בעדכון.
5. בונה את index.html: קובץ אחד עם כל הקוד, Leaflet והנתונים.
6. כותב דוח שינויים (data/changes.md) ורשימת תחנות שלא אותרו (data/not-found.md).

הרצה: python scripts/update.py      (NO_GEOCODE=1 לדילוג על Geocoding)
"""
import datetime as dt
import glob
import hashlib
import json
import os
import re
import shutil
import sys
import time
import urllib.parse
import urllib.request

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
P = lambda *a: os.path.join(ROOT, *a)

STATIONS = P("data", "stations.json")
OVERRIDES = P("data", "overrides.jsonl")
GEOCACHE = P("data", "geocache.json")
CONFIG = P("data", "config.json")
CHANGES = P("data", "changes.md")
HISTORY = P("data", "history.md")
NOTFOUND = P("data", "not-found.md")
TEMPLATE = P("src", "template.html")
OUT = P("index.html")

NOMINATIM = "https://nominatim.openstreetmap.org/search"
FIELDS_COMPARE = [("street", "כתובת"), ("phone", "טלפון"), ("services", "שירותים")]


# ---------- כלים ----------
def read_json(path, default):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return default


def write_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
        f.write("\n")


def clean(s):
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()


def norm_key(s):
    s = clean(s).lower()
    s = re.sub(r"[\"'׳״`\-–.,()]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def clean_phone(p):
    digits = re.sub(r"\D", "", clean(p))
    if not digits:
        return ""
    if not digits.startswith("0") and len(digits) in (8, 9):
        digits = "0" + digits
    if len(digits) == 10 and digits.startswith("05"):
        return f"{digits[:3]}-{digits[3:]}"
    if len(digits) == 10 and digits.startswith("07"):
        return f"{digits[:3]}-{digits[3:]}"
    if len(digits) == 9:
        return f"{digits[:2]}-{digits[2:]}"
    return clean(p)


def station_id(name, city):
    return hashlib.sha1(f"{norm_key(name)}|{norm_key(city)}".encode()).hexdigest()[:10]


def file_hash(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        h.update(f.read())
    return h.hexdigest()


def fmt_date(iso):
    y, m, d = iso.split("-")
    return f"{d}/{m}/{y}"


# ---------- קריאת Excel ----------
def parse_excel(path):
    ws = openpyxl.load_workbook(path, data_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    report_date = None
    header_idx = None
    for i, r in enumerate(rows):
        first = clean(r[0] if r else "")
        m = re.search(r"(\d{1,2})[./](\d{1,2})[./](\d{4})", first)
        if m and report_date is None and "תאריך" in first:
            report_date = f"{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"
        if first == "שם תחנה":
            header_idx = i
            break
    if header_idx is None:
        raise SystemExit(f"לא נמצאה שורת כותרות 'שם תחנה' בקובץ {os.path.basename(path)}")
    headers = [clean(h) for h in rows[header_idx]]
    col = {h: j for j, h in enumerate(headers)}

    def get(r, name):
        j = col.get(name)
        return clean(r[j]) if j is not None and j < len(r) else ""

    stations, seen = [], {}
    for r in rows[header_idx + 1:]:
        name = get(r, "שם תחנה")
        if not name:
            continue
        city = get(r, "יישוב")
        sid = station_id(name, city)
        if sid in seen:  # שם ויישוב זהים – נוסיף סיומת
            seen[sid] += 1
            sid = f"{sid}-{seen[sid]}"
        else:
            seen[sid] = 1
        stations.append({
            "id": sid,
            "name": name,
            "city": city,
            "street": get(r, "רחוב"),
            "phone": clean_phone(get(r, "טלפון")),
            "services": get(r, "שירותים"),
        })
    if not report_date:
        report_date = dt.date.fromtimestamp(os.path.getmtime(path)).isoformat()
    return report_date, stations



# ---------- ניקוי כתובות לאיתור מיקום ----------
AREA_TYPES = ("city", "town", "village", "municipality", "state", "country", "county",
              "hamlet", "suburb", "neighbourhood", "quarter", "district", "region")
MAX_KM_FROM_CITY = 12
GEO_VERSION = 2
STATION_WORDS = r"(?:ב?ה?תחנת?(?:\s+(?:ה?דלק|פז|סונול|דור\s*אלון|דלק|ילו|yellow|יעד|טן|מנטה))?|ב?תחנה|מתחם|צומת|א\.?ת\.?|אזור\s+תעשיה|כביש|מול|ליד|בכניסה\s+ל\S+)"


def km(a, b):
    import math
    R = 6371
    la1, lo1, la2, lo2 = map(math.radians, (a["lat"], a["lng"], b["lat"], b["lng"]))
    h = math.sin((la2 - la1) / 2) ** 2 + math.cos(la1) * math.cos(la2) * math.sin((lo2 - lo1) / 2) ** 2
    return 2 * R * math.asin(math.sqrt(h))


def clean_name(n):
    n = re.sub(r"\(.*?\)", " ", n)
    n = re.sub(r"בע\\?\"?מ|בעמ", " ", n)
    return clean(n.replace('"', " "))


def street_candidates(street):
    """מחזיר גרסאות מנוקות של הכתובת, מהמדויקת לכללית."""
    s = re.sub(r"\(.*?\)", " ", street or "")
    s = s.replace('"', " ")
    s = re.sub(r"\bשד['׳]?\s", "שדרות ", s)
    s = re.sub(r"\bרח['׳]?\s|\bרחוב\s", " ", s)
    s = re.sub(r"\bדר['׳]\s", "דרך ", s)
    s = clean(s)
    out = []
    m = re.search(r"([א-ת][א-ת'׳\- ]*?[א-ת])\s*(\d+)", s)
    if m:
        nm = clean(re.sub(STATION_WORDS, " ", m.group(1)))
        nm = re.sub(r"^(.+?) \1$", r"\1", nm)
        if nm:
            out.append(f"{nm} {m.group(2)}")
            out.append(nm)
    rest = clean(re.sub(STATION_WORDS, " ", re.sub(r"[-–,]", " ", s)))
    rest = clean(re.sub(r"\d+", " ", rest))
    if rest and rest not in out and len(rest) > 2:
        out.append(rest)
    return out[:3]


# ---------- Geocoding ----------
class Geocoder:
    def __init__(self):
        self.cache = read_json(GEOCACHE, {})
        self.last = 0.0
        repo = os.environ.get("GITHUB_REPOSITORY", "personal")
        self.ua = f"pazomat-stations-personal/1.0 (https://github.com/{repo})"
        self.requests = 0

    def _query(self, params):
        key = json.dumps(params, ensure_ascii=False, sort_keys=True)
        if key in self.cache:
            return self.cache[key]
        wait = 1.1 - (time.time() - self.last)
        if wait > 0:
            time.sleep(wait)
        q = dict(params, format="jsonv2", countrycodes="il", limit=1, **{"accept-language": "he"})
        url = NOMINATIM + "?" + urllib.parse.urlencode(q)
        req = urllib.request.Request(url, headers={"User-Agent": self.ua})
        try:
            with urllib.request.urlopen(req, timeout=20) as resp:
                data = json.load(resp)
        except Exception as e:  # תקלה זמנית – לא שומרים במטמון
            print(f"  ! Geocoding נכשל ({e})", file=sys.stderr)
            self.last = time.time()
            return None
        finally:
            self.requests += 1
        self.last = time.time()
        res = None
        if data:
            d = data[0]
            res = {"lat": round(float(d["lat"]), 6), "lng": round(float(d["lon"]), 6),
                   "type": d.get("addresstype") or d.get("type", "")}
        self.cache[key] = res
        return res

    def locate(self, s):
        street, city, name = s["street"], s["city"], s["name"]
        center = self._query({"q": city}) if city else None
        cands = street_candidates(street)
        tries = [({"street": c, "city": city}) for c in cands]
        tries += [({"q": f"{c}, {city}"}) for c in cands]
        tries.append({"q": f"{clean_name(name)}, {city}"})
        for params in tries:
            r = self._query(params)
            if not r or r["type"] in AREA_TYPES:
                continue
            # בדיקת סבירות: התוצאה חייבת להיות קרובה ליישוב
            if center and km(r, center) > MAX_KM_FROM_CITY:
                continue
            return r["lat"], r["lng"], "exact"
        if center:
            return center["lat"], center["lng"], "approx"
        return None, None, "none"

    def save(self):
        write_json(GEOCACHE, self.cache)


# ---------- עריכות ידניות ----------
def read_overrides():
    result, bad = {}, []
    if not os.path.exists(OVERRIDES):
        return result, bad
    with open(OVERRIDES, encoding="utf-8") as f:
        for n, line in enumerate(f, 1):
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            try:
                o = json.loads(line)
                sid = o["id"]
            except Exception:
                bad.append(n)
                continue
            cur = result.setdefault(sid, {})
            for k in ("lat", "lng", "hours"):
                if k in o:
                    cur[k] = o[k]
    return result, bad


# ---------- השוואה ----------
def diff(old, new):
    old_by = {s["id"]: s for s in old}
    new_by = {s["id"]: s for s in new}
    added = [s for s in new if s["id"] not in old_by]
    removed = [s for s in old if s["id"] not in new_by]
    changed = []
    for s in new:
        o = old_by.get(s["id"])
        if not o:
            continue
        for f, label in FIELDS_COMPARE:
            if clean(o.get(f)) != clean(s.get(f)):
                changed.append((s, label, o.get(f) or "—", s.get(f) or "—"))
    return added, removed, changed


def report_md(prev_date, new_date, count, added, removed, changed):
    L = [f"# דוח שינויים – {fmt_date(new_date)}", ""]
    L.append(f"השוואה לדוח הקודם: {fmt_date(prev_date) if prev_date else 'אין (טעינה ראשונה)'}")
    L.append(f"מספר תחנות בדוח החדש: {count}")
    L.append("")
    if not prev_date:
        L.append("זו הטעינה הראשונה, ולכן אין השוואה.")
        return "\n".join(L) + "\n"
    if not (added or removed or changed):
        L.append("לא נמצאו שינויים ברשימת התחנות.")
        return "\n".join(L) + "\n"
    if added:
        L += [f"## תחנות חדשות ({len(added)})", ""]
        L += [f"- {s['name']} – {s['street']}, {s['city']}" for s in added]
        L.append("")
    if removed:
        L += [f"## תחנות שהוסרו ({len(removed)})", ""]
        L += [f"- {s['name']} – {s['street']}, {s['city']}" for s in removed]
        L.append("")
    if changed:
        L += [f"## שינויים בפרטים ({len(changed)})", ""]
        L += [f"- {s['name']} ({s['city']}) – {label}: היה \"{old}\", עכשיו \"{new}\"" for s, label, old, new in changed]
        L.append("")
    return "\n".join(L) + "\n"


# ---------- ראשי ----------
def main():
    prev = read_json(STATIONS, {"meta": {}, "stations": []})
    meta = dict(prev.get("meta", {}))
    stations = prev.get("stations", [])
    new_report = False

    files = [f for f in glob.glob(P("upload", "*.xls*")) if not os.path.basename(f).startswith("~")]
    if files:
        parsed = []
        for f in files:
            d, st = parse_excel(f)
            parsed.append((d, os.path.getmtime(f), f, st))
        parsed.sort()
        date, _, path, new_st = parsed[-1]
        h = file_hash(path)
        if h != meta.get("sourceHash"):
            if meta.get("reportDate") and date < meta["reportDate"]:
                print(f"הדוח שהועלה ({date}) ישן יותר מהנוכחי ({meta['reportDate']}) – מדלג.")
            else:
                old_by = {s["id"]: s for s in stations}
                for s in new_st:  # שמירת קואורדינטות קיימות אם הכתובת לא השתנתה
                    o = old_by.get(s["id"])
                    if o and o.get("street") == s["street"] and o.get("city") == s["city"] and o.get("geo") not in (None, "none"):
                        s["lat"], s["lng"], s["geo"] = o["lat"], o["lng"], o["geo"]
                added, removed, changed = diff(stations, new_st)
                md = report_md(meta.get("reportDate"), date, len(new_st), added, removed, changed)
                with open(CHANGES, "w", encoding="utf-8") as fh:
                    fh.write(md)
                hist = open(HISTORY, encoding="utf-8").read() if os.path.exists(HISTORY) else ""
                with open(HISTORY, "w", encoding="utf-8") as fh:
                    fh.write(md.replace("# דוח", "## דוח", 1).replace("\n## ", "\n### ") + "\n---\n\n" + hist)
                stations = new_st
                meta.update(reportDate=date, sourceHash=h, sourceFile=os.path.basename(path))
                new_report = True
                print(f"דוח חדש מתאריך {date}: {len(new_st)} תחנות, "
                      f"{len(added)} חדשות, {len(removed)} הוסרו, {len(changed)} שינויים")
        os.makedirs(P("upload", "archive"), exist_ok=True)
        for d, _, f, _ in parsed:
            dest = P("upload", "archive", f"report-{d}{os.path.splitext(f)[1]}")
            if os.path.exists(dest):
                os.remove(dest)
            shutil.move(f, dest)

    if not stations:
        raise SystemExit("אין נתוני תחנות. העלה קובץ Excel לתיקייה upload/.")

    # Geocoding
    if not os.environ.get("NO_GEOCODE"):
        g = Geocoder()
        if meta.get("geoVersion") != GEO_VERSION:
            todo = list(stations)
            meta["geoVersion"] = GEO_VERSION
        else:
            todo = [s for s in stations if s.get("geo") in (None, "none")]
        if todo:
            print(f"מאתר מיקום ל-{len(todo)} תחנות...")
        for i, s in enumerate(todo, 1):
            lat, lng, q = g.locate(s)
            s["lat"], s["lng"], s["geo"] = lat, lng, q
            if i % 10 == 0:
                g.save()
        g.save()
        print(f"בקשות Geocoding: {g.requests}")

    for s in stations:
        s.setdefault("lat", None)
        s.setdefault("lng", None)
        s.setdefault("geo", "none")

    write_json(STATIONS, {"meta": meta, "stations": stations})

    # החלת עריכות ידניות
    overrides, bad = read_overrides()
    final = []
    for s in stations:
        s2 = {k: s[k] for k in ("id", "name", "city", "street", "phone", "services", "lat", "lng", "geo")}
        o = overrides.get(s["id"], {})
        if o.get("lat") is not None and o.get("lng") is not None:
            s2["lat"], s2["lng"], s2["geo"] = o["lat"], o["lng"], "manual"
        if "hours" in o:
            s2["hours"] = o["hours"]
        final.append(s2)

    # רשימת תחנות שלא אותרו
    missing = [s for s in final if s["geo"] == "none"]
    approx = [s for s in final if s["geo"] == "approx"]
    L = ["# תחנות שלא אותרו במדויק", ""]
    L.append("אפשר לקבוע להן מיקום באפליקציה: פותחים את התחנה ← עריכה ← קבע מיקום במפה.")
    L.append("")
    L += [f"## ללא מיקום ({len(missing)})", ""] + [f"- {s['name']} – {s['street']}, {s['city']}" for s in missing] + [""]
    L += [f"## מיקום משוער לפי יישוב ({len(approx)})", ""] + [f"- {s['name']} – {s['street']}, {s['city']}" for s in approx] + [""]
    if bad:
        L += ["## שורות לא תקינות בקובץ העריכות", "", f"שורות: {', '.join(map(str, bad))}", ""]
    with open(NOTFOUND, "w", encoding="utf-8") as fh:
        fh.write("\n".join(L))

    # בניית index.html
    cfg = read_json(CONFIG, {})
    data = {
        "meta": {
            "reportDate": meta.get("reportDate"),
            "builtAt": dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%MZ"),
            "sourceName": cfg.get("sourceName", ""),
            "sourceUrl": cfg.get("sourceUrl", ""),
            "count": len(final),
        },
        "stations": final,
    }
    blob = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    tpl = open(TEMPLATE, encoding="utf-8").read()
    css = open(P("src", "vendor", "leaflet.css"), encoding="utf-8").read()
    js = open(P("src", "vendor", "leaflet.js"), encoding="utf-8").read()
    html = (tpl.replace("/*{{LEAFLET_CSS}}*/", css)
               .replace("/*{{LEAFLET_JS}}*/", js)
               .replace("{{DATA}}", blob))
    with open(OUT, "w", encoding="utf-8") as fh:
        fh.write(html)
    print(f"index.html נבנה: {len(final)} תחנות, {len(missing)} ללא מיקום, {len(approx)} משוערות.")

    out = os.environ.get("GITHUB_OUTPUT")
    if out:
        with open(out, "a") as fh:
            fh.write(f"new_report={'true' if new_report else 'false'}\n")
            fh.write(f"report_date={fmt_date(meta['reportDate']) if meta.get('reportDate') else ''}\n")


if __name__ == "__main__":
    main()
