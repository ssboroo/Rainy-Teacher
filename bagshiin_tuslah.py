"""Багшийн Туслах — offline Windows desktop app for Mongolian teachers."""
from __future__ import annotations

import csv
import base64
import html
import json
import mimetypes
import os
import shutil
import sqlite3
import tempfile
import webbrowser
from datetime import date, datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

APP_NAME = "Багшийн Туслах"
APP_VERSION = "1.2.0"
BG = "#F3F6FB"
NAV = "#13213A"
BLUE = "#356AE6"
INK = "#17243A"
MUTED = "#718096"
WHITE = "#FFFFFF"
GREEN = "#1B8B67"
RED = "#D35454"


def data_dir() -> Path:
    root = Path(os.environ.get("APPDATA", Path.home())) / "BagshiinTuslah"
    root.mkdir(parents=True, exist_ok=True)
    return root


class Store:
    def __init__(self, path: Path | None = None):
        self.path = path or data_dir() / "bagshiin_tuslah.db"
        self.db = sqlite3.connect(self.path)
        self.db.row_factory = sqlite3.Row
        self.db.execute("PRAGMA foreign_keys=ON")
        self.db.executescript("""
        CREATE TABLE IF NOT EXISTS classes (
          id INTEGER PRIMARY KEY, name TEXT NOT NULL UNIQUE, school_year TEXT NOT NULL DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS students (
          id INTEGER PRIMARY KEY, class_id INTEGER NOT NULL REFERENCES classes(id) ON DELETE CASCADE,
          code TEXT NOT NULL DEFAULT '', name TEXT NOT NULL, parent TEXT NOT NULL DEFAULT '',
          phone TEXT NOT NULL DEFAULT '', note TEXT NOT NULL DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS attendance (
          id INTEGER PRIMARY KEY, student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
          day TEXT NOT NULL, status TEXT NOT NULL DEFAULT 'Ирсэн', note TEXT NOT NULL DEFAULT '',
          UNIQUE(student_id, day)
        );
        CREATE TABLE IF NOT EXISTS grades (
          id INTEGER PRIMARY KEY, student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
          subject TEXT NOT NULL, title TEXT NOT NULL DEFAULT '', score REAL NOT NULL,
          max_score REAL NOT NULL DEFAULT 100, day TEXT NOT NULL, note TEXT NOT NULL DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS certificate_templates (
          id INTEGER PRIMARY KEY, name TEXT NOT NULL UNIQUE, style TEXT NOT NULL,
          title TEXT NOT NULL, subtitle TEXT NOT NULL, body TEXT NOT NULL,
          logo BLOB, logo_mime TEXT NOT NULL DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS certificates (
          id INTEGER PRIMARY KEY, template_id INTEGER, template_name TEXT NOT NULL,
          style TEXT NOT NULL, title TEXT NOT NULL, subtitle TEXT NOT NULL, body TEXT NOT NULL,
          recipient TEXT NOT NULL, achievement TEXT NOT NULL DEFAULT '', teacher TEXT NOT NULL DEFAULT '',
          organization TEXT NOT NULL DEFAULT '', issue_date TEXT NOT NULL, logo BLOB, logo_mime TEXT NOT NULL DEFAULT '',
          saved_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
          FOREIGN KEY(template_id) REFERENCES certificate_templates(id) ON DELETE SET NULL
        );
        CREATE TABLE IF NOT EXISTS exams (
          id INTEGER PRIMARY KEY, class_id INTEGER REFERENCES classes(id) ON DELETE SET NULL,
          title TEXT NOT NULL, subject TEXT NOT NULL, instructions TEXT NOT NULL DEFAULT '',
          created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS exam_questions (
          id INTEGER PRIMARY KEY, exam_id INTEGER NOT NULL REFERENCES exams(id) ON DELETE CASCADE,
          position INTEGER NOT NULL, prompt TEXT NOT NULL, options TEXT NOT NULL,
          correct_index INTEGER NOT NULL, points REAL NOT NULL DEFAULT 1
        );
        CREATE INDEX IF NOT EXISTS idx_student_class ON students(class_id);
        CREATE INDEX IF NOT EXISTS idx_attendance_day ON attendance(day);
        """)
        self._seed_certificate_templates()
        self.db.commit()

    def _seed_certificate_templates(self):
        defaults = [
            ("Алтан Эрдэнэ", "gold", "ӨРГӨМЖЛӨЛ", "Эрдмийн амжилтаа үнэлүүлэв", "{achievement} амжилт гаргасныг үнэлж энэхүү өргөмжлөлийг гардуулав."),
            ("Хөх Тэнгэр", "sky", "БАТЛАМЖ", "Мэдлэг, хичээл зүтгэлийн батламж", "{achievement} амжилтыг гаргасныг батламжилж энэхүү бичгийг олгов."),
            ("Улаан Тамга", "ruby", "ТАЛАРХАЛ", "Хичээл зүтгэл, оролцоонд", "{achievement} ажилд идэвх санаачилгатай оролцсонд талархал илэрхийлэв."),
            ("Ногоон Хээ", "jade", "ӨРГӨМЖЛӨЛ", "Шилдэг сурагч", "{achievement} амжилтаар манлайлан суралцсанд энэхүү өргөмжлөлийг гардуулав."),
            ("Нил ягаан Эрдэм", "violet", "ТУСГАЙ БАЙР", "Авьяас, чадвараа нээн харуулсанд", "{achievement} амжилт гаргасныг үнэлж тусгай өргөмжлөл гардуулав."),
            ("Зэс Хүрээ", "copper", "АВАРГЫН ӨРГӨМЖЛӨЛ", "Спортын тэмцээний амжилтад", "{achievement} амжилт гаргаж, хурд хүчээрээ манлайлсанд гардуулав."),
            ("Шөнийн Тэнгэр", "midnight", "ОЛИМПИАДЫН БАТЛАМЖ", "Оюун ухаан, мэдлэгээ сорьсонд", "{achievement} амжилт үзүүлж, мэдлэг чадвараа батлан харуулсанд олгов."),
            ("Номин Далай", "teal", "СУДАЛГААНЫ ӨРГӨМЖЛӨЛ", "Шинэ санаа, нээлтэд", "{achievement} бүтээл, судалгаагаар мэдлэгийн санг баяжуулсанд гардуулав."),
            ("Шүрэн Цэцэг", "coral", "УРЛАГИЙН ӨРГӨМЖЛӨЛ", "Уран бүтээлийн амжилтад", "{achievement} уран бүтээлээр авьяас билгээ нээн харуулсанд гардуулав."),
            ("Хөх Индранил", "indigo", "МАНЛАЙЛАГЧ СУРАГЧ", "Эерэг үлгэр дуурайлалд", "{achievement} манлайлал гаргаж, анги хамт олноо үлгэрлэн оролцсонд олгов."),
            ("Маргад Эрдэнэ", "emerald", "НОМЫН САЙН НӨХӨР", "Уншихуйн амжилтад", "{achievement} уншиж судлан, эрдэм мэдлэгээ тасралтгүй ахиулсанд гардуулав."),
            ("Мөнгөн Хээ", "silver", "ХИЧЭЭЛИЙН ИРЦИЙН БАТЛАМЖ", "Хариуцлагатай суралцсанд", "Хичээлдээ тогтмол хамрагдаж, суралцах хариуцлагаа эрхэмлэн {achievement} амжилт үзүүлснийг үнэлэв."),
            ("Нарны Туяа", "amber", "САЙН ҮЙЛСИЙН ӨРГӨМЖЛӨЛ", "Найрсаг, тусч зан чанарт", "{achievement} сайн үйлс, тусч сэтгэлээр хамт олондоо үлгэрлэснийг үнэлэв."),
            ("Сарнай", "rose", "ТАЛАРХАЛ", "Сэтгэл зүтгэл, дэмжлэгт", "{achievement} хувь нэмэр оруулж, хамт олны амжилтад үнэтэй дэмжлэг үзүүлсэнд талархав."),
            ("Хөвсгөл Нуур", "ocean", "БАЙГАЛЬ ХАМГААЛАГЧ", "Эх дэлхийгээ хайрлан хамгаалсанд", "{achievement} санаачилга гарган, байгаль орчноо хамгаалах ажилд идэвхтэй оролцсонд олгов."),
            ("Нил Ягаан Хилэн", "plum", "ШИЛДЭГ БАГ", "Хамтын ажиллагааны амжилтад", "{achievement} хамтын амжилт гаргаж, эв нэгдэлтэй ажилласныг үнэлэв."),
            ("Гааны Сэрүүн", "mint", "ДУРСГАЛЫН БИЧИГ", "Эрдмийн баярын дурсгалд", "Эрдмийн баяраа амжилттай тэмдэглэж, {achievement} үйлсийг бүтээсний дурсгалд гардуулав."),
            ("Хааны Хөх", "navy", "АНГИЙН АВАРГА", "Тууштай хөдөлмөр, амжилтад", "{achievement} амжилт гарган анги хамт олноо манлайлсныг үнэлж олгов."),
            ("Хүрэл Тамга", "bronze", "БҮТЭЭЛЧ САНААНЫ ӨРГӨМЖЛӨЛ", "Шинийг санаачилсанд", "{achievement} санаачилга, бүтээлч шийдлээрээ бусдыгаа уриалан оролцсонд гардуулав."),
            ("Сувдан Гэгээ", "pearl", "ОНЦГОЙ АМЖИЛТЫН ӨРГӨМЖЛӨЛ", "Өндөр амжилт гаргасанд", "{achievement} онцгой амжилтаар хамт олныхоо бахархал болсонд гардуулав."),
        ]
        for item in defaults:
            self.db.execute("INSERT OR IGNORE INTO certificate_templates(name,style,title,subtitle,body) VALUES(?,?,?,?,?)", item)

    def all(self, sql: str, args=()):
        return self.db.execute(sql, args).fetchall()

    def run(self, sql: str, args=()):
        cur = self.db.execute(sql, args)
        self.db.commit()
        return cur

    def close(self):
        self.db.close()


CERT_DESIGNS = {
    "gold":    {"colors":("#a77920","#f5e8bc","#fffaf0"),"motif":"✦　❖　✦","border":"double","layout":"classic","pattern":"ulzii"},
    "sky":     {"colors":("#2469a8","#d6eafd","#f4faff"),"motif":"☼　✧　☼","border":"solid","layout":"ribbon","pattern":"cloud"},
    "ruby":    {"colors":("#9e334b","#f3d7dc","#fff7f8"),"motif":"❧　✦　❧","border":"double","layout":"corners","pattern":"ever"},
    "jade":    {"colors":("#287c63","#d2eee3","#f4fcf8"),"motif":"❋　❖　❋","border":"ridge","layout":"lattice","pattern":"alkhan"},
    "violet":  {"colors":("#7056a4","#e4dcf3","#faf8ff"),"motif":"✧　✦　✧","border":"double","layout":"seal","pattern":"lotus"},
    "copper":  {"colors":("#a85c37","#f1dccb","#fff9f3"),"motif":"❖　◈　❖","border":"groove","layout":"corners","pattern":"ever"},
    "midnight":{"colors":("#243b68","#d8e1f2","#f6f8fd"),"motif":"☾　✦　☽","border":"solid","layout":"night","pattern":"ulzii"},
    "teal":    {"colors":("#167b83","#d2eef0","#f3fcfc"),"motif":"≈　✧　≈","border":"double","layout":"wave","pattern":"cloud"},
    "coral":   {"colors":("#c35a4c","#f8dfd7","#fff8f5"),"motif":"❀　✦　❀","border":"ridge","layout":"floral","pattern":"lotus"},
    "indigo":  {"colors":("#4b55a0","#e0e3f8","#f8f8ff"),"motif":"✧　◇　✧","border":"double","layout":"ribbon","pattern":"alkhan"},
    "emerald": {"colors":("#26734b","#d8eddb","#f5fbf3"),"motif":"❈　❖　❈","border":"solid","layout":"lattice","pattern":"ulzii"},
    "silver":  {"colors":("#687786","#e2e8ed","#fbfcfd"),"motif":"◇　✧　◇","border":"double","layout":"minimal","pattern":"alkhan"},
    "amber":   {"colors":("#b87519","#f7e7c7","#fffaf1"),"motif":"☀　❖　☀","border":"ridge","layout":"sun","pattern":"ever"},
    "rose":    {"colors":("#a94f78","#f4dce9","#fff8fc"),"motif":"❀　❧　❀","border":"double","layout":"floral","pattern":"lotus"},
    "ocean":   {"colors":("#176b91","#d4eaf2","#f3fbff"),"motif":"≋　✦　≋","border":"solid","layout":"wave","pattern":"cloud"},
    "plum":    {"colors":("#75456f","#ecddec","#fcf7fc"),"motif":"❦　✧　❦","border":"groove","layout":"seal","pattern":"ever"},
    "mint":    {"colors":("#46856e","#dcefe6","#f6fcf8"),"motif":"❋　◇　❋","border":"ridge","layout":"lattice","pattern":"alkhan"},
    "navy":    {"colors":("#203d5d","#d9e4ef","#f6f9fc"),"motif":"⚜　✦　⚜","border":"double","layout":"classic","pattern":"ulzii"},
    "bronze":  {"colors":("#8d623c","#ede0cd","#fcf8f1"),"motif":"◈　❖　◈","border":"groove","layout":"corners","pattern":"ever"},
    "pearl":   {"colors":("#837248","#eee9d6","#fffdf5"),"motif":"✺　◇　✺","border":"double","layout":"minimal","pattern":"lotus"},
}

CERT_STYLE_LABELS = {
    "Алтан Эрдэнэ":"gold", "Хөх Тэнгэр":"sky", "Улаан Тамга":"ruby", "Ногоон Хээ":"jade", "Нил ягаан Эрдэм":"violet",
    "Зэс Хүрээ":"copper", "Шөнийн Тэнгэр":"midnight", "Номин Далай":"teal", "Шүрэн Цэцэг":"coral", "Хөх Индранил":"indigo",
    "Маргад Эрдэнэ":"emerald", "Мөнгөн Хээ":"silver", "Нарны Туяа":"amber", "Сарнай":"rose", "Хөвсгөл Нуур":"ocean",
    "Нил Ягаан Хилэн":"plum", "Гааны Сэрүүн":"mint", "Хааны Хөх":"navy", "Хүрэл Тамга":"bronze", "Сувдан Гэгээ":"pearl",
}


MONGOLIAN_PATTERN_PATHS = {
    # Angular continuous knots inspired by ulzii and alhan khее.
    "ulzii": "M1 18h8V8h10V1h18v7h10v10h8v18h-8v10H37v7H19v-7H9V36H1z M19 8v10h10v10h10V18h8 M9 28h10v10h10v10",
    "alkhan": "M1 1h12v12h12V1h12v24h12V13h12v12h12V1h12 M1 49h12V37h12v12h12V25h12v12h12v12h12",
    "ever": "M2 25c0-13 18-18 25-8 6 8-1 18-9 14-5-3-2-10 3-9 M35 25c0-13 18-18 25-8 6 8-1 18-9 14-5-3-2-10 3-9 M2 3h10v8h10 M35 39h10v8h10",
    "cloud": "M1 35c6 0 6-12 13-12s7 12 14 12 7-12 14-12 7 12 14 12 7-12 14-12 7 12 14 12 M1 43h14m4 0h14m4 0h14m4 0h14",
    "lotus": "M1 38c8-3 10-13 18-19 3 9 3 17 0 24 7-8 13-13 18-15 5 2 11 7 18 15-3-7-3-15 0-24 8 6 10 16 18 19 M1 45h72",
}


def mongolian_border_svg(pattern: str, accent: str) -> str:
    """Create a self-contained, printer-friendly Mongolian geometric border."""
    path = MONGOLIAN_PATTERN_PATHS.get(pattern, MONGOLIAN_PATTERN_PATHS["ulzii"])
    return f"""<svg class='mongol-border' aria-hidden='true' viewBox='0 0 1000 700' preserveAspectRatio='none' xmlns='http://www.w3.org/2000/svg'>
    <defs><pattern id='hee' width='76' height='50' patternUnits='userSpaceOnUse'><path d='{path}' fill='none' stroke='{accent}' stroke-width='2' stroke-linecap='round' stroke-linejoin='round' opacity='.76'/></pattern></defs>
    <rect x='17' y='17' width='966' height='666' fill='none' stroke='{accent}' stroke-width='5'/>
    <rect x='27' y='27' width='946' height='646' fill='none' stroke='{accent}' stroke-width='1.5' opacity='.72'/>
    <rect x='39' y='39' width='922' height='34' fill='url(#hee)'/>
    <rect x='39' y='627' width='922' height='34' fill='url(#hee)'/>
    <g transform='translate(39 661) rotate(-90)'><rect width='588' height='34' fill='url(#hee)'/></g>
    <g transform='translate(961 39) rotate(90)'><rect width='588' height='34' fill='url(#hee)'/></g>
    <g fill='white' stroke='{accent}' stroke-width='3'><circle cx='56' cy='56' r='12'/><circle cx='944' cy='56' r='12'/><circle cx='56' cy='644' r='12'/><circle cx='944' cy='644' r='12'/></g>
    <g fill='{accent}'><path d='M56 49l2 5 5 2-5 2-2 5-2-5-5-2 5-2z'/><path d='M944 49l2 5 5 2-5 2-2 5-2-5-5-2 5-2z'/><path d='M56 637l2 5 5 2-5 2-2 5-2-5-5-2 5-2z'/><path d='M944 637l2 5 5 2-5 2-2 5-2-5-5-2 5-2z'/></g>
    </svg>"""


def certificate_html(record: dict) -> str:
    design = CERT_DESIGNS.get(record.get("style", "gold"), CERT_DESIGNS["gold"])
    accent, pale, paper = design["colors"]
    motif, border, layout = design["motif"], design["border"], design["layout"]
    border_svg = mongolian_border_svg(design["pattern"], accent)
    vals = {k: str(record.get(k, "") or "") for k in ("recipient", "achievement", "teacher", "organization", "issue_date", "title", "subtitle", "body")}
    body = vals["body"].replace("{name}", vals["recipient"]).replace("{achievement}", vals["achievement"])
    logo = ""
    logo_bytes = record.get("logo")
    if logo_bytes:
        logo_mime = record.get("logo_mime") or "image/png"
        logo = f"<img class='logo' src='data:{html.escape(logo_mime)};base64,{base64.b64encode(logo_bytes).decode('ascii')}'>"
    return f"""<!doctype html><html lang='mn'><meta charset='utf-8'><title>{html.escape(vals['title'])} · {html.escape(vals['recipient'])}</title>
    <style>
    @page{{size:A4 landscape;margin:0}}*{{box-sizing:border-box}}body{{margin:0;background:#e9edf3;color:#27344a;font-family:'Georgia','Times New Roman',serif}}
    .sheet{{width:297mm;height:210mm;margin:12px auto;padding:12mm;background:{paper};position:relative;overflow:hidden;box-shadow:0 8px 36px #9aa5b5}}
    .frame{{height:100%;border:3px {border} {accent};outline:1px solid {accent};outline-offset:-9px;padding:12mm 17mm;text-align:center;position:relative;background:radial-gradient(ellipse at 50% 0%,{pale} 0,transparent 42%),{paper}}}
    .mongol-border{{position:absolute;inset:0;width:100%;height:100%;pointer-events:none;z-index:0}}
    .frame>:not(.mongol-border){{position:relative;z-index:1}}
    .frame:before,.frame:after{{content:'{motif}';position:absolute;left:0;right:0;color:{accent};font-size:16px;letter-spacing:7px;opacity:.82}}
    .frame:before{{top:6mm}}.frame:after{{bottom:6mm}}
    .logo{{max-width:22mm;max-height:22mm;object-fit:contain;margin:0 auto 3mm;display:block}}
    .ornament{{color:{accent};font-size:17px;letter-spacing:12px;margin:0 0 2mm}}
    h1{{color:{accent};font-size:30pt;letter-spacing:5px;margin:2mm 0 1mm;font-weight:700}}
    .subtitle{{color:#627086;font-family:'Segoe UI',Arial,sans-serif;font-size:12pt;letter-spacing:2px;text-transform:uppercase;margin:2mm 0 7mm}}
    .recipient{{font-size:25pt;font-weight:bold;color:#1c2940;border-bottom:1px solid {accent};display:inline-block;min-width:95mm;padding:0 12mm 3mm;margin:0 0 6mm}}
    .bodytext{{font-family:'Segoe UI',Arial,sans-serif;font-size:14pt;line-height:1.7;max-width:218mm;margin:0 auto;color:#43516a}}
    .award{{color:{accent};font-weight:700}}
    .footer{{position:absolute;left:24mm;right:24mm;bottom:21mm;display:flex;justify-content:space-between;font:11pt 'Segoe UI',Arial,sans-serif;color:#4d5c72}}
    .footer span{{min-width:58mm;border-top:1px solid #aab4c2;padding-top:3mm}}
    .layout-ribbon .subtitle{{display:inline-block;background:{accent};color:#fff;padding:2.5mm 12mm;border-radius:2mm}}
    .layout-corners{{border-style:none;outline:2px solid {accent};outline-offset:-10px}}.layout-corners:before,.layout-corners:after{{font-size:21px;letter-spacing:20px}}
    .layout-lattice{{background:repeating-linear-gradient(45deg,{paper} 0 18px,{pale} 19px 20px,{paper} 21px 38px)}}
    .layout-seal .ornament{{display:inline-block;border:1px solid {accent};border-radius:50%;padding:5mm 8mm;background:{pale}}}
    .layout-night{{background:radial-gradient(ellipse at 50% 0%,{pale},transparent 65%),linear-gradient(135deg,{paper},#fff)}}
    .layout-wave{{background:repeating-linear-gradient(170deg,{paper} 0 23mm,{pale} 23.5mm 24mm,{paper} 24.5mm 42mm)}}
    .layout-floral .recipient{{border-bottom:3px double {accent};font-style:italic}}.layout-floral .ornament{{font-size:21px}}
    .layout-minimal{{border-width:1px;outline:none;padding:17mm 22mm}}.layout-minimal:before,.layout-minimal:after{{display:none}}.layout-minimal .ornament{{font-size:12px;letter-spacing:18px}}
    .layout-sun{{background:radial-gradient(circle at 50% 4%,{pale} 0 22%,transparent 50%),{paper}}}.layout-sun .ornament{{font-size:24px}}
    .print{{position:fixed;right:20px;top:18px;background:{accent};color:white;border:0;border-radius:8px;padding:12px 20px;font:14px Arial;cursor:pointer}}
    @media print{{body{{background:white}}.sheet{{margin:0;box-shadow:none}}.print{{display:none}}}}
    @media screen and (max-width:900px){{.sheet{{transform-origin:top left;transform:scale(.65);margin:8px}}}}
    </style><button class='print' onclick='window.print()'>Хэвлэх / PDF хадгалах</button>
    <main class='sheet'><section class='frame layout-{layout} style-{html.escape(str(record.get('style', 'gold')))}'>{border_svg}{logo}<div class='ornament'>{motif}</div><h1>{html.escape(vals['title'])}</h1>
    <div class='subtitle'>{html.escape(vals['subtitle'])}</div><div class='recipient'>{html.escape(vals['recipient'])}</div>
    <div class='bodytext'>{html.escape(body).replace(chr(10), '<br>')}</div>
    <div class='footer'><span>{html.escape(vals['organization'])}</span><span>{html.escape(vals['issue_date'])}</span><span>{html.escape(vals['teacher'])}</span></div>
    </section></main></html>"""


def exam_html(exam: dict, questions: list[dict], answer_key: bool = False) -> str:
    title = html.escape(str(exam.get("title", "Шалгалт")))
    subject = html.escape(str(exam.get("subject", "")))
    instruction = html.escape(str(exam.get("instructions", "")))
    rows = []
    for index, q in enumerate(questions, 1):
        prompt = html.escape(str(q.get("prompt", "")))
        options = q.get("options", [])
        points = html.escape(str(q.get("points", 1)))
        answer = int(q.get("correct_index", 0))
        choices = "".join(f"<div class='choice'>{chr(65+i)}. {html.escape(str(option))}</div>" for i, option in enumerate(options) if str(option).strip())
        key = f"<b class='key'>Зөв хариулт: {chr(65+answer)} · {points} оноо</b>" if answer_key else f"<span class='points'>{points} оноо</span>"
        rows.append(f"<article><div class='q'><b>{index}.</b> {prompt} {key}</div><div class='choices'>{choices}</div></article>")
    key_title = "Хариултын түлхүүр" if answer_key else "Сурагчийн хувилбар"
    return f"""<!doctype html><html lang='mn'><meta charset='utf-8'><title>{title} · {key_title}</title><style>
    @page{{size:A4 portrait;margin:16mm}}body{{font:12pt 'Segoe UI',Arial,sans-serif;color:#1d2b42;max-width:900px;margin:30px auto;padding:30px;background:white}}
    .head{{border-bottom:3px solid #356ae6;padding-bottom:14px;margin-bottom:20px}}h1{{font-size:22pt;margin:0 0 6px;color:#17315f}}.meta{{color:#5b6e8b}}.instructions{{background:#f2f6ff;border-left:4px solid #356ae6;padding:10px 14px;margin:16px 0}}
    article{{padding:12px 0;border-bottom:1px solid #dfe5ef;break-inside:avoid}}.q{{font-size:12pt;line-height:1.6}}.choices{{display:grid;grid-template-columns:1fr 1fr;gap:6px 18px;margin:9px 0 0 25px;color:#3e4e68}}.points{{color:#75839a;font-size:10pt;white-space:nowrap}}.key{{color:#14805e;margin-left:8px;white-space:nowrap}}.print{{position:fixed;right:20px;top:18px;background:#356ae6;color:white;border:0;border-radius:8px;padding:12px 20px;font-size:14px;cursor:pointer}}
    @media print{{.print{{display:none}}}}@media screen{{body{{box-shadow:0 6px 30px #d6dce6}}}}</style>
    <button class='print' onclick='window.print()'>Хэвлэх / PDF хадгалах</button><header class='head'><h1>{title}</h1><div class='meta'>{subject} · {key_title}</div></header>
    <p>Сурагчийн нэр: __________________________________　 Анги: ______________　 Огноо: ______________</p>
    {f'<div class="instructions">{instruction}</div>' if instruction else ''}{''.join(rows)}</html>"""


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} · {APP_VERSION}")
        self.geometry("1180x760")
        self.minsize(980, 650)
        self.configure(bg=BG)
        self.store = Store()
        self.active_class: int | None = None
        self.page = "Нүүр"
        self._style()
        self._layout()
        self.show_page("Нүүр")
        self.protocol("WM_DELETE_WINDOW", self._close)

    def _style(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("Treeview", background=WHITE, fieldbackground=WHITE, foreground=INK,
                    rowheight=34, borderwidth=0, font=("Segoe UI", 10))
        s.configure("Treeview.Heading", background="#EAF0FA", foreground=INK,
                    font=("Segoe UI", 10, "bold"), relief="flat", padding=(8, 9))
        s.map("Treeview", background=[("selected", "#DCE8FF")], foreground=[("selected", INK)])
        s.configure("TCombobox", padding=7, font=("Segoe UI", 10))
        s.configure("TEntry", padding=7, font=("Segoe UI", 10))
        s.configure("TButton", padding=(12, 8), font=("Segoe UI", 10))

    def _layout(self):
        self.sidebar = tk.Frame(self, bg=NAV, width=218)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)
        tk.Label(self.sidebar, text="БАГШИЙН\nТУСЛАХ", bg=NAV, fg=WHITE,
                 font=("Segoe UI", 19, "bold"), justify="left", pady=24).pack(anchor="w", padx=24)
        tk.Label(self.sidebar, text="Хичээлээ цэгцтэй удирд", bg=NAV, fg="#A9B8D2",
                 font=("Segoe UI", 9)).pack(anchor="w", padx=24, pady=(0, 22))
        self.nav_buttons = {}
        for label, icon in [("Нүүр", "⌂"), ("Сурагчид", "♙"), ("Ирц", "✓"), ("Дүн", "▤"),
                            ("Өргөмжлөл", "✦"), ("Шалгалтын материал", "✎"), ("Тайлан ба хэвлэх", "▧")]:
            b = tk.Button(self.sidebar, text=f"  {icon}    {label}", anchor="w", relief="flat",
                          bd=0, bg=NAV, fg="#D5DEED", activebackground="#263956",
                          activeforeground=WHITE, font=("Segoe UI", 10), padx=20, pady=13,
                          command=lambda p=label: self.show_page(p))
            b.pack(fill="x", padx=10, pady=2)
            self.nav_buttons[label] = b
        tk.Frame(self.sidebar, bg="#344460", height=1).pack(fill="x", padx=18, pady=(24, 14))
        tk.Button(self.sidebar, text="  ⤓    Нөөц хуулбар", anchor="w", relief="flat", bd=0,
                  bg=NAV, fg="#D5DEED", activebackground="#263956", activeforeground=WHITE,
                  font=("Segoe UI", 10), padx=20, pady=13, command=self.backup).pack(fill="x", padx=10)
        tk.Label(self.sidebar, text="Офлайн · Таны компьютерт хадгална", bg=NAV, fg="#8FA0BA",
                 font=("Segoe UI", 8), wraplength=175, justify="left").pack(side="bottom", anchor="w", padx=23, pady=20)

        self.main = tk.Frame(self, bg=BG)
        self.main.pack(side="left", fill="both", expand=True)
        top = tk.Frame(self.main, bg=WHITE, height=72)
        top.pack(fill="x")
        top.pack_propagate(False)
        self.page_title = tk.Label(top, text="Нүүр", bg=WHITE, fg=INK, font=("Segoe UI", 17, "bold"))
        self.page_title.pack(side="left", padx=28)
        tk.Label(top, text="АНГИ", bg=WHITE, fg=MUTED, font=("Segoe UI", 8, "bold")).pack(side="right", padx=(8, 24))
        self.class_var = tk.StringVar()
        self.class_combo = ttk.Combobox(top, textvariable=self.class_var, state="readonly", width=18)
        self.class_combo.pack(side="right", pady=17)
        self.class_combo.bind("<<ComboboxSelected>>", self._select_class)
        self._button(top, "+ Анги", self.add_class).pack(side="right", padx=10)
        self.body = tk.Frame(self.main, bg=BG)
        self.body.pack(fill="both", expand=True, padx=26, pady=22)
        self.refresh_classes()

    def refresh_classes(self):
        rows = self.store.all("SELECT * FROM classes ORDER BY name")
        self.class_rows = rows
        self.class_combo["values"] = [r["name"] for r in rows]
        if rows:
            selected = next((r for r in rows if r["id"] == self.active_class), rows[0])
            self.active_class = selected["id"]
            self.class_var.set(selected["name"])
        else:
            self.active_class = None
            self.class_var.set("")

    def _select_class(self, _event=None):
        name = self.class_var.get()
        self.active_class = next((r["id"] for r in self.class_rows if r["name"] == name), None)
        self.show_page(self.page)

    def show_page(self, page: str):
        self.page = page
        self.page_title.config(text=page)
        for name, b in self.nav_buttons.items():
            b.config(bg="#263956" if name == page else NAV, fg=WHITE if name == page else "#D5DEED")
        for child in self.body.winfo_children(): child.destroy()
        if page == "Нүүр": self._dashboard()
        elif page == "Сурагчид": self._students_page()
        elif page == "Ирц": self._attendance_page()
        elif page == "Дүн": self._grades_page()
        elif page == "Өргөмжлөл": self._certificates_page()
        elif page == "Шалгалтын материал": self._exams_page()
        else: self._reports_page()

    def _card(self, parent, title, value, hint, accent=BLUE):
        card = tk.Frame(parent, bg=WHITE, highlightthickness=1, highlightbackground="#E7ECF4")
        tk.Frame(card, bg=accent, width=4).pack(side="left", fill="y")
        content = tk.Frame(card, bg=WHITE, padx=17, pady=14); content.pack(fill="both", expand=True)
        tk.Label(content, text=title, bg=WHITE, fg=MUTED, font=("Segoe UI", 9)).pack(anchor="w")
        tk.Label(content, text=value, bg=WHITE, fg=INK, font=("Segoe UI", 23, "bold")).pack(anchor="w", pady=3)
        tk.Label(content, text=hint, bg=WHITE, fg=MUTED, font=("Segoe UI", 8)).pack(anchor="w")
        return card

    def _section(self, parent, title, subtitle=None):
        row = tk.Frame(parent, bg=BG); row.pack(fill="x", pady=(0, 14))
        tk.Label(row, text=title, bg=BG, fg=INK, font=("Segoe UI", 14, "bold")).pack(anchor="w")
        if subtitle: tk.Label(row, text=subtitle, bg=BG, fg=MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(3, 0))

    def _button(self, parent, text, command, primary=True):
        return tk.Button(parent, text=text, command=command, relief="flat", bd=0,
                         bg=BLUE if primary else WHITE, fg=WHITE if primary else INK,
                         activebackground="#2856C4" if primary else "#EAF0FA",
                         activeforeground=WHITE if primary else INK, padx=14, pady=9,
                         font=("Segoe UI", 9, "bold" if primary else "normal"), cursor="hand2")

    def _dashboard(self):
        self._section(self.body, "Сайн байна уу 👋", "Анги, ирц, дүнгээ нэг дороос удирдаарай.")
        if not self.active_class:
            box = tk.Frame(self.body, bg=WHITE, padx=28, pady=26); box.pack(fill="x", pady=10)
            tk.Label(box, text="Эхлээд ангиа үүсгээрэй", bg=WHITE, fg=INK, font=("Segoe UI", 14, "bold")).pack(anchor="w")
            tk.Label(box, text="Анги үүсгэсний дараа сурагч бүртгэж, ирц болон дүн хөтөлж эхэлнэ.", bg=WHITE, fg=MUTED, font=("Segoe UI", 10)).pack(anchor="w", pady=8)
            self._button(box, "+ Анги үүсгэх", self.add_class).pack(anchor="w", pady=(7, 0))
            return
        cid = self.active_class
        total = self.store.all("SELECT COUNT(*) n FROM students WHERE class_id=?", (cid,))[0]["n"]
        today = date.today().isoformat()
        marked = self.store.all("SELECT COUNT(*) n FROM attendance a JOIN students s ON s.id=a.student_id WHERE s.class_id=? AND a.day=?", (cid, today))[0]["n"]
        present = self.store.all("SELECT COUNT(*) n FROM attendance a JOIN students s ON s.id=a.student_id WHERE s.class_id=? AND a.day=? AND a.status='Ирсэн'", (cid, today))[0]["n"]
        cards = tk.Frame(self.body, bg=BG); cards.pack(fill="x", pady=(0, 22))
        for title, val, hint, color in [("Нийт сурагч", str(total), "Сонгосон ангид", BLUE),
                                       ("Өнөөдрийн ирц", f"{marked}/{total}" if total else "0/0", f"Ирсэн: {present}", GREEN),
                                       ("Анги", self.class_var.get(), "Одоогийн сонголт", "#7A58C1")]:
            self._card(cards, title, val, hint, color).pack(side="left", fill="both", expand=True, padx=(0, 12))
        self._section(self.body, "Шуурхай үйлдэл", "Түгээмэл ажлаа шууд эхлүүлээрэй.")
        quick = tk.Frame(self.body, bg=BG); quick.pack(fill="x")
        for title, desc, page in [("Сурагч нэмэх", "Шинэ сурагчийн мэдээлэл бүртгэх", "Сурагчид"),
                                  ("Өнөөдрийн ирц", "Ирц тэмдэглэж хадгалах", "Ирц"),
                                  ("Дүн оруулах", "Шалгалт, даалгаврын оноо нэмэх", "Дүн"),
                                  ("Тайлан хэвлэх", "Жагсаалт ба тайлан хэвлэх", "Тайлан ба хэвлэх")]:
            tile = tk.Frame(quick, bg=WHITE, padx=18, pady=17, highlightthickness=1, highlightbackground="#E7ECF4")
            tile.pack(side="left", fill="both", expand=True, padx=(0, 12))
            tk.Label(tile, text=title, bg=WHITE, fg=INK, font=("Segoe UI", 11, "bold")).pack(anchor="w")
            tk.Label(tile, text=desc, bg=WHITE, fg=MUTED, font=("Segoe UI", 8), wraplength=145, justify="left").pack(anchor="w", pady=(6, 12))
            self._button(tile, "Нээх →", lambda p=page: self.show_page(p), False).pack(anchor="w")

    def _tree(self, parent, columns, headings, widths=None):
        frame = tk.Frame(parent, bg=WHITE, highlightthickness=1, highlightbackground="#E7ECF4")
        frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(frame, columns=columns, show="headings")
        for i, (col, head) in enumerate(zip(columns, headings)):
            tree.heading(col, text=head)
            tree.column(col, width=(widths[i] if widths else 130), anchor="w")
        y = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=y.set)
        tree.pack(side="left", fill="both", expand=True); y.pack(side="right", fill="y")
        return tree

    def _students_page(self):
        self._section(self.body, "Сурагчдын бүртгэл", "Сурагчдын мэдээллээ хадгалж, жагсаалтаа хэвлэж болно.")
        bar = tk.Frame(self.body, bg=BG); bar.pack(fill="x", pady=(0, 12))
        self._button(bar, "+ Сурагч нэмэх", self.add_student).pack(side="left")
        self._button(bar, "Засах", self.edit_student, False).pack(side="left", padx=8)
        self._button(bar, "Устгах", self.delete_student, False).pack(side="left")
        self._button(bar, "CSV импортлох", self.import_csv, False).pack(side="left", padx=8)
        self._button(bar, "Жагсаалт хэвлэх", lambda: self.print_report("students"), False).pack(side="right")
        self.student_search = tk.StringVar()
        search = ttk.Entry(bar, textvariable=self.student_search, width=22)
        search.pack(side="right", padx=8)
        search.insert(0, "")
        tk.Label(bar, text="Хайх", bg=BG, fg=MUTED, font=("Segoe UI", 9)).pack(side="right")
        self.student_search.trace_add("write", lambda *_: self._load_students())
        self.student_tree = self._tree(self.body, ("no", "code", "name", "parent", "phone"),
                                       ("№", "Код", "Овог нэр", "Эцэг эх/асран хамгаалагч", "Утас"), (55, 100, 250, 250, 150))
        self._load_students()

    def _load_students(self):
        if not hasattr(self, "student_tree"): return
        self.student_tree.delete(*self.student_tree.get_children())
        if not self.active_class: return
        term = getattr(self, "student_search", tk.StringVar()).get().strip()
        rows = self.store.all("SELECT * FROM students WHERE class_id=? AND (name LIKE ? OR code LIKE ? OR parent LIKE ? OR phone LIKE ?) ORDER BY name", (self.active_class, f"%{term}%", f"%{term}%", f"%{term}%", f"%{term}%"))
        for i, r in enumerate(rows, 1): self.student_tree.insert("", "end", iid=str(r["id"]), values=(i, r["code"], r["name"], r["parent"], r["phone"]))

    def _student_form(self, row=None):
        if not self.active_class:
            self.add_class(); return
        win = tk.Toplevel(self); win.title("Сурагчийн мэдээлэл"); win.configure(bg=BG); win.resizable(False, False)
        win.transient(self); win.grab_set()
        tk.Label(win, text="Сурагчийн мэдээлэл", bg=BG, fg=INK, font=("Segoe UI", 15, "bold")).grid(row=0, column=0, columnspan=2, padx=22, pady=(20, 12), sticky="w")
        fields = [("Сурагчийн код", "code"), ("Овог, нэр *", "name"), ("Эцэг эх/асран хамгаалагч", "parent"), ("Холбоо барих утас", "phone"), ("Тэмдэглэл", "note")]
        vars_ = {}
        for i, (label, key) in enumerate(fields, 1):
            tk.Label(win, text=label, bg=BG, fg=MUTED, font=("Segoe UI", 9)).grid(row=i, column=0, padx=22, pady=5, sticky="w")
            var = tk.StringVar(value=row[key] if row else ""); vars_[key] = var
            ttk.Entry(win, textvariable=var, width=36).grid(row=i, column=1, padx=(0, 22), pady=5)
        def save():
            if not vars_["name"].get().strip(): messagebox.showwarning("Мэдээлэл дутуу", "Сурагчийн нэрийг оруулна уу.", parent=win); return
            vals = [vars_[k].get().strip() for k in ("code", "name", "parent", "phone", "note")]
            if row: self.store.run("UPDATE students SET code=?,name=?,parent=?,phone=?,note=? WHERE id=?", (*vals, row["id"]))
            else: self.store.run("INSERT INTO students(class_id,code,name,parent,phone,note) VALUES(?,?,?,?,?,?)", (self.active_class, *vals))
            win.destroy(); self._load_students(); self.show_page("Нүүр") if self.page == "Нүүр" else None
        footer = tk.Frame(win, bg=BG); footer.grid(row=6, column=0, columnspan=2, padx=22, pady=18, sticky="e")
        self._button(footer, "Болих", win.destroy, False).pack(side="left", padx=5); self._button(footer, "Хадгалах", save).pack(side="left")

    def add_student(self): self._student_form()
    def _selected_student(self):
        tree = getattr(self, "student_tree", None)
        sel = tree.selection() if tree else ()
        if not sel: messagebox.showinfo("Сурагч сонгоно уу", "Жагсаалтаас сурагч сонгоно уу.", parent=self); return None
        return int(sel[0])
    def edit_student(self):
        sid = self._selected_student()
        if sid is not None: self._student_form(self.store.all("SELECT * FROM students WHERE id=?", (sid,))[0])
    def delete_student(self):
        sid = self._selected_student()
        if sid is not None and messagebox.askyesno("Устгах уу?", "Сурагч болон түүнтэй холбоотой ирц, дүнг устгах уу?", parent=self):
            self.store.run("DELETE FROM students WHERE id=?", (sid,)); self._load_students()

    def add_class(self):
        name = simpledialog.askstring("Анги үүсгэх", "Ангийн нэр (жишээ нь 7А):", parent=self)
        if not name or not name.strip(): return
        try:
            self.store.run("INSERT INTO classes(name,school_year) VALUES(?,?)", (name.strip(), str(date.today().year)))
            self.refresh_classes(); self.show_page(self.page)
        except sqlite3.IntegrityError: messagebox.showwarning("Анги бүртгэлтэй", "Ийм нэртэй анги байна.", parent=self)

    def _attendance_page(self):
        self._section(self.body, "Ирц бүртгэл", "Өдрөө сонгоод сурагч бүрийн ирцийг тэмдэглэнэ үү.")
        if not self.active_class:
            self._button(self.body, "+ Анги үүсгэх", self.add_class).pack(anchor="w"); return
        bar = tk.Frame(self.body, bg=BG); bar.pack(fill="x", pady=(0, 12))
        tk.Label(bar, text="Огноо", bg=BG, fg=MUTED, font=("Segoe UI", 9)).pack(side="left")
        self.att_date = tk.StringVar(value=date.today().isoformat())
        ttk.Entry(bar, textvariable=self.att_date, width=14).pack(side="left", padx=8)
        tk.Label(bar, text="YYYY-MM-DD", bg=BG, fg=MUTED, font=("Segoe UI", 8)).pack(side="left")
        self._button(bar, "Ирц ачаалах", self._load_attendance, False).pack(side="left", padx=10)
        self._button(bar, "Ирц хадгалах", self.save_attendance).pack(side="left")
        self._button(bar, "Тайлан хэвлэх", lambda: self.print_report("attendance"), False).pack(side="right")
        self.att_tree = self._tree(self.body, ("no", "name", "status", "note"), ("№", "Овог нэр", "Төлөв", "Тэмдэглэл"), (55, 280, 180, 300))
        self.att_tree.bind("<Double-1>", self._cycle_attendance)
        tk.Label(self.body, text="Төлөв дээр хоёр товшиход: Ирсэн → Өвчтэй → Чөлөөтэй → Тасалсан гэж солигдоно.", bg=BG, fg=MUTED, font=("Segoe UI", 8)).pack(anchor="w", pady=8)
        self._load_attendance()

    def _load_attendance(self):
        if not hasattr(self, "att_tree") or not self.active_class: return
        try: datetime.strptime(self.att_date.get().strip(), "%Y-%m-%d")
        except ValueError: messagebox.showwarning("Огноо буруу", "Огноог YYYY-MM-DD хэлбэрээр оруулна уу.", parent=self); return
        self.att_tree.delete(*self.att_tree.get_children())
        rows = self.store.all("SELECT s.id,s.name,COALESCE(a.status,'Ирсэн') status,COALESCE(a.note,'') note FROM students s LEFT JOIN attendance a ON a.student_id=s.id AND a.day=? WHERE s.class_id=? ORDER BY s.name", (self.att_date.get().strip(), self.active_class))
        for i, r in enumerate(rows, 1): self.att_tree.insert("", "end", iid=str(r["id"]), values=(i, r["name"], r["status"], r["note"]))

    def _cycle_attendance(self, event):
        item = self.att_tree.identify_row(event.y); col = self.att_tree.identify_column(event.x)
        if item and col == "#3":
            vals = list(self.att_tree.item(item, "values")); statuses = ["Ирсэн", "Өвчтэй", "Чөлөөтэй", "Тасалсан"]
            vals[2] = statuses[(statuses.index(vals[2]) + 1) % len(statuses)] if vals[2] in statuses else statuses[0]
            self.att_tree.item(item, values=vals)
        elif item and col == "#4":
            vals = list(self.att_tree.item(item, "values")); value = simpledialog.askstring("Тэмдэглэл", "Ирцийн тэмдэглэл:", initialvalue=vals[3], parent=self)
            if value is not None: vals[3] = value; self.att_tree.item(item, values=vals)

    def save_attendance(self):
        if not self.active_class: return
        try: datetime.strptime(self.att_date.get().strip(), "%Y-%m-%d")
        except ValueError: messagebox.showwarning("Огноо буруу", "YYYY-MM-DD хэлбэр ашиглана уу.", parent=self); return
        for iid in self.att_tree.get_children():
            _, _, status, note = self.att_tree.item(iid, "values")
            self.store.run("INSERT INTO attendance(student_id,day,status,note) VALUES(?,?,?,?) ON CONFLICT(student_id,day) DO UPDATE SET status=excluded.status,note=excluded.note", (int(iid), self.att_date.get().strip(), status, note))
        messagebox.showinfo("Хадгаллаа", "Ирц амжилттай хадгалагдлаа.", parent=self)
        self.show_page("Нүүр")

    def _grades_page(self):
        self._section(self.body, "Дүнгийн бүртгэл", "Шалгалт, даалгаврын оноог бүртгээд сурагчийн дундажийг харна.")
        if not self.active_class:
            self._button(self.body, "+ Анги үүсгэх", self.add_class).pack(anchor="w"); return
        bar = tk.Frame(self.body, bg=BG); bar.pack(fill="x", pady=(0, 12))
        self._button(bar, "+ Дүн нэмэх", self.add_grade).pack(side="left")
        self._button(bar, "Сонгосон дүн устгах", self.delete_grade, False).pack(side="left", padx=8)
        self._button(bar, "Дүнгийн тайлан хэвлэх", lambda: self.print_report("grades"), False).pack(side="right")
        self.grade_tree = self._tree(self.body, ("id", "day", "student", "subject", "title", "score", "percent"),
                                     ("№", "Огноо", "Сурагч", "Хичээл", "Ажил", "Оноо", "%"), (45, 105, 205, 135, 180, 100, 70))
        self.grade_tree.column("id", width=45, anchor="center")
        self._load_grades()

    def _load_grades(self):
        if not hasattr(self, "grade_tree") or not self.active_class: return
        self.grade_tree.delete(*self.grade_tree.get_children())
        rows = self.store.all("SELECT g.*,s.name FROM grades g JOIN students s ON s.id=g.student_id WHERE s.class_id=? ORDER BY g.day DESC,g.id DESC", (self.active_class,))
        for i, r in enumerate(rows, 1):
            pct = round(100 * r["score"] / r["max_score"]) if r["max_score"] else 0
            self.grade_tree.insert("", "end", iid=str(r["id"]), values=(i, r["day"], r["name"], r["subject"], r["title"], f'{r["score"]:g}/{r["max_score"]:g}', f"{pct}%"))

    def add_grade(self):
        students = self.store.all("SELECT id,name FROM students WHERE class_id=? ORDER BY name", (self.active_class,))
        if not students: messagebox.showinfo("Сурагч алга", "Эхлээд сурагч бүртгэнэ үү.", parent=self); return
        win = tk.Toplevel(self); win.title("Дүн нэмэх"); win.configure(bg=BG); win.resizable(False, False); win.transient(self); win.grab_set()
        tk.Label(win, text="Шинэ дүн бүртгэх", bg=BG, fg=INK, font=("Segoe UI", 15, "bold")).grid(row=0, column=0, columnspan=2, padx=22, pady=(20, 12), sticky="w")
        labels = ["Сурагч", "Хичээл", "Шалгалт/даалгавар", "Авсан оноо", "Нийт оноо", "Огноо", "Тэмдэглэл"]
        tk.Label(win, text="Сурагч", bg=BG, fg=MUTED).grid(row=1, column=0, padx=22, pady=5, sticky="w")
        names = [r["name"] for r in students]; stu = ttk.Combobox(win, values=names, state="readonly", width=32); stu.current(0); stu.grid(row=1, column=1, padx=22, pady=5)
        fields = {}
        defaults = {"Хичээл":"", "Шалгалт/даалгавар":"", "Авсан оноо":"", "Нийт оноо":"100", "Огноо":date.today().isoformat(), "Тэмдэглэл":""}
        for i, label in enumerate(labels[1:], 2):
            tk.Label(win, text=label, bg=BG, fg=MUTED).grid(row=i, column=0, padx=22, pady=5, sticky="w")
            var = tk.StringVar(value=defaults[label]); fields[label] = var
            ttk.Entry(win, textvariable=var, width=35).grid(row=i, column=1, padx=22, pady=5)
        def save():
            try:
                score = float(fields["Авсан оноо"].get()); maximum = float(fields["Нийт оноо"].get())
                datetime.strptime(fields["Огноо"].get(), "%Y-%m-%d")
                if maximum <= 0 or score < 0 or score > maximum: raise ValueError
                if not fields["Хичээл"].get().strip(): raise KeyError
            except KeyError: messagebox.showwarning("Мэдээлэл дутуу", "Хичээлийн нэрийг оруулна уу.", parent=win); return
            except ValueError: messagebox.showwarning("Мэдээлэл буруу", "Оноо болон огноог шалгана уу. Авсан оноо 0-ээс их, нийт онооноос хэтрэхгүй байна.", parent=win); return
            sid = students[names.index(stu.get())]["id"]
            self.store.run("INSERT INTO grades(student_id,subject,title,score,max_score,day,note) VALUES(?,?,?,?,?,?,?)", (sid, fields["Хичээл"].get().strip(), fields["Шалгалт/даалгавар"].get().strip(), score, maximum, fields["Огноо"].get().strip(), fields["Тэмдэглэл"].get().strip()))
            win.destroy(); self._load_grades()
        f = tk.Frame(win, bg=BG); f.grid(row=8, column=0, columnspan=2, padx=22, pady=18, sticky="e")
        self._button(f, "Болих", win.destroy, False).pack(side="left", padx=5); self._button(f, "Хадгалах", save).pack(side="left")

    def delete_grade(self):
        sel = self.grade_tree.selection()
        if not sel: messagebox.showinfo("Дүн сонгоно уу", "Устгах дүнгээ сонгоно уу.", parent=self); return
        if messagebox.askyesno("Дүн устгах", "Сонгосон дүнг устгах уу?", parent=self): self.store.run("DELETE FROM grades WHERE id=?", (int(sel[0]),)); self._load_grades()

    def _certificates_page(self):
        self._section(self.body, "Өргөмжлөл, батламж", "Монгол загвараа сонгож засварлаад лого, сурагчийн нэртэйгээр хэвлэнэ.")
        bar = tk.Frame(self.body, bg=BG); bar.pack(fill="x", pady=(0, 12))
        self.cert_template_var = tk.StringVar()
        self.cert_template_combo = ttk.Combobox(bar, textvariable=self.cert_template_var, state="readonly", width=28)
        self.cert_template_combo.pack(side="left")
        self.cert_template_combo.bind("<<ComboboxSelected>>", self._certificate_template_selected)
        self._button(bar, "Загварын бичвэр, өнгө засах", self.edit_certificate_template, False).pack(side="left", padx=8)
        self._button(bar, "Лого оруулах", self.choose_certificate_logo, False).pack(side="left")
        self._button(bar, "Лого арилгах", self.remove_certificate_logo, False).pack(side="left", padx=8)
        self.cert_logo_label = tk.Label(bar, text="Лого сонгоогүй", bg=BG, fg=MUTED, font=("Segoe UI", 8))
        self.cert_logo_label.pack(side="right")

        form = tk.Frame(self.body, bg=WHITE, padx=18, pady=16, highlightthickness=1, highlightbackground="#E7ECF4")
        form.pack(fill="x", pady=(0, 14))
        tk.Label(form, text="Сурагчид олгох мэдээлэл", bg=WHITE, fg=INK, font=("Segoe UI", 11, "bold")).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 10))
        self.cert_vars = {k: tk.StringVar(value=v) for k, v in {
            "recipient":"", "achievement":"", "organization":"", "teacher":"", "issue_date":date.today().isoformat()
        }.items()}
        tk.Label(form, text="Ангиас нэр сонгох", bg=WHITE, fg=MUTED, font=("Segoe UI", 9)).grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        self.cert_student_combo = ttk.Combobox(form, state="readonly", width=25)
        self.cert_student_combo.grid(row=1, column=1, sticky="ew", padx=(0, 16), pady=4)
        self.cert_student_combo.bind("<<ComboboxSelected>>", self._certificate_student_selected)
        tk.Label(form, text="Сурагчийн овог, нэр", bg=WHITE, fg=MUTED, font=("Segoe UI", 9)).grid(row=1, column=2, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(form, textvariable=self.cert_vars["recipient"], width=29).grid(row=1, column=3, sticky="ew", pady=4)
        for rownum, (left, key1, right, key2) in enumerate([
            ("Амжилт / шалтгаан", "achievement", "Сургууль / байгууллага", "organization"),
            ("Багшийн нэр", "teacher", "Огноо", "issue_date")], 2):
            tk.Label(form, text=left, bg=WHITE, fg=MUTED, font=("Segoe UI", 9)).grid(row=rownum, column=0, sticky="w", padx=(0, 8), pady=4)
            ttk.Entry(form, textvariable=self.cert_vars[key1], width=29).grid(row=rownum, column=1, sticky="ew", padx=(0, 16), pady=4)
            tk.Label(form, text=right, bg=WHITE, fg=MUTED, font=("Segoe UI", 9)).grid(row=rownum, column=2, sticky="w", padx=(0, 8), pady=4)
            ttk.Entry(form, textvariable=self.cert_vars[key2], width=29).grid(row=rownum, column=3, sticky="ew", pady=4)
        form.grid_columnconfigure(1, weight=1); form.grid_columnconfigure(3, weight=1)
        actions = tk.Frame(form, bg=WHITE); actions.grid(row=4, column=0, columnspan=4, sticky="w", pady=(12, 0))
        self._button(actions, "Хадгалаад хэвлэх", lambda: self.save_certificate(True)).pack(side="left")
        self._button(actions, "Зөвхөн хадгалах", lambda: self.save_certificate(False), False).pack(side="left", padx=8)

        self._section(self.body, "Хадгалсан өргөмжлөлүүд", "Дараа нь дахин нээж хэвлэх эсвэл хэрэггүйг устгана.")
        self.cert_tree = self._tree(self.body, ("id", "date", "name", "award", "template"),
                                    ("№", "Огноо", "Сурагч", "Амжилт", "Загвар"), (55, 100, 200, 260, 170))
        actions = tk.Frame(self.body, bg=BG); actions.pack(fill="x", pady=(8, 0))
        self._button(actions, "Сонгосныг дахин хэвлэх", self.reprint_certificate, False).pack(side="left")
        self._button(actions, "Устгах", self.delete_certificate, False).pack(side="left", padx=8)
        self._refresh_certificate_templates()
        self._refresh_saved_certificates()

    def _refresh_certificate_templates(self):
        self.cert_templates = self.store.all("SELECT * FROM certificate_templates ORDER BY id")
        self.cert_template_combo["values"] = [r["name"] for r in self.cert_templates]
        if self.cert_templates:
            chosen = next((r for r in self.cert_templates if r["name"] == self.cert_template_var.get()), self.cert_templates[0])
            self.cert_template_var.set(chosen["name"])
            self._load_certificate_template(chosen)
        names = [r["name"] for r in self.store.all("SELECT name FROM students WHERE class_id=? ORDER BY name", (self.active_class,))] if self.active_class else []
        self.cert_student_combo["values"] = names

    def _load_certificate_template(self, row):
        self.active_cert_template = row
        logo = row["logo"]
        self.cert_logo_label.config(text="Лого оруулсан ✓" if logo else "Лого сонгоогүй")

    def _certificate_template_selected(self, _event=None):
        row = next((r for r in self.cert_templates if r["name"] == self.cert_template_var.get()), None)
        if row: self._load_certificate_template(row)

    def _certificate_student_selected(self, _event=None):
        self.cert_vars["recipient"].set(self.cert_student_combo.get())

    def choose_certificate_logo(self):
        if not hasattr(self, "active_cert_template"): return
        path = filedialog.askopenfilename(parent=self, title="Сургуулийн лого сонгох", filetypes=[("Зураг", "*.png *.jpg *.jpeg *.gif *.bmp"), ("Бүх файл", "*.*")])
        if not path: return
        try:
            data = Path(path).read_bytes()
            if len(data) > 5 * 1024 * 1024: raise ValueError("Логоны хэмжээ 5 MB-аас бага байх ёстой.")
            mime = mimetypes.guess_type(path)[0] or "image/png"
            if mime not in ("image/png", "image/jpeg", "image/gif", "image/bmp"): raise ValueError("PNG эсвэл JPG зураг сонгоно уу.")
            tid = self.active_cert_template["id"]
            self.store.run("UPDATE certificate_templates SET logo=?,logo_mime=? WHERE id=?", (data, mime, tid))
            self._refresh_certificate_templates()
        except (OSError, ValueError) as exc: messagebox.showerror("Лого оруулж чадсангүй", str(exc), parent=self)

    def remove_certificate_logo(self):
        if not hasattr(self, "active_cert_template"): return
        tid = self.active_cert_template["id"]
        self.store.run("UPDATE certificate_templates SET logo=NULL,logo_mime='' WHERE id=?", (tid,))
        self._refresh_certificate_templates()

    def edit_certificate_template(self):
        row = getattr(self, "active_cert_template", None)
        if not row: return
        win = tk.Toplevel(self); win.title("Өргөмжлөлийн загвар засах"); win.configure(bg=BG); win.resizable(False, False); win.transient(self); win.grab_set()
        tk.Label(win, text="Загвараа өөрийнхөөрөө тохируулна", bg=BG, fg=INK, font=("Segoe UI", 14, "bold")).grid(row=0, column=0, columnspan=2, padx=22, pady=(20, 12), sticky="w")
        values = {"name":row["name"], "title":row["title"], "subtitle":row["subtitle"], "body":row["body"], "style":row["style"]}
        labels = [("Загварын нэр", "name"), ("Гарчиг", "title"), ("Дэд гарчиг", "subtitle")]
        vars_ = {}
        for i, (label, key) in enumerate(labels, 1):
            tk.Label(win, text=label, bg=BG, fg=MUTED).grid(row=i, column=0, padx=22, pady=5, sticky="w")
            vars_[key] = tk.StringVar(value=values[key]); ttk.Entry(win, textvariable=vars_[key], width=42).grid(row=i, column=1, padx=(0, 22), pady=5)
        tk.Label(win, text="Хэв загвар", bg=BG, fg=MUTED).grid(row=4, column=0, padx=22, pady=5, sticky="w")
        style_labels = CERT_STYLE_LABELS
        style_var = tk.StringVar(value=next((k for k,v in style_labels.items() if v==values["style"]), "Алтан Эрдэнэ"))
        ttk.Combobox(win, textvariable=style_var, state="readonly", values=list(style_labels), width=39).grid(row=4, column=1, padx=(0,22), pady=5)
        tk.Label(win, text="Тайлбарын өгүүлбэр", bg=BG, fg=MUTED).grid(row=5, column=0, padx=22, pady=5, sticky="nw")
        body_box = tk.Text(win, width=42, height=4, font=("Segoe UI", 10), wrap="word"); body_box.insert("1.0", values["body"]); body_box.grid(row=5, column=1, padx=(0,22), pady=5)
        tk.Label(win, text="Нэрийг {name}, амжилтыг {achievement} гэж бичвэл автоматаар орно.", bg=BG, fg=MUTED, font=("Segoe UI", 8)).grid(row=6, column=1, padx=(0,22), sticky="w")
        def save():
            name, title = vars_["name"].get().strip(), vars_["title"].get().strip()
            if not name or not title: messagebox.showwarning("Мэдээлэл дутуу", "Загварын нэр, гарчгийг оруулна уу.", parent=win); return
            try:
                self.store.run("UPDATE certificate_templates SET name=?,title=?,subtitle=?,body=?,style=? WHERE id=?", (name,title,vars_["subtitle"].get().strip(),body_box.get("1.0","end-1c").strip(),style_labels[style_var.get()],row["id"]))
            except sqlite3.IntegrityError: messagebox.showwarning("Нэр давхардсан", "Өөр загвартай ижил нэр сонгосон байна.", parent=win); return
            win.destroy(); self._refresh_certificate_templates()
        foot=tk.Frame(win,bg=BG); foot.grid(row=7,column=0,columnspan=2,padx=22,pady=18,sticky="e")
        self._button(foot,"Болих",win.destroy,False).pack(side="left",padx=5); self._button(foot,"Загварыг хадгалах",save).pack(side="left")

    def _certificate_payload(self):
        row = self.active_cert_template
        return {"template_id":row["id"],"template_name":row["name"],"style":row["style"],"title":row["title"],"subtitle":row["subtitle"],"body":row["body"],
                "recipient":self.cert_vars["recipient"].get().strip(),"achievement":self.cert_vars["achievement"].get().strip(),
                "teacher":self.cert_vars["teacher"].get().strip(),"organization":self.cert_vars["organization"].get().strip(),
                "issue_date":self.cert_vars["issue_date"].get().strip(),"logo":row["logo"],"logo_mime":row["logo_mime"]}

    def save_certificate(self, print_after=False):
        p = self._certificate_payload()
        if not p["recipient"] or not p["achievement"]:
            messagebox.showwarning("Мэдээлэл дутуу", "Сурагчийн нэр болон амжилт/шалтгааныг оруулна уу.", parent=self); return
        try: datetime.strptime(p["issue_date"], "%Y-%m-%d")
        except ValueError: messagebox.showwarning("Огноо буруу", "Огноог YYYY-MM-DD хэлбэрээр оруулна уу.", parent=self); return
        self.store.run("INSERT INTO certificates(template_id,template_name,style,title,subtitle,body,recipient,achievement,teacher,organization,issue_date,logo,logo_mime) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",
                       tuple(p[k] for k in ("template_id","template_name","style","title","subtitle","body","recipient","achievement","teacher","organization","issue_date","logo","logo_mime")))
        cert = dict(p); cert["id"] = self.store.db.execute("SELECT last_insert_rowid()").fetchone()[0]
        self._refresh_saved_certificates()
        if print_after: self._open_certificate(cert)
        else: messagebox.showinfo("Хадгаллаа", "Өргөмжлөлийн мэдээлэл хадгалагдлаа.", parent=self)

    def _open_certificate(self, record):
        path=Path(tempfile.gettempdir())/f"bagsh_urgumjlul_{record.get('id','preview')}_{datetime.now().strftime('%Y%m%d%H%M%S')}.html"
        path.write_text(certificate_html(dict(record)),encoding="utf-8"); webbrowser.open(path.as_uri())

    def _refresh_saved_certificates(self):
        if not hasattr(self,"cert_tree"): return
        self.cert_tree.delete(*self.cert_tree.get_children())
        rows=self.store.all("SELECT id,issue_date,recipient,achievement,template_name FROM certificates ORDER BY id DESC LIMIT 500")
        for r in rows: self.cert_tree.insert("","end",iid=str(r["id"]),values=(r["id"],r["issue_date"],r["recipient"],r["achievement"],r["template_name"]))

    def reprint_certificate(self):
        if not self.cert_tree.selection(): messagebox.showinfo("Сонгоно уу","Хэвлэх өргөмжлөлөө сонгоно уу.",parent=self); return
        row=self.store.all("SELECT * FROM certificates WHERE id=?",(int(self.cert_tree.selection()[0]),))[0]
        self._open_certificate(dict(row))

    def delete_certificate(self):
        if not self.cert_tree.selection(): messagebox.showinfo("Сонгоно уу","Устгах өргөмжлөлөө сонгоно уу.",parent=self); return
        if messagebox.askyesno("Устгах уу?","Сонгосон өргөмжлөлийн бүртгэлийг устгах уу?",parent=self):
            self.store.run("DELETE FROM certificates WHERE id=?",(int(self.cert_tree.selection()[0]),)); self._refresh_saved_certificates()

    def _exams_page(self):
        self._section(self.body, "Шалгалтын материалын сан", "Сонгох тестээ үүсгэж хадгалаад сурагчийн хувилбар болон хариултын түлхүүрийг хэвлэнэ.")
        bar=tk.Frame(self.body,bg=BG); bar.pack(fill="x",pady=(0,12))
        self._button(bar,"+ Шалгалт үүсгэх",lambda:self.open_exam_editor()).pack(side="left")
        self._button(bar,"Засах",self.edit_selected_exam,False).pack(side="left",padx=8)
        self._button(bar,"Устгах",self.delete_exam,False).pack(side="left")
        self._button(bar,"Сурагчийн хувилбар хэвлэх",lambda:self.print_selected_exam(False),False).pack(side="right")
        self._button(bar,"Хариултын түлхүүр",lambda:self.print_selected_exam(True),False).pack(side="right",padx=8)
        self.exam_tree=self._tree(self.body,("id","subject","title","questions","updated"),("№","Хичээл","Шалгалтын нэр","Асуултын тоо","Зассан огноо"),(55,150,300,130,155))
        self.exam_tree.column("id",width=55,anchor="center")
        self.exam_tree.bind("<Double-1>",lambda _e:self.edit_selected_exam())
        self._refresh_exams()

    def _refresh_exams(self):
        if not hasattr(self,"exam_tree"): return
        self.exam_tree.delete(*self.exam_tree.get_children())
        rows=self.store.all("SELECT e.*,COUNT(q.id) question_count FROM exams e LEFT JOIN exam_questions q ON q.exam_id=e.id GROUP BY e.id ORDER BY e.updated_at DESC,e.id DESC")
        for r in rows:self.exam_tree.insert("","end",iid=str(r["id"]),values=(r["id"],r["subject"],r["title"],r["question_count"],r["updated_at"][:10]))

    def open_exam_editor(self, exam_id=None):
        existing=None; questions=[]
        if exam_id is not None:
            existing=self.store.all("SELECT * FROM exams WHERE id=?",(exam_id,))[0]
            for q in self.store.all("SELECT * FROM exam_questions WHERE exam_id=? ORDER BY position",(exam_id,)):
                questions.append({"prompt":q["prompt"],"options":json.loads(q["options"]),"correct_index":q["correct_index"],"points":q["points"]})
        win=tk.Toplevel(self); win.title("Шалгалтын материал засах" if existing else "Шалгалтын материал үүсгэх"); win.configure(bg=BG); win.geometry("780x680"); win.minsize(680,560); win.transient(self); win.grab_set()
        outer=tk.Frame(win,bg=BG,padx=20,pady=18); outer.pack(fill="both",expand=True)
        tk.Label(outer,text="Шалгалтын материал",bg=BG,fg=INK,font=("Segoe UI",16,"bold")).pack(anchor="w")
        tk.Label(outer,text="Асуулт, дөрвөн сонголт, зөв хариулт болон оноогоо тохируулна.",bg=BG,fg=MUTED,font=("Segoe UI",9)).pack(anchor="w",pady=(3,12))
        top=tk.Frame(outer,bg=BG); top.pack(fill="x")
        tk.Label(top,text="Материалын нэр",bg=BG,fg=MUTED).grid(row=0,column=0,sticky="w",padx=(0,8),pady=4)
        title_var=tk.StringVar(value=existing["title"] if existing else "")
        ttk.Entry(top,textvariable=title_var,width=35).grid(row=0,column=1,sticky="ew",pady=4,padx=(0,14))
        tk.Label(top,text="Хичээл",bg=BG,fg=MUTED).grid(row=0,column=2,sticky="w",padx=(0,8),pady=4)
        subject_var=tk.StringVar(value=existing["subject"] if existing else "")
        ttk.Entry(top,textvariable=subject_var,width=22).grid(row=0,column=3,sticky="ew",pady=4)
        top.grid_columnconfigure(1,weight=2);top.grid_columnconfigure(3,weight=1)
        tk.Label(outer,text="Заавар / хугацаа",bg=BG,fg=MUTED).pack(anchor="w",pady=(9,3))
        inst=tk.Text(outer,height=3,font=("Segoe UI",10),wrap="word");inst.pack(fill="x")
        if existing: inst.insert("1.0",existing["instructions"])
        row=tk.Frame(outer,bg=BG);row.pack(fill="x",pady=(14,7))
        tk.Label(row,text="Асуултууд",bg=BG,fg=INK,font=("Segoe UI",11,"bold")).pack(side="left")
        self._button(row,"+ Асуулт нэмэх",lambda:edit_question()).pack(side="right")
        qframe=tk.Frame(outer,bg=WHITE,highlightthickness=1,highlightbackground="#E7ECF4");qframe.pack(fill="both",expand=True)
        qtree=ttk.Treeview(qframe,columns=("no","question","answer","points"),show="headings")
        for c,h,w in [("no","№",50),("question","Асуулт",430),("answer","Зөв хариулт",120),("points","Оноо",80)]:qtree.heading(c,text=h);qtree.column(c,width=w,anchor="w")
        qtree.pack(fill="both",expand=True)
        def refresh_questions():
            qtree.delete(*qtree.get_children())
            for i,q in enumerate(questions,1):
                ans=chr(65+q["correct_index"])
                qtree.insert("","end",iid=str(i-1),values=(i,q["prompt"],ans,q["points"]))
        refresh_questions()
        def edit_question(index=None):
            current=questions[index] if index is not None else None
            qw=tk.Toplevel(win);qw.title("Асуулт засах" if current else "Шинэ асуулт");qw.configure(bg=BG);qw.resizable(False,False);qw.transient(win);qw.grab_set()
            tk.Label(qw,text="Асуултын текст",bg=BG,fg=MUTED).grid(row=0,column=0,padx=18,pady=(18,4),sticky="w")
            prompt=tk.Text(qw,width=64,height=4,font=("Segoe UI",10),wrap="word");prompt.grid(row=1,column=0,columnspan=2,padx=18,pady=(0,8))
            if current:prompt.insert("1.0",current["prompt"])
            vars_=[]
            for i in range(4):
                tk.Label(qw,text=f"{chr(65+i)} сонголт",bg=BG,fg=MUTED).grid(row=i+2,column=0,padx=18,pady=4,sticky="w")
                v=tk.StringVar(value=current["options"][i] if current and i<len(current["options"]) else "");vars_.append(v)
                ttk.Entry(qw,textvariable=v,width=54).grid(row=i+2,column=1,padx=(0,18),pady=4)
            tk.Label(qw,text="Зөв хариулт",bg=BG,fg=MUTED).grid(row=6,column=0,padx=18,pady=6,sticky="w")
            ansvar=tk.StringVar(value=chr(65+current["correct_index"]) if current else "A")
            ttk.Combobox(qw,textvariable=ansvar,values=["A","B","C","D"],state="readonly",width=8).grid(row=6,column=1,padx=18,pady=6,sticky="w")
            tk.Label(qw,text="Оноо",bg=BG,fg=MUTED).grid(row=7,column=0,padx=18,pady=6,sticky="w")
            pts=tk.StringVar(value=str(current["points"]) if current else "1");ttk.Entry(qw,textvariable=pts,width=12).grid(row=7,column=1,padx=18,pady=6,sticky="w")
            def qsave():
                text=prompt.get("1.0","end-1c").strip();opts=[v.get().strip() for v in vars_]
                try:score=float(pts.get());correct=ord(ansvar.get())-65
                except (ValueError,TypeError):messagebox.showwarning("Утга буруу","Оноог тоогоор оруулна уу.",parent=qw);return
                if not text or sum(bool(x) for x in opts)<2 or score<=0 or not opts[correct]:
                    messagebox.showwarning("Мэдээлэл дутуу","Асуулт, дор хаяж 2 сонголт, зөв хариулт, эерэг оноо оруулна уу.",parent=qw);return
                q={"prompt":text,"options":opts,"correct_index":correct,"points":score}
                if index is None:questions.append(q)
                else:questions[index]=q
                qw.destroy();refresh_questions()
            foot=tk.Frame(qw,bg=BG);foot.grid(row=8,column=0,columnspan=2,padx=18,pady=15,sticky="e")
            self._button(foot,"Болих",qw.destroy,False).pack(side="left",padx=5);self._button(foot,"Асуулт хадгалах",qsave).pack(side="left")
        def on_double(_event=None):
            sel=qtree.selection()
            if sel:edit_question(int(sel[0]))
        qtree.bind("<Double-1>",on_double)
        qactions=tk.Frame(outer,bg=BG);qactions.pack(fill="x",pady=(7,0))
        self._button(qactions,"Сонгосон асуултыг засах",lambda:on_double(),False).pack(side="left")
        def remove_question():
            if not qtree.selection():return
            questions.pop(int(qtree.selection()[0]));refresh_questions()
        self._button(qactions,"Асуулт устгах",remove_question,False).pack(side="left",padx=8)
        def save_exam():
            title=title_var.get().strip();subject=subject_var.get().strip();instructions=inst.get("1.0","end-1c").strip()
            if not title or not subject:messagebox.showwarning("Мэдээлэл дутуу","Материалын нэр, хичээлийг бөглөнө үү.",parent=win);return
            now=datetime.now().isoformat(timespec="minutes")
            if existing:
                eid=existing["id"]
                self.store.run("UPDATE exams SET title=?,subject=?,instructions=?,updated_at=? WHERE id=?",(title,subject,instructions,now,eid))
                self.store.run("DELETE FROM exam_questions WHERE exam_id=?",(eid,))
            else:
                cur=self.store.run("INSERT INTO exams(class_id,title,subject,instructions,created_at,updated_at) VALUES(?,?,?,?,?,?)",(self.active_class,title,subject,instructions,now,now));eid=cur.lastrowid
            for i,q in enumerate(questions):self.store.run("INSERT INTO exam_questions(exam_id,position,prompt,options,correct_index,points) VALUES(?,?,?,?,?,?)",(eid,i,q["prompt"],json.dumps(q["options"],ensure_ascii=False),q["correct_index"],q["points"]))
            win.destroy();self._refresh_exams();messagebox.showinfo("Хадгаллаа",f"‘{title}’ шалгалтын материалыг хадгаллаа.",parent=self)
        bottom=tk.Frame(outer,bg=BG);bottom.pack(fill="x",pady=(12,0))
        self._button(bottom,"Болих",win.destroy,False).pack(side="right",padx=(8,0));self._button(bottom,"Материалыг хадгалах",save_exam).pack(side="right")

    def _selected_exam_id(self):
        sel=self.exam_tree.selection() if hasattr(self,"exam_tree") else ()
        if not sel:messagebox.showinfo("Шалгалт сонгоно уу","Жагсаалтаас материал сонгоно уу.",parent=self);return None
        return int(sel[0])

    def edit_selected_exam(self):
        eid=self._selected_exam_id()
        if eid is not None:self.open_exam_editor(eid)

    def delete_exam(self):
        eid=self._selected_exam_id()
        if eid is not None and messagebox.askyesno("Материал устгах уу?","Шалгалт болон асуултуудыг устгах уу?",parent=self):
            self.store.run("DELETE FROM exams WHERE id=?",(eid,));self._refresh_exams()

    def print_selected_exam(self,answer_key=False):
        eid=self._selected_exam_id()
        if eid is None:return
        exam=self.store.all("SELECT * FROM exams WHERE id=?",(eid,))[0]
        questions=[dict(q) for q in self.store.all("SELECT * FROM exam_questions WHERE exam_id=? ORDER BY position",(eid,))]
        for q in questions:q["options"]=json.loads(q["options"])
        content=exam_html(dict(exam),questions,answer_key)
        filename="hariult" if answer_key else "shalgalt"
        path=Path(tempfile.gettempdir())/f"bagsh_{filename}_{eid}_{datetime.now().strftime('%Y%m%d%H%M%S')}.html"
        path.write_text(content,encoding="utf-8");webbrowser.open(path.as_uri())

    def _reports_page(self):
        self._section(self.body, "Тайлан ба хэвлэх", "Тайланг хэвлэхийн өмнө урьдчилан харах цонх нээгдэнэ. Тэндээс принтер эсвэл PDF сонгоно.")
        for title, desc, key in [("Сурагчдын жагсаалт", "Ангийн сурагчид, код, холбоо барих мэдээлэл", "students"),
                                 ("Ирцийн тайлан", "Огнооны хүрээний ирцийн нэгтгэл", "attendance"),
                                 ("Дүнгийн тайлан", "Сурагч тус бүрийн хичээлийн дундаж", "grades")]:
            box = tk.Frame(self.body, bg=WHITE, padx=20, pady=17, highlightthickness=1, highlightbackground="#E7ECF4")
            box.pack(fill="x", pady=(0, 12))
            right = tk.Frame(box, bg=WHITE); right.pack(side="right")
            self._button(right, "Урьдчилан харах / хэвлэх", lambda k=key: self.print_report(k), False).pack()
            tk.Label(box, text=title, bg=WHITE, fg=INK, font=("Segoe UI", 11, "bold")).pack(anchor="w")
            tk.Label(box, text=desc, bg=WHITE, fg=MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(4, 0))
        tk.Label(self.body, text="Хэвлэх цонхноос принтерээ сонгох эсвэл “PDF болгон хадгалах” сонголтыг ашиглана.", bg=BG, fg=MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=4)
        self._button(self.body, "Сурагчдын мэдээллийг CSV болгон гаргах", self.export_csv, False).pack(anchor="w", pady=6)

    def _class_name(self):
        return next((r["name"] for r in self.class_rows if r["id"] == self.active_class), "Анги")

    def print_report(self, kind):
        if not self.active_class: messagebox.showinfo("Анги сонгоно уу", "Эхлээд анги үүсгэж/сонгоно уу.", parent=self); return
        cname = html.escape(self._class_name())
        if kind == "students":
            rows = self.store.all("SELECT * FROM students WHERE class_id=? ORDER BY name", (self.active_class,))
            title = "Сурагчдын жагсаалт"
            heads = ["№", "Код", "Овог нэр", "Эцэг эх/асран хамгаалагч", "Утас"]
            body = [[str(i), r["code"], r["name"], r["parent"], r["phone"]] for i, r in enumerate(rows, 1)]
        elif kind == "attendance":
            start = simpledialog.askstring("Ирцийн тайлан", "Эхлэх огноо (YYYY-MM-DD):", initialvalue=date.today().replace(day=1).isoformat(), parent=self)
            if not start: return
            end = simpledialog.askstring("Ирцийн тайлан", "Дуусах огноо (YYYY-MM-DD):", initialvalue=date.today().isoformat(), parent=self)
            if not end: return
            try:
                datetime.strptime(start, "%Y-%m-%d"); datetime.strptime(end, "%Y-%m-%d")
            except ValueError: messagebox.showwarning("Огноо буруу", "YYYY-MM-DD хэлбэр ашиглана уу.", parent=self); return
            rows = self.store.all("SELECT s.name, SUM(a.status='Ирсэн') present, SUM(a.status='Өвчтэй') sick, SUM(a.status='Чөлөөтэй') excused, SUM(a.status='Тасалсан') absent, COUNT(a.id) total FROM students s LEFT JOIN attendance a ON a.student_id=s.id AND a.day BETWEEN ? AND ? WHERE s.class_id=? GROUP BY s.id ORDER BY s.name", (start, end, self.active_class))
            title = f"Ирцийн тайлан · {start} — {end}"
            heads = ["№", "Овог нэр", "Ирсэн", "Өвчтэй", "Чөлөөтэй", "Тасалсан", "Бүртгэл"]
            body = [[str(i), r["name"], str(r["present"] or 0), str(r["sick"] or 0), str(r["excused"] or 0), str(r["absent"] or 0), str(r["total"] or 0)] for i, r in enumerate(rows, 1)]
        else:
            rows = self.store.all("SELECT s.name,g.subject,ROUND(AVG(100.0*g.score/g.max_score),1) avg,COUNT(g.id) count FROM students s LEFT JOIN grades g ON g.student_id=s.id WHERE s.class_id=? GROUP BY s.id,g.subject ORDER BY s.name,g.subject", (self.active_class,))
            title = "Дүнгийн тайлан"
            heads = ["№", "Овог нэр", "Хичээл", "Дундаж (%)", "Ажлын тоо"]
            body = [[str(i), r["name"], r["subject"] or "—", f'{r["avg"]}%' if r["avg"] is not None else "—", str(r["count"])] for i, r in enumerate(rows, 1)]
        table_head = "".join(f"<th>{html.escape(h)}</th>" for h in heads)
        table_rows = "".join("<tr>" + "".join(f"<td>{html.escape(str(v))}</td>" for v in row) + "</tr>" for row in body)
        content = f"""<!doctype html><html lang='mn'><meta charset='utf-8'><title>{html.escape(title)}</title>
        <style>@page{{size:A4 portrait;margin:16mm}}body{{font:12px 'Arial',sans-serif;color:#18243a}}h1{{font-size:22px;margin:0 0 6px}}.meta{{color:#63718a;margin-bottom:20px}}table{{border-collapse:collapse;width:100%}}th{{background:#eaf0fa;text-align:left}}th,td{{border:1px solid #d8e0ed;padding:8px 7px}}tr:nth-child(even){{background:#f8faff}}.sign{{display:flex;justify-content:space-between;margin-top:45px}}@media screen{{body{{max-width:900px;margin:32px auto;padding:30px;background:white;box-shadow:0 4px 28px #ccd3df}}.print{{position:fixed;right:24px;top:20px;background:#356ae6;color:white;border:0;padding:12px 20px;border-radius:8px;font-size:14px;cursor:pointer}}}}@media print{{.print{{display:none}}}}</style>
        <button class='print' onclick='window.print()'>Хэвлэх / PDF хадгалах</button><h1>{html.escape(title)}</h1><div class='meta'>{cname} анги · Гаргасан: {date.today().isoformat()} · Сурагч: {len(body)}</div>
        <table><thead><tr>{table_head}</tr></thead><tbody>{table_rows}</tbody></table><div class='sign'><span>Багш: ____________________</span><span>Гарын үсэг: ______________</span></div></html>"""
        path = Path(tempfile.gettempdir()) / f"bagshiin_tuslah_{kind}_{datetime.now().strftime('%Y%m%d%H%M%S')}.html"
        path.write_text(content, encoding="utf-8")
        webbrowser.open(path.as_uri())

    def export_csv(self):
        if not self.active_class: return
        path = filedialog.asksaveasfilename(parent=self, title="CSV хадгалах", defaultextension=".csv", initialfile=f"{self._class_name()}_suraltsagch.csv", filetypes=[("CSV файл", "*.csv")])
        if not path: return
        rows = self.store.all("SELECT code,name,parent,phone,note FROM students WHERE class_id=? ORDER BY name", (self.active_class,))
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f); w.writerow(["Код", "Овог нэр", "Эцэг эх/асран хамгаалагч", "Утас", "Тэмдэглэл"])
            for r in rows: w.writerow(list(r))
        messagebox.showinfo("Амжилттай", "CSV файл хадгалагдлаа.", parent=self)

    def import_csv(self):
        if not self.active_class:
            messagebox.showinfo("Анги сонгоно уу", "Эхлээд анги үүсгэнэ үү.", parent=self); return
        path = filedialog.askopenfilename(parent=self, title="Сурагчдын CSV сонгох", filetypes=[("CSV файл", "*.csv"), ("Бүх файл", "*.*")])
        if not path: return
        added = skipped = 0
        try:
            with open(path, "r", encoding="utf-8-sig", newline="") as f:
                reader = csv.DictReader(f)
                required = {"Овог нэр"}
                if not reader.fieldnames or not required.issubset(set(reader.fieldnames)):
                    raise ValueError("CSV файлд ‘Овог нэр’ багана байх ёстой. Өөрийн гаргасан CSV-г ашиглана уу.")
                for row in reader:
                    name = (row.get("Овог нэр") or "").strip()
                    if not name: skipped += 1; continue
                    exists = self.store.all("SELECT id FROM students WHERE class_id=? AND name=?", (self.active_class, name))
                    if exists: skipped += 1; continue
                    self.store.run("INSERT INTO students(class_id,code,name,parent,phone,note) VALUES(?,?,?,?,?,?)", (
                        self.active_class, (row.get("Код") or "").strip(), name,
                        (row.get("Эцэг эх/асран хамгаалагч") or "").strip(),
                        (row.get("Утас") or "").strip(), (row.get("Тэмдэглэл") or "").strip()))
                    added += 1
        except (OSError, UnicodeError, csv.Error, ValueError) as e:
            messagebox.showerror("CSV уншиж чадсангүй", str(e), parent=self); return
        self._load_students()
        messagebox.showinfo("Импорт дууслаа", f"Нэмсэн: {added}\nАлгассан/давхардсан: {skipped}", parent=self)

    def backup(self):
        target = filedialog.asksaveasfilename(parent=self, title="Нөөц хуулбар хадгалах", defaultextension=".db", initialfile=f"bagshiin_tuslah_backup_{date.today().isoformat()}.db", filetypes=[("Өгөгдлийн нөөц", "*.db")])
        if not target: return
        with sqlite3.connect(target) as dest: self.store.db.backup(dest)
        messagebox.showinfo("Нөөц хуулбар", "Өгөгдлийн нөөц хуулбар амжилттай хадгалагдлаа.", parent=self)

    def _close(self):
        self.store.close(); self.destroy()


if __name__ == "__main__":
    App().mainloop()
