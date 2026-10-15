"""Багшийн Туслах — offline Windows desktop app for Mongolian teachers."""
from __future__ import annotations

import csv
import base64
import html
import json
import mimetypes
import os
import re
import shutil
import sqlite3
import sys
import tempfile
import webbrowser
from io import BytesIO
from datetime import date, datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from powerpoint_export import convert_pptx_to_legacy_ppt, save_certificate_pptx, save_presentation_pptx
from design_editor import SlideDesignEditor
try:
    from PIL import Image, ImageTk
except ImportError:
    Image = ImageTk = None

APP_NAME = "Багшийн Туслах"
APP_VERSION = "1.7.0"
BG = "#F4F7FB"
NAV = "#101B31"
BLUE = "#3478F6"
INK = "#17243A"
MUTED = "#748198"
WHITE = "#FFFFFF"
GREEN = "#18A77A"
RED = "#D35454"
BORDER = "#E5EAF2"
TEAL = "#27AFA5"


def data_dir() -> Path:
    root = Path(os.environ.get("APPDATA", Path.home())) / "BagshiinTuslah"
    root.mkdir(parents=True, exist_ok=True)
    return root


def resource_path(relative_path: str) -> Path:
    bundle_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return bundle_dir / relative_path


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
          logo BLOB, logo_mime TEXT NOT NULL DEFAULT '', design_json TEXT NOT NULL DEFAULT '{}'
        );
        CREATE TABLE IF NOT EXISTS certificates (
          id INTEGER PRIMARY KEY, template_id INTEGER, template_name TEXT NOT NULL,
          style TEXT NOT NULL, title TEXT NOT NULL, subtitle TEXT NOT NULL, body TEXT NOT NULL,
          recipient TEXT NOT NULL, achievement TEXT NOT NULL DEFAULT '', teacher TEXT NOT NULL DEFAULT '',
          organization TEXT NOT NULL DEFAULT '', issue_date TEXT NOT NULL, logo BLOB, logo_mime TEXT NOT NULL DEFAULT '',
          design_json TEXT NOT NULL DEFAULT '{}',
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
        CREATE TABLE IF NOT EXISTS presentations (
          id INTEGER PRIMARY KEY, title TEXT NOT NULL, subject TEXT NOT NULL DEFAULT '',
          theme TEXT NOT NULL DEFAULT 'steppe', slides_json TEXT NOT NULL DEFAULT '[]',
          updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        CREATE INDEX IF NOT EXISTS idx_student_class ON students(class_id);
        CREATE INDEX IF NOT EXISTS idx_attendance_day ON attendance(day);
        """)
        for table in ("certificate_templates", "certificates"):
            columns={row[1] for row in self.db.execute(f"PRAGMA table_info({table})")}
            if "design_json" not in columns:
                self.db.execute(f"ALTER TABLE {table} ADD COLUMN design_json TEXT NOT NULL DEFAULT '{{}}'")
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
            ("Хөх Алт · Монгол хээ", "navy", "ӨРГӨМЖЛӨЛ", "Амжилтаа үнэлүүлэв", "{achievement} амжилт гаргасныг үнэлж энэхүү өргөмжлөлийг гардуулав."),
            ("Хүрэл Тамга", "bronze", "БҮТЭЭЛЧ САНААНЫ ӨРГӨМЖЛӨЛ", "Шинийг санаачилсанд", "{achievement} санаачилга, бүтээлч шийдлээрээ бусдыгаа уриалан оролцсонд гардуулав."),
            ("Сувдан Гэгээ", "pearl", "ОНЦГОЙ АМЖИЛТЫН ӨРГӨМЖЛӨЛ", "Өндөр амжилт гаргасанд", "{achievement} онцгой амжилтаар хамт олныхоо бахархал болсонд гардуулав."),
        ]
        if not self.db.execute("SELECT 1 FROM certificate_templates WHERE name='Хөх Алт · Монгол хээ'").fetchone():
            self.db.execute("UPDATE certificate_templates SET name='Хөх Алт · Монгол хээ' WHERE name='Хааны Хөх'")
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
    "navy":    {"colors":("#183f73","#efdfb3","#fffaf0"),"motif":"✦　❖　✦","border":"double","layout":"classic","pattern":"royal"},
    "bronze":  {"colors":("#8d623c","#ede0cd","#fcf8f1"),"motif":"◈　❖　◈","border":"groove","layout":"corners","pattern":"ever"},
    "pearl":   {"colors":("#837248","#eee9d6","#fffdf5"),"motif":"✺　◇　✺","border":"double","layout":"minimal","pattern":"lotus"},
}

CERT_STYLE_LABELS = {
    "Алтан Эрдэнэ":"gold", "Хөх Тэнгэр":"sky", "Улаан Тамга":"ruby", "Ногоон Хээ":"jade", "Нил ягаан Эрдэм":"violet",
    "Зэс Хүрээ":"copper", "Шөнийн Тэнгэр":"midnight", "Номин Далай":"teal", "Шүрэн Цэцэг":"coral", "Хөх Индранил":"indigo",
    "Маргад Эрдэнэ":"emerald", "Мөнгөн Хээ":"silver", "Нарны Туяа":"amber", "Сарнай":"rose", "Хөвсгөл Нуур":"ocean",
    "Нил Ягаан Хилэн":"plum", "Гааны Сэрүүн":"mint", "Хөх Алт · Монгол хээ":"navy", "Хүрэл Тамга":"bronze", "Сувдан Гэгээ":"pearl",
}

PRESENTATION_THEMES = {
    "steppe": {"name":"Талын хөх · Монгол хээ", "accent":"#1C5D9D", "bg":"#F5F8FC", "ink":"#102B48", "soft":"#DCEAF7", "mode":"steppe"},
    "gold": {"name":"Алтан эрдэм", "accent":"#A87819", "bg":"#FCF8EE", "ink":"#342711", "soft":"#F1E4C4", "mode":"classic"},
    "jade": {"name":"Ногоон мэдлэг", "accent":"#197A67", "bg":"#F1FAF6", "ink":"#163B36", "soft":"#D8EEE6", "mode":"split"},
    "coral": {"name":"Сарны ягаан", "accent":"#C4536C", "bg":"#FFF5F6", "ink":"#45232D", "soft":"#F6DDE3", "mode":"shape"},
    "midnight": {"name":"Шөнийн тэнгэр", "accent":"#75B5FF", "bg":"#101D35", "ink":"#F5F8FF", "soft":"#263A5C", "mode":"dark"},
    "sunrise": {"name":"Нар мандахуй", "accent":"#E27436", "bg":"#FFF7F0", "ink":"#3D291E", "soft":"#F9E0C9", "mode":"shape"},
    "violet": {"name":"Бүтээлч сэтгэлгээ", "accent":"#7054B8", "bg":"#F7F5FF", "ink":"#292141", "soft":"#E7E0FA", "mode":"split"},
    "minimal": {"name":"Цэвэр минимал", "accent":"#2563EB", "bg":"#FFFFFF", "ink":"#111827", "soft":"#EFF4FC", "mode":"minimal"},
    "ocean": {"name":"Далайн эрдэм", "accent":"#1685A3", "bg":"#F1FBFD", "ink":"#12333F", "soft":"#D8F0F5", "mode":"wave"},
    "forest": {"name":"Ойн нууц", "accent":"#66874A", "bg":"#F6F8F0", "ink":"#2D3923", "soft":"#E6ECD8", "mode":"classic"},
}


def presentation_html(deck: dict) -> str:
    """Create print-ready 16:9 slides that browsers can save as a multi-page PDF."""
    theme = PRESENTATION_THEMES.get(deck.get("theme", "steppe"), PRESENTATION_THEMES["steppe"])
    slides = deck.get("slides", [])
    if isinstance(slides, str):
        try: slides = json.loads(slides)
        except (json.JSONDecodeError, TypeError): slides = []
    title = html.escape(str(deck.get("title", "Presentation")))
    subject = html.escape(str(deck.get("subject", "")))
    accent, bg, ink, soft = theme["accent"], theme["bg"], theme["ink"], theme["soft"]
    mode = theme["mode"]
    cover = f"""<section class='slide cover mode-{mode}'><div class='cover-mark'>RAINY TEACHER <span>•</span> СУРГАЛТЫН ТАНИЛЦУУЛГА</div><div class='cover-copy'><div class='eyebrow'>{subject or 'БАГШИЙН PRESENTATION'}</div><h1>{title}</h1><div class='cover-line'></div><p>Мэдлэгийг ойлгомжтой, сонирхолтой хүргэе.</p></div><div class='cover-foot'>БАГШИЙН ТУСЛАХ <span>01</span></div></section>"""
    pages = []
    for index, slide in enumerate(slides, 2):
        stitle = html.escape(str(slide.get("title", "Шинэ слайд")))
        elements = slide.get("elements")
        if isinstance(elements, list) and elements:
            def safe_color(value, fallback):
                value=str(value or fallback)
                return value if re.fullmatch(r"#[0-9a-fA-F]{6}",value) else fallback
            nodes=[]
            for el in elements:
                try:
                    x=float(el.get("x",0))/960*100; y=float(el.get("y",0))/540*100
                    w=float(el.get("w",100))/960*100; h=float(el.get("h",40))/540*100
                    style=f"left:{x:.3f}%;top:{y:.3f}%;width:{w:.3f}%;height:{h:.3f}%;"
                    kind=el.get("type","text")
                    if kind in ("rect","ellipse"):
                        color=safe_color(el.get("fill"),accent)
                        radius="50%" if kind=="ellipse" else "3pt"
                        nodes.append(f"<div class='canvas-el' style='{style}background:{color};border-radius:{radius}'></div>")
                    elif kind=="image" and el.get("data"):
                        data=html.escape(str(el["data"]),quote=True)
                        mime=el.get("mime") if el.get("mime") in ("image/png","image/jpeg","image/webp") else "image/png"
                        nodes.append(f"<img class='canvas-el' style='{style}object-fit:contain' src='data:{mime};base64,{data}'>")
                    else:
                        color=safe_color(el.get("color"),ink)
                        size=max(8,min(96,int(float(el.get("font_size",24)))))
                        weight="700" if el.get("bold") else "400"
                        text=html.escape(str(el.get("text",""))).replace("\n","<br>")
                        nodes.append(f"<div class='canvas-el' style='{style}font-size:{size}pt;line-height:1.15;color:{color};font-weight:{weight};white-space:pre-wrap'>{text}</div>")
                except (ValueError,TypeError,KeyError):
                    continue
            slide_bg=safe_color(slide.get("background"),bg)
            pages.append(f"<section class='slide canvas-slide' style='background:{slide_bg}'>{''.join(nodes)}</section>")
            continue
        points = slide.get("points", [])
        if isinstance(points, str): points = [p.strip() for p in points.splitlines() if p.strip()]
        bullets = "".join(f"<li><i></i><span>{html.escape(str(point))}</span></li>" for point in points if str(point).strip())
        pages.append(f"<section class='slide content-slide mode-{mode}'><div class='topline'><b>RAINY TEACHER</b><span>{subject}</span></div><div class='slide-kicker'>ХИЧЭЭЛИЙН ТАНИЛЦУУЛГА</div><h2>{stitle}</h2><div class='rule'></div><ul>{bullets}</ul><div class='bottomline'><span>{title}</span><b>{index:02d}</b></div></section>")
    if not pages:
        pages.append("<section class='slide content-slide'><div class='topline'><b>RAINY TEACHER</b><span></span></div><div class='slide-kicker'>ЭХЛЭХЭД БЭЛЭН</div><h2>Эхний слайдаа нэмээрэй</h2><div class='rule'></div><ul><li><i></i><span>Слайдын гарчиг болон гол санаагаа оруулна уу.</span></li></ul><div class='bottomline'><span>Танилцуулга</span><b>02</b></div></section>")
    return f"""<!doctype html><html lang='mn'><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>{title} · Rainy Teacher</title><style>
    *{{box-sizing:border-box}}@page{{size:13.333in 7.5in;margin:0}}body{{margin:0;background:#e9edf4;color:{ink};font-family:'Segoe UI','Arial',sans-serif;-webkit-print-color-adjust:exact;print-color-adjust:exact}}.toolbar{{position:fixed;z-index:9;top:18px;right:22px}}.toolbar button{{border:0;border-radius:10px;background:{accent};color:#fff;padding:12px 18px;font-weight:700;box-shadow:0 6px 20px #0002;cursor:pointer}}.slide{{width:13.333in;height:7.5in;margin:24px auto;background:{bg};position:relative;overflow:hidden;padding:.68in .82in;box-shadow:0 10px 38px #182b4630;page-break-after:always;break-after:page}}.slide:last-of-type{{page-break-after:auto;break-after:auto}}.cover{{background:radial-gradient(ellipse at 82% 15%,{soft} 0 19%,transparent 19.2%),linear-gradient(125deg,{bg} 0 63%,{soft} 63% 100%);display:flex;flex-direction:column;justify-content:space-between}}.cover:before{{content:'';position:absolute;width:4.8in;height:4.8in;border:1px solid {accent}55;border-radius:50%;right:-1.2in;bottom:-2.5in;box-shadow:0 0 0 18px {bg}66,0 0 0 19px {accent}35,0 0 0 42px {bg}40,0 0 0 43px {accent}22}}.cover-mark,.topline{{position:relative;display:flex;justify-content:space-between;align-items:center;font-size:10pt;letter-spacing:2px;color:{accent};font-weight:700}}.cover-mark span{{font-size:20pt}}.cover-copy{{position:relative;max-width:8.7in;margin-top:-.12in}}.eyebrow,.slide-kicker{{font-size:11pt;letter-spacing:3px;color:{accent};font-weight:700;text-transform:uppercase}}h1{{font-size:42pt;line-height:1.14;margin:.24in 0;color:{ink};max-width:8.5in}}.cover-line,.rule{{width:1.2in;height:5px;background:{accent};border-radius:5px}}.cover-copy p{{font-size:17pt;color:{ink};opacity:.72;margin-top:.25in}}.cover-foot,.bottomline{{position:relative;display:flex;justify-content:space-between;border-top:1px solid {accent}44;padding-top:12px;color:{ink};opacity:.7;font-size:10pt;letter-spacing:1px}}.content-slide{{background:linear-gradient(120deg,{bg}, {bg} 75%,{soft})}}.content-slide:before{{content:'';position:absolute;left:0;top:0;bottom:0;width:14px;background:{accent}}}.content-slide:after{{content:'';position:absolute;right:.55in;top:.55in;width:.22in;height:.22in;border:2px solid {accent}77;transform:rotate(45deg)}}.mode-steppe.cover:after{{content:'';position:absolute;left:.5in;top:1.35in;width:.34in;height:4.2in;border-left:2px solid {accent}66;border-right:2px solid {accent}66;background:repeating-linear-gradient(0deg,transparent 0 .16in,{accent}55 .17in .19in,transparent .2in .34in)}}.mode-classic.cover{{outline:2px solid {accent}66;outline-offset:-.22in}}.mode-split .cover-copy{{border-left:7px solid {accent};padding-left:.34in}}.mode-shape.cover:before{{width:5.4in;height:5.4in;right:-1.8in;bottom:-2.8in}}.mode-dark.content-slide:before{{width:.12in}}.mode-dark .cover-copy p{{opacity:.84}}.mode-minimal.cover:before,.mode-minimal.cover:after{{display:none}}.mode-minimal.content-slide:before{{width:4px}}.mode-wave.cover:before{{height:2.3in;width:7in;border:0;border-radius:55% 45% 0 0;background:{soft};right:-.4in;bottom:-1.6in}}.mode-wave.cover:after{{content:'';position:absolute;right:-.35in;bottom:.65in;width:7in;height:1.6in;border-top:3px solid {accent}66;border-radius:55% 50% 0 0;transform:rotate(-7deg)}}.topline{{padding-bottom:15px;border-bottom:1px solid {accent}44}}.slide-kicker{{margin-top:.62in;font-size:9pt}}h2{{font-size:32pt;line-height:1.16;margin:.18in 0 .24in;color:{ink}}}.rule{{width:.85in;height:4px}}ul{{list-style:none;padding:0;margin:.38in 0 0;max-width:10.2in}}li{{display:flex;align-items:flex-start;gap:18px;margin:0 0 .2in;font-size:20pt;line-height:1.35;color:{ink}}}li i{{flex:0 0 11px;height:11px;border-radius:50%;background:{accent};margin-top:12px;box-shadow:0 0 0 6px {soft}}}.bottomline{{position:absolute;left:.82in;right:.82in;bottom:.48in}}.bottomline b{{color:{accent};opacity:1}}@media screen{{body{{padding:1px}}}}@media print{{body{{background:white}}.slide{{margin:0;box-shadow:none}}.toolbar{{display:none}}}}
    .canvas-slide{{padding:0!important}}.canvas-el{{position:absolute;overflow:hidden;box-sizing:border-box}}
    </style><div class='toolbar'><button onclick='window.print()'>PDF болгон хадгалах / хэвлэх</button></div>{cover}{''.join(pages)}</html>"""


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
    if pattern == "royal":
        blue, gold = accent, "#b58a35"
        return f"""<svg class='mongol-border' aria-hidden='true' viewBox='0 0 1000 700' preserveAspectRatio='none' xmlns='http://www.w3.org/2000/svg'>
        <rect x='16' y='16' width='968' height='668' fill='none' stroke='{blue}' stroke-width='11'/>
        <rect x='30' y='30' width='940' height='640' fill='none' stroke='{gold}' stroke-width='3'/>
        <rect x='42' y='42' width='916' height='616' fill='none' stroke='{blue}' stroke-width='2'/>
        <g fill='none' stroke='{blue}' stroke-width='5' stroke-linejoin='miter'>
          <path d='M44 146V44h102v28H72v74zm912 0V44H854v28h74v74zM44 554v102h102v-28H72v-74zm912 0v102H854v-28h74v-74z'/>
          <path d='M55 55h40v15H70v40H55zm890 0h-40v15h25v40h15zM55 645h40v-15H70v-40H55zm890 0h-40v-15h25v-40h15z'/>
        </g>
        <g fill='none' stroke='{gold}' stroke-width='3'>
          <path d='M50 50h132v4H54v132h-4zm900 0H818v4h128v132h4zM50 650h132v-4H54V514h-4zm900 0H818v-4h128V514h4z'/>
          <path d='M48 220h12v-12h12v12h12v12H72v12H60v-12H48zm904 0h-12v-12h-12v12h-12v12h12v12h12v-12h12z'/>
          <path d='M48 480h12v12h12v-12h12v-12H72v-12H60v12H48zm904 0h-12v12h-12v-12h-12v-12h12v-12h12v12h12z'/>
          <path d='M452 58l48 28 48-28 18 12-66 40-66-40zM452 642l48-28 48 28 18-12-66-40-66 40z'/>
        </g>
        <g fill='{blue}' stroke='{gold}' stroke-width='3'>
          <path d='M500 50l18 18-18 18-18-18z'/><path d='M500 614l18 18-18 18-18-18z'/>
          <path d='M54 330l17-17 17 17-17 17z'/><path d='M946 330l-17-17-17 17 17 17z'/>
        </g>
        <path d='M70 582c72-52 140-52 212 0s140 52 212 0 140-52 212 0 140 52 212 0' fill='none' stroke='{blue}' stroke-width='15'/>
        <path d='M70 590c72-52 140-52 212 0s140 52 212 0 140-52 212 0 140 52 212 0' fill='none' stroke='{gold}' stroke-width='4'/>
        <path d='M70 606c72-52 140-52 212 0s140 52 212 0 140-52 212 0 140 52 212 0' fill='none' stroke='{blue}' stroke-width='6'/>
        </svg>"""
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
    try: custom=json.loads(record.get("design_json") or "{}")
    except (TypeError,json.JSONDecodeError):custom={}
    if custom.get("elements"):
        values={"{name}":str(record.get("recipient") or "Сурагчийн нэр"),
                "{achievement}":str(record.get("achievement") or "Амжилтын тайлбар"),
                "{organization}":str(record.get("organization") or "Сургуулийн нэр"),
                "{teacher}":str(record.get("teacher") or "Багшийн нэр"),
                "{date}":str(record.get("issue_date") or "2026-10-15")}
        nodes=[]
        for el in custom["elements"]:
            try:
                x,y,w,h=[float(el.get(k,d)) for k,d in (("x",0),("y",0),("w",100),("h",40))]
                style=f"left:{x/960*100:.4f}%;top:{y/679*100:.4f}%;width:{w/960*100:.4f}%;height:{h/679*100:.4f}%;"
                kind=el.get("type","text")
                if kind in ("rect","ellipse"):
                    color=str(el.get("fill",accent))
                    color=color if re.fullmatch(r"#[0-9a-fA-F]{6}",color) else accent
                    nodes.append(f"<div style='position:absolute;{style}background:{color};border-radius:{'50%' if kind=='ellipse' else '2pt'}'></div>")
                elif kind=="image" and el.get("data"):
                    mime=el.get("mime") if el.get("mime") in ("image/png","image/jpeg","image/webp") else "image/png"
                    nodes.append(f"<img style='position:absolute;{style}object-fit:contain' src='data:{mime};base64,{html.escape(str(el['data']),quote=True)}'>")
                else:
                    color=str(el.get("color",accent))
                    color=color if re.fullmatch(r"#[0-9a-fA-F]{6}",color) else accent
                    size=max(8,min(96,int(float(el.get("font_size",22)))))
                    text=str(el.get("text",""))
                    for token,value in values.items():text=text.replace(token,value)
                    weight="700" if el.get("bold") else "400"
                    nodes.append(f"<div style='position:absolute;{style}font-size:{size}pt;line-height:1.15;color:{color};font-weight:{weight};white-space:pre-wrap;overflow:hidden'>{html.escape(text).replace(chr(10),'<br>')}</div>")
            except (ValueError,TypeError,KeyError):continue
        background=str(custom.get("background",paper))
        if not re.fullmatch(r"#[0-9a-fA-F]{6}",background):background=paper
        return f"""<!doctype html><html lang='mn'><meta charset='utf-8'><title>{html.escape(str(record.get('title','ӨРГӨМЖЛӨЛ')))} · {html.escape(values['{name}'])}</title><style>
        @page{{size:A4 landscape;margin:0}}*{{box-sizing:border-box}}body{{margin:0;background:#e9edf3;font-family:'Segoe UI',Arial,sans-serif;-webkit-print-color-adjust:exact;print-color-adjust:exact}}
        .sheet{{width:297mm;height:210mm;margin:12px auto;background:{background};position:relative;overflow:hidden;box-shadow:0 8px 36px #9aa5b5}}
        @media print{{body{{background:white}}.sheet{{margin:0;box-shadow:none}}}}
        </style><main class='sheet'>{''.join(nodes)}</main></html>"""
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
    @page{{size:A4 landscape;margin:0}}*{{box-sizing:border-box}}body{{margin:0;background:#e9edf3;color:#27344a;font-family:'Segoe UI','Noto Serif',Georgia,'Times New Roman',serif;-webkit-print-color-adjust:exact;print-color-adjust:exact}}
    .sheet{{width:297mm;height:210mm;margin:12px auto;padding:12mm;background:{paper};position:relative;overflow:hidden;box-shadow:0 8px 36px #9aa5b5}}
    .frame{{height:100%;border:3px {border} {accent};outline:1px solid {accent};outline-offset:-9px;padding:12mm 17mm;text-align:center;position:relative;background:radial-gradient(ellipse at 50% 0%,{pale} 0,transparent 42%),{paper}}}
    .mongol-border{{position:absolute;inset:0;width:100%;height:100%;pointer-events:none;z-index:0}}
    .frame>:not(.mongol-border){{position:relative;z-index:1}}
    .frame:before,.frame:after{{content:'{motif}';position:absolute;left:0;right:0;color:{accent};font-size:16px;letter-spacing:7px;opacity:.82}}
    .frame:before{{top:6mm}}.frame:after{{bottom:6mm}}
    .logo{{max-width:25mm;max-height:23mm;object-fit:contain;margin:0 auto 3mm;display:block}}
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
        self.title(f"Rainy Teacher · {APP_NAME}")
        self.geometry("1360x900")
        self.minsize(1120, 760)
        self.configure(bg=BG)
        try:
            self.iconbitmap(str(resource_path("assets/rainy_teacher.ico")))
        except tk.TclError:
            pass
        try:
            self.window_icon = tk.PhotoImage(data=resource_path("assets/rainy_teacher_mark.png.b64").read_text(encoding="ascii"))
            self.iconphoto(True, self.window_icon)
        except tk.TclError:
            self.window_icon = None
        self.store = Store()
        self.active_class: int | None = None
        self.page = "Нүүр"
        self._style()
        self._layout()
        self.show_page("Нүүр")
        self.protocol("WM_DELETE_WINDOW", self._close)

    def _style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Treeview", background=WHITE, fieldbackground=WHITE, foreground=INK,
                        rowheight=38, borderwidth=0, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", background="#F4F7FB", foreground="#5D6B82",
                        font=("Segoe UI", 9, "bold"), relief="flat", padding=(12, 12))
        style.map("Treeview", background=[("selected", "#DDEAFE")], foreground=[("selected", INK)])
        style.configure("TCombobox", padding=8, font=("Segoe UI", 10), fieldbackground=WHITE)
        style.map("TCombobox", fieldbackground=[("readonly", WHITE)], selectbackground=[("readonly", WHITE)])
        style.configure("TEntry", padding=9, font=("Segoe UI", 10))
        style.configure("TButton", padding=(13, 9), font=("Segoe UI", 10))
        style.configure("Vertical.TScrollbar", background="#E7ECF4", troughcolor=BG, bordercolor=BG, arrowsize=12)
    def _layout(self):
        self.sidebar = tk.Frame(self, bg=NAV, width=244)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        brand = tk.Frame(self.sidebar, bg=NAV)
        brand.pack(fill="x", padx=22, pady=(25, 6))
        try:
            self.brand_icon = self.window_icon.subsample(8, 8)
            tk.Label(brand, image=self.brand_icon, bg=NAV).pack(side="left", padx=(0, 11))
        except tk.TclError:
            self.brand_icon = None
        brand_text = tk.Frame(brand, bg=NAV)
        brand_text.pack(side="left", anchor="center")
        tk.Label(brand_text, text="RAINY", bg=NAV, fg=WHITE, font=("Segoe UI", 14, "bold")).pack(anchor="w")
        tk.Label(brand_text, text="TEACHER", bg=NAV, fg="#62DCCB", font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(self.sidebar, text="БАГШИЙН АЖЛЫН НЭГДСЭН ОРЧИН", bg=NAV, fg="#91A1BA",
                 font=("Segoe UI", 7, "bold")).pack(anchor="w", padx=24, pady=(0, 26))

        self.nav_buttons = {}
        sections = [
            ("ЕРӨНХИЙ", [("Нүүр", "home")]),
            ("СУРАЛЦАГЧ", [("Сурагчид", "students"), ("Ирц", "attendance"), ("Дүн", "grades")]),
            ("БИЧИГ БАРИМТ", [("Өргөмжлөл", "award"), ("Шалгалтын материал", "exam"), ("Танилцуулга", "slides"), ("Тайлан ба хэвлэх", "print")]),
        ]
        for section, items in sections:
            tk.Label(self.sidebar, text=section, bg=NAV, fg="#7485A0",
                     font=("Segoe UI", 8, "bold")).pack(anchor="w", padx=25, pady=(8, 8))
            for label, icon in items:
                row = tk.Frame(self.sidebar, bg=NAV, cursor="hand2")
                row.pack(fill="x", padx=12, pady=2)
                glyph = tk.Canvas(row, width=22, height=22, bg=NAV, highlightthickness=0)
                glyph.pack(side="left", padx=(15, 10), pady=9)
                button = tk.Button(
                    row, text=label, anchor="w", relief="flat", bd=0,
                    bg=NAV, fg="#C8D2E1", activebackground="#1E3150", activeforeground=WHITE,
                    font=("Segoe UI", 10, "bold" if label == self.page else "normal"),
                    padx=0, pady=12, command=lambda page=label: self.show_page(page), cursor="hand2"
                )
                button.pack(side="left", fill="x", expand=True)
                row.bind("<Button-1>", lambda _event, page=label: self.show_page(page))
                glyph.bind("<Button-1>", lambda _event, page=label: self.show_page(page))
                self.nav_buttons[label] = (row, glyph, button, icon)
                self._draw_nav_icon(glyph, icon, label == self.page)

        tk.Frame(self.sidebar, bg="#2B3B55", height=1).pack(fill="x", padx=22, pady=(18, 12))
        self._button(self.sidebar, "↓   Нөөц хуулбар", self.backup, False).pack(fill="x", padx=14, pady=(0, 14))
        status = tk.Frame(self.sidebar, bg="#172743", padx=14, pady=13)
        status.pack(side="bottom", fill="x", padx=14, pady=18)
        tk.Label(status, text="●  ОФЛАЙН ХАДГАЛАЛТ", bg="#172743", fg="#70DCC4",
                 font=("Segoe UI", 8, "bold")).pack(anchor="w")
        tk.Label(status, text="Мэдээлэл таны компьютерт хадгалагдана", bg="#172743", fg="#A4B2C7",
                 font=("Segoe UI", 8), wraplength=185, justify="left").pack(anchor="w", pady=(5, 0))

        self.main = tk.Frame(self, bg=BG)
        self.main.pack(side="left", fill="both", expand=True)
        top = tk.Frame(self.main, bg=WHITE, height=82)
        top.pack(fill="x")
        top.pack_propagate(False)
        self.page_title = tk.Label(top, text="Нүүр", bg=WHITE, fg=INK, font=("Segoe UI", 19, "bold"))
        self.page_title.pack(side="left", padx=30)
        right = tk.Frame(top, bg=WHITE)
        right.pack(side="right", padx=25)
        tk.Label(right, text="АНГИ", bg=WHITE, fg=MUTED, font=("Segoe UI", 8, "bold")).pack(side="left", padx=(0, 9))
        self.class_var = tk.StringVar()
        self.class_combo = ttk.Combobox(right, textvariable=self.class_var, state="readonly", width=17)
        self.class_combo.pack(side="left", ipady=2)
        self.class_combo.bind("<<ComboboxSelected>>", self._select_class)
        self._button(right, "+ Анги", self.add_class).pack(side="left", padx=(10, 0))
        self.body = tk.Frame(self.main, bg=BG)
        self.body.pack(fill="both", expand=True, padx=30, pady=26)
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
        for name, parts in self.nav_buttons.items():
            row, glyph, button, icon = parts
            active = name == page
            row.config(bg="#263956" if active else NAV)
            glyph.config(bg="#263956" if active else NAV)
            button.config(bg="#263956" if active else NAV, fg=WHITE if active else "#D5DEED",
                          font=("Segoe UI", 10, "bold" if active else "normal"))
            self._draw_nav_icon(glyph, icon, active)
        for child in self.body.winfo_children(): child.destroy()
        if page == "Нүүр": self._dashboard()
        elif page == "Сурагчид": self._students_page()
        elif page == "Ирц": self._attendance_page()
        elif page == "Дүн": self._grades_page()
        elif page == "Өргөмжлөл": self._certificates_page()
        elif page == "Шалгалтын материал": self._exams_page()
        elif page == "Танилцуулга": self._presentations_page()
        else: self._reports_page()

    def _draw_nav_icon(self, canvas, kind, active=False):
        """Draw crisp sidebar icons from lines and shapes; avoid missing emoji glyphs."""
        canvas.delete("all")
        color = "#7FE0D2" if active else "#9DADC3"
        shape_opts = {"outline": color, "width": 1.8}
        line_opts = {"fill": color, "width": 1.8, "capstyle": "round", "joinstyle": "round"}
        if kind == "home":
            canvas.create_line(3, 10, 11, 3, 19, 10, **line_opts)
            canvas.create_line(5, 9, 5, 19, 17, 19, 17, 9, **line_opts)
            canvas.create_rectangle(9, 13, 13, 19, outline=color, width=1.5)
        elif kind == "students":
            canvas.create_oval(7, 3, 14, 10, **shape_opts)
            canvas.create_arc(3, 10, 18, 22, start=0, extent=180, style="arc", **shape_opts)
            canvas.create_oval(15, 5, 20, 10, outline=color, width=1.5)
        elif kind == "attendance":
            canvas.create_oval(3, 3, 19, 19, **shape_opts)
            canvas.create_line(6, 11, 9, 14, 16, 7, **line_opts)
        elif kind == "grades":
            canvas.create_rectangle(4, 3, 18, 20, **shape_opts)
            for y in (8, 12, 16):
                canvas.create_line(7, y, 10, y, **line_opts)
                canvas.create_line(12, y, 15, y, **line_opts)
        elif kind == "award":
            canvas.create_polygon(11, 2, 13, 7, 19, 8, 15, 12, 16, 18, 11, 15, 6, 18, 7, 12, 3, 8, 9, 7, fill="", **shape_opts)
            canvas.create_line(8, 16, 7, 21, 11, 19, 15, 21, 14, 16, **line_opts)
        elif kind == "exam":
            canvas.create_rectangle(5, 3, 18, 20, **shape_opts)
            for y in (8, 12, 16):
                canvas.create_rectangle(7, y-1, 9, y+1, outline=color, width=1.2)
                canvas.create_line(11, y, 16, y, **line_opts)
        elif kind == "slides":
            canvas.create_rectangle(3, 4, 19, 16, **shape_opts)
            canvas.create_line(11, 16, 11, 20, **line_opts)
            canvas.create_line(7, 20, 15, 20, **line_opts)
            canvas.create_line(6, 13, 9, 10, 12, 12, 16, 7, **line_opts)
        elif kind == "print":
            canvas.create_rectangle(6, 3, 16, 8, **shape_opts)
            canvas.create_rectangle(4, 9, 18, 16, **shape_opts)
            canvas.create_rectangle(6, 14, 16, 20, **shape_opts)
            canvas.create_line(8, 17, 14, 17, **line_opts)

    def _card(self, parent, title, value, hint, accent=BLUE):
        card = tk.Frame(parent, bg=WHITE, highlightthickness=1, highlightbackground="#E7ECF4")
        tk.Frame(card, bg=accent, width=4).pack(side="left", fill="y")
        content = tk.Frame(card, bg=WHITE, padx=17, pady=14); content.pack(fill="both", expand=True)
        tk.Label(content, text=title, bg=WHITE, fg=MUTED, font=("Segoe UI", 9)).pack(anchor="w")
        tk.Label(content, text=value, bg=WHITE, fg=INK, font=("Segoe UI", 23, "bold")).pack(anchor="w", pady=3)
        tk.Label(content, text=hint, bg=WHITE, fg=MUTED, font=("Segoe UI", 8)).pack(anchor="w")
        return card

    def _section(self, parent, title, subtitle=None):
        row = tk.Frame(parent, bg=BG)
        row.pack(fill="x", pady=(0, 18))
        tk.Label(row, text=title, bg=BG, fg=INK, font=("Segoe UI", 16, "bold")).pack(anchor="w")
        if subtitle:
            tk.Label(row, text=subtitle, bg=BG, fg=MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(4, 0))
    def _button(self, parent, text, command, primary=True):
        return tk.Button(parent, text=text, command=command, relief="flat", bd=0,
                         bg=BLUE if primary else WHITE, fg=WHITE if primary else INK,
                         activebackground="#2463D4" if primary else "#EEF3FA",
                         activeforeground=WHITE if primary else INK,
                         padx=15, pady=10, font=("Segoe UI", 9, "bold" if primary else "normal"),
                         cursor="hand2", highlightthickness=1 if not primary else 0,
                         highlightbackground=BORDER)
    def _dashboard(self):
        if not self.active_class:
            self._section(self.body, "Тавтай морилно уу", "Rainy Teacher-д ангиа нэмээд өдөр тутмын ажлаа эхлүүлээрэй.")
            welcome = tk.Frame(self.body, bg=NAV, padx=30, pady=28)
            welcome.pack(fill="x", pady=(0, 22))
            tk.Label(welcome, text="Багшийн ажлаа нэг дороос", bg=NAV, fg=WHITE,
                     font=("Segoe UI", 21, "bold")).pack(anchor="w")
            tk.Label(welcome, text="Сурагч, ирц, дүн, өргөмжлөл, шалгалтын материалаа нэг офлайн орчинд.",
                     bg=NAV, fg="#C1CEE0", font=("Segoe UI", 10)).pack(anchor="w", pady=(7, 16))
            self._button(welcome, "+  Анги үүсгэх", self.add_class).pack(anchor="w")
            return
        cid = self.active_class
        total = self.store.all("SELECT COUNT(*) n FROM students WHERE class_id=?", (cid,))[0]["n"]
        today = date.today().isoformat()
        marked = self.store.all("SELECT COUNT(*) n FROM attendance a JOIN students s ON s.id=a.student_id WHERE s.class_id=? AND a.day=?", (cid, today))[0]["n"]
        present = self.store.all("SELECT COUNT(*) n FROM attendance a JOIN students s ON s.id=a.student_id WHERE s.class_id=? AND a.day=? AND a.status='Ирсэн'", (cid, today))[0]["n"]
        welcome = tk.Frame(self.body, bg=NAV, padx=26, pady=22)
        welcome.pack(fill="x", pady=(0, 20))
        welcome_left = tk.Frame(welcome, bg=NAV); welcome_left.pack(side="left", fill="x", expand=True)
        tk.Label(welcome_left, text=f"Сайн байна уу, {self.class_var.get()} анги 👋", bg=NAV, fg=WHITE,
                 font=("Segoe UI", 19, "bold")).pack(anchor="w")
        tk.Label(welcome_left, text="Өнөөдрийн ажлаа үргэлжлүүлээрэй. Таны мэдээлэл энэ компьютерт хадгалагдана.",
                 bg=NAV, fg="#BECDE0", font=("Segoe UI", 9)).pack(anchor="w", pady=(6, 0))
        tk.Label(welcome, text=date.today().strftime("%Y · %m · %d"), bg="#1B2C48", fg="#D8E3F1",
                 font=("Segoe UI", 10, "bold"), padx=14, pady=10).pack(side="right")
        cards = tk.Frame(self.body, bg=BG); cards.pack(fill="x", pady=(0, 22))
        for title, val, hint, color in [
            ("Нийт сурагч", str(total), "Сонгосон ангид", BLUE),
            ("Өнөөдрийн ирц", f"{marked}/{total}" if total else "0/0", f"Ирсэн: {present}", GREEN),
            ("Бүртгэлтэй анги", str(len(self.class_rows)), "Таны сургалтын сан", "#7A58C1")
        ]:
            self._card(cards, title, val, hint, color).pack(side="left", fill="both", expand=True, padx=(0, 12))
        self._section(self.body, "Шуурхай үйлдэл", "Түгээмэл ажлаа нэг товшилтоор нээнэ.")
        quick = tk.Frame(self.body, bg=BG); quick.pack(fill="x")
        shortcuts = [
            ("Сурагч нэмэх", "Шинэ сурагчийн мэдээлэл", "Сурагчид", "＋", BLUE),
            ("Ирц тэмдэглэх", "Өнөөдрийн ирцээ хөтлөх", "Ирц", "✓", GREEN),
            ("Дүн оруулах", "Шалгалт, даалгаврын оноо", "Дүн", "↗", "#7A58C1"),
            ("Шалгалт үүсгэх", "Материалаа хадгалж хэвлэх", "Шалгалтын материал", "✎", "#DB8A39")
        ]
        for title, desc, page, symbol, color in shortcuts:
            tile = tk.Frame(quick, bg=WHITE, padx=16, pady=17, highlightthickness=1, highlightbackground=BORDER)
            tile.pack(side="left", fill="both", expand=True, padx=(0, 12))
            tk.Label(tile, text=symbol, bg="#EEF4FF", fg=color, font=("Segoe UI", 16, "bold"),
                     width=3, pady=5).pack(anchor="w")
            tk.Label(tile, text=title, bg=WHITE, fg=INK, font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(12, 3))
            tk.Label(tile, text=desc, bg=WHITE, fg=MUTED, font=("Segoe UI", 8),
                     wraplength=160, justify="left").pack(anchor="w")
            self._button(tile, "Нээх  →", lambda page=page: self.show_page(page), False).pack(anchor="w", pady=(11, 0))
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
        self._section(self.body, "Өргөмжлөл, батламжийн студи",
                      "Монгол хээтэй 20 загвараас сонгож, preview дээр шалгаад хэвлэнэ.")
        self.cert_template_var = tk.StringVar()
        self.cert_vars = {key: tk.StringVar(value=value) for key, value in {
            "recipient": "", "achievement": "", "organization": "", "teacher": "",
            "issue_date": date.today().isoformat()
        }.items()}
        self.cert_vars["recipient"].trace_add("write", lambda *_: self._draw_certificate_preview())
        self.cert_vars["achievement"].trace_add("write", lambda *_: self._draw_certificate_preview())
        self.cert_vars["teacher"].trace_add("write", lambda *_: self._draw_certificate_preview())
        self.cert_vars["organization"].trace_add("write", lambda *_: self._draw_certificate_preview())
        self.cert_vars["issue_date"].trace_add("write", lambda *_: self._draw_certificate_preview())

        toolbar = tk.Frame(self.body, bg=BG); toolbar.pack(fill="x", pady=(0, 12))
        self.cert_template_name = tk.Label(toolbar, text="Загвар сонгоогүй", bg=BG, fg=INK,
                                           font=("Segoe UI", 10, "bold"))
        self.cert_template_name.pack(side="left")
        self._button(toolbar, "▦   20 загвар харах", self.open_certificate_gallery, False).pack(side="left", padx=(12, 7))
        self._button(toolbar, "Загвар засах", self.edit_certificate_template, False).pack(side="left")
        self._button(toolbar, "Дизайныг чирж засах", self.design_certificate_template, False).pack(side="left", padx=7)
        self._button(toolbar, "Лого нэмэх", self.choose_certificate_logo, False).pack(side="right", padx=(7, 0))
        self._button(toolbar, "Лого арилгах", self.remove_certificate_logo, False).pack(side="right")
        self.cert_logo_label = tk.Label(toolbar, text="Лого сонгоогүй", bg=BG, fg=MUTED, font=("Segoe UI", 8))
        self.cert_logo_label.pack(side="right", padx=10)

        stage = tk.Frame(self.body, bg=BG); stage.pack(fill="x", pady=(0, 15))
        preview_card = tk.Frame(stage, bg=WHITE, highlightthickness=1, highlightbackground=BORDER)
        preview_card.pack(side="left", fill="both", expand=True, padx=(0, 14))
        preview_head = tk.Frame(preview_card, bg=WHITE); preview_head.pack(fill="x", padx=15, pady=(12, 0))
        tk.Label(preview_head, text="ШУУД PREVIEW", bg=WHITE, fg=BLUE,
                 font=("Segoe UI", 8, "bold")).pack(side="left")
        tk.Label(preview_head, text="A4 · Хэвлэхэд бэлэн", bg=WHITE, fg=MUTED,
                 font=("Segoe UI", 8)).pack(side="right")
        self.cert_preview = tk.Canvas(preview_card, height=255, bg="#F6F8FC", highlightthickness=0)
        self.cert_preview.pack(fill="x", padx=14, pady=(10, 14))
        self.cert_preview.bind("<Configure>", lambda _event: self._draw_certificate_preview())

        form = tk.Frame(stage, bg=WHITE, width=345, padx=17, pady=15,
                        highlightthickness=1, highlightbackground=BORDER)
        form.pack(side="left", fill="y")
        form.pack_propagate(False)
        tk.Label(form, text="Өргөмжлөл бэлдэх", bg=WHITE, fg=INK,
                 font=("Segoe UI", 12, "bold")).pack(anchor="w")
        tk.Label(form, text="Сурагчийн мэдээллийг оруулна уу", bg=WHITE, fg=MUTED,
                 font=("Segoe UI", 8)).pack(anchor="w", pady=(3, 12))
        tk.Label(form, text="АНГИАС СОНГОХ", bg=WHITE, fg=MUTED,
                 font=("Segoe UI", 7, "bold")).pack(anchor="w")
        self.cert_student_combo = ttk.Combobox(form, state="readonly")
        self.cert_student_combo.pack(fill="x", pady=(4, 9))
        self.cert_student_combo.bind("<<ComboboxSelected>>", self._certificate_student_selected)
        fields = [
            ("Сурагчийн нэр", "recipient", "Жишээ: Бат-Эрдэнэ"),
            ("Амжилт / шалтгаан", "achievement", "Жишээ: Олимпиадад тэргүүлсэн"),
            ("Сургууль / байгууллага", "organization", ""),
            ("Багшийн нэр", "teacher", ""),
            ("Огноо", "issue_date", "YYYY-MM-DD")
        ]
        for label, key, placeholder in fields:
            tk.Label(form, text=label, bg=WHITE, fg=MUTED,
                     font=("Segoe UI", 8, "bold")).pack(anchor="w")
            entry = ttk.Entry(form, textvariable=self.cert_vars[key])
            entry.pack(fill="x", pady=(3, 8))
            if placeholder:
                entry.insert(0, "") if not self.cert_vars[key].get() else None
        actions = tk.Frame(form, bg=WHITE); actions.pack(fill="x", pady=(3, 0))
        self._button(actions, "Хадгалаад хэвлэх", lambda: self.save_certificate(True)).pack(fill="x", pady=(0, 7))
        self._button(actions, "Зөвхөн хадгалах", lambda: self.save_certificate(False), False).pack(fill="x", pady=(0, 7))
        self._button(actions, "PowerPoint файл хадгалах", self.export_current_certificate_pptx, False).pack(fill="x")

        saved = tk.Frame(self.body, bg=BG); saved.pack(fill="both", expand=True)
        self._section(saved, "Хадгалсан өргөмжлөлүүд", "Хадгалсан загвараа дахин хэвлэх эсвэл PowerPoint файл болгоно.")
        self.cert_tree = self._tree(saved, ("id", "date", "name", "award", "template"),
                                    ("№", "Огноо", "Сурагч", "Амжилт", "Загвар"), (48, 100, 180, 250, 160))
        actions = tk.Frame(saved, bg=BG); actions.pack(fill="x", pady=(7, 0))
        self._button(actions, "Сонгосныг дахин хэвлэх", self.reprint_certificate, False).pack(side="left")
        self._button(actions, "Сонгосныг PowerPoint болгох", self.export_selected_certificate_pptx, False).pack(side="left", padx=8)
        self._button(actions, "Устгах", self.delete_certificate, False).pack(side="left")
        self._refresh_certificate_templates()
        self._refresh_saved_certificates()

    def _draw_certificate_art(self, canvas, row, recipient="", achievement="", teacher="", organization="", issue_date=""):
        row=dict(row)
        canvas.delete("all")
        width=max(canvas.winfo_width(),220); height=max(canvas.winfo_height(),160)
        accent,pale,paper=CERT_DESIGNS.get(row["style"],CERT_DESIGNS["gold"])["colors"]
        try: custom=json.loads(row.get("design_json") or "{}")
        except (TypeError,json.JSONDecodeError): custom={}
        if custom.get("elements"):
            sx,sy=width/960,height/679
            canvas.configure(bg=custom.get("background",paper))
            self.preview_logo_photos=[]
            values={"{name}":recipient or "Сурагчийн нэр","{achievement}":achievement or "Амжилтын тайлбар",
                    "{organization}":organization or "Сургуулийн нэр","{teacher}":teacher or "Багшийн нэр",
                    "{date}":issue_date or "2026-10-15"}
            for el in custom["elements"]:
                try:
                    x,y,w,h=[float(el.get(k,d)) for k,d in (("x",0),("y",0),("w",100),("h",40))]
                    coords=(x*sx,y*sy,(x+w)*sx,(y+h)*sy);kind=el.get("type","text")
                    if kind in ("rect","ellipse"):
                        if kind=="ellipse":canvas.create_oval(*coords,fill=el.get("fill",accent),outline="")
                        else:canvas.create_rectangle(*coords,fill=el.get("fill",accent),outline="")
                    elif kind=="image" and el.get("data") and Image and ImageTk:
                        image=Image.open(BytesIO(base64.b64decode(el["data"]))).convert("RGBA")
                        image.thumbnail((max(1,int(w*sx)),max(1,int(h*sy))),Image.Resampling.LANCZOS)
                        photo=ImageTk.PhotoImage(image);self.preview_logo_photos.append(photo);canvas.create_image(coords[0]+(coords[2]-coords[0])/2,coords[1]+(coords[3]-coords[1])/2,image=photo)
                    else:
                        text=str(el.get("text",""))
                        for token,value in values.items():text=text.replace(token,value)
                        canvas.create_text(coords[0],coords[1],anchor="nw",text=text,fill=el.get("color",accent),
                            width=coords[2]-coords[0],font=("Segoe UI",max(6,int(float(el.get("font_size",22))*min(sx,sy))),"bold" if el.get("bold") else "normal"))
                except (ValueError,TypeError,tk.TclError,OSError):continue
            return
        canvas.configure(bg=paper)
        pad=18 if width>340 else 9
        x1,y1,x2,y2=pad,12,width-pad,height-12
        canvas.create_rectangle(x1,y1,x2,y2,fill=paper,outline=accent,width=3)
        inset=7 if width>340 else 4
        canvas.create_rectangle(x1+inset,y1+inset,x2-inset,y2-inset,outline=accent,width=1)
        motif=CERT_DESIGNS[row["style"]]["motif"]
        fs=max(7,min(13,int(width/42)))
        for yy in (y1+12,y2-13):
            canvas.create_text(width/2,yy,text=motif,fill=accent,font=("Segoe UI Symbol",fs,"bold"))
        for xx in (x1+14,x2-14):
            canvas.create_text(xx,(y1+y2)/2,text="❖",fill=accent,font=("Segoe UI Symbol",fs+2,"bold"))
        # Corner rosettes recall the continuous ulzii and alhan patterns on print.
        for xx in (x1+10,x2-10):
            for yy in (y1+10,y2-10):
                r=max(3,fs//2)
                canvas.create_polygon(xx,yy-r,xx+r,yy,xx,yy+r,xx-r,yy,
                                     fill=pale,outline=accent,width=1)
        # Repeating angular knot marks echo the ulzii border used by the print view.
        unit=max(2, int(width/105))
        step=unit*4
        line_width=max(1, unit//2)
        for xx in range(int(x1+24), int(x2-18), step):
            points=[xx,y1+5, xx+unit,y1+5, xx+unit,y1+5+unit,
                    xx+2*unit,y1+5+unit, xx+2*unit,y1+5,
                    xx+3*unit,y1+5, xx+3*unit,y1+5+3*unit,
                    xx,y1+5+3*unit, xx,y1+5]
            canvas.create_line(*points,fill=accent,width=line_width,joinstyle="miter")
            bottom_points=[xx,y2-5, xx+unit,y2-5, xx+unit,y2-5-unit,
                           xx+2*unit,y2-5-unit, xx+2*unit,y2-5,
                           xx+3*unit,y2-5, xx+3*unit,y2-5-3*unit,
                           xx,y2-5-3*unit, xx,y2-5]
            canvas.create_line(*bottom_points,fill=accent,width=line_width,joinstyle="miter")
        if row["style"] == "navy":
            gold = "#b58a35"
            canvas.create_rectangle(x1+5,y1+5,x2-5,y2-5,outline=gold,width=1)
            for left in (True,False):
                for top in (True,False):
                    cx=x1+13 if left else x2-13; cy=y1+13 if top else y2-13
                    sx=1 if left else -1; sy=1 if top else -1
                    canvas.create_line(cx,cy+sy*28,cx,cy,cx+sx*28,cy,fill=gold,width=2)
                    canvas.create_line(cx+sx*6,cy+sy*22,cx+sx*6,cy+sy*6,cx+sx*22,cy+sy*6,fill=gold,width=1)
            cx=width/2
            for yy in (y1+12,y2-12):
                canvas.create_polygon(cx,yy-6,cx+6,yy,cx,yy+6,cx-6,yy,
                                      fill=paper,outline=gold,width=2)
        title=row["title"] or "ӨРГӨМЖЛӨЛ"
        subtitle=row["subtitle"] or "Амжилтаа үнэлүүлэв"
        scale=max(.65,min(1.0,width/560))
        title_y=height*.30
        if row["logo"] and width>340:
            try:
                if Image and ImageTk:
                    image=Image.open(BytesIO(row["logo"])).convert("RGBA")
                    image.thumbnail((110,52),Image.Resampling.LANCZOS)
                    self.preview_logo_photo=ImageTk.PhotoImage(image)
                else:
                    self.preview_logo_photo=tk.PhotoImage(data=base64.b64encode(row["logo"]).decode("ascii"),format="png")
                    factor=max(1, int(max(self.preview_logo_photo.width()/110,self.preview_logo_photo.height()/52)))
                    self.preview_logo_photo=self.preview_logo_photo.subsample(factor,factor)
                canvas.create_image(width/2,height*.19,image=self.preview_logo_photo)
                title_y=height*.34
            except (tk.TclError, OSError, ValueError):
                self.preview_logo_photo=None
        canvas.create_text(width/2,title_y,text=title,fill=accent,
                           font=("Georgia",max(10,int(19*scale)),"bold"),width=width*.78)
        canvas.create_text(width/2,title_y+height*.10,text=subtitle,fill=MUTED,
                           font=("Segoe UI",max(7,int(10*scale))),width=width*.75)
        name=recipient or "Сурагчийн нэр"
        canvas.create_text(width/2,height*.56,text=name,fill=INK,
                           font=("Segoe UI",max(10,int(17*scale)),"bold"),width=width*.78)
        canvas.create_line(width*.28,height*.62,width*.72,height*.62,fill=accent,width=1)
        achievement=achievement or "Амжилт, баяр хүргэх шалтгаан"
        canvas.create_text(width/2,height*.71,text=achievement,fill=MUTED,
                           font=("Segoe UI",max(7,int(9*scale))),width=width*.70)
        meta="  •  ".join(v for v in (organization,teacher,issue_date) if v)
        if not meta: meta="Сургууль  •  Багш  •  Огноо"
        canvas.create_text(width/2,height*.86,text=meta,fill=accent,
                           font=("Segoe UI",max(6,int(8*scale))),width=width*.78)

    def _draw_certificate_preview(self):
        if not hasattr(self,"cert_preview") or not hasattr(self,"active_cert_template"): return
        self._draw_certificate_art(
            self.cert_preview,self.active_cert_template,
            self.cert_vars["recipient"].get(),self.cert_vars["achievement"].get(),
            self.cert_vars["teacher"].get(),self.cert_vars["organization"].get(),
            self.cert_vars["issue_date"].get()
        )

    def open_certificate_gallery(self):
        if not hasattr(self,"cert_templates"): self._refresh_certificate_templates()
        win=tk.Toplevel(self);win.title("Монгол загварын сан · 20 өргөмжлөл")
        win.configure(bg=BG);win.geometry("1120x760");win.minsize(900,620);win.transient(self)
        header=tk.Frame(win,bg=WHITE,padx=22,pady=16);header.pack(fill="x")
        tk.Label(header,text="Өргөмжлөлийн загварын сан",bg=WHITE,fg=INK,
                 font=("Segoe UI",17,"bold")).pack(anchor="w")
        tk.Label(header,text="Загвар дээр дарж сонгоно уу · Сонгосон загвар шууд preview дээр харагдана",
                 bg=WHITE,fg=MUTED,font=("Segoe UI",9)).pack(anchor="w",pady=(4,0))
        shell=tk.Frame(win,bg=BG);shell.pack(fill="both",expand=True,padx=18,pady=18)
        gallery=tk.Canvas(shell,bg=BG,highlightthickness=0)
        scrollbar=ttk.Scrollbar(shell,orient="vertical",command=gallery.yview,style="Vertical.TScrollbar")
        content=tk.Frame(gallery,bg=BG)
        window_id=gallery.create_window((0,0),window=content,anchor="nw")
        gallery.configure(yscrollcommand=scrollbar.set)
        gallery.pack(side="left",fill="both",expand=True);scrollbar.pack(side="right",fill="y")
        def resize_content(_event=None):
            gallery.itemconfigure(window_id,width=gallery.winfo_width())
            gallery.configure(scrollregion=gallery.bbox("all"))
        content.bind("<Configure>",resize_content);gallery.bind("<Configure>",resize_content)
        def wheel_scroll(event): gallery.yview_scroll(int(-1*(event.delta/120)),"units")
        gallery.bind_all("<MouseWheel>",wheel_scroll)
        def close_gallery():
            gallery.unbind_all("<MouseWheel>")
            win.destroy()
        columns=4
        for i,row in enumerate(self.cert_templates):
            r,c=divmod(i,columns)
            selected=row["id"]==self.active_cert_template["id"]
            card=tk.Frame(content,bg=WHITE,padx=11,pady=10,
                          highlightthickness=2 if selected else 1,
                          highlightbackground=BLUE if selected else BORDER,cursor="hand2")
            card.grid(row=r,column=c,sticky="nsew",padx=7,pady=7)
            thumb=tk.Canvas(card,width=220,height=128,bg="#F5F7FB",highlightthickness=0)
            thumb.pack(fill="x")
            self._draw_certificate_art(thumb,row,"Сурагчийн нэр","Амжилтын тайлбар")
            tk.Label(card,text=row["name"],bg=WHITE,fg=INK,font=("Segoe UI",10,"bold")).pack(anchor="w",pady=(8,1))
            tk.Label(card,text=row["title"],bg=WHITE,fg=MUTED,font=("Segoe UI",7)).pack(anchor="w")
            def choose(template=row):
                self.cert_template_var.set(template["name"]);self._load_certificate_template(template);close_gallery()
            card.bind("<Button-1>",lambda _e,fn=choose:fn())
            for child in card.winfo_children():child.bind("<Button-1>",lambda _e,fn=choose:fn())
            content.grid_columnconfigure(c,weight=1,uniform="certificate")
        footer=tk.Frame(win,bg=WHITE,padx=20,pady=12);footer.pack(fill="x")
        self._button(footer,"Болих",close_gallery,False).pack(side="right")
    def _refresh_certificate_templates(self):
        self.cert_templates = self.store.all("SELECT * FROM certificate_templates ORDER BY id")
        current_name=self.cert_template_var.get() if hasattr(self,"cert_template_var") else ""
        chosen=next((row for row in self.cert_templates if row["name"]==current_name),None)
        if chosen is None and self.cert_templates: chosen=self.cert_templates[0]
        names=[row["name"] for row in self.store.all("SELECT name FROM students WHERE class_id=? ORDER BY name",(self.active_class,))] if self.active_class else []
        if hasattr(self,"cert_student_combo"): self.cert_student_combo["values"]=names
        if chosen: self.cert_template_var.set(chosen["name"]);self._load_certificate_template(chosen)
    def _load_certificate_template(self, row):
        self.active_cert_template=row
        if hasattr(self,"cert_template_name"): self.cert_template_name.config(text=f"{row['name']}  ·  {row['title']}")
        if hasattr(self,"cert_logo_label"): self.cert_logo_label.config(text="Лого оруулсан ✓" if row["logo"] else "Лого сонгоогүй")
        self._draw_certificate_preview()
    def _certificate_template_selected(self, _event=None):
        row=next((template for template in self.cert_templates if template["name"]==self.cert_template_var.get()),None)
        if row:self._load_certificate_template(row)
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

    def design_certificate_template(self):
        row=getattr(self,"active_cert_template",None)
        if not row:return
        row=dict(row)
        accent,pale,paper=CERT_DESIGNS.get(row["style"],CERT_DESIGNS["gold"])["colors"]
        try: design=json.loads(row.get("design_json") or "{}")
        except (TypeError,json.JSONDecodeError):design={}
        if not design.get("elements"):
            elements=[
                {"type":"rect","x":24,"y":24,"w":912,"h":4,"fill":accent},
                {"type":"rect","x":24,"y":651,"w":912,"h":4,"fill":accent},
                {"type":"rect","x":24,"y":24,"w":4,"h":631,"fill":accent},
                {"type":"rect","x":932,"y":24,"w":4,"h":631,"fill":accent},
                {"type":"text","text":row["title"],"x":100,"y":82,"w":760,"h":62,"font_size":34,"color":accent,"bold":True},
                {"type":"text","text":row["subtitle"],"x":120,"y":158,"w":720,"h":42,"font_size":17,"color":"#627086","bold":False},
                {"type":"text","text":"{name}","x":100,"y":245,"w":760,"h":62,"font_size":32,"color":"#1C2940","bold":True},
                {"type":"text","text":row["body"],"x":105,"y":337,"w":750,"h":116,"font_size":19,"color":"#43516A","bold":False},
                {"type":"text","text":"{organization}","x":65,"y":552,"w":250,"h":45,"font_size":14,"color":"#43516A","bold":False},
                {"type":"text","text":"{teacher}","x":355,"y":552,"w":250,"h":45,"font_size":14,"color":"#43516A","bold":False},
                {"type":"text","text":"{date}","x":645,"y":552,"w":250,"h":45,"font_size":14,"color":"#43516A","bold":False},
            ]
            if row.get("logo"):
                try:
                    logo=Image.open(BytesIO(row["logo"])).convert("RGBA");stream=BytesIO();logo.save(stream,format="PNG")
                    logo_data=stream.getvalue()
                except Exception:logo_data=row["logo"]
                elements.append({"type":"image","data":base64.b64encode(logo_data).decode("ascii"),"mime":"image/png","x":440,"y":28,"w":80,"h":48})
            design={"background":paper,"elements":elements}
        def saved(slides):
            updated=slides[0]
            payload={"background":updated.get("background",paper),"elements":updated.get("elements",[])}
            self.store.run("UPDATE certificate_templates SET design_json=? WHERE id=?",(json.dumps(payload,ensure_ascii=False),row["id"]))
            self._refresh_certificate_templates()
            messagebox.showinfo("Загвар хадгалагдлаа","Өргөмжлөлийн дизайн хадгалагдлаа. Шинээр гаргах болон дахин хэвлэхэд энэ загвар хэрэглэгдэнэ.",parent=self)
        theme={"name":row["name"],"accent":accent,"bg":design.get("background",paper),"ink":"#1C2940","soft":pale}
        SlideDesignEditor(self,[{"title":row["title"],"elements":design["elements"],"background":design.get("background",paper)}],theme,saved,canvas_height=679,single_slide=True)

    def _certificate_payload(self):
        row = self.active_cert_template
        return {"template_id":row["id"],"template_name":row["name"],"style":row["style"],"title":row["title"],"subtitle":row["subtitle"],"body":row["body"],
                "recipient":self.cert_vars["recipient"].get().strip(),"achievement":self.cert_vars["achievement"].get().strip(),
                "teacher":self.cert_vars["teacher"].get().strip(),"organization":self.cert_vars["organization"].get().strip(),
                "issue_date":self.cert_vars["issue_date"].get().strip(),"logo":row["logo"],"logo_mime":row["logo_mime"],"design_json":row["design_json"]}

    def save_certificate(self, print_after=False):
        p = self._certificate_payload()
        if not p["recipient"] or not p["achievement"]:
            messagebox.showwarning("Мэдээлэл дутуу", "Сурагчийн нэр болон амжилт/шалтгааныг оруулна уу.", parent=self); return
        try: datetime.strptime(p["issue_date"], "%Y-%m-%d")
        except ValueError: messagebox.showwarning("Огноо буруу", "Огноог YYYY-MM-DD хэлбэрээр оруулна уу.", parent=self); return
        self.store.run("INSERT INTO certificates(template_id,template_name,style,title,subtitle,body,recipient,achievement,teacher,organization,issue_date,logo,logo_mime,design_json) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                       tuple(p[k] for k in ("template_id","template_name","style","title","subtitle","body","recipient","achievement","teacher","organization","issue_date","logo","logo_mime","design_json")))
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

    def _save_certificate_powerpoint(self, record):
        safe_name = "".join(ch if ch.isalnum() or ch in "-_ " else "_" for ch in str(record.get("recipient", "Өргөмжлөл"))).strip() or "Өргөмжлөл"
        target = filedialog.asksaveasfilename(parent=self, title="Өргөмжлөлийг PowerPoint файлаар хадгалах",
            initialfile=f"{safe_name}_urgumjlul.pptx", defaultextension=".pptx",
            filetypes=[("PowerPoint файл (*.pptx)", "*.pptx"), ("PowerPoint 97–2003 (*.ppt)", "*.ppt")])
        if not target: return
        path = Path(target)
        try:
            if path.suffix.lower() == ".ppt":
                temp_pptx = Path(tempfile.gettempdir()) / f"rainy_certificate_{datetime.now().strftime('%Y%m%d%H%M%S%f')}.pptx"
                try:
                    save_certificate_pptx(record, temp_pptx, CERT_DESIGNS)
                    convert_pptx_to_legacy_ppt(temp_pptx, path)
                finally:
                    temp_pptx.unlink(missing_ok=True)
            else:
                if path.suffix.lower() != ".pptx": path = path.with_suffix(".pptx")
                save_certificate_pptx(record, path, CERT_DESIGNS)
            messagebox.showinfo("Амжилттай", f"PowerPoint файл хадгалагдлаа:\n{path}", parent=self)
        except Exception as exc:
            messagebox.showerror("PowerPoint файл үүсгэсэнгүй", str(exc), parent=self)

    def export_current_certificate_pptx(self):
        record = self._certificate_payload()
        if not record.get("recipient") or not record.get("achievement"):
            messagebox.showwarning("Мэдээлэл дутуу", "Сурагчийн нэр болон амжилт/шалтгааныг оруулна уу.", parent=self); return
        try: datetime.strptime(record["issue_date"], "%Y-%m-%d")
        except ValueError: messagebox.showwarning("Огноо буруу", "Огноог YYYY-MM-DD хэлбэрээр оруулна уу.", parent=self); return
        self._save_certificate_powerpoint(record)

    def export_selected_certificate_pptx(self):
        if not self.cert_tree.selection():
            messagebox.showinfo("Сонгоно уу", "PowerPoint болгон хадгалах өргөмжлөлөө сонгоно уу.", parent=self); return
        row = self.store.all("SELECT * FROM certificates WHERE id=?", (int(self.cert_tree.selection()[0]),))
        if row: self._save_certificate_powerpoint(dict(row[0]))

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

    def _presentations_page(self):
        self._section(self.body, "Presentation студи", "10 загвараар слайдаа бэлдэж, засварлагдах PowerPoint эсвэл хэвлэх PDF болгон гаргана.")
        hero=tk.Frame(self.body,bg=NAV,padx=22,pady=17);hero.pack(fill="x",pady=(0,14))
        tk.Label(hero,text="Хичээлээ гоё танилцуул",bg=NAV,fg=WHITE,font=("Segoe UI",15,"bold")).pack(anchor="w")
        tk.Label(hero,text="Загвар сонгох · Слайдын агуулга бичих · Хэвлэх цонхоор PDF гаргах",bg=NAV,fg="#B8C8DC",font=("Segoe UI",9)).pack(anchor="w",pady=(4,0))
        controls=tk.Frame(self.body,bg=BG);controls.pack(fill="x",pady=(0,10))
        self._button(controls,"＋ Шинэ presentation",lambda:self.open_presentation_editor()).pack(side="left")
        self._button(controls,"Сонгосныг засах",self.edit_selected_presentation,False).pack(side="left",padx=8)
        self._button(controls,"Canva маягаар дизайн засах",self.design_selected_presentation,False).pack(side="left",padx=8)
        self._button(controls,"PowerPoint (.pptx) хадгалах",self.export_selected_presentation_pptx,False).pack(side="left",padx=8)
        self._button(controls,"PDF / хэвлэх",self.print_selected_presentation,False).pack(side="left")
        self._button(controls,"Устгах",self.delete_selected_presentation,False).pack(side="right")
        list_frame=tk.Frame(self.body,bg=WHITE,highlightthickness=1,highlightbackground=BORDER);list_frame.pack(fill="both",expand=True)
        self.presentation_tree=ttk.Treeview(list_frame,columns=("id","title","subject","theme","slides","updated"),show="headings",height=12)
        for key,label,width in [("id","№",50),("title","Presentation-ийн нэр",300),("subject","Хичээл / сэдэв",190),("theme","Загвар",230),("slides","Слайд",70),("updated","Шинэчилсэн",150)]:
            self.presentation_tree.heading(key,text=label);self.presentation_tree.column(key,width=width,anchor="w")
        scroll=ttk.Scrollbar(list_frame,orient="vertical",command=self.presentation_tree.yview);self.presentation_tree.configure(yscrollcommand=scroll.set)
        self.presentation_tree.pack(side="left",fill="both",expand=True,padx=(8,0),pady=8);scroll.pack(side="right",fill="y",pady=8,padx=(0,8))
        self.presentation_tree.bind("<Double-1>",lambda _e:self.edit_selected_presentation())
        self._refresh_presentations()

    def _refresh_presentations(self):
        if not hasattr(self,"presentation_tree"):return
        self.presentation_tree.delete(*self.presentation_tree.get_children())
        for row in self.store.all("SELECT * FROM presentations ORDER BY updated_at DESC,id DESC"):
            try: slides=json.loads(row["slides_json"] or "[]")
            except json.JSONDecodeError: slides=[]
            theme=PRESENTATION_THEMES.get(row["theme"],PRESENTATION_THEMES["steppe"])
            self.presentation_tree.insert("","end",iid=str(row["id"]),values=(row["id"],row["title"],row["subject"],theme["name"],len(slides),row["updated_at"]))

    def open_presentation_editor(self, presentation_id=None):
        existing=None
        if presentation_id:
            found=self.store.all("SELECT * FROM presentations WHERE id=?",(presentation_id,))
            if found:
                existing=dict(found[0])
        try: initial_slides=json.loads(existing["slides_json"]) if existing else []
        except (TypeError,json.JSONDecodeError): initial_slides=[]
        win=tk.Toplevel(self);win.title("Presentation засах" if existing else "Шинэ presentation");win.configure(bg=BG);win.geometry("1040x760");win.minsize(900,650);win.transient(self)
        header=tk.Frame(win,bg=WHITE,padx=24,pady=16);header.pack(fill="x")
        tk.Label(header,text="Хичээлийн presentation",bg=WHITE,fg=INK,font=("Segoe UI",17,"bold")).pack(anchor="w")
        tk.Label(header,text="Нэг слайдын гарчиг болон гол санааг бич. ‘---’ мөрөөр дараагийн слайд эхлүүлнэ.",bg=WHITE,fg=MUTED,font=("Segoe UI",9)).pack(anchor="w",pady=(4,0))
        body=tk.Frame(win,bg=BG);body.pack(fill="both",expand=True,padx=22,pady=18)
        left=tk.Frame(body,bg=BG);left.pack(side="left",fill="both",expand=True,padx=(0,16))
        right=tk.Frame(body,bg=WHITE,width=300,padx=16,pady=16,highlightthickness=1,highlightbackground=BORDER);right.pack(side="right",fill="y");right.pack_propagate(False)
        formtop=tk.Frame(left,bg=BG);formtop.pack(fill="x")
        title_var=tk.StringVar(value=existing["title"] if existing else "")
        subject_var=tk.StringVar(value=existing["subject"] if existing else "")
        tk.Label(formtop,text="Presentation-ийн нэр",bg=BG,fg=MUTED,font=("Segoe UI",8,"bold")).grid(row=0,column=0,sticky="w")
        ttk.Entry(formtop,textvariable=title_var).grid(row=1,column=0,sticky="ew",padx=(0,10),pady=(4,10))
        tk.Label(formtop,text="Хичээл / сэдэв",bg=BG,fg=MUTED,font=("Segoe UI",8,"bold")).grid(row=0,column=1,sticky="w")
        ttk.Entry(formtop,textvariable=subject_var).grid(row=1,column=1,sticky="ew",pady=(4,10))
        formtop.grid_columnconfigure(0,weight=1);formtop.grid_columnconfigure(1,weight=1)
        tk.Label(left,text="ЗАГВАРАА СОНГОХ · 10 БЭЛЭН ХЭВ МАЯГ",bg=BG,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w",pady=(4,6))
        selected_theme=tk.StringVar(value=existing["theme"] if existing else "steppe")
        themes=tk.Frame(left,bg=BG);themes.pack(fill="x",pady=(0,10))
        preview=tk.Canvas(right,width=268,height=164,highlightthickness=0);preview.pack(fill="x",pady=(0,14))
        preview_title=tk.Label(right,text="Слайдын нүүр",bg=WHITE,fg=INK,font=("Segoe UI",11,"bold"));preview_title.pack(anchor="w")
        preview_text=tk.Label(right,text="Сонгосон загвараа энд харна",bg=WHITE,fg=MUTED,font=("Segoe UI",8),wraplength=260,justify="left");preview_text.pack(anchor="w",pady=(5,14))
        count_label=tk.Label(right,text=f"{len(initial_slides)} слайд",bg="#F3F6FB",fg=INK,font=("Segoe UI",9,"bold"),padx=10,pady=7);count_label.pack(anchor="w")
        tk.Label(right,text="ЗАГВАРЫН БҮРЭЛДЭХҮҮН",bg=WHITE,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w",pady=(20,7))
        for line in ("• 16:9 widescreen", "• Монгол кирилл үсэг", "• Хэвлэхэд PDF", "• Офлайн ажиллана"):
            tk.Label(right,text=line,bg=WHITE,fg=INK,font=("Segoe UI",9)).pack(anchor="w",pady=3)
        def draw_cover():
            key=selected_theme.get();theme=PRESENTATION_THEMES[key]
            preview.delete("all");w=max(preview.winfo_width(),268);h=164
            preview.create_rectangle(0,0,w,h,fill=theme["bg"],outline=theme["accent"],width=2)
            preview.create_rectangle(0,0,9,h,fill=theme["accent"],outline=theme["accent"])
            preview.create_oval(w-80,-30,w+15,65,fill=theme["soft"],outline="")
            preview.create_text(20,25,text="RAINY TEACHER  ·  СУРГАЛТ",anchor="w",fill=theme["accent"],font=("Segoe UI",7,"bold"))
            preview.create_text(20,78,text=(title_var.get().strip() or "Таны хичээлийн сэдэв"),anchor="w",fill=theme["ink"],font=("Segoe UI",13,"bold"),width=w-55)
            preview.create_line(20,113,70,113,fill=theme["accent"],width=4)
            preview.create_text(20,137,text=(subject_var.get().strip() or "ХИЧЭЭЛИЙН PRESENTATION"),anchor="w",fill=theme["ink"],font=("Segoe UI",7))
            preview_text.config(text=theme["name"]+" · танилцуулгын нүүрний загвар")
        def choose_theme(key):
            selected_theme.set(key)
            for card_key,card in theme_cards.items():card.config(highlightbackground=BLUE if card_key==key else BORDER,highlightthickness=2 if card_key==key else 1)
            draw_cover()
        theme_cards={}
        for i,(key,theme) in enumerate(PRESENTATION_THEMES.items()):
            r,c=divmod(i,5)
            card=tk.Frame(themes,bg=WHITE,padx=8,pady=7,highlightthickness=2,highlightbackground=BORDER,cursor="hand2")
            card.grid(row=r,column=c,sticky="nsew",padx=(0,7),pady=(0,7))
            strip=tk.Canvas(card,width=112,height=38,bg=theme["bg"],highlightthickness=0);strip.pack(fill="x")
            strip.create_rectangle(0,0,7,38,fill=theme["accent"],outline=theme["accent"])
            strip.create_oval(82,-18,125,25,fill=theme["soft"],outline="")
            tk.Label(card,text=theme["name"],bg=WHITE,fg=INK,font=("Segoe UI",7,"bold"),wraplength=120,justify="left").pack(anchor="w",pady=(4,0))
            card.bind("<Button-1>",lambda _e,k=key:choose_theme(k))
            for child in card.winfo_children():child.bind("<Button-1>",lambda _e,k=key:choose_theme(k))
            themes.grid_columnconfigure(c,weight=1,uniform="theme")
            theme_cards[key]=card
        choose_theme(selected_theme.get() if selected_theme.get() in PRESENTATION_THEMES else "steppe")
        title_var.trace_add("write",lambda *_:draw_cover());subject_var.trace_add("write",lambda *_:draw_cover())
        tk.Label(left,text="СЛАЙДЫН АГУУЛГА",bg=BG,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w",pady=(6,5))
        content=tk.Text(left,height=17,font=("Segoe UI",10),wrap="word",relief="flat",bd=0,padx=12,pady=12,highlightthickness=1,highlightbackground=BORDER)
        content.pack(fill="both",expand=True)
        if initial_slides:
            parts=[]
            for slide in initial_slides:
                parts.append(str(slide.get("title", "Слайдын гарчиг"))+"\n"+"\n".join(str(p) for p in slide.get("points",[])))
            content.insert("1.0","\n\n---\n\n".join(parts))
        else: content.insert("1.0","Сэдвийн тойм\nӨнөөдрийн хичээлийн зорилго\nГол ойлголт\n\n---\n\nГол ойлголт\nЭхний чухал санаа\nЖишээ болон тайлбар\n\n---\n\nДүгнэлт\nСурсан зүйлээ нэгтгэх\nАсуулт, ярилцлага")
        def update_count(_event=None):
            blocks=[part for part in content.get("1.0","end-1c").split("---") if part.strip()]
            count_label.config(text=f"{len(blocks)} слайд + нүүр")
        content.bind("<KeyRelease>",update_count)
        actions=tk.Frame(win,bg=WHITE,padx=22,pady=13);actions.pack(fill="x")
        self._button(actions,"Болих",win.destroy,False).pack(side="right",padx=(8,0))
        def save_deck():
            title=title_var.get().strip();subject=subject_var.get().strip()
            blocks=[part.strip() for part in content.get("1.0","end-1c").split("---") if part.strip()]
            slides=[]
            for block in blocks:
                lines=[line.strip() for line in block.splitlines() if line.strip()]
                if lines:slides.append({"title":lines[0][:120],"points":lines[1:] or ["Гол санаагаа энд бичнэ үү."]})
            if not title or not subject:
                messagebox.showwarning("Мэдээлэл дутуу","Presentation-ийн нэр болон хичээл / сэдвийг бөглөнө үү.",parent=win);return
            if not slides:
                messagebox.showwarning("Слайд хоосон байна","Дор хаяж нэг слайдын гарчиг болон агуулгыг оруулна уу.",parent=win);return
            now=datetime.now().isoformat(timespec="minutes")
            serialized=json.dumps(slides,ensure_ascii=False)
            if existing:
                self.store.run("UPDATE presentations SET title=?,subject=?,theme=?,slides_json=?,updated_at=? WHERE id=?",(title,subject,selected_theme.get(),serialized,now,existing["id"]))
            else:
                self.store.run("INSERT INTO presentations(title,subject,theme,slides_json,updated_at) VALUES(?,?,?,?,?)",(title,subject,selected_theme.get(),serialized,now))
            win.destroy();self._refresh_presentations()
        self._button(actions,"Хадгалах",save_deck).pack(side="right")

    def _selected_presentation(self):
        selection=self.presentation_tree.selection() if hasattr(self,"presentation_tree") else ()
        if not selection:
            messagebox.showinfo("Presentation сонгоно уу","Жагсаалтаас presentation сонгоно уу.",parent=self);return None
        rows=self.store.all("SELECT * FROM presentations WHERE id=?",(int(selection[0]),))
        return dict(rows[0]) if rows else None

    def edit_selected_presentation(self):
        deck=self._selected_presentation()
        if deck:self.open_presentation_editor(deck["id"])

    def design_selected_presentation(self):
        deck=self._selected_presentation()
        if not deck:return
        try: slides=json.loads(deck.get("slides_json") or "[]")
        except (TypeError,json.JSONDecodeError): slides=[]
        if not slides:
            messagebox.showinfo("Слайд алга байна","Эхлээд presentation-дээ слайдын агуулга нэмнэ үү.",parent=self);return
        theme=PRESENTATION_THEMES.get(deck.get("theme"),PRESENTATION_THEMES["steppe"])
        def save_design(updated_slides):
            serialized=json.dumps(updated_slides,ensure_ascii=False)
            now=datetime.now().isoformat(timespec="minutes")
            self.store.run("UPDATE presentations SET slides_json=?,updated_at=? WHERE id=?",(serialized,now,deck["id"]))
            self._refresh_presentations()
            messagebox.showinfo("Дизайн хадгалагдлаа","Слайдын өөрчлөлтүүд хадгалагдлаа. PPTX болгон гаргахад элемент бүр PowerPoint дээр засварлагдана.",parent=self)
        SlideDesignEditor(self,slides,theme,save_design)

    def print_selected_presentation(self):
        deck=self._selected_presentation()
        if not deck:return
        deck["slides"]=json.loads(deck["slides_json"] or "[]")
        path=Path(tempfile.gettempdir())/f"rainy_presentation_{deck['id']}_{datetime.now().strftime('%Y%m%d%H%M%S')}.html"
        path.write_text(presentation_html(deck),encoding="utf-8");webbrowser.open(path.as_uri())

    def export_selected_presentation_pptx(self):
        deck = self._selected_presentation()
        if not deck: return
        try: deck["slides"] = json.loads(deck["slides_json"] or "[]")
        except (TypeError, json.JSONDecodeError): deck["slides"] = []
        safe_name = "".join(ch if ch.isalnum() or ch in "-_ " else "_" for ch in deck["title"]).strip() or "Presentation"
        target = filedialog.asksaveasfilename(parent=self, title="Presentation-ээ PowerPoint файлаар хадгалах",
            initialfile=f"{safe_name}.pptx", defaultextension=".pptx",
            filetypes=[("PowerPoint файл (*.pptx)", "*.pptx"), ("PowerPoint 97–2003 (*.ppt)", "*.ppt")])
        if not target: return
        path = Path(target)
        temp_pptx = None
        try:
            theme = PRESENTATION_THEMES.get(deck.get("theme"), PRESENTATION_THEMES["steppe"])
            if path.suffix.lower() == ".ppt":
                temp_pptx = Path(tempfile.gettempdir()) / f"rainy_presentation_{datetime.now().strftime('%Y%m%d%H%M%S%f')}.pptx"
                save_presentation_pptx(deck, temp_pptx, theme)
                convert_pptx_to_legacy_ppt(temp_pptx, path)
            else:
                if path.suffix.lower() != ".pptx": path = path.with_suffix(".pptx")
                save_presentation_pptx(deck, path, theme)
            messagebox.showinfo("Амжилттай", f"PowerPoint файл хадгалагдлаа:\n{path}", parent=self)
        except Exception as exc:
            messagebox.showerror("PowerPoint файл үүсгэсэнгүй", str(exc), parent=self)
        finally:
            if temp_pptx: temp_pptx.unlink(missing_ok=True)

    def delete_selected_presentation(self):
        deck=self._selected_presentation()
        if deck and messagebox.askyesno("Устгах уу?",f"‘{deck['title']}’ presentation-ийг устгах уу?",parent=self):
            self.store.run("DELETE FROM presentations WHERE id=?",(deck["id"],));self._refresh_presentations()

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
