"""Simple offline drag-and-drop slide designer for Rainy Teacher."""
from __future__ import annotations

import base64
import tkinter as tk
from tkinter import ttk, filedialog, colorchooser, messagebox
from io import BytesIO

from PIL import Image, ImageTk


class SlideDesignEditor:
    """Edits a deck's slides as independently movable PowerPoint elements."""

    W = 960

    def __init__(self, parent, slides, theme, on_save, *, canvas_height=540, single_slide=False):
        self.parent = parent
        self.slides = slides
        self.theme = theme
        self.on_save = on_save
        self.H = canvas_height
        self.single_slide = single_slide
        self.index = 0
        self.elements = []
        self.selected = None
        self.scale = 1.0
        self.drag_origin = None
        self.photo_refs = []

        self.win = tk.Toplevel(parent)
        self.win.title("Слайдын дизайн засварлагч · Rainy Teacher")
        self.win.geometry("1240x800")
        self.win.minsize(1050, 680)
        self.win.configure(bg="#F2F5FA")
        self.win.transient(parent)

        bar = tk.Frame(self.win, bg="#111D34", padx=16, pady=11)
        bar.pack(fill="x")
        tk.Label(bar, text="СЛАЙДЫН ДИЗАЙН", bg="#111D34", fg="white",
                 font=("Segoe UI", 12, "bold")).pack(side="left")
        self.page_label = tk.Label(bar, text="", bg="#111D34", fg="#CCD8E9",
                                   font=("Segoe UI", 9))
        self.page_label.pack(side="left", padx=18)
        tools = [("＋ Текст", self.add_text), ("＋ Дүрс", self.add_shape), ("＋ Зураг", self.add_image)]
        if not self.single_slide: tools.append(("＋ Слайд", self.add_slide))
        tools.append(("Устгах", self.delete_selected))
        for label, command in tools:
            tk.Button(bar, text=label, command=command, bg="#263651", fg="white",
                      activebackground="#365070", activeforeground="white", relief="flat",
                      padx=11, pady=7, font=("Segoe UI", 9, "bold"), cursor="hand2").pack(side="left", padx=4)
        tk.Button(bar, text="Хадгалах", command=self.save, bg="#3478F6", fg="white",
                  activebackground="#2465DD", relief="flat", padx=16, pady=7,
                  font=("Segoe UI", 9, "bold"), cursor="hand2").pack(side="right")

        work = tk.Frame(self.win, bg="#F2F5FA", padx=14, pady=14)
        work.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(work, bg="white", highlightthickness=1,
                                highlightbackground="#D8E0EC", cursor="arrow")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.bind("<Configure>", lambda _e: self.draw())
        self.canvas.bind("<Button-1>", self.pointer_down)
        self.canvas.bind("<B1-Motion>", self.pointer_move)
        self.canvas.bind("<ButtonRelease-1>", self.pointer_up)

        panel = tk.Frame(work, bg="white", width=250, padx=16, pady=16,
                         highlightthickness=1, highlightbackground="#D8E0EC")
        panel.pack(side="right", fill="y", padx=(14, 0))
        panel.pack_propagate(False)
        tk.Label(panel, text="ЭЛЕМЕНТИЙН ТОХИРГОО", bg="white", fg="#6C7A90",
                 font=("Segoe UI", 8, "bold")).pack(anchor="w")
        tk.Label(panel, text="Текст", bg="white", fg="#33415A",
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(16, 4))
        self.text_entry = tk.Text(panel, height=5, wrap="word", relief="flat", bd=0,
                                  padx=7, pady=6, font=("Segoe UI", 9),
                                  highlightthickness=1, highlightbackground="#D8E0EC")
        self.text_entry.pack(fill="x")
        self.text_entry.bind("<KeyRelease>", lambda _e: self.apply_properties())
        tk.Label(panel, text="Үсгийн хэмжээ", bg="white", fg="#33415A",
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(13, 4))
        self.size_var = tk.IntVar(value=28)
        self.size_spin = ttk.Spinbox(panel, from_=8, to=96, textvariable=self.size_var,
                                     command=self.apply_properties, width=9)
        self.size_spin.pack(anchor="w")
        self.size_spin.bind("<KeyRelease>", lambda _e: self.apply_properties())
        self.bold_var = tk.BooleanVar(value=False)
        tk.Checkbutton(panel, text="Тод үсэг", variable=self.bold_var, command=self.apply_properties,
                       bg="white", fg="#33415A", activebackground="white").pack(anchor="w", pady=8)
        tk.Label(panel, text="Өнгө", bg="white", fg="#33415A",
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(7, 4))
        color_row = tk.Frame(panel, bg="white"); color_row.pack(fill="x")
        self.color_swatch = tk.Label(color_row, text="      ", bg=theme.get("accent", "#3478F6"),
                                     relief="solid", bd=1)
        self.color_swatch.pack(side="left")
        tk.Button(color_row, text="Өнгө сонгох", command=self.choose_color, relief="flat",
                  bg="#EDF2F8", padx=8, pady=5).pack(side="left", padx=8)
        tk.Button(panel, text="Слайдын дэвсгэр өнгө", command=self.choose_background,
                  relief="flat", bg="#EDF2F8", padx=8, pady=7).pack(anchor="w", pady=(12, 0))
        tk.Label(panel, text="", bg="white").pack(expand=True, fill="both")
        if not self.single_slide:
            nav = tk.Frame(panel, bg="white"); nav.pack(fill="x", side="bottom")
            tk.Button(nav, text="← Өмнөх", command=lambda: self.change_slide(-1), relief="flat",
                      bg="#EDF2F8", pady=7).pack(side="left", fill="x", expand=True, padx=(0, 4))
            tk.Button(nav, text="Дараах →", command=lambda: self.change_slide(1), relief="flat",
                      bg="#EDF2F8", pady=7).pack(side="left", fill="x", expand=True, padx=(4, 0))
        else:
            self.page_label.config(text="A4 хэвтээ · Өргөмжлөл")
        tk.Label(panel, text="Элементийг сонгоод чирж байрлуулна.\nХадгалсны дараа файлдаа тус тусад нь засна.",
                 bg="white", fg="#758198", justify="left", font=("Segoe UI", 8),
                 wraplength=215).pack(side="bottom", anchor="w", pady=(12, 8))

        self.load_slide()

    def _default_elements(self, slide):
        elements = [
            {"type": "text", "text": slide.get("title", "Слайдын гарчиг"),
             "x": 62, "y": 66, "w": 820, "h": 75, "font_size": 32,
             "color": self.theme.get("ink", "#102B48"), "bold": True},
        ]
        points = slide.get("points", [])
        if isinstance(points, str): points = points.splitlines()
        for i, point in enumerate(points[:7]):
            elements.append({"type": "text", "text": "• " + str(point),
                             "x": 72, "y": 165 + i * 48, "w": 810, "h": 42,
                             "font_size": 19, "color": self.theme.get("ink", "#102B48"), "bold": False})
        return elements

    def load_slide(self):
        if not self.slides: self.slides.append({"title": "Шинэ слайд", "points": ["Гол санаа"]})
        slide = self.slides[self.index]
        self.elements = [dict(e) for e in slide.get("elements", [])] or self._default_elements(slide)
        self.selected = None
        self.page_label.config(text=f"Слайд {self.index + 1} / {len(self.slides)}")
        self._sync_properties(); self.draw()

    def _store_slide(self):
        self.slides[self.index]["elements"] = [dict(e) for e in self.elements]

    def change_slide(self, delta):
        self._store_slide()
        new_index = max(0, min(len(self.slides) - 1, self.index + delta))
        if new_index != self.index:
            self.index = new_index
            self.load_slide()

    def _scale(self):
        cw, ch = max(self.canvas.winfo_width() - 28, 320), max(self.canvas.winfo_height() - 28, 180)
        return min(cw / self.W, ch / self.H)

    def draw(self):
        if not hasattr(self, "canvas") or not self.canvas.winfo_exists(): return
        self.photo_refs = []
        self.canvas.delete("all")
        self.scale = self._scale()
        ox = (self.canvas.winfo_width() - self.W * self.scale) / 2
        oy = (self.canvas.winfo_height() - self.H * self.scale) / 2
        self.origin = (ox, oy)
        background = self.slides[self.index].get("background", self.theme.get("bg", "#FFFFFF"))
        self.canvas.create_rectangle(ox, oy, ox + self.W*self.scale, oy + self.H*self.scale,
                                     fill=background, outline="#C7D1E0")
        for idx, el in enumerate(self.elements):
            x, y, w, h = [float(el.get(k, default)) * self.scale for k, default in (("x", 60), ("y", 60), ("w", 300), ("h", 60))]
            x += ox; y += oy
            typ = el.get("type", "text")
            if typ in ("rect", "ellipse"):
                shape_fn = self.canvas.create_oval if typ == "ellipse" else self.canvas.create_rectangle
                shape_fn(x, y, x+w, y+h, fill=el.get("fill", self.theme.get("accent", "#3478F6")),
                         outline=el.get("color", ""), width=1, tags=(f"el{idx}",))
            elif typ == "image" and el.get("data"):
                try:
                    img = Image.open(BytesIO(base64.b64decode(el["data"])))
                    img.thumbnail((max(1, int(w)), max(1, int(h))))
                    photo = ImageTk.PhotoImage(img)
                    self.photo_refs.append(photo)
                    self.canvas.create_image(x+w/2, y+h/2, image=photo, tags=(f"el{idx}",))
                except Exception:
                    self.canvas.create_rectangle(x, y, x+w, y+h, outline="#D35454", tags=(f"el{idx}",))
            else:
                self.canvas.create_text(x, y, anchor="nw", text=el.get("text", "Текст"),
                    fill=el.get("color", "#17243A"), width=w,
                    font=("Segoe UI", max(7, int(float(el.get("font_size", 24))*self.scale)),
                          "bold" if el.get("bold") else "normal"), tags=(f"el{idx}",))
            if self.selected == idx:
                self.canvas.create_rectangle(x-3, y-3, x+w+3, y+h+3, outline="#3478F6",
                                             width=2, dash=(4, 2), tags=(f"sel{idx}",))

    def pointer_down(self, event):
        item = self.canvas.find_withtag("current")
        tags = self.canvas.gettags(item[-1]) if item else ()
        selected = next((int(t[2:]) for t in tags if t.startswith("el") and t[2:].isdigit()), None)
        self.selected = selected
        self.drag_origin = (event.x, event.y)
        self._sync_properties(); self.draw()

    def pointer_move(self, event):
        if self.selected is None or not self.drag_origin: return
        dx = (event.x - self.drag_origin[0]) / self.scale
        dy = (event.y - self.drag_origin[1]) / self.scale
        el = self.elements[self.selected]
        el["x"] = max(0, min(self.W - float(el.get("w", 100)), float(el.get("x", 0)) + dx))
        el["y"] = max(0, min(self.H - float(el.get("h", 40)), float(el.get("y", 0)) + dy))
        self.drag_origin = (event.x, event.y); self.draw()

    def pointer_up(self, _event): self.drag_origin = None

    def _sync_properties(self):
        el = self.elements[self.selected] if self.selected is not None and self.selected < len(self.elements) else None
        editable = bool(el and el.get("type") == "text")
        self.text_entry.configure(state="normal" if editable else "disabled")
        self.size_spin.configure(state="normal" if editable else "disabled")
        self.bold_var.set(bool(el and el.get("bold")))
        if editable:
            self.text_entry.configure(state="normal")
            self.text_entry.delete("1.0", "end")
            self.text_entry.insert("1.0", el.get("text", ""))
            try: self.size_var.set(int(el.get("font_size", 24)))
            except (ValueError, TypeError): self.size_var.set(24)
        else:
            self.text_entry.configure(state="normal")
            self.text_entry.delete("1.0", "end")
            self.text_entry.configure(state="disabled")
        color = (el or {}).get("color", (el or {}).get("fill", self.theme.get("accent", "#3478F6")))
        self.color_swatch.configure(bg=color if isinstance(color, str) else self.theme.get("accent", "#3478F6"))

    def apply_properties(self):
        if self.selected is None or self.selected >= len(self.elements): return
        el = self.elements[self.selected]
        if el.get("type") != "text": return
        el["text"] = self.text_entry.get("1.0", "end-1c")
        try: el["font_size"] = max(8, min(96, int(self.size_var.get())))
        except (ValueError, tk.TclError): pass
        el["bold"] = self.bold_var.get()
        self.draw()

    def choose_color(self):
        if self.selected is None: return
        el = self.elements[self.selected]
        initial = el.get("color", el.get("fill", self.theme.get("accent", "#3478F6")))
        chosen = colorchooser.askcolor(initialcolor=initial, parent=self.win)[1]
        if chosen:
            if el.get("type") in ("rect", "ellipse"): el["fill"] = chosen
            else: el["color"] = chosen
            self._sync_properties(); self.draw()

    def choose_background(self):
        slide = self.slides[self.index]
        initial = slide.get("background", self.theme.get("bg", "#FFFFFF"))
        chosen = colorchooser.askcolor(initialcolor=initial, parent=self.win)[1]
        if chosen:
            slide["background"] = chosen
            self.draw()

    def add_slide(self):
        self._store_slide()
        self.slides.append({"title":"Шинэ слайд", "points":["Шинэ санаагаа бичнэ үү."], "elements":[]})
        self.index = len(self.slides)-1
        self.load_slide()

    def add_text(self):
        self.elements.append({"type":"text", "text":"Шинэ текст", "x":300, "y":230,
                              "w":380, "h":60, "font_size":28,
                              "color":self.theme.get("ink", "#102B48"), "bold":True})
        self.selected = len(self.elements)-1; self._sync_properties(); self.draw()
        self.text_entry.focus_set(); self.text_entry.tag_add("sel", "1.0", "end-1c")

    def add_shape(self):
        self.elements.append({"type":"rect", "x":390, "y":245, "w":180, "h":80,
                              "fill":self.theme.get("accent", "#3478F6"), "color":""})
        self.selected = len(self.elements)-1; self._sync_properties(); self.draw()

    def add_image(self):
        filename = filedialog.askopenfilename(parent=self.win, title="Зураг сонгох",
                    filetypes=[("Зураг", "*.png *.jpg *.jpeg *.webp *.bmp"), ("Бүх файл", "*.*")])
        if not filename: return
        try:
            img = Image.open(filename).convert("RGBA")
            img.thumbnail((1400, 1400))
            stream = BytesIO(); img.save(stream, format="PNG", optimize=True)
            self.elements.append({"type":"image", "data":base64.b64encode(stream.getvalue()).decode("ascii"),
                                  "mime":"image/png", "x":340, "y":150, "w":280, "h":220})
            self.selected = len(self.elements)-1; self._sync_properties(); self.draw()
        except Exception as exc:
            messagebox.showerror("Зураг нээж чадсангүй", str(exc), parent=self.win)

    def delete_selected(self):
        if self.selected is None or self.selected >= len(self.elements): return
        self.elements.pop(self.selected); self.selected = None; self._sync_properties(); self.draw()

    def save(self):
        self._store_slide()
        self.on_save(self.slides)
        self.win.destroy()

