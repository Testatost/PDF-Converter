import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk

# Optional: Drag & Drop
HAVE_DND = True
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
except Exception:
    HAVE_DND = False
    TkinterDnD = tk
    DND_FILES = None


# --- DIN A4 exakte ISO-216-Maße ---
MM_PER_INCH = 25.4       # 1 Zoll = 25,4 mm
A4_WIDTH_MM = 210         # Breite in mm
A4_HEIGHT_MM = 297        # Höhe in mm
DPI = 300                 # gewünschte Druckauflösung (96, 150, 300 … DPI)

def mm_to_px(mm, dpi=DPI):
    """Umrechnung Millimeter → Pixel"""
    return round((mm / MM_PER_INCH) * dpi)



class ImageToPDFApp(TkinterDnD.Tk):
    A4_BORDER = 20
    GRID_SPACING_PAGE = 100

    def __init__(self):
        super().__init__()
        self.title("Bild zu PDF Konverter")
        self.geometry("1100x827")
        self.minsize(962, 827)
        self.configure(bg="#1e1e1e")

        # App-Zustand
        self.files = []
        self.img_cache = {}
        self.place = {}
        self.output_dir = self.get_desktop_path()

        self.orientation = tk.StringVar(value="Hochformat")  # Neu
        self.grid_enabled = tk.BooleanVar(value=True)
        self.a4_export = tk.BooleanVar(value=True)

        self.page_rect = (0, 0, 0, 0)
        self.page_scale = 1.0

        # Maussteuerung
        self.dragging = False
        self.last_canvas_xy = (0, 0)

        # Styles
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TButton", font=("Segoe UI", 10), padding=8, background="#0078D7", foreground="white")
        style.map("TButton", background=[("active", "#005a9e")])
        style.configure("TLabel", background="#1e1e1e", foreground="white", font=("Segoe UI", 10))
        style.configure("TFrame", background="#1e1e1e")
        style.configure("TCheckbutton", background="#1e1e1e", foreground="white", font=("Segoe UI", 10))
        style.configure("TRadiobutton", background="#1e1e1e", foreground="white", font=("Segoe UI", 10))

        # --- Neues, flexibles Layout ---
        main = ttk.Frame(self)
        main.grid(row=0, column=0, sticky="nsew", padx=14, pady=14)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        main.columnconfigure(0, weight=0)  # linke Spalte (Buttons)
        main.columnconfigure(1, weight=1)  # rechte Seite (Vorschau)
        main.rowconfigure(0, weight=1)

        left = ttk.Frame(main)
        left.grid(row=0, column=0, sticky="ns", padx=(0, 12))
        right = ttk.Frame(main)
        right.grid(row=0, column=1, sticky="nsew")

        # --- Linke Seite ---
        ttk.Label(left, text="Warteschlange").pack(anchor="w", pady=(0, 6))
        self.listbox = tk.Listbox(left, selectmode=tk.SINGLE, bg="#252526", fg="white",
                                  relief="flat", font=("Segoe UI", 9), height=22)
        self.listbox.pack(fill="both", expand=True)
        self.listbox.bind("<<ListboxSelect>>", self.on_select_file)
        if HAVE_DND:
            self.listbox.drop_target_register(DND_FILES)
            self.listbox.dnd_bind("<<Drop>>", self.drop_files)

        # Buttons
        btns = ttk.Frame(left)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text="Dateien auswählen", command=self.add_files).pack(fill="x", pady=2)
        ttk.Button(btns, text="Löschen", command=self.delete_selected).pack(fill="x", pady=2)
        ttk.Button(btns, text="Zielordner wählen", command=self.choose_output_folder).pack(fill="x", pady=2)
        ttk.Button(btns, text="Start", command=self.convert_to_pdf).pack(fill="x", pady=(10, 0))

        self.folder_label = ttk.Label(left, text=f"Ziel: {self.output_dir}", wraplength=260)
        self.folder_label.pack(pady=(8, 4))

        # Optionen
        opts = ttk.Frame(left)
        opts.pack(fill="x", pady=(4, 0))
        ttk.Checkbutton(opts, text="Raster anzeigen", variable=self.grid_enabled,
                        command=lambda: self.redraw_canvas()).pack(anchor="w")
        ttk.Checkbutton(left, text="Als DIN A4 speichern",
                        variable=self.a4_export,
                        command=lambda: self.on_toggle_a4_export()).pack(anchor="w", pady=(6, 0))

        # Orientierungsauswahl
        oriframe = ttk.LabelFrame(left, text="Ausrichtung")
        oriframe.pack(fill="x", pady=(10, 0))
        ttk.Radiobutton(oriframe, text="Hochformat", value="Hochformat",
                        variable=self.orientation, command=self.on_orientation_change).pack(anchor="w")
        ttk.Radiobutton(oriframe, text="Querformat", value="Querformat",
                        variable=self.orientation, command=self.on_orientation_change).pack(anchor="w")

        ttk.Label(left, text="Hinweis: Ziehen zum Verschieben, Mausrad zum Zoomen").pack(anchor="w", pady=(8, 0))

        # --- Rechte Seite ---
        ttk.Label(right, text="Vorschau").pack(anchor="w", pady=(0, 6))
        self.canvas = tk.Canvas(right, bg="#111111", highlightthickness=0, cursor="tcross")
        self.canvas.pack(fill="both", expand=True)

        # Canvas-Bindings
        self.canvas.bind("<Configure>", self.on_canvas_resize)
        self.canvas.bind("<Button-1>", self.on_canvas_press)
        self.canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_canvas_release)
        self.canvas.bind("<MouseWheel>", self.on_mousewheel)
        self.canvas.bind("<Button-4>", lambda e: self.on_mousewheel(e, delta=120))
        self.canvas.bind("<Button-5>", lambda e: self.on_mousewheel(e, delta=-120))

        # Shortcuts
        self.bind("<Return>", lambda e: self.convert_to_pdf())
        self.bind("<Delete>", lambda e: self.delete_selected())

        self.redraw_canvas()

    # ----------------- Utils -----------------
    def get_desktop_path(self):
        if os.name == "nt":
            return os.path.join(os.environ.get("USERPROFILE", os.path.expanduser("~")), "Desktop")
        else:
            path = os.path.join(os.path.expanduser("~"), "Schreibtisch")
            return path if os.path.isdir(path) else os.path.join(os.path.expanduser("~"), "Desktop")

    def get_a4_dimensions(self):
        """Aktuelle A4-Abmessungen in Pixeln, abhängig von der Orientierung"""
        if self.orientation.get() == "Querformat":
            return mm_to_px(A4_HEIGHT_MM), mm_to_px(A4_WIDTH_MM)
        return mm_to_px(A4_WIDTH_MM), mm_to_px(A4_HEIGHT_MM)

    # ----------------- Dateioperationen -----------------
    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Bilder auswählen",
            filetypes=[("Bilddateien", "*.png;*.jpg;*.jpeg;*.bmp;*.tiff;*.webp")]
        )
        self._add_to_list(paths)

    def drop_files(self, event):
        if not HAVE_DND:
            return
        files = self.tk.splitlist(event.data)
        self._add_to_list(files)

    def _add_to_list(self, files):
        added_any = False
        for f in files:
            if not (os.path.isfile(f) and f.lower().endswith((".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".webp"))):
                continue
            if f in self.files:
                continue
            self.files.append(f)
            self.listbox.insert(tk.END, os.path.basename(f))
            self.place[f] = {"scale": None, "x": None, "y": None}
            added_any = True

        if added_any and self.listbox.size() == len(files):
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(0)
            self.on_select_file()

    def delete_selected(self):
        idxs = self.listbox.curselection()
        if not idxs:
            return
        i = idxs[0]
        f = self.files[i]
        self.listbox.delete(i)
        del self.files[i]
        self.img_cache.pop(f, None)
        self.place.pop(f, None)
        self.redraw_canvas()
        if self.files:
            i = min(i, len(self.files) - 1)
            self.listbox.selection_set(i)
            self.on_select_file()

    def choose_output_folder(self):
        folder = filedialog.askdirectory(title="Zielordner auswählen", initialdir=self.output_dir)
        if folder:
            self.output_dir = folder
            self.folder_label.config(text=f"Ziel: {folder}")

    # ----------------- Canvas -----------------
    def on_canvas_resize(self, event=None):
        self.redraw_canvas()

    def redraw_canvas(self):
        self.canvas.delete("all")
        cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
        if cw <= 1 or ch <= 1:
            return

        A4_W, A4_H = self.get_a4_dimensions()
        a4_ratio = A4_W / A4_H

        if cw / ch > a4_ratio:
            page_h = ch * 0.96
            page_w = page_h * a4_ratio
        else:
            page_w = cw * 0.96
            page_h = page_w / a4_ratio

        x0 = (cw - page_w) / 2
        y0 = (ch - page_h) / 2
        x1 = x0 + page_w
        y1 = y0 + page_h
        self.page_rect = (x0, y0, x1, y1)
        self.page_scale = page_w / A4_W

        self.canvas.create_rectangle(0, 0, cw, ch, fill="#111111", outline="")
        self.canvas.create_rectangle(x0, y0, x1, y1, fill="white", outline="#dddddd")

        if self.grid_enabled.get():
            self.draw_grid(A4_W, A4_H)

        cur = self.current_file()
        if cur:
            self.draw_image_on_page(cur, A4_W, A4_H)

    def draw_grid(self, A4_W, A4_H):
        step = self.GRID_SPACING_PAGE * self.page_scale
        x0, y0, x1, y1 = self.page_rect
        shade = "#999999"
        for i in range(1, int((x1 - x0) / step)):
            x = x0 + i * step
            self.canvas.create_line(x, y0, x, y1, fill=shade)
        for j in range(1, int((y1 - y0) / step)):
            y = y0 + j * step
            self.canvas.create_line(x0, y, x1, y, fill=shade)
        border = self.A4_BORDER * self.page_scale
        self.canvas.create_rectangle(x0 + border, y0 + border, x1 - border, y1 - border, outline="#bbbbbb")

    # ----------------- Bild & Interaktion -----------------
    def current_file(self):
        sel = self.listbox.curselection()
        if not sel:
            return None
        return self.files[sel[0]]

    def ensure_loaded(self, f):
        if f in self.img_cache:
            return self.img_cache[f]
        try:
            img = Image.open(f).convert("RGB")
            self.img_cache[f] = img
            if self.place[f]["scale"] is None:
                A4_W, A4_H = self.get_a4_dimensions()
                max_w = A4_W - 2 * self.A4_BORDER
                max_h = A4_H - 2 * self.A4_BORDER
                s = min(max_w / img.width, max_h / img.height, 1.0)
                w = img.width * s
                h = img.height * s
                self.place[f] = {"scale": s, "x": (A4_W - w) / 2, "y": (A4_H - h) / 2}
            return img
        except Exception as e:
            messagebox.showerror("Fehler", f"Bild konnte nicht geladen werden:\n{f}\n{e}")
            return None

    def max_scale_for_a4(self, img):
        A4_W, A4_H = self.get_a4_dimensions()
        max_w = A4_W - 2 * self.A4_BORDER
        max_h = A4_H - 2 * self.A4_BORDER
        return min(max_w / img.width, max_h / img.height)

    def draw_image_on_page(self, f, A4_W, A4_H):
        img = self.ensure_loaded(f)
        if img is None:
            return
        p = self.place[f]
        if self.a4_export.get():
            cap = self.max_scale_for_a4(img)
            if p["scale"] > cap:
                p["scale"] = cap
        scale = p["scale"]
        x_page, y_page = p["x"], p["y"]
        w_page = img.width * scale
        h_page = img.height * scale

        x0, y0, _, _ = self.page_rect
        sx = self.page_scale
        cx = x0 + x_page * sx
        cy = y0 + y_page * sx
        cw, ch = int(w_page * sx), int(h_page * sx)

        if cw < 2 or ch < 2:
            return
        preview = img.resize((int(w_page), int(h_page)), Image.LANCZOS)
        preview = preview.resize((cw, ch), Image.LANCZOS)
        self._tk_img = ImageTk.PhotoImage(preview)
        self.canvas.create_image(cx, cy, image=self._tk_img, anchor="nw")

    # ----------------- Maussteuerung -----------------
    def canvas_to_page(self, x, y):
        x0, y0, _, _ = self.page_rect
        return (x - x0) / self.page_scale, (y - y0) / self.page_scale

    def in_page(self, x, y):
        x0, y0, x1, y1 = self.page_rect
        return x0 <= x <= x1 and y0 <= y <= y1

    def on_canvas_press(self, event):
        if self.current_file() and self.in_page(event.x, event.y):
            self.dragging = True
            self.last_canvas_xy = (event.x, event.y)

    def on_canvas_drag(self, event):
        if not self.dragging:
            return
        f = self.current_file()
        if not f:
            return
        dx, dy = event.x - self.last_canvas_xy[0], event.y - self.last_canvas_xy[1]
        self.last_canvas_xy = (event.x, event.y)
        pdx, pdy = dx / self.page_scale, dy / self.page_scale
        p = self.place[f]
        p["x"] += pdx
        p["y"] += pdy
        if self.a4_export.get():
            self.keep_inside_page(f)
        self.redraw_canvas()

    def on_canvas_release(self, event):
        self.dragging = False

    def on_mousewheel(self, event, delta=None):
        f = self.current_file()
        if not f or not self.in_page(event.x, event.y):
            return
        img = self.ensure_loaded(f)
        if img is None:
            return
        d = delta if delta is not None else event.delta
        z = 1.1 if d > 0 else 1 / 1.1
        p = self.place[f]
        max_scale = self.max_scale_for_a4(img) if self.a4_export.get() else 5.0
        new_scale = max(0.05, min(max_scale, p["scale"] * z))
        if abs(new_scale - p["scale"]) < 1e-6:
            return
        mx, my = self.canvas_to_page(event.x, event.y)
        old_w, old_h = img.width * p["scale"], img.height * p["scale"]
        new_w, new_h = img.width * new_scale, img.height * new_scale
        rel_x = (mx - p["x"]) / max(old_w, 1)
        rel_y = (my - p["y"]) / max(old_h, 1)
        p["scale"] = new_scale
        p["x"] = mx - rel_x * new_w
        p["y"] = my - rel_y * new_h
        if self.a4_export.get():
            self.keep_inside_page(f)
        self.redraw_canvas()

    def keep_inside_page(self, f):
        img = self.img_cache.get(f)
        if not img:
            return
        A4_W, A4_H = self.get_a4_dimensions()
        p = self.place[f]
        w, h = img.width * p["scale"], img.height * p["scale"]
        p["x"] = min(max(p["x"], self.A4_BORDER), A4_W - self.A4_BORDER - w)
        p["y"] = min(max(p["y"], self.A4_BORDER), A4_H - self.A4_BORDER - h)

    # ----------------- Events -----------------
    def on_select_file(self, event=None):
        f = self.current_file()
        if not f:
            self.redraw_canvas()
            return
        def loader():
            self.ensure_loaded(f)
            self.after(0, self.redraw_canvas)
        threading.Thread(target=loader, daemon=True).start()

    def on_orientation_change(self):
        """Beim Umschalten Hoch/Quer: Seite neu berechnen und Bild ggf. wieder einpassen."""
        f = self.current_file()
        if f:
            img = self.ensure_loaded(f)
            if img is not None and self.a4_export.get():
                cap = self.max_scale_for_a4(img)
                p = self.place[f]
                if p["scale"] > cap:
                    p["scale"] = cap
                self.keep_inside_page(f)
        self.redraw_canvas()

    def on_toggle_a4_export(self):
        f = self.current_file()
        if not f:
            self.redraw_canvas()
            return
        img = self.ensure_loaded(f)
        if img is None:
            self.redraw_canvas()
            return
        if self.a4_export.get():
            cap = self.max_scale_for_a4(img)
            p = self.place[f]
            if p["scale"] > cap:
                p["scale"] = cap
            self.keep_inside_page(f)
        self.redraw_canvas()

    # ----------------- Export -----------------
    def convert_to_pdf(self):
        if not self.files:
            messagebox.showwarning("Keine Dateien", "Bitte zuerst Bilder hinzufügen.")
            return

        errors = 0
        for f in self.files:
            try:
                img = self.ensure_loaded(f)
                if img is None:
                    errors += 1
                    continue

                if self.a4_export.get():
                    A4_W, A4_H = self.get_a4_dimensions()
                    p = self.place[f]

                    # Begrenzung sicherheitshalber anwenden
                    cap = self.max_scale_for_a4(img)
                    scale = max(0.05, min(cap, p["scale"]))
                    w, h = int(img.width * scale), int(img.height * scale)
                    if w < 1 or h < 1:
                        errors += 1
                        continue
                    resized = img.resize((w, h), Image.LANCZOS)

                    x = int(max(self.A4_BORDER, min(p["x"], A4_W - self.A4_BORDER - w)))
                    y = int(max(self.A4_BORDER, min(p["y"], A4_H - self.A4_BORDER - h)))

                    page = Image.new("RGB", (A4_W, A4_H), "white")
                    page.paste(resized, (x, y))

                    base = os.path.splitext(os.path.basename(f))[0]
                    out_path = os.path.join(self.output_dir, f"{base}.pdf")
                    page.save(out_path, "PDF")
                else:
                    base = os.path.splitext(os.path.basename(f))[0]
                    out_path = os.path.join(self.output_dir, f"{base}.pdf")
                    img.convert("RGB").save(out_path, "PDF")

            except Exception as e:
                print("Fehler beim Export:", e)
                errors += 1

        if errors == 0:
            messagebox.showinfo("Fertig", f"Alle PDFs wurden in\n{self.output_dir}\nerstellt.")
        else:
            messagebox.showwarning("Hinweis", f"Einige Dateien konnten nicht konvertiert werden. Fehler: {errors}")


if __name__ == "__main__":
    app = ImageToPDFApp()
    app.mainloop()
