import os
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk

# Check for python-docx
try:
    from docx import Document
    from docx.shared import Inches
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Photo Selector & Report Generator v10")
        self.root.geometry("1500x900")
        self.photos, self.thumbs, self.checks = [], {}, {}
        self.preview_img, self.zoom_level, self.original_image = None, 1.0, None
        self.pan_x, self.pan_y = 0, 0
        self.crop_t, self.crop_b, self.crop_l, self.crop_r = 0, 0, 0, 0
        self.current_photo = None
        self.selected_for_rename, self.output_folder, self.current_rename_index = [], "", 0
        self.renamed_photos = []  # Store renamed photo paths for report generation
        
        # Top frame
        top = ttk.Frame(root, padding=10)
        top.pack(fill=tk.X)
        ttk.Label(top, text="Boring ID:").pack(side=tk.LEFT)
        self.vb_id = tk.StringVar()
        ttk.Entry(top, textvariable=self.vb_id, width=15).pack(side=tk.LEFT, padx=(5,20))
        ttk.Label(top, text="Folder:").pack(side=tk.LEFT)
        self.path = tk.StringVar()
        ttk.Entry(top, textvariable=self.path, width=40).pack(side=tk.LEFT, padx=5)
        ttk.Button(top, text="Browse", command=self.browse).pack(side=tk.LEFT)
        ttk.Button(top, text="Load", command=self.load).pack(side=tk.LEFT, padx=5)
        ttk.Button(top, text="Select All", command=self.sel_all).pack(side=tk.LEFT, padx=10)
        ttk.Button(top, text="Clear All", command=self.clr_all).pack(side=tk.LEFT)
        
        # Main frame
        main = ttk.Frame(root)
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Left - photo grid
        left = ttk.LabelFrame(main, text="Photos - Click to Preview", padding=5)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.canvas = tk.Canvas(left, bg="white")
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.inner = ttk.Frame(self.canvas)
        self.canvas.create_window((0,0), window=self.inner, anchor=tk.NW)
        
        # Right panel
        right = ttk.Frame(main, width=450)
        right.pack(side=tk.RIGHT, fill=tk.Y, padx=(10,0))
        right.pack_propagate(False)
        
        # Preview
        pf = ttk.LabelFrame(right, text="Preview", padding=5)
        pf.pack(fill=tk.BOTH, expand=True)
        self.pcv = tk.Canvas(pf, bg="#e0e0e0", width=400, height=220)
        self.pcv.pack(fill=tk.BOTH, expand=True)
        
        # Zoom buttons
        zf = ttk.Frame(pf)
        zf.pack(fill=tk.X, pady=2)
        ttk.Button(zf, text="Zoom+", command=self.zoom_in, width=6).pack(side=tk.LEFT, padx=2)
        ttk.Button(zf, text="Zoom-", command=self.zoom_out, width=6).pack(side=tk.LEFT, padx=2)
        ttk.Button(zf, text="Reset", command=self.zoom_reset, width=6).pack(side=tk.LEFT, padx=2)
        
        # Pan buttons
        pnf = ttk.Frame(pf)
        pnf.pack(fill=tk.X, pady=2)
        ttk.Button(pnf, text="Up", width=5, command=self.pan_up).pack(side=tk.LEFT, padx=2)
        ttk.Button(pnf, text="Down", width=5, command=self.pan_down).pack(side=tk.LEFT, padx=2)
        ttk.Button(pnf, text="Left", width=5, command=self.pan_left).pack(side=tk.LEFT, padx=2)
        ttk.Button(pnf, text="Right", width=5, command=self.pan_right).pack(side=tk.LEFT, padx=2)
        
        # Crop tool
        cpf = ttk.LabelFrame(right, text="Crop Tool (%)", padding=5)
        cpf.pack(fill=tk.X, pady=(5,0))
        
        cf1 = ttk.Frame(cpf)
        cf1.pack(fill=tk.X, pady=2)
        ttk.Label(cf1, text="T:").pack(side=tk.LEFT)
        ttk.Button(cf1, text="+", width=2, command=lambda: self.adj_crop('t',5)).pack(side=tk.LEFT)
        ttk.Button(cf1, text="-", width=2, command=lambda: self.adj_crop('t',-5)).pack(side=tk.LEFT)
        ttk.Label(cf1, text=" B:").pack(side=tk.LEFT)
        ttk.Button(cf1, text="+", width=2, command=lambda: self.adj_crop('b',5)).pack(side=tk.LEFT)
        ttk.Button(cf1, text="-", width=2, command=lambda: self.adj_crop('b',-5)).pack(side=tk.LEFT)
        ttk.Label(cf1, text=" L:").pack(side=tk.LEFT)
        ttk.Button(cf1, text="+", width=2, command=lambda: self.adj_crop('l',5)).pack(side=tk.LEFT)
        ttk.Button(cf1, text="-", width=2, command=lambda: self.adj_crop('l',-5)).pack(side=tk.LEFT)
        ttk.Label(cf1, text=" R:").pack(side=tk.LEFT)
        ttk.Button(cf1, text="+", width=2, command=lambda: self.adj_crop('r',5)).pack(side=tk.LEFT)
        ttk.Button(cf1, text="-", width=2, command=lambda: self.adj_crop('r',-5)).pack(side=tk.LEFT)
        
        cf2 = ttk.Frame(cpf)
        cf2.pack(fill=tk.X, pady=2)
        ttk.Button(cf2, text="Square", command=self.crop_square, width=7).pack(side=tk.LEFT, padx=2)
        ttk.Button(cf2, text="Reset", command=self.crop_reset, width=7).pack(side=tk.LEFT, padx=2)
        
        self.crop_lbl = ttk.Label(cpf, text="Crop: T0 B0 L0 R0 %")
        self.crop_lbl.pack(anchor=tk.W)
        
        # Info
        inf = ttk.LabelFrame(right, text="Properties", padding=5)
        inf.pack(fill=tk.X, pady=(5,0))
        self.lbl_n = ttk.Label(inf, text="Name: -")
        self.lbl_n.pack(anchor=tk.W)
        self.lbl_s = ttk.Label(inf, text="Size: -")
        self.lbl_s.pack(anchor=tk.W)
        
        # Selected list
        sf = ttk.LabelFrame(right, text="Selected", padding=5)
        sf.pack(fill=tk.BOTH, expand=True, pady=(5,0))
        self.slist = tk.Listbox(sf, height=4)
        self.slist.pack(fill=tk.BOTH, expand=True)
        
        # Bottom
        bot = ttk.Frame(root, padding=10)
        bot.pack(fill=tk.X)
        self.status = ttk.Label(bot, text="No photos")
        self.status.pack(side=tk.LEFT)
        self.scnt = ttk.Label(bot, text="Selected: 0")
        self.scnt.pack(side=tk.LEFT, padx=20)
        ttk.Button(bot, text="Generate Report", command=self.generate_report).pack(side=tk.RIGHT, padx=5)
        ttk.Button(bot, text="Rename and Save", command=self.start_rename).pack(side=tk.RIGHT, padx=5)

    def browse(self):
        f = filedialog.askdirectory()
        if f: self.path.set(f)
    
    def sel_all(self):
        for v in self.checks.values(): v.set(True)
        self.upd_sel()
    
    def clr_all(self):
        for v in self.checks.values(): v.set(False)
        self.upd_sel()
    
    def zoom_in(self): self.zoom_level = min(5, self.zoom_level * 1.3); self.upd_pv()
    def zoom_out(self): self.zoom_level = max(0.3, self.zoom_level / 1.3); self.upd_pv()
    def zoom_reset(self): self.zoom_level = 1.0; self.pan_x = 0; self.pan_y = 0; self.upd_pv()
    def pan_up(self): self.pan_y -= 30; self.upd_pv()
    def pan_down(self): self.pan_y += 30; self.upd_pv()
    def pan_left(self): self.pan_x -= 30; self.upd_pv()
    def pan_right(self): self.pan_x += 30; self.upd_pv()
    
    def adj_crop(self, side, val):
        if side == 't': self.crop_t = max(0, self.crop_t + val)
        elif side == 'b': self.crop_b = max(0, self.crop_b + val)
        elif side == 'l': self.crop_l = max(0, self.crop_l + val)
        elif side == 'r': self.crop_r = max(0, self.crop_r + val)
        self.upd_pv()
    
    def crop_reset(self):
        self.crop_t = self.crop_b = self.crop_l = self.crop_r = 0
        self.upd_pv()
    
    def crop_square(self):
        if not self.original_image: return
        w, h = self.original_image.size
        if w > h:
            self.crop_l = self.crop_r = (w - h) * 50 // w
            self.crop_t = self.crop_b = 0
        else:
            self.crop_t = self.crop_b = (h - w) * 50 // h
            self.crop_l = self.crop_r = 0
        self.upd_pv()
    
    def upd_sel(self):
        self.slist.delete(0, tk.END)
        sel = [p for p, v in self.checks.items() if v.get()]
        self.scnt.config(text="Selected: " + str(len(sel)))
        for p in sel: self.slist.insert(tk.END, os.path.basename(p))
    
    def load(self):
        p = self.path.get()
        if not p or not os.path.isdir(p): return
        for w in self.inner.winfo_children(): w.destroy()
        self.photos, self.thumbs, self.checks = [], {}, {}
        exts = {'.jpg', '.jpeg', '.png', '.gif', '.bmp'}
        for f in sorted(os.listdir(p)):
            if os.path.splitext(f)[1].lower() in exts:
                self.photos.append(os.path.join(p, f))
        r, c = 0, 0
        for ph in self.photos:
            fr = ttk.Frame(self.inner, padding=3)
            fr.grid(row=r, column=c, padx=5, pady=5)
            try:
                im = Image.open(ph)
                im.thumbnail((100, 100))
                tk_im = ImageTk.PhotoImage(im)
                self.thumbs[ph] = tk_im
                tk.Button(fr, image=tk_im, command=lambda x=ph: self.show(x)).pack()
            except: pass
            ttk.Label(fr, text=os.path.basename(ph)[:15]).pack()
            v = tk.BooleanVar()
            self.checks[ph] = v
            ttk.Checkbutton(fr, text="Select", variable=v, command=self.upd_sel).pack()
            c += 1
            if c > 5: c, r = 0, r + 1
        self.status.config(text="Loaded " + str(len(self.photos)))
    
    def show(self, path):
        self.current_photo = path
        self.zoom_level = 1.0
        self.pan_x = self.pan_y = 0
        self.crop_t = self.crop_b = self.crop_l = self
