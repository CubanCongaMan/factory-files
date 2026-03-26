#03/25/2026 version at 12:12 pm of Photo Selector Version 9 Author: RLS & Factory AI Droids
import os
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Photo Selector - Core Samples")
        self.root.geometry("1400x850")
        self.photos, self.thumbs, self.checks = [], {}, {}
        self.preview_img, self.zoom_level, self.original_image = None, 1.0, None
        self.pan_x, self.pan_y = 0, 0
        self.selected_for_rename, self.output_folder, self.current_rename_index = [], "", 0
        self.current_photo_path = None
        self.crop_mode = False
        self.crop_rect = None
        self.crop_handles = []
        self.crop_start = None
        self.active_handle = None
        self.crop_coords = [50, 50, 200, 200]
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
        main = ttk.Frame(root)
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        left = ttk.LabelFrame(main, text="Photos - Click to Preview", padding=5)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.canvas = tk.Canvas(left, bg="white")
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.inner = ttk.Frame(self.canvas)
        self.canvas.create_window((0,0), window=self.inner, anchor=tk.NW)
        right = ttk.Frame(main, width=420)
        right.pack(side=tk.RIGHT, fill=tk.Y, padx=(10,0))
        right.pack_propagate(False)
        pf = ttk.LabelFrame(right, text="Preview", padding=5)
        pf.pack(fill=tk.BOTH, expand=True)
        self.pcv = tk.Canvas(pf, bg="#e0e0e0", width=380, height=280)
        self.pcv.pack(fill=tk.BOTH, expand=True)
        zf = ttk.Frame(pf)
        zf.pack(fill=tk.X, pady=3)
        ttk.Button(zf, text="Zoom +", command=self.zoom_in).pack(side=tk.LEFT, padx=3)
        ttk.Button(zf, text="Zoom -", command=self.zoom_out).pack(side=tk.LEFT, padx=3)
        ttk.Button(zf, text="Reset", command=self.zoom_reset).pack(side=tk.LEFT, padx=3)
        self.crop_btn = ttk.Button(zf, text="Crop", command=self.toggle_crop_mode)
        self.crop_btn.pack(side=tk.LEFT, padx=10)
        self.crop_frame = ttk.Frame(pf)
        self.apply_crop_btn = ttk.Button(self.crop_frame, text="Apply Crop", command=self.apply_crop)
        self.apply_crop_btn.pack(side=tk.LEFT, padx=3)
        self.cancel_crop_btn = ttk.Button(self.crop_frame, text="Cancel", command=self.cancel_crop)
        self.cancel_crop_btn.pack(side=tk.LEFT, padx=3)
        pnf = ttk.Frame(pf)
        pnf.pack(fill=tk.X, pady=3)
        ttk.Button(pnf, text="Up", width=5, command=self.pan_up).pack(side=tk.LEFT, padx=2)
        ttk.Button(pnf, text="Down", width=5, command=self.pan_down).pack(side=tk.LEFT, padx=2)
        ttk.Button(pnf, text="Left", width=5, command=self.pan_left).pack(side=tk.LEFT, padx=2)
        ttk.Button(pnf, text="Right", width=5, command=self.pan_right).pack(side=tk.LEFT, padx=2)
        inf = ttk.LabelFrame(right, text="Properties", padding=5)
        inf.pack(fill=tk.X, pady=(5,0))
        self.lbl_n = ttk.Label(inf, text="Name: -")
        self.lbl_n.pack(anchor=tk.W)
        self.lbl_s = ttk.Label(inf, text="Size: -")
        self.lbl_s.pack(anchor=tk.W)
        self.lbl_f = ttk.Label(inf, text="Format: -")
        self.lbl_f.pack(anchor=tk.W)
        sf = ttk.LabelFrame(right, text="Selected", padding=5)
        sf.pack(fill=tk.BOTH, expand=True, pady=(5,0))
        self.slist = tk.Listbox(sf, height=5)
        self.slist.pack(fill=tk.BOTH, expand=True)
        bot = ttk.Frame(root, padding=10)
        bot.pack(fill=tk.X)
        self.status = ttk.Label(bot, text="No photos")
        self.status.pack(side=tk.LEFT)
        self.scnt = ttk.Label(bot, text="Selected: 0")
        self.scnt.pack(side=tk.LEFT, padx=20)
        ttk.Button(bot, text="Rename and Save", command=self.start_rename).pack(side=tk.RIGHT)

    def browse(self): f=filedialog.askdirectory(); self.path.set(f) if f else None
    def sel_all(self): [v.set(True) for v in self.checks.values()]; self.upd_sel()
    def clr_all(self): [v.set(False) for v in self.checks.values()]; self.upd_sel()

    def zoom_in(self):
        old_zoom = self.zoom_level
        self.zoom_level = min(5, self.zoom_level * 1.3)
        self.adjust_crop_for_zoom(old_zoom, self.zoom_level)
        self.upd_pv()
        if self.crop_mode:
            self.draw_crop_rect()

    def zoom_out(self):
        old_zoom = self.zoom_level
        self.zoom_level = max(0.3, self.zoom_level / 1.3)
        self.adjust_crop_for_zoom(old_zoom, self.zoom_level)
        self.upd_pv()
        if self.crop_mode:
            self.draw_crop_rect()

    def zoom_reset(self):
        self.zoom_level = 1.0
        self.pan_x = 0
        self.pan_y = 0
        self.crop_coords = [50, 50, 200, 200]
        self.upd_pv()
        if self.crop_mode:
            self.draw_crop_rect()

    def pan_up(self):
        self.pan_y -= 30
        self.adjust_crop_for_pan(0, -30)
        self.upd_pv()
        if self.crop_mode:
            self.draw_crop_rect()

    def pan_down(self):
        self.pan_y += 30
        self.adjust_crop_for_pan(0, 30)
        self.upd_pv()
        if self.crop_mode:
            self.draw_crop_rect()

    def pan_left(self):
        self.pan_x -= 30
        self.adjust_crop_for_pan(-30, 0)
        self.upd_pv()
        if self.crop_mode:
            self.draw_crop_rect()

    def pan_right(self):
        self.pan_x += 30
        self.adjust_crop_for_pan(30, 0)
        self.upd_pv()
        if self.crop_mode:
            self.draw_crop_rect()

    def adjust_crop_for_zoom(self, old_zoom, new_zoom):
        if not self.crop_mode:
            return
        cv_w, cv_h = self.pcv.winfo_width(), self.pcv.winfo_height()
        center_x, center_y = cv_w // 2 + self.pan_x, cv_h // 2 + self.pan_y
        scale = new_zoom / old_zoom
        x1, y1, x2, y2 = self.crop_coords
        self.crop_coords = [
            int(center_x + (x1 - center_x) * scale),
            int(center_y + (y1 - center_y) * scale),
            int(center_x + (x2 - center_x) * scale),
            int(center_y + (y2 - center_y) * scale)
        ]

    def adjust_crop_for_pan(self, dx, dy):
        if not self.crop_mode:
            return
        x1, y1, x2, y2 = self.crop_coords
        self.crop_coords = [x1 + dx, y1 + dy, x2 + dx, y2 + dy]

    def toggle_crop_mode(self):
        if not self.original_image:
            messagebox.showwarning("Warning", "Load a photo first")
            return
        self.crop_mode = not self.crop_mode
        if self.crop_mode:
            self.crop_btn.config(text="Cropping...")
            self.crop_frame.pack(fill=tk.X, pady=3)
            self.crop_coords = [50, 50, 200, 200]
            self.upd_pv()
            self.draw_crop_rect()
            self.pcv.bind("<Button-1>", self.crop_mouse_down)
            self.pcv.bind("<B1-Motion>", self.crop_mouse_drag)
            self.pcv.bind("<ButtonRelease-1>", self.crop_mouse_up)
        else:
            self.cancel_crop()

    def draw_crop_rect(self):
        self.pcv.delete("crop")
        x1, y1, x2, y2 = self.crop_coords
        self.crop_rect = self.pcv.create_rectangle(x1, y1, x2, y2, outline="red", width=2, tags="crop", dash=(4, 2))
        handle_size = 12
        handles_pos = [
            (x1, y1), (x2, y1), (x1, y2), (x2, y2),
            ((x1+x2)//2, y1), ((x1+x2)//2, y2), (x1, (y1+y2)//2), (x2, (y1+y2)//2)
        ]
        self.crop_handles = []
        for i, (hx, hy) in enumerate(handles_pos):
            if i < 4:
                fill_color = "blue"
            else:
                fill_color = "green"
            h = self.pcv.create_rectangle(hx-handle_size//2, hy-handle_size//2, 
                                          hx+handle_size//2, hy+handle_size//2,
                                          fill=fill_color, outline="white", width=2, tags="crop")
            self.crop_handles.append(h)

    def crop_mouse_down(self, event):
        x, y = event.x, event.y
        handle_size = 15
        x1, y1, x2, y2 = self.crop_coords
        handles_pos = [
            (x1, y1), (x2, y1), (x1, y2), (x2, y2),
            ((x1+x2)//2, y1), ((x1+x2)//2, y2), (x1, (y1+y2)//2), (x2, (y1+y2)//2)
        ]
        for i, (hx, hy) in enumerate(handles_pos):
            if abs(x - hx) < handle_size and abs(y - hy) < handle_size:
                self.active_handle = i
                self.crop_start = (x, y)
                return
        min_x, max_x = min(x1, x2), max(x1, x2)
        min_y, max_y = min(y1, y2), max(y1, y2)
        if min_x < x < max_x and min_y < y < max_y:
            self.active_handle = "move"
            self.crop_start = (x, y)

    def crop_mouse_drag(self, event):
        if not self.crop_start or self.active_handle is None:
            return
        x, y = event.x, event.y
        dx, dy = x - self.crop_start[0], y - self.crop_start[1]
        x1, y1, x2, y2 = self.crop_coords
        if self.active_handle == "move":
            self.crop_coords = [x1+dx, y1+dy, x2+dx, y2+dy]
        elif self.active_handle == 0:
            self.crop_coords = [x1+dx, y1+dy, x2, y2]
        elif self.active_handle == 1:
            self.crop_coords = [x1, y1+dy, x2+dx, y2]
        elif self.active_handle == 2:
            self.crop_coords = [x1+dx, y1, x2, y2+dy]
        elif self.active_handle == 3:
            self.crop_coords = [x1, y1, x2+dx, y2+dy]
        elif self.active_handle == 4:
            self.crop_coords = [x1, y1+dy, x2, y2]
        elif self.active_handle == 5:
            self.crop_coords = [x1, y1, x2, y2+dy]
        elif self.active_handle == 6:
            self.crop_coords = [x1+dx, y1, x2, y2]
        elif self.active_handle == 7:
            self.crop_coords = [x1, y1, x2+dx, y2]
        self.crop_start = (x, y)
        self.draw_crop_rect()

    def crop_mouse_up(self, event):
        self.active_handle = None
        self.crop_start = None

    def apply_crop(self):
        if not self.original_image or not self.current_photo_path:
            return
        cv_w, cv_h = self.pcv.winfo_width(), self.pcv.winfo_height()
        img_w, img_h = self.original_image.size
        
        # Calculate the displayed image size at current zoom level
        max_disp = int(360 * self.zoom_level)
        scale = min(max_disp / img_w, max_disp / img_h)
        disp_w = int(img_w * scale)
        disp_h = int(img_h * scale)
        
        # Calculate where the image top-left corner is on the canvas
        img_center_x = 190 + self.pan_x
        img_center_y = 140 + self.pan_y
        img_left = img_center_x - disp_w // 2
        img_top = img_center_y - disp_h // 2
        
        # Get crop rectangle coordinates
        x1, y1, x2, y2 = self.crop_coords
        if x1 > x2: x1, x2 = x2, x1
        if y1 > y2: y1, y2 = y2, y1
        
        # Convert canvas coordinates to relative image coordinates (0 to 1)
        rel_x1 = (x1 - img_left) / disp_w
        rel_y1 = (y1 - img_top) / disp_h
        rel_x2 = (x2 - img_left) / disp_w
        rel_y2 = (y2 - img_top) / disp_h
        
        # Clamp to valid range [0, 1]
        rel_x1 = max(0, min(1, rel_x1))
        rel_y1 = max(0, min(1, rel_y1))
        rel_x2 = max(0, min(1, rel_x2))
        rel_y2 = max(0, min(1, rel_y2))
        
        # Convert to actual image pixels
        real_x1 = int(rel_x1 * img_w)
        real_y1 = int(rel_y1 * img_h)
        real_x2 = int(rel_x2 * img_w)
        real_y2 = int(rel_y2 * img_h)
        
        if real_x2 <= real_x1 or real_y2 <= real_y1:
            messagebox.showwarning("Warning", "Invalid crop area - make sure the crop box is over the image")
            return
        
        cropped = self.original_image.crop((real_x1, real_y1, real_x2, real_y2))
        
        # Show crop dimensions in the dialog
        crop_w, crop_h = cropped.size
        choice = messagebox.askyesnocancel("Save Cropped Image",
            f"Crop size: {crop_w} x {crop_h} pixels\n\nYes = Replace original file\nNo = Save as new file\nCancel = Abort")
        if choice is None:
            return
        elif choice:
            cropped.save(self.current_photo_path)
            messagebox.showinfo("Saved", "Original file replaced with cropped image")
        else:
            base, ext = os.path.splitext(self.current_photo_path)
            new_path = base + "_cropped" + ext
            save_path = filedialog.asksaveasfilename(initialfile=os.path.basename(new_path),
                defaultextension=ext, filetypes=[("Image files", "*.jpg *.jpeg *.png *.gif *.bmp")])
            if save_path:
                cropped.save(save_path)
                messagebox.showinfo("Saved", f"Cropped image saved to {save_path}")
        self.cancel_crop()
        self.original_image = Image.open(self.current_photo_path)
        self.upd_pv()
        self.load()

    def cancel_crop(self):
        self.crop_mode = False
        self.crop_btn.config(text="Crop")
        self.crop_frame.pack_forget()
        self.pcv.delete("crop")
        self.pcv.unbind("<Button-1>")
        self.pcv.unbind("<B1-Motion>")
        self.pcv.unbind("<ButtonRelease-1>")
        self.upd_pv()

    def upd_sel(self):
        self.slist.delete(0, tk.END)
        sel=[p for p,v in self.checks.items() if v.get()]
        self.scnt.config(text="Selected: "+str(len(sel)))
        for p in sel: self.slist.insert(tk.END, os.path.basename(p))

    def load(self):
        p=self.path.get()
        if not p or not os.path.isdir(p): return
        for w in self.inner.winfo_children(): w.destroy()
        self.photos, self.thumbs, self.checks = [], {}, {}
        exts={'.jpg','.jpeg','.png','.gif','.bmp'}
        for f in sorted(os.listdir(p)):
            if os.path.splitext(f)[1].lower() in exts: self.photos.append(os.path.join(p,f))
        r,c=0,0
        for ph in self.photos:
            fr=ttk.Frame(self.inner, padding=3)
            fr.grid(row=r, column=c, padx=5, pady=5)
            try:
                im=Image.open(ph); im.thumbnail((100,100))
                tk_im=ImageTk.PhotoImage(im); self.thumbs[ph]=tk_im
                tk.Button(fr, image=tk_im, command=lambda x=ph: self.show(x)).pack()
            except: pass
            ttk.Label(fr, text=os.path.basename(ph)[:15]).pack()
            v=tk.BooleanVar(); self.checks[ph]=v
            ttk.Checkbutton(fr, text="Select", variable=v, command=self.upd_sel).pack()
            c+=1
            if c>5: c,r=0,r+1
        self.status.config(text="Loaded "+str(len(self.photos)))

    def show(self, path):
        if self.crop_mode:
            self.cancel_crop()
        self.zoom_level=1.0; self.pan_x=0; self.pan_y=0
        self.current_photo_path = path
        try:
            self.original_image=Image.open(path)
            w,h=self.original_image.size
            self.lbl_n.config(text="Name: "+os.path.basename(path))
            self.lbl_s.config(text="Size: "+str(w)+"x"+str(h))
            self.lbl_f.config(text="Format: "+str(self.original_image.format))
            self.upd_pv()
        except: pass

    def upd_pv(self):
        if not self.original_image: return
        im=self.original_image.copy(); im.thumbnail((int(360*self.zoom_level),int(360*self.zoom_level)))
        self.preview_img=ImageTk.PhotoImage(im)
        self.pcv.delete("all"); self.pcv.create_image(190+self.pan_x,140+self.pan_y,image=self.preview_img)

    def start_rename(self):
        vb=self.vb_id.get().strip()
        if not vb: messagebox.showwarning("Warning","Enter Boring ID"); return
        self.selected_for_rename=[p for p,v in self.checks.items() if v.get()]
        if not self.selected_for_rename: messagebox.showinfo("Info","No photos selected"); return
        self.output_folder=filedialog.askdirectory(title="Output Folder")
        if not self.output_folder: return
        self.current_rename_index=0
        self.rename_dlg()

    def rename_dlg(self):
        if self.current_rename_index>=len(self.selected_for_rename): messagebox.showinfo("Done","All photos saved!"); return
        ph=self.selected_for_rename[self.current_rename_index]
        d=tk.Toplevel(self.root); d.title("Rename "+str(self.current_rename_index+1)+"/"+str(len(self.selected_for_rename)))
        d.geometry("950x750"); d.grab_set()
        ttk.Label(d, text="Original: "+os.path.basename(ph)).pack(pady=5)
        mf=ttk.Frame(d); mf.pack(fill=tk.BOTH, expand=True, padx=10)
        imgf=ttk.LabelFrame(mf, text="Photo Preview", padding=5)
        imgf.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        cv=tk.Canvas(imgf, bg="#d0d0d0", width=500, height=500)
        cv.pack(fill=tk.BOTH, expand=True)
        d.zm=1.0; d.px=0; d.py=0; d.tki=None; d.oimg=None
        try: d.oimg=Image.open(ph)
        except: cv.create_text(250,250,text="Cannot load")
        def upi():
            if not d.oimg: return
            i=d.oimg.copy(); i.thumbnail((int(480*d.zm),int(480*d.zm))); d.tki=ImageTk.PhotoImage(i)
            cv.delete("all"); cv.create_image(250+d.px,250+d.py,image=d.tki)
        zf=ttk.Frame(imgf); zf.pack(fill=tk.X, pady=3)
        ttk.Button(zf, text="Zoom +", command=lambda: (setattr(d,'zm',min(5,d.zm*1.3)), upi())).pack(side=tk.LEFT, padx=3)
        ttk.Button(zf, text="Zoom -", command=lambda: (setattr(d,'zm',max(0.3,d.zm/1.3)), upi())).pack(side=tk.LEFT, padx=3)
        ttk.Button(zf, text="Reset", command=lambda: (setattr(d,'zm',1.0), setattr(d,'px',0), setattr(d,'py',0), upi())).pack(side=tk.LEFT, padx=3)
        pf=ttk.Frame(imgf); pf.pack(fill=tk.X, pady=3)
        ttk.Button(pf, text="Up", width=5, command=lambda: (setattr(d,'py',d.py-30), upi())).pack(side=tk.LEFT, padx=2)
        ttk.Button(pf, text="Down", width=5, command=lambda: (setattr(d,'py',d.py+30), upi())).pack(side=tk.LEFT, padx=2)
        ttk.Button(pf, text="Left", width=5, command=lambda: (setattr(d,'px',d.px-30), upi())).pack(side=tk.LEFT, padx=2)
        ttk.Button(pf, text="Right", width=5, command=lambda: (setattr(d,'px',d.px+30), upi())).pack(side=tk.LEFT, padx=2)
        upi()
        of=ttk.LabelFrame(mf, text="New Name", padding=10)
        of.pack(side=tk.RIGHT, fill=tk.Y, padx=(10,0))
        ttk.Label(of, text="Boring ID:").pack(anchor=tk.W)
        ttk.Label(of, text=self.vb_id.get(), foreground="blue").pack(anchor=tk.W,pady=(0,10))
        ttk.Label(of, text="Run Number:").pack(anchor=tk.W)
        rf=ttk.Frame(of); rf.pack(fill=tk.X)
        rn=tk.StringVar(value="01")
        rc=ttk.Combobox(rf, textvariable=rn, width=4, state="readonly")
        rc["values"]=[str(i).zfill(2) for i in range(1,21)]; rc.pack(side=tk.LEFT)
        ttk.Label(rf, text=" - ").pack(side=tk.LEFT)
        rs=tk.StringVar(value="")
        sc=ttk.Combobox(rf, textvariable=rs, width=4, state="readonly")
        sc["values"]=["","A","B","C"]; sc.pack(side=tk.LEFT)
        ttk.Label(of, text="Category:").pack(anchor=tk.W,pady=(10,0))
        ct=tk.StringVar(value="Front")
        ttk.Radiobutton(of, text="Front", variable=ct, value="Front").pack(anchor=tk.W)
        ttk.Radiobutton(of, text="Back", variable=ct, value="Back").pack(anchor=tk.W)
        ttk.Radiobutton(of, text="Close-up", variable=ct, value="Close-up").pack(anchor=tk.W)
        ttk.Label(of, text="New Filename:").pack(anchor=tk.W,pady=(15,0))
        nm=tk.StringVar()
        ttk.Label(of, textvariable=nm, foreground="green").pack(anchor=tk.W)
        def upn(*a): nm.set(self.vb_id.get()+"_Run"+rn.get()+("-"+rs.get() if rs.get() else "")+"_"+ct.get()+os.path.splitext(ph)[1].lower())
        rn.trace("w",upn); rs.trace("w",upn); ct.trace("w",upn); upn()
        bf=ttk.Frame(of); bf.pack(pady=20)
        def sav():
            np=os.path.join(self.output_folder,nm.get())
            if os.path.exists(np) and not messagebox.askyesno("Exists","Overwrite?"): return
            shutil.copy2(ph,np); d.destroy(); self.current_rename_index+=1; self.rename_dlg()
        def skp(): d.destroy(); self.current_rename_index+=1; self.rename_dlg()
        ttk.Button(bf, text="Save Next", command=sav).pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="Skip", command=skp).pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="Cancel", command=d.destroy).pack(side=tk.LEFT, padx=5)

root=tk.Tk()
App(root)
root.mainloop()
