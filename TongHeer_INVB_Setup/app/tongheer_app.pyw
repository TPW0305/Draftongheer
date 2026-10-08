"""Tong Heer INVB Converter - หน้าจอโปรแกรม (ดับเบิลคลิกไฟล์นี้เพื่อเปิด)"""
import os
import sys
import traceback
from tkinter import messagebox, filedialog

import customtkinter as ctk

import tongheer_invb as core

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception:  # ไม่มี tkinterdnd2 ก็ยังใช้ปุ่มเลือกไฟล์ได้
    TkinterDnD = None

FONT = "Leelawadee UI"
BLUE, BLUE_H = "#1f6feb", "#1a5fd0"
GREEN, GREEN_H = "#1a7f37", "#166c2f"
BG, CARD, LINE = "#f3f5f8", "#ffffff", "#d8dee6"
MISSING = "#fff3bf"
OK_BG, OK_FG = "#dafbe1", "#116329"
WARN_BG, WARN_FG = "#fff1c2", "#7a4b00"
MUTED = "#57606a"
# (หัวคอลัมน์, ความกว้าง px, ชิด)
COLS = [("#", 34, "center"), ("ชื่อสินค้า / MODEL", 300, "w"), ("Net (KGM)", 90, "e"), ("Gross (KGM)", 90, "e"),
        ("Amount (USD)", 100, "e"), ("PARTNO", 220, "w"), ("SUBGROUP", 180, "w")]


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("light")
        self.title("Tong Heer - แปลง Export Summary Report เป็น INVB (ECS)")
        self.geometry("1180x800")
        self.minsize(1100, 680)
        self.configure(fg_color=BG)
        ico = os.path.join(core.APP_DIR, "app.ico")
        if os.path.exists(ico):  # customtkinter ตั้งไอคอนของตัวเองหลังเปิด 200ms จึงตั้งทับทีหลัง
            self.after(300, lambda: self.iconbitmap(ico))
        self.cfg = core.load_config()
        self.rep = None
        self.rows = []       # [(item, partno_var, sub_var, combo, sub_entry, state)]
        self.last_out = None
        self.f = lambda size=14, weight="normal": ctk.CTkFont(FONT, size, weight)
        self._build()
        if TkinterDnD:
            try:
                self.TkdndVersion = TkinterDnD._require(self)
                self.drop_target_register(DND_FILES)
                self.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass

    # ---------- สร้างหน้าจอ ----------
    def _card(self, parent, **grid):
        card = ctk.CTkFrame(parent, fg_color=CARD, corner_radius=10, border_width=1, border_color=LINE)
        card.grid(sticky="nsew", padx=16, **grid)
        return card

    def _build(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # ขั้นที่ 1: เลือกไฟล์
        top = self._card(self, row=0, column=0, pady=(14, 6))
        top.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(top, text="①  เลือกไฟล์รายงาน", font=self.f(15, "bold"), height=44, width=210, fg_color=BLUE,
                      hover_color=BLUE_H, command=self._pick).grid(row=0, column=0, rowspan=2, padx=14, pady=14)
        self.lbl_file = ctk.CTkLabel(top, text="ยังไม่ได้เลือกไฟล์", font=self.f(15, "bold"), anchor="w")
        self.lbl_file.grid(row=0, column=1, sticky="sw", pady=(12, 0))
        self.lbl_info = ctk.CTkLabel(top, text="กดปุ่มทางซ้าย หรือลากไฟล์ Export Summary Report (.xlsx) มาวางในหน้าต่างนี้",
                                     font=self.f(13), text_color=MUTED, anchor="w")
        self.lbl_info.grid(row=1, column=1, sticky="nw", pady=(0, 12))

        # ขั้นที่ 2: ข้อมูลหัวอินวอย
        head = self._card(self, row=1, column=0, pady=6)
        ctk.CTkLabel(head, text="②  ตรวจข้อมูล  (ช่องสีเหลือง = ต้องกรอก)", font=self.f(14, "bold")).grid(
            row=0, column=0, columnspan=5, sticky="w", padx=14, pady=(10, 2))
        self.v_inv, self.v_date, self.v_mark, self.v_cons, self.v_pkg = (ctk.StringVar() for _ in range(5))
        fields = [("เลขอินวอย (INV_NO)", self.v_inv, 250), ("วันที่ (วว/ดด/ปปปป ค.ศ.)", self.v_date, 160),
                  ("Marking (MARK1)", self.v_mark, 330), ("CONSIGNEE (รหัสใน ECS)", self.v_cons, 180),
                  ("จำนวนพาเลท (NUMPKG)", self.v_pkg, 140)]
        self.entries = []
        for c, (label, var, w) in enumerate(fields):
            ctk.CTkLabel(head, text=label, font=self.f(12), text_color=MUTED).grid(row=1, column=c, sticky="w", padx=(14, 0))
            e = ctk.CTkEntry(head, textvariable=var, width=w, height=34, font=self.f(14), border_color=LINE)
            e.grid(row=2, column=c, sticky="w", padx=(14, 0), pady=(0, 12))
            var.trace_add("write", lambda *_: self._paint())
            self.entries.append((e, var))

        # รายการสินค้า
        mid = self._card(self, row=2, column=0, pady=6)
        mid.grid_columnconfigure(0, weight=1)
        mid.grid_rowconfigure(1, weight=1)
        hdr = ctk.CTkFrame(mid, fg_color="#eef1f5", corner_radius=6)
        hdr.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 0))
        for c, (text, w, anchor) in enumerate(COLS):
            ctk.CTkLabel(hdr, text=text, width=w, anchor=anchor, font=self.f(12, "bold"), text_color=MUTED).grid(
                row=0, column=c, padx=6, pady=5)
        self.body = ctk.CTkScrollableFrame(mid, fg_color=CARD)
        self.body.grid(row=1, column=0, sticky="nsew", padx=4, pady=(0, 4))
        self.lbl_empty = ctk.CTkLabel(self.body, text="รายการสินค้าจะแสดงที่นี่หลังเลือกไฟล์", font=self.f(14), text_color=MUTED)
        self.lbl_empty.grid(row=0, column=0, padx=20, pady=40)
        self.chips = ctk.CTkFrame(mid, fg_color="transparent")
        self.chips.grid(row=2, column=0, sticky="w", padx=10, pady=(0, 10))

        # จุดที่ต้องตรวจ
        self.txt_warn = ctk.CTkTextbox(self, height=108, font=self.f(13), fg_color=CARD, border_width=1, border_color=LINE,
                                       corner_radius=10, wrap="word")
        self.txt_warn.grid(row=3, column=0, sticky="ew", padx=16, pady=6)
        self._set_warn([])

        # ขั้นที่ 3: สร้างไฟล์
        bot = ctk.CTkFrame(self, fg_color="transparent")
        bot.grid(row=4, column=0, sticky="ew", padx=16, pady=(6, 14))
        bot.grid_columnconfigure(1, weight=1)
        self.v_price = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(bot, text="เพิ่มคอลัมน์ PRICE (ใช้ทดสอบปัญหาราคาต่อหน่วยใน ECS)", variable=self.v_price,
                        font=self.f(13), checkbox_width=20, checkbox_height=20).grid(row=0, column=0, sticky="w")
        self.lbl_status = ctk.CTkLabel(bot, text="", font=self.f(13, "bold"), text_color=OK_FG, anchor="e")
        self.lbl_status.grid(row=0, column=1, sticky="e", padx=10)
        self.btn_folder = ctk.CTkButton(bot, text="เปิดโฟลเดอร์", font=self.f(13), height=44, width=120, fg_color="#e7ebf0",
                                        hover_color="#d8dee6", text_color="#24292f", command=self._open_folder)
        self.btn_open = ctk.CTkButton(bot, text="เปิดไฟล์", font=self.f(13), height=44, width=100, fg_color="#e7ebf0",
                                      hover_color="#d8dee6", text_color="#24292f", command=self._open_file)
        self.btn_gen = ctk.CTkButton(bot, text="③  สร้างไฟล์ INVB", font=self.f(16, "bold"), height=44, width=230,
                                     fg_color=GREEN, hover_color=GREEN_H, state="disabled", command=self.generate)
        self.btn_gen.grid(row=0, column=4, sticky="e", padx=(8, 0))

    # ---------- โหลดรายงาน ----------
    def _pick(self):
        start = os.path.dirname(self.rep.path) if self.rep else core.APP_DIR
        path = filedialog.askopenfilename(title="เลือกไฟล์ Export Summary Report", initialdir=start,
                                          filetypes=[("Excel", "*.xlsx")])
        if path:
            self.load(path)

    def _on_drop(self, event):
        files = self.tk.splitlist(event.data)
        if files:
            self.load(files[0])

    def load(self, path):
        try:
            self.cfg = core.load_config()
            rep = core.resolve(core.parse_report(os.path.abspath(path)), self.cfg)
        except Exception as e:
            messagebox.showerror("อ่านไฟล์ไม่ได้", str(e))
            return
        self.rep = rep
        self.last_out = None
        self.lbl_status.configure(text="")
        self.btn_open.grid_forget()
        self.btn_folder.grid_forget()
        self.lbl_file.configure(text=os.path.basename(rep.path))
        self.lbl_info.configure(text=f"ลูกค้า: {rep.customer or '-'}     สาขา (EXP_BRANCH): {rep.branch or '-'}     "
                                     f"BOI: {rep.boi or '-'}     ETD: {core.date_text(rep.etd) or '-'}")
        self.v_inv.set(rep.inv_no)
        self.v_date.set(core.date_text(rep.date))
        self.v_mark.set(rep.marking)
        self.v_cons.set(rep.consignee)
        self.v_pkg.set("" if rep.pallets is None else str(rep.pallets))

        for w in self.body.winfo_children():
            w.destroy()
        self.rows = []
        known = core.known_partnos(self.cfg, rep.prefix)
        for n, it in enumerate(rep.items, 1):
            r = n - 1
            ctk.CTkLabel(self.body, text=str(n), width=COLS[0][1], font=self.f(13)).grid(row=r, column=0, padx=6, pady=3)
            ctk.CTkLabel(self.body, text=f"{it.name}\n{it.model_line}", width=COLS[1][1], anchor="w", justify="left",
                         font=self.f(12), wraplength=COLS[1][1]).grid(row=r, column=1, padx=6, pady=3)
            for c, v in ((2, it.net), (3, it.gross), (4, it.amount)):
                ctk.CTkLabel(self.body, text=core.fmt2(v), width=COLS[c][1], anchor="e", font=self.f(13)).grid(
                    row=r, column=c, padx=6, pady=3)
            pv, sv = ctk.StringVar(value=it.partno), ctk.StringVar(value=it.subgroup)
            choices = it.candidates + [k for k in known if k not in it.candidates]
            combo = ctk.CTkComboBox(self.body, values=choices, variable=pv, width=COLS[5][1], height=32, font=self.f(13),
                                    dropdown_font=self.f(13), border_color=LINE, button_color=LINE, button_hover_color="#c3ccd7")
            combo.grid(row=r, column=5, padx=6, pady=3)
            sub = ctk.CTkEntry(self.body, textvariable=sv, width=COLS[6][1], height=32, font=self.f(13), border_color=LINE)
            sub.grid(row=r, column=6, padx=6, pady=3)
            state = {"auto": it.subgroup}
            pv.trace_add("write", lambda *_, pv=pv, sv=sv, state=state: self._partno_changed(pv, sv, state))
            self.rows.append((it, pv, sv, combo, sub))
        self._chips()
        self._set_warn(self._warnings())
        self.btn_gen.configure(state="normal")
        self._paint()

    def _partno_changed(self, pv, sv, state):
        # SUBGROUP ตาม PARTNO อัตโนมัติ เว้นแต่ผู้ใช้พิมพ์ SUBGROUP เองแล้ว
        new = core.default_subgroup(self.cfg, self.rep.prefix, pv.get())
        if sv.get() in ("", state["auto"]):
            sv.set(new)
        state["auto"] = new
        self._paint()

    def _chips(self):
        for w in self.chips.winfo_children():
            w.destroy()
        ctk.CTkLabel(self.chips, text=f"ยอดรวม {len(self.rep.items)} รายการ เทียบรายงาน:", font=self.f(13, "bold")).pack(
            side="left", padx=(0, 8))
        for label, s, ref, diff in self.rep.total_checks():
            if ref is None:
                text, bg, fg = f"{label} {core.fmt2(s)}  (ไม่พบยอดในรายงาน)", WARN_BG, WARN_FG
            elif abs(diff) < 0.005:
                text, bg, fg = f"✓ {label} {core.fmt2(s)}", OK_BG, OK_FG
            else:
                text, bg, fg = f"⚠ {label} {core.fmt2(s)}  รายงาน {core.fmt2(ref)}  (ต่าง {diff:+.2f})", WARN_BG, WARN_FG
            ctk.CTkLabel(self.chips, text=text, font=self.f(13, "bold"), fg_color=bg, text_color=fg, corner_radius=6,
                         height=28).pack(side="left", padx=4, ipadx=6)

    def _warnings(self):
        w = list(self.rep.warnings)
        for n, (it, pv, *_rest) in enumerate(self.rows, 1):
            if not pv.get().strip():
                w.append(f"รายการ {n}: ยังไม่รู้ PARTNO ของ \"{it.name} {it.model_line}\" - เลือกจากรายการหรือพิมพ์ให้ตรง "
                         "DATA_TONGHEER (โปรแกรมจะจำไว้)")
        return w + core.subgroup_warnings(self.rep, self.cfg)

    def _set_warn(self, items):
        self.txt_warn.configure(state="normal")
        self.txt_warn.delete("1.0", "end")
        if not self.rep:
            self.txt_warn.insert("end", "จุดที่ต้องตรวจจะแสดงที่นี่")
        elif not items:
            self.txt_warn.insert("end", "✓ ไม่มีจุดที่ต้องตรวจเพิ่ม")
        else:
            self.txt_warn.insert("end", f"จุดที่ต้องตรวจ ({len(items)})\n" + "\n".join("•  " + i for i in items))
        self.txt_warn.configure(state="disabled")

    def _paint(self):
        for e, var in self.entries:
            e.configure(fg_color=MISSING if self.rep and not var.get().strip() else CARD)
        for _it, pv, sv, combo, sub in self.rows:
            miss = not pv.get().strip()
            combo.configure(fg_color=MISSING if miss else CARD)
            sub.configure(fg_color=MISSING if miss and not sv.get().strip() else CARD)

    # ---------- สร้างไฟล์ ----------
    def generate(self, out_override=None):
        rep = self.rep
        if not rep:
            return None
        err = []
        inv_no, mark, cons = self.v_inv.get().strip(), self.v_mark.get().strip(), self.v_cons.get().strip()
        date = core.parse_date_text(self.v_date.get())
        if not inv_no:
            err.append("ยังไม่ได้กรอกเลขอินวอย (INV_NO)")
        if not date:
            err.append("วันที่ต้องเป็นรูปแบบ วว/ดด/ปปปป (ปี ค.ศ.) เช่น 07/10/2026")
        if not mark:
            err.append("ยังไม่ได้กรอก Marking (MARK1)")
        if not cons:
            err.append("ยังไม่ได้กรอก CONSIGNEE (รหัสตาม ECS)")
        if not self.v_pkg.get().strip().isdigit():
            err.append("จำนวนพาเลท (NUMPKG) ต้องเป็นจำนวนเต็ม")
        if not rep.branch:
            err.append(f'ไม่รู้จัก prefix อินวอย "{rep.prefix}" (ต้องเป็น 1THA / 2THA / 3THA)')
        for n, (it, pv, sv, *_w) in enumerate(self.rows, 1):
            it.partno = pv.get()
            it.subgroup = sv.get() if sv.get().strip() else it.partno
            if not it.partno.strip():
                err.append(f"รายการ {n}: ยังไม่ได้เลือก PARTNO")
        if err:
            self._paint()
            if not out_override:
                messagebox.showwarning("ข้อมูลยังไม่ครบ", "\n".join(err))
            return None
        out = out_override or core.out_path_for(rep, inv_no)
        if not out_override and os.path.exists(out):
            if not messagebox.askyesno("ยืนยัน", f"มีไฟล์ {os.path.basename(out)} อยู่แล้ว\nต้องการเขียนทับหรือไม่?"):
                return None
        try:
            core.write_invb(rep, self.cfg, out, inv_no, date, mark, cons, int(self.v_pkg.get()), self.v_price.get())
        except PermissionError:
            messagebox.showerror("เขียนไฟล์ไม่ได้", f"ไฟล์ {os.path.basename(out)} เปิดค้างอยู่ใน Excel\nปิดไฟล์ก่อนแล้วกดสร้างอีกครั้ง")
            return None
        except Exception as e:
            messagebox.showerror("เขียนไฟล์ไม่ได้", str(e))
            return None
        if not out_override:
            core.learn(rep, self.cfg, mark, cons)
        self.last_out = out
        self._set_warn(self._warnings())
        self.lbl_status.configure(text=f"✓ สร้างแล้ว: {os.path.basename(out)}")
        self.btn_open.grid(row=0, column=2, padx=(8, 0))
        self.btn_folder.grid(row=0, column=3, padx=(8, 0))
        return out

    def _open_file(self):
        if self.last_out:
            os.startfile(self.last_out)

    def _open_folder(self):
        if self.last_out:
            os.startfile(os.path.dirname(self.last_out))

    def report_callback_exception(self, exc, val, tb):
        messagebox.showerror("ผิดพลาด", "".join(traceback.format_exception_only(exc, val)))


def selftest(app, report, out, png):
    """ทดสอบหน้าจอโดยไม่ต้องคลิก: โหลดไฟล์ สร้างผลลัพธ์ เก็บภาพหน้าจอ"""
    from PIL import ImageGrab
    app.attributes("-topmost", True)
    app.load(report)
    app.update()
    result = app.generate(out) if out != "-" else None
    for _ in range(20):
        app.update()
        app.after(30)
    x, y = app.winfo_rootx(), app.winfo_rooty()
    ImageGrab.grab(bbox=(x, y, x + app.winfo_width(), y + app.winfo_height())).save(png)
    print("selftest out:", result)
    app.destroy()


if __name__ == "__main__":
    app = App()
    if len(sys.argv) >= 5 and sys.argv[1] == "--selftest":
        app.after(400, lambda: selftest(app, *sys.argv[2:5]))
    elif len(sys.argv) > 1:
        app.after(200, lambda: app.load(sys.argv[1]))
    app.mainloop()
