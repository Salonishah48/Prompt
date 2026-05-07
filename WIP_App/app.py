"""WIP Dashboard Generator — US Accounts Team"""
import json, os, subprocess, sys, threading, webbrowser
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

BASE_DIR=Path(__file__).parent.resolve()
sys.path.insert(0,str(BASE_DIR))
from wip_parser import parse_all, MASTER_FILENAME

OUTPUT_DIR=BASE_DIR/"output"; RAW_DIR=BASE_DIR/"Raw Data"; TEMPLATE=BASE_DIR/"dashboard_template.html"
BG="#f0f4ff"; CARD="#ffffff"; NAVY="#1a3280"; BLUE="#4f6ef7"; GREEN="#16a34a"
AMBER="#f59e0b"; RED="#ef4444"; TEXT="#1e293b"; MUTED="#64748b"; BORDER="#e2e8f0"

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("WIP Dashboard Generator  —  US Accounts Team")
        self.geometry("860x660"); self.minsize(740,560)
        self.configure(bg=BG); self.resizable(True,True)
        OUTPUT_DIR.mkdir(exist_ok=True); RAW_DIR.mkdir(exist_ok=True)
        self._last=None; self._build(); self._refresh()

    def _build(self):
        # Header
        h=tk.Frame(self,bg=NAVY,height=58); h.pack(fill="x"); h.pack_propagate(False)
        tk.Label(h,text="  WIP",bg=NAVY,fg="#fff",font=("Segoe UI",18,"bold")).pack(side="left",padx=(18,4))
        tk.Label(h,text="Dashboard Generator",bg=NAVY,fg="#c7d4ff",font=("Segoe UI",12)).pack(side="left")
        tk.Label(h,text="US Accounts Team  ",bg=NAVY,fg="#8fa8e8",font=("Segoe UI",10)).pack(side="right")
        # Body
        body=tk.Frame(self,bg=BG); body.pack(fill="both",expand=True,padx=20,pady=16)
        left=tk.Frame(body,bg=BG); left.pack(side="left",fill="both",expand=True,padx=(0,12))
        right=tk.Frame(body,bg=BG,width=255); right.pack(side="right",fill="y"); right.pack_propagate(False)
        self._paths_card(left); self._files_card(left); self._log_card(left); self._right(right)

    def _card(self,p,title):
        f=tk.Frame(p,bg=CARD,highlightthickness=1,highlightbackground=BORDER); f.pack(fill="x",pady=(0,10))
        tk.Label(f,text=title,bg=CARD,fg=NAVY,font=("Segoe UI",10,"bold")).pack(anchor="w",padx=14,pady=(10,5))
        tk.Frame(f,bg=BORDER,height=1).pack(fill="x",padx=14); return f

    def _paths_card(self,p):
        c=self._card(p,"📁  Folder Setup"); f=tk.Frame(c,bg=CARD); f.pack(fill="x",padx=14,pady=10)
        self.raw_var=tk.StringVar(value=str(RAW_DIR))
        self.out_var=tk.StringVar(value=str(OUTPUT_DIR))
        for r,(lbl,var,cmd) in enumerate([("Raw Data Folder:",self.raw_var,self._braw),("Output Folder:",self.out_var,self._bout)]):
            tk.Label(f,text=lbl,bg=CARD,fg=MUTED,font=("Segoe UI",9)).grid(row=r,column=0,sticky="w",pady=3)
            tk.Entry(f,textvariable=var,width=44,font=("Segoe UI",9),relief="flat",bg="#f8fafc",fg=TEXT).grid(row=r,column=1,padx=6,sticky="ew")
            tk.Button(f,text="Browse",command=cmd,bg=BLUE,fg="white",relief="flat",cursor="hand2",font=("Segoe UI",9),padx=9).grid(row=r,column=2)
        f.columnconfigure(1,weight=1)
        tk.Label(c,text=f'  ⚠  Master file must be named  "{MASTER_FILENAME}"  inside Raw Data',bg=CARD,fg=AMBER,font=("Segoe UI",8,"italic")).pack(anchor="w",padx=14,pady=(0,8))

    def _files_card(self,p):
        c=self._card(p,"📋  Detected Files"); row=tk.Frame(c,bg=CARD); row.pack(fill="x",padx=14,pady=(8,6))
        self.lbl_wip   =self._bdg(row,"0","WIP Files",BLUE)
        self.lbl_master=self._bdg(row,"—","Master File",GREEN)
        self.lbl_skip  =self._bdg(row,"0","Unmatched",AMBER)
        tf=tk.Frame(c,bg=CARD); tf.pack(fill="x",padx=14,pady=(0,4))
        self.tree=ttk.Treeview(tf,columns=("File","Type","Status"),show="headings",height=5)
        for col,w in [("File",290),("Type",80),("Status",100)]:
            self.tree.heading(col,text=col); self.tree.column(col,width=w,minwidth=60)
        st=ttk.Style(); st.configure("Treeview",rowheight=22,font=("Segoe UI",9))
        st.configure("Treeview.Heading",font=("Segoe UI",9,"bold"))
        vsb=ttk.Scrollbar(tf,orient="vertical",command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set); self.tree.pack(side="left",fill="x",expand=True); vsb.pack(side="right",fill="y")
        tk.Button(c,text="🔄  Refresh",command=self._refresh,bg=CARD,fg=NAVY,relief="flat",cursor="hand2",font=("Segoe UI",9),padx=8,pady=3).pack(anchor="e",padx=14,pady=(0,6))

    def _bdg(self,p,val,label,color):
        f=tk.Frame(p,bg=color,padx=12,pady=6); f.pack(side="left",padx=(0,8))
        v=tk.Label(f,text=val,bg=color,fg="white",font=("Segoe UI",15,"bold")); v.pack()
        tk.Label(f,text=label,bg=color,fg="white",font=("Segoe UI",8)).pack(); return v

    def _log_card(self,p):
        c=self._card(p,"📜  Generation Log")
        self.log=scrolledtext.ScrolledText(c,height=7,font=("Consolas",9),bg="#0f172a",fg="#94a3b8",insertbackground="white",relief="flat",state="disabled")
        self.log.pack(fill="x",padx=14,pady=(6,12))
        for tag,fg in [("ok","#22c55e"),("warn","#f59e0b"),("err","#ef4444"),("info","#94a3b8"),("head","#a5b4fc")]:
            self.log.tag_config(tag,foreground=fg)

    def _right(self,p):
        self.gen_btn=tk.Button(p,text="⚡  Generate Dashboard",command=self._generate,bg=NAVY,fg="white",activebackground=BLUE,font=("Segoe UI",12,"bold"),relief="flat",cursor="hand2",pady=14)
        self.gen_btn.pack(fill="x",pady=(0,10))
        self.open_btn=tk.Button(p,text="🌐  Open Last Dashboard",command=self._open,bg=GREEN,fg="white",font=("Segoe UI",10),relief="flat",cursor="hand2",pady=9,state="disabled")
        self.open_btn.pack(fill="x",pady=(0,8))
        for lbl,cmd in [("📂  Open Output Folder",self._oout),("📂  Open Raw Data Folder",self._oraw)]:
            tk.Button(p,text=lbl,command=cmd,bg=CARD,fg=NAVY,relief="flat",cursor="hand2",font=("Segoe UI",10),pady=7,highlightthickness=1,highlightbackground=BORDER).pack(fill="x",pady=(0,7))
        sf=tk.Frame(p,bg=CARD,highlightthickness=1,highlightbackground=BORDER); sf.pack(fill="x",pady=(6,0))
        tk.Label(sf,text="Status",bg=CARD,fg=MUTED,font=("Segoe UI",9,"bold")).pack(anchor="w",padx=12,pady=(8,2))
        self.slbl=tk.Label(sf,text="Ready",bg=CARD,fg=GREEN,font=("Segoe UI",10,"bold"),wraplength=210); self.slbl.pack(anchor="w",padx=12,pady=(0,4))
        self.prog=ttk.Progressbar(sf,mode="indeterminate"); self.prog.pack(fill="x",padx=12,pady=(0,8))
        tk.Label(sf,text="Last generated:",bg=CARD,fg=MUTED,font=("Segoe UI",8)).pack(anchor="w",padx=12)
        self.tslbl=tk.Label(sf,text="—",bg=CARD,fg=TEXT,font=("Segoe UI",8)); self.tslbl.pack(anchor="w",padx=12,pady=(0,10))

    def _lg(self,msg,tag="info"):
        self.log.configure(state="normal")
        self.log.insert("end",f"[{datetime.now().strftime('%H:%M:%S')}]  {msg}\n",tag)
        self.log.see("end"); self.log.configure(state="disabled")

    def _ss(self,t,c=GREEN): self.slbl.configure(text=t,fg=c)
    def _braw(self):
        d=filedialog.askdirectory(initialdir=self.raw_var.get())
        if d: self.raw_var.set(d); self._refresh()
    def _bout(self):
        d=filedialog.askdirectory(initialdir=self.out_var.get())
        if d: self.out_var.set(d)
    def _open(self):
        if self._last and os.path.exists(self._last): webbrowser.open(f"file:///{self._last}")
        else: messagebox.showinfo("No output","No dashboard generated yet.")
    def _ofolder(self,path):
        path=str(path)
        if sys.platform=="win32": os.startfile(path)
        elif sys.platform=="darwin": subprocess.Popen(["open",path])
        else: subprocess.Popen(["xdg-open",path])
    def _oout(self): self._ofolder(self.out_var.get())
    def _oraw(self): self._ofolder(self.raw_var.get())

    def _refresh(self):
        raw=Path(self.raw_var.get())
        for r in self.tree.get_children(): self.tree.delete(r)
        wip_cnt=0; master_ok=False
        if not raw.exists(): self._lg(f"Folder not found: {raw}","warn"); return
        for f in sorted(raw.iterdir()):
            if f.suffix.lower() not in (".xlsx",".xls"): continue
            if f.name.lower()==MASTER_FILENAME.lower():
                self.tree.insert("","end",values=(f.name,"Master","✅ Found"),tags=("m",))
                master_ok=True
            else:
                self.tree.insert("","end",values=(f.name,"WIP","✅ Ready"),tags=("w",))
                wip_cnt+=1
        self.tree.tag_configure("m",background="#f0fdf4")
        self.lbl_wip.configure(text=str(wip_cnt))
        self.lbl_master.configure(text="✅" if master_ok else "❌")
        self._lg(f"Scanned '{raw.name}': {wip_cnt} WIP files, master={'found' if master_ok else 'MISSING'}","ok" if master_ok else "warn")

    def _generate(self):
        raw=Path(self.raw_var.get()); out=Path(self.out_var.get()); master=raw/MASTER_FILENAME
        if not raw.exists(): messagebox.showerror("Error",f"Raw Data folder not found:\n{raw}"); return
        if not master.exists(): messagebox.showerror("Error",f'Master file "{MASTER_FILENAME}" not found in:\n{raw}'); return
        if not TEMPLATE.exists(): messagebox.showerror("Error",f"dashboard_template.html not found:\n{TEMPLATE}"); return
        self.gen_btn.configure(state="disabled"); self.prog.start(10)
        self._ss("Generating…",AMBER); self._lg("━━━  Starting generation  ━━━","head")
        threading.Thread(target=self._worker,args=(raw,out,master),daemon=True).start()

    def _worker(self,raw,out,master):
        try:
            employees,week_dates,stats=parse_all(raw,master,log_fn=self._lg)
            self._lg(f"Parsed {stats['total']} employees — {stats['with_wip']} WIP, {stats['no_wip']} missing","ok")
            self._lg(f"Week dates from files: {week_dates}","info")
            ts=datetime.now().strftime("%Y%m%d_%H%M%S"); out_file=out/f"WIP_Dashboard_{ts}.html"
            self._lg("Injecting data into template…","info")
            with open(TEMPLATE,encoding="utf-8") as f: tmpl=f.read()
            html=tmpl.replace("%%WEEK_DATA%%",json.dumps(week_dates)).replace("%%EMP_DATA%%",json.dumps(employees))
            out.mkdir(exist_ok=True)
            with open(out_file,"w",encoding="utf-8") as f: f.write(html)
            self._lg(f"Saved: {out_file.name}  ({out_file.stat().st_size//1024} KB)","ok")
            self._lg("━━━  Done!  ━━━","head")
            self._last=str(out_file); self.after(0,self._done,str(out_file),stats)
        except Exception as e:
            import traceback; self._lg(f"ERROR: {e}","err"); self._lg(traceback.format_exc(),"err")
            self.after(0,self._fail,str(e))

    def _done(self,path,stats):
        self.prog.stop(); self.gen_btn.configure(state="normal"); self.open_btn.configure(state="normal")
        self._ss(f"✅ Done!\n{stats['with_wip']} employees · {stats['no_wip']} missing WIP",GREEN)
        self.tslbl.configure(text=datetime.now().strftime("%d %b %Y   %H:%M:%S"))
        if messagebox.askyesno("Dashboard Ready",f"Generated successfully!\n{stats['with_wip']} employees · {stats['no_wip']} missing WIP\n\nOpen in browser now?"):
            webbrowser.open(f"file:///{path}")

    def _fail(self,msg):
        self.prog.stop(); self.gen_btn.configure(state="normal")
        self._ss("❌ Error occurred",RED); messagebox.showerror("Generation Failed",msg)

if __name__=="__main__":
    App().mainloop()
