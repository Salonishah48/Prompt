"""
wip_parser.py — Column mapping (fixed):
  Col J(9)=W1 | J6(idx5,9)=W1 date label
  Col K(10)=W2 | K6(idx5,10)=W2 date label
  Col L(11)=W3 | L6(idx5,11)=W3 date label
  Col M(12)=W4 | M6(idx5,12)=W4 date label
"""
import warnings
from pathlib import Path
import pandas as pd
warnings.filterwarnings("ignore")

MASTER_FILENAME = "US_Accounts_Team_Details.xls"

SM_MAP = {
    "shah saloni":"Saloni Shah","shaikh firoz rehman":"Firoz Shaikh",
    "jainam  appa":"Jainam Appa","jainam appa":"Jainam Appa",
    "soni pragneshkumar m":"Pragnesh Soni","pragneshkumar soni":"Pragnesh Soni",
    "pragnesh soni":"Pragnesh Soni",
    "anantkumar prahladbhai suthar":"Anant Suthar","anantkumar suthar":"Anant Suthar",
    "nidhi chintankumar pandya":"Nidhi Pandya",
    "sandeepkumar bhailalbhai chavda":"Sandeepkumar Chavda",
    "sandeepkumar chavda":"Sandeepkumar Chavda",
    "modi yug":"Yug Modi",
    "vora parshwa prakashkumar":"Parshwa Vora","parshwa vora":"Parshwa Vora",
    "prajapati piyush dilipbhai":"Piyush Prajapati",
    "viralkumar jayantibhai patel":"Viral Patel",
    "1001 - vivek - shah":"Vivek Shah",
    "soni rajan prakashbhai":"Rajan Soni",
    "sanghavi meet shaileshbhai":"Meet Sanghavi",
    "gadariya rajendra subhashsinh":"Rajendra Gadariya",
    "kulabkar malay mahendra":"Malay Kulabkar",
}
NAME_FIXES = {
    "manthan chauhan":"manthan chuhan","mohdakram shaikh":"mohadakram shaikh",
    "divya karukaparambil":"divya kamaljeet",
    "viralkumar patel":None,"piyush prajapati":None,"parshwa vora":None,
}

def _t1(s):
    s=str(s).lower()
    if "parshwa" in s: return "Parshwa Vora"
    if "piyush"  in s: return "Piyush Prajapati"
    return "Viral Patel"

def _sm(s): return SM_MAP.get(str(s).strip().lower(), str(s).strip())

def _sf(v):
    try:
        f=float(v); return 0.0 if f!=f else round(f,2)
    except: return 0.0

def parse_wip_file(filepath):
    wb=pd.read_excel(filepath, sheet_name=0, header=None)
    if wb.shape[0]<8: return None  # need at least 8 rows
    name   =str(wb.iloc[1,1]).strip() if pd.notna(wb.iloc[1,1]) else ""
    desig  =str(wb.iloc[2,1]).strip() if pd.notna(wb.iloc[2,1]) else ""
    if not name or name=="nan": return None

    # Week dates from row 6 (index 5), cols J-M (9-12)
    wd={}
    for wk,ci in [("w1",9),("w2",10),("w3",11),("w4",12)]:
        try: wd[wk]=pd.Timestamp(wb.iloc[5,ci]).strftime("%Y-%m-%d")
        except: wd[wk]=None

    tasks,nb,in_nb=[],[],False
    for i in range(7,len(wb)):
        row=wb.iloc[i]
        ca=str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        cb=str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        if any(x in ca.lower() for x in ["total","weekly","available"]): continue
        if "non billable" in ca.lower() or "non-billable" in ca.lower():
            in_nb=True; continue
        if in_nb:
            if cb and cb not in ("","nan","_"):
                nb.append({"cat":cb,
                    "w1":_sf(row.iloc[9])  if wb.shape[1]>9  else 0,
                    "w2":_sf(row.iloc[10]) if wb.shape[1]>10 else 0,
                    "w3":_sf(row.iloc[11]) if wb.shape[1]>11 else 0,
                    "w4":_sf(row.iloc[12]) if wb.shape[1]>12 else 0})
            continue
        try: float(ca)
        except: continue
        w1h=_sf(row.iloc[9]) if wb.shape[1]>9 else 0
        w2h=_sf(row.iloc[10]) if wb.shape[1]>10 else 0
        w3h=_sf(row.iloc[11]) if wb.shape[1]>11 else 0
        w4h=_sf(row.iloc[12]) if wb.shape[1]>12 else 0
        if w1h+w2h+w3h+w4h==0: continue
        tasks.append({
            "c":str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else "",
            "p":str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else "",
            "w":str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else "",
            "s":str(row.iloc[7]).strip() if pd.notna(row.iloc[7]) else "",
            "w1":w1h,"w2":w2h,"w3":w3h,"w4":w4h,"h":round(w1h+w2h+w3h+w4h,2)})

    def tb(k):  return round(sum(t[k] for t in tasks),2)
    def tnb(k): return round(sum(n[k] for n in nb),2)
    return {"name":name,"desig":desig,"j6":wd.get("w1"),"wd":wd,"tasks":tasks,"nb":nb,
        "tb1":tb("w1"),"tnb1":tnb("w1"),"w1":round(tb("w1")+tnb("w1"),2),
        "tb2":tb("w2"),"tnb2":tnb("w2"),"w2":round(tb("w2")+tnb("w2"),2),
        "tb3":tb("w3"),"tnb3":tnb("w3"),"w3":round(tb("w3")+tnb("w3"),2),
        "tb4":tb("w4"),"tnb4":tnb("w4"),"w4":round(tb("w4")+tnb("w4"),2),
        "tb":round(tb("w1")+tb("w2")+tb("w3")+tb("w4"),2),
        "tnb":round(tnb("w1")+tnb("w2")+tnb("w3")+tnb("w4"),2),
        "th":round(tb("w1")+tb("w2")+tb("w3")+tb("w4")+tnb("w1")+tnb("w2")+tnb("w3")+tnb("w4"),2)}

def parse_all(raw_dir, master_path, log_fn=print):
    raw_dir=Path(raw_dir); master_name=Path(master_path).name.lower()
    wip={};common_wd=None
    files=sorted(f for f in raw_dir.iterdir()
                 if f.suffix.lower() in (".xlsx",".xls") and f.name.lower()!=master_name)
    log_fn(f"Found {len(files)} WIP file(s)","info")
    for fp in files:
        try:
            r=parse_wip_file(str(fp))
            if r:
                wip[r["name"].lower()]=r
                if common_wd is None and r.get("wd"): common_wd=r["wd"]
                log_fn(f"  \u2713 {fp.name} \u2192 {r['name']} (W1={r['w1']} W2={r['w2']} W3={r['w3']} W4={r['w4']})","ok")
            else: log_fn(f"  \u26a0 {fp.name} \u2014 skipped","warn")
        except Exception as e: log_fn(f"  \u2717 {fp.name} \u2014 {e}","err")
    if common_wd is None: common_wd={"w1":None,"w2":None,"w3":None,"w4":None}
    log_fn(f"Loading master: {Path(master_path).name}","info")
    try: master=pd.read_excel(str(master_path))
    except Exception as e: raise RuntimeError(f"Cannot read master: {e}")
    emps=[]
    for _,row in master.iterrows():
        n=str(row["Employe Name"]).strip(); eid=str(row["Employee ID"]).strip()
        t1=_t1(str(row["T1"])); sm=_sm(str(row["Manager"]))
        if "piyush" in n.lower(): continue
        if n=="Parshwa Vora": t1="Vivek Shah"
        nk=n.lower(); fix=NAME_FIXES.get(nk); lookup=fix if fix is not None else nk
        w=wip.get(lookup) if lookup else None
        emps.append({"n":n,"i":eid,"d":w["desig"] if w else "","t1":t1,"sm":sm,"w":w is not None,
            "j6":w["j6"] if w else None,
            "tb1":w["tb1"] if w else 0,"tnb1":w["tnb1"] if w else 0,"w1":w["w1"] if w else 0,
            "tb2":w["tb2"] if w else 0,"tnb2":w["tnb2"] if w else 0,"w2":w["w2"] if w else 0,
            "tb3":w["tb3"] if w else 0,"tnb3":w["tnb3"] if w else 0,"w3":w["w3"] if w else 0,
            "tb4":w["tb4"] if w else 0,"tnb4":w["tnb4"] if w else 0,"w4":w["w4"] if w else 0,
            "tb":w["tb"] if w else 0,"tnb":w["tnb"] if w else 0,"th":w["th"] if w else 0,
            "tasks":[{"c":t["c"],"p":t["p"],"w":t["w"],"s":t["s"],"w1":t["w1"],"w2":t["w2"],"w3":t["w3"],"w4":t["w4"],"h":t["h"]} for t in (w["tasks"] if w else [])],
            "nb":[{"cat":nb["cat"],"w1":nb["w1"],"w2":nb["w2"],"w3":nb["w3"],"w4":nb["w4"]} for nb in (w["nb"] if w else [])]})
    stats={"total":len(emps),"with_wip":sum(1 for e in emps if e["w"]),"no_wip":sum(1 for e in emps if not e["w"])}
    log_fn(f"Done: {stats['total']} employees, {stats['with_wip']} WIP, {stats['no_wip']} missing","ok")
    return emps, common_wd, stats
