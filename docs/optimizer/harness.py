"""Phase-0 evaluation harness: reconstruct a historical day and compare the
CURRENT bot allocation vs the MANUAL (ground truth) for one user."""
import sys, io, re, math
import pandas as pd, numpy as np
sys.path.insert(0,".")
import bot

HIST=sys.argv[1]; MASTER=sys.argv[2]; USER=sys.argv[3]; NDAYS=int(sys.argv[4])

def pfx(r):
    m=re.match(r'\s*([A-Z]{1,3}\d+[A-Z]?)',str(r)); return m.group(1) if m else ""

h=pd.read_excel(HIST); h.columns=[str(c).strip().upper() for c in h.columns]
h["LICENSE"]=h["LICENSE"].astype(str).str.strip().str.upper()
h=h[~h["LICENSE"].isin(["","NAN","NONE"])].copy()
h["DATE"]=pd.to_datetime(h["DATE"],errors="coerce"); h=h[h["DATE"].notna()]
h["W"]=pd.to_numeric(h["GROSS WEIGHT"],errors="coerce").fillna(0)/1000.0

m=pd.read_excel(MASTER); m.columns=[str(c).strip().upper() for c in m.columns]
cap={str(r["LORRY"]).strip().upper():float(r["TON"]) for _,r in m.iterrows()}
usr={str(r["LORRY"]).strip().upper():str(r["USER"]).strip().upper() for _,r in m.iterrows()}

days=sorted(h["DATE"].dt.date.unique())
# sample evenly
idx=np.linspace(0,len(days)-1,min(NDAYS,len(days))).astype(int)
sample=[days[i] for i in idx]

rows=[]
for d in sample:
    day=h[h["DATE"].dt.date==d]
    # DOs whose manual lorry belongs to USER
    dayu=day[day["LICENSE"].map(lambda l: usr.get(l)==USER)].copy()
    if len(dayu)<5: continue
    manual_lorries=set(dayu["LICENSE"])&set(cap)
    if not manual_lorries: continue
    # manual metrics
    man_used=len(manual_lorries)
    man_load={l:dayu[dayu["LICENSE"]==l]["W"].sum() for l in manual_lorries}
    man_util=np.mean([man_load[l]/cap[l] for l in manual_lorries if cap[l]>0])
    # run bot with the same fleet (mark available = manual_lorries), same DOs
    mm=m.copy()
    mm["Status"]=mm["LORRY"].map(lambda l: "Available" if str(l).strip().upper() in manual_lorries else "Block")
    mb=io.BytesIO(); mm.to_excel(mb,index=False); mb.seek(0)
    bot.sessions.pop("t",None); s=bot.get_session("t")
    bot._handle_user_id("t",s,USER); 
    if s.get("state")=="AWAIT_MASTER_UPLOAD": bot._handle_master_upload("t",s,mb.read())
    s["_ignore_schedule"]=True; bot._handle_trip_day("t",s,"today")
    do=dayu.drop(columns=["W"]).copy(); do["LICENSE"]=""
    b=io.BytesIO(); do.to_excel(b,index=False); b.seek(0)
    try:
        msg=bot._handle_excel_upload("t",s,b.read())
    except Exception as e:
        rows.append((d,len(dayu),man_used,round(man_util*100),"CRASH",str(e)[:30])); continue
    items=s["items"]
    un=sum(1 for it in items if it.get("LORRY") in ("NO_LORRY","NO_ELIGIBLE_LORRY"))
    bl={}
    for it in items:
        l=it.get("LORRY")
        if l in cap: bl[l]=bl.get(l,0)+it["WEIGHT"]
    bot_used=len(bl)
    bot_util=np.mean([bl[l]/cap[l] for l in bl if cap[l]>0]) if bl else 0
    rows.append((d,len(dayu),man_used,round(man_util*100),bot_used,round(bot_util*100),un))

print(f"=== {USER}: manual vs bot (sample {len(rows)} days) ===")
print(f"{'date':12} {'DOs':4} {'MANlor':6} {'MANutil':7} {'BOTlor':6} {'BOTutil':7} {'BOTunassigned':6}")
tot_un=0
for r in rows:
    print("  ".join(str(x) for x in r))
    if isinstance(r[-1],int): tot_un+=r[-1]
print(f"\nTOTAL bot unassigned across sample: {tot_un}")
