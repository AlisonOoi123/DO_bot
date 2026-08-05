"""Full-corpus shadow run: optimizer vs manual, per user, with rule-invariant
checks. Uses the DOs the manual actually assigned each day (ground truth) and
the fleet the manual used that day. Single-trip only."""
import sys, re, math, json
sys.path.insert(0,"docs/optimizer"); sys.path.insert(0,".")
import pandas as pd, numpy as np
from optimizer import (Cluster, Lorry, optimize, OVERLOAD_CAP,
                       MAX_DOS_PER_LORRY, URBAN_MIX_DIST_DEG)

HIST, MASTER, USER, MAXDAYS = sys.argv[1], sys.argv[2], sys.argv[3], int(sys.argv[4])
URB={"SELANGOR","KUALA LUMPUR","WILAYAH PERSEKUTUAN","W.P. KUALA LUMPUR"}
def pfx(r):
    m=re.match(r'\s*([A-Z]{1,3}\d+[A-Z]?)',str(r)); return m.group(1) if m else ""
def ll(v):
    try: p=str(v).split(); return float(p[0]),float(p[1])
    except: return (None,None)

h=pd.read_excel(HIST); h.columns=[str(c).strip().upper() for c in h.columns]
h["LICENSE"]=h["LICENSE"].astype(str).str.strip().str.upper()
h=h[~h["LICENSE"].isin(["","NAN","NONE"])]; h["DATE"]=pd.to_datetime(h["DATE"],errors="coerce")
h=h[h["DATE"].notna()]; h["W"]=pd.to_numeric(h["GROSS WEIGHT"],errors="coerce").fillna(0)/1000.0
m=pd.read_excel(MASTER); m.columns=[str(c).strip().upper() for c in m.columns]
cap={str(r["LORRY"]).strip().upper():float(r["TON"]) for _,r in m.iterrows()}
usr={str(r["LORRY"]).strip().upper():str(r["USER"]).strip().upper() for _,r in m.iterrows()}

def build_clusters(dayu, maxcap):
    cl=[]
    for code,g in dayu.groupby(dayu["ROUTE"].apply(pfx)):
        rows=sorted(g.to_dict("records"),key=lambda r:-r["W"])
        lat=[ll(r["LONGITUD"])[0] for r in rows if ll(r["LONGITUD"])[0] is not None]
        lon=[ll(r["LONGITUD"])[1] for r in rows if ll(r["LONGITUD"])[1] is not None]
        clat=sum(lat)/len(lat) if lat else None; clon=sum(lon)/len(lon) if lon else None
        st=str(rows[0].get("STATE","")).upper(); route=str(rows[0]["ROUTE"]).upper()
        iu= st in URB; ik= route.startswith("PH09") or "KUANTAN" in route
        chunk=[];cw=0;idx=0
        for r in rows:
            if (cw+r["W"]>maxcap or len(chunk)>=MAX_DOS_PER_LORRY) and chunk:
                cl.append(Cluster(f"{code}#{idx}",round(cw,3),clat,clon,iu,ik,None,len(chunk)));idx+=1;chunk=[];cw=0
            chunk.append(r);cw+=r["W"]
        if chunk: cl.append(Cluster(f"{code}#{idx}",round(cw,3),clat,clon,iu,ik,None,len(chunk)))
    return cl

def violations(clusters, assign, cmap):
    v=0; per={}
    for cid,pl in assign.items():
        if not pl: continue
        per.setdefault(pl,[]).append(cmap[cid])
    for pl,cs in per.items():
        load=sum(c.weight for c in cs); nd=sum(c.ndos for c in cs)
        if load>cap[pl]*OVERLOAD_CAP+1e-6: v+=1
        if nd>MAX_DOS_PER_LORRY: v+=1
        if any(c.is_urban for c in cs) and any(not c.is_urban for c in cs): v+=1
        if any(c.is_kuantan for c in cs) and len(cs)>1: v+=1
        urb=[c for c in cs if c.is_urban and c.lat is not None]
        for i in range(len(urb)):
            for j in range(i+1,len(urb)):
                if math.hypot(urb[i].lat-urb[j].lat,urb[i].lon-urb[j].lon)>URBAN_MIX_DIST_DEG+1e-9: v+=1
    return v

days=sorted(h["DATE"].dt.date.unique())
idx=np.linspace(0,len(days)-1,min(MAXDAYS,len(days))).astype(int)
sample=[days[i] for i in idx]
res=[]
for d in sample:
    dayu=h[(h["DATE"].dt.date==d)&(h["LICENSE"].map(lambda l:usr.get(l)==USER))].copy()
    fp=set(dayu["LICENSE"])&set(cap)
    if len(dayu)<5 or not fp: continue
    fleet=[Lorry(p,cap[p]) for p in fp]; maxcap=max(cap[p] for p in fp)
    cl=build_clusters(dayu,maxcap); cmap={c.id:c for c in cl}
    tw=sum(c.weight for c in cl)
    o=optimize(cl,fleet,3.0)
    load={};aw=0
    for cid,pl in o["assign"].items():
        if pl: load[pl]=load.get(pl,0)+cmap[cid].weight; aw+=cmap[cid].weight
    v=violations(cl,o["assign"],cmap)
    man_l=len(fp); man_util=np.mean([dayu[dayu["LICENSE"]==p]["W"].sum()/cap[p] for p in fp])
    fleet_cap=sum(cap[p] for p in fp)*OVERLOAD_CAP
    res.append(dict(date=str(d),dos=len(dayu),totW=round(tw,1),
                    fleetCapx=round(fleet_cap,1), fits=(tw<=fleet_cap),
                    man_lor=man_l, man_util=round(man_util*100),
                    opt_lor=len(load), opt_util=round(np.mean([load[p]/cap[p] for p in load])*100) if load else 0,
                    opt_unW=round(tw-aw,2), viol=v, status=o["status"]))
# aggregate
import statistics as st
fit=[r for r in res if r["fits"]]
print(f"=== {USER}: {len(res)} days ({len(fit)} single-trip-feasible) ===")
print(f"rule violations total: {sum(r['viol'] for r in res)}")
if fit:
    print(f"on feasible days: opt fully-assigned = {sum(1 for r in fit if r['opt_unW']<0.5)}/{len(fit)}")
    print(f"  avg lorries: manual {st.mean(r['man_lor'] for r in fit):.1f} vs optim {st.mean(r['opt_lor'] for r in fit):.1f}")
    print(f"  avg util:    manual {st.mean(r['man_util'] for r in fit):.0f}% vs optim {st.mean(r['opt_util'] for r in fit):.0f}%")
infeasible=[r for r in res if not r["fits"]]
print(f"days needing >single-trip capacity (manual used 2nd trips): {len(infeasible)}")
json.dump(res, open(f"docs/optimizer/shadow_{USER}.json","w"), indent=1)
print("wrote docs/optimizer/shadow_%s.json"%USER)
