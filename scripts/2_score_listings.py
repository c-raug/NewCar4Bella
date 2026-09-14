import json, os, datetime, statistics, re
ROOT=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA=os.path.join(ROOT,'data')
rows=json.load(open(os.path.join(DATA,'carfax_listings_latest.json')))
TODAY=datetime.date(2026,9,14)
MILE_CAP={'Toyota':100000,'Honda':70000,'Subaru':50000}

def optset(l):
    s=set()
    for k in ('topOptions','atomTopOptions','atomOtherOptions'):
        for o in (l.get(k) or []): s.add(str(o).lower())
    return s

FEATURES=[  # label, matcher, points
 ('Heated Seats',      lambda o: any('heated seat' in x or 'heated front seat' in x for x in o), 0),
 ('Heated Steering',   lambda o: any('heated steering' in x for x in o), 10),
 ('Sunroof/Moonroof',  lambda o: any('sunroof' in x or 'moonroof' in x for x in o), 14),
 ('Blind Spot',        lambda o: any('blind spot' in x for x in o), 12),
 ('Power Liftgate',    lambda o: any('power liftgate' in x for x in o), 12),
 ('Leather',           lambda o: any(('leather' in x and 'leatherette' not in x and 'steering' not in x) for x in o), 12),
 ('Remote Start',      lambda o: any('remote start' in x for x in o), 10),
 ('Navigation',        lambda o: any('navigation' in x for x in o), 8),
 ('Wireless Charging', lambda o: any('wireless charg' in x for x in o), 8),
 ('Parking Sensors',   lambda o: any('parking sensor' in x for x in o), 6),
 ('Power Driver Seat', lambda o: any('power seat' in x or 'power driver' in x for x in o), 4),
 ('Premium Audio',     lambda o: any('premium' in x and ('audio' in x or 'sound' in x or 'speaker' in x) for x in o), 4),
]
FEAT_MAX=sum(p for _,_,p in FEATURES)

# ---- market price model: per model, least squares price ~ a + b*year + c*mileage ----
def fit(group):
    n=len(group)
    if n<8: return None
    import itertools
    X=[[1.0,float(g['year']),float(g['mileage'] or 0)/1000.0] for g in group]
    y=[float(g['currentPrice'] or g['listPrice'] or 0) for g in group]
    # normal equations 3x3
    A=[[sum(X[i][r]*X[i][c] for i in range(n)) for c in range(3)] for r in range(3)]
    b=[sum(X[i][r]*y[i] for i in range(n)) for r in range(3)]
    # gaussian elimination
    M=[A[r][:]+[b[r]] for r in range(3)]
    for c in range(3):
        p=max(range(c,3), key=lambda r: abs(M[r][c]))
        if abs(M[p][c])<1e-9: return None
        M[c],M[p]=M[p],M[c]
        for r in range(3):
            if r!=c:
                f=M[r][c]/M[c][c]
                for k in range(c,4): M[r][k]-=f*M[c][k]
    return [M[i][3]/M[i][i] for i in range(3)]

by_model={}
for l in rows: by_model.setdefault(l['model'],[]).append(l)
coefs={m:fit(g) for m,g in by_model.items()}

# reliability fallback medians per model
relscore_raw={}
def rel_component(l):
    r=l.get('reliability') or {}
    if not r: return None
    badge={'GREAT':100,'GOOD':72,'FAIR':50}.get(r.get('overallReliabilityBadge'),60)
    cost={'LOW':100,'AVERAGE':65,'HIGH':30}.get(r.get('costBadge'),65)
    risk={'LOW':100,'AVERAGE':65,'HIGH':30}.get(r.get('riskBadge'),65)
    pl=r.get('reliabilityPlacement')
    place=100 if pl==25 else (72 if pl==50 else (50 if pl else 65))
    return 0.45*badge+0.20*cost+0.20*risk+0.15*place
for l in rows:
    c=rel_component(l)
    if c is not None: relscore_raw.setdefault(l['model'],[]).append(c)
rel_fallback={m:statistics.median(v) for m,v in relscore_raw.items()}

out=[]
for l in rows:
    make,model=l['make'],l['model']
    price=l.get('currentPrice') or l.get('listPrice') or 0
    miles=l.get('mileage') or 0
    year=l.get('year')
    o=optset(l)
    cap=MILE_CAP.get(make,100000)

    # --- market delta ---
    co=coefs.get(model)
    pred=None; delta=None
    if co and price:
        pred=co[0]+co[1]*year+co[2]*(miles/1000.0)
        if pred>5000: delta=(pred-price)/pred
    badge=l.get('badge')
    badge_s={'GREAT':100,'GOOD':78,'FAIR':55}.get(badge,62)
    if delta is None: delta_s=62
    else: delta_s=max(0.0,min(100.0,(delta+0.15)/0.30*100))
    value_sub=0.55*badge_s+0.45*delta_s

    # --- mileage ---
    cap_s=max(0.0,min(100.0,(1-miles/cap)*100))
    age=max(TODAY.year-year,0.5)
    ann=miles/age
    ann_s=max(0.0,min(100.0,(1-(ann-6000)/12000)*100))
    mile_sub=0.7*cap_s+0.3*ann_s

    # --- reliability (model rating 60% + this car's history 40%) ---
    rc=rel_component(l)
    rc_est = rc is None
    if rc is None: rc=rel_fallback.get(model,70.0)
    sh=(l.get('serviceHistory') or {})
    nsvc=sh.get('number') or 0
    hist=0.0
    hist+= 34 if l.get('noAccidents') else 0
    hist+= 22 if l.get('oneOwner') else 0
    hist+= 14 if l.get('personalUse') else 0
    hist+= min(30.0, nsvc*3.0)
    rel_sub=0.6*rc+0.4*hist

    # --- features ---
    feats={}
    pts=0
    for label,fn,p in FEATURES:
        has=fn(o); feats[label]=has
        if has: pts+=p
    feat_sub=min(100.0,pts/FEAT_MAX*100) if FEAT_MAX else 0

    dist=l.get('distanceToDealer') or 0
    dist_sub=max(0.0,min(100.0,(1-dist/75.0)*100))
    cpo_sub=100.0 if l.get('certified') else 0.0

    fs=l.get('firstSeen')
    days=None
    if fs:
        try: days=(TODAY-datetime.date.fromisoformat(fs)).days
        except Exception: pass
    ph=l.get('priceHistory') or []
    drop=0
    if len(ph)>1:
        first=ph[-1].get('listPrice'); 
        if first and price: drop=first-price
    mpe=(l.get('monthlyPaymentEstimate') or {}).get('monthlyPayment')
    fees=sum(f.get('fee',0) for f in (l.get('fees') or []) if f.get('feeType')=='document_fee')
    oh=(l.get('ownerHistory') or {})
    nown=len(oh.get('history') or []) or (1 if l.get('oneOwner') else None)
    r=l.get('reliability') or {}

    out.append(dict(
      vin=l['vin'], year=year, make=make, model=model,
      trim=(l.get('trim') or '')+('' if (l.get('subTrim') in (None,'Unspecified')) else ' '+l['subTrim']),
      price=price, miles=miles, drivetype=l.get('drivetype'), fuel=l.get('fuel'),
      body=l.get('bodytype'), ext=l.get('exteriorColor'), int_=l.get('interiorColor'),
      engine=(l.get('displacement') or '')+' '+(l.get('engine') or ''),
      mpg=l.get('mpgCombined'), certified=bool(l.get('certified')),
      no_accidents=bool(l.get('noAccidents')), owners=nown,
      personal=bool(l.get('personalUse')), svc_records=nsvc,
      value_badge=badge or 'None', rel_badge=r.get('overallReliabilityBadge') or 'n/a',
      repair_cost=r.get('averageCost'), rel_estimated=rc_est,
      feats=feats, dealer=(l.get('dealer') or {}).get('name'),
      city=(l.get('dealer') or {}).get('city'), state=(l.get('dealer') or {}).get('state'),
      phone=(l.get('dealer') or {}).get('phone'),
      dealer_rating=(l.get('dealer') or {}).get('dealerAverageRating'),
      dealer_reviews=(l.get('dealer') or {}).get('dealerReviewCount'),
      distance=round(dist,1), days_listed=days, price_drop=drop,
      monthly=mpe, doc_fee=fees, pred_price=round(pred) if pred else None,
      delta_pct=round(delta*100,1) if delta is not None else None,
      s_rel=round(rel_sub,1), s_val=round(value_sub,1), s_mile=round(mile_sub,1),
      s_feat=round(feat_sub,1), s_cpo=cpo_sub, s_dist=round(dist_sub,1),
      url=l.get('vdpUrl'),
    ))

json.dump(out, open(os.path.join(DATA,'scored_latest.json'),'w'), indent=0)
print('rows:',len(out))
hard=[r for r in out if r['drivetype'] in ('AWD','4WD','4X4') and r['no_accidents'] and r['svc_records']>0
      and r['miles']<=MILE_CAP[r['make']] and r['value_badge'] in ('GOOD','GREAT') and r['feats']['Heated Seats']]
print('meets all (ex-CPO):',len(hard),' of which CPO:',sum(1 for r in hard if r['certified']))
