import json, os
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule

ROOT=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA=os.path.join(ROOT,'data')
BASE=os.path.join(ROOT,'output')
os.makedirs(BASE,exist_ok=True)
OUT=os.path.join(BASE,'Bella_Car_Search.xlsx')
D=json.load(open(os.path.join(DATA,'scored_latest.json')))
LAY=json.load(open(os.path.join(DATA,'layout.json')))
H=LAY['H']; FIRST=LAY['first']
FONT='Arial'; NAVY='1F3864'; ACC='2E5C8A'; YEL='FFF2CC'
HDR=PatternFill('solid',fgColor=NAVY); SUB=PatternFill('solid',fgColor=ACC)
thin=Side(style='thin',color='BFBFBF'); BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
wb=load_workbook(OUT)
AC="'All Cars'"
MILE_CAP={'Toyota':100000,'Honda':70000,'Subaru':50000}

D.sort(key=lambda r:-(0.26*r['s_rel']+0.22*r['s_val']+0.18*r['s_mile']+0.15*r['s_cpo']+0.14*r['s_feat']+0.05*r['s_dist']))
rowof={d['vin']:FIRST+i for i,d in enumerate(D)}
def L(n): return get_column_letter(H[n])
def ref(vin,col): return f"={AC}!{L(col)}{rowof[vin]}"

passing=[d for d in D if d['drivetype'] in ('AWD','4WD','4X4') and d['no_accidents'] and d['svc_records']>0
         and d['miles']<=MILE_CAP[d['make']] and d['value_badge'] in ('GOOD','GREAT') and d['feats']['Heated Seats']]

def style_header(ws,row,ncols,fill=HDR,h=30):
    for c in range(1,ncols+1):
        cell=ws.cell(row=row,column=c)
        cell.font=Font(name=FONT,size=10,bold=True,color='FFFFFF'); cell.fill=fill
        cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); cell.border=BOX
    ws.row_dimensions[row].height=h
def title(ws,t,sub=None):
    ws['A1']=t; ws['A1'].font=Font(name=FONT,size=16,bold=True,color=NAVY)
    if sub:
        ws['A2']=sub; ws['A2'].font=Font(name=FONT,size=10,italic=True,color='595959')

# ============================== SHORTLIST ==============================
sl=wb.create_sheet('Shortlist',2)
title(sl,f'Shortlist - {len(passing)} Cars That Meet Every Criterion',
      'AWD/4WD, no accidents, service history on file, under your mileage cap for the make, GOOD or GREAT value, heated seats, Apple CarPlay. Ranked best-value first. Certified cars are marked and scored higher.')
SLC=[('Rank',7),('Score',9),('Still Meets',11),('Year',7),('Make',9),('Model',11),('Trim',17),('Price',11),
     ('Mileage',10),('Miles/Yr',9),('CPO',7),('Value Badge',12),('vs Market',10),('Reliability',11),
     ('Owners',8),('Service Recs',12),('Sunroof/Moonroof',12),('Leather',10),('Power Liftgate',12),
     ('Blind Spot',10),('Fuel',10),('MPG',7),('Dealer',28),('City',14),('Miles Away',11),
     ('Est. Monthly',12),('CarFax Link',30),('Listing URL',46)]
SH={n:i+1 for i,(n,_) in enumerate(SLC)}
HR=4
for n,w in SLC:
    sl.cell(row=HR,column=SH[n],value=n); sl.column_dimensions[get_column_letter(SH[n])].width=w
style_header(sl,HR,len(SLC))
sl.freeze_panes='D5'
sf=HR+1
for i,d in enumerate(passing):
    R=sf+i
    def put(name,val,fmt=None,bold=False,align='center'):
        c=sl.cell(row=R,column=SH[name],value=val)
        c.font=Font(name=FONT,size=10,bold=bold)
        if fmt: c.number_format=fmt
        c.alignment=Alignment(horizontal=align); c.border=BOX; return c
    sc=get_column_letter(SH['Score'])
    put('Score',ref(d['vin'],'Score'),'0.0',bold=True)
    put('Rank',f"=RANK({sc}{R},${sc}${sf}:${sc}${sf+len(passing)-1})",'0')
    put('Still Meets',ref(d['vin'],'Meets All'))
    for n in ['Year','Make','Model','Trim','Price','Mileage','Miles/Yr','CPO','Value Badge','vs Market',
              'Reliability','Owners','Service Recs','Sunroof/Moonroof','Leather','Power Liftgate',
              'Blind Spot','Fuel','MPG','Dealer','City','Miles Away','Est. Monthly']:
        fmt=None
        if n in ('Price','Est. Monthly'): fmt='$#,##0'
        elif n in ('Mileage','Miles/Yr'): fmt='#,##0'
        elif n=='vs Market': fmt='0.0%;(0.0%);-'
        elif n=='Miles Away': fmt='0.0'
        elif n in ('Year','MPG','Owners','Service Recs'): fmt='0'
        put(n,ref(d['vin'],n),fmt,align='left' if n in ('Trim','Dealer','City','Model') else 'center')
    c=sl.cell(row=R,column=SH['CarFax Link'],value=f"{d['year']} {d['make']} {d['model']} {d['trim']}".strip())
    c.hyperlink=d['url']; c.font=Font(name=FONT,size=10,color='0563C1',underline='single'); c.border=BOX
    c.alignment=Alignment(horizontal='left')
    u=sl.cell(row=R,column=SH['Listing URL'],value=d['url'])
    u.font=Font(name=FONT,size=9,color='0563C1'); u.border=BOX
slast=sf+len(passing)-1
sl.auto_filter.ref=f"A{HR}:{get_column_letter(len(SLC))}{slast}"
scl=get_column_letter(SH['Score'])
sl.conditional_formatting.add(f"{scl}{sf}:{scl}{slast}",
    ColorScaleRule(start_type='min',start_color='F8696B',mid_type='percentile',mid_value=50,
                   mid_color='FFEB84',end_type='max',end_color='63BE7B'))
cp=get_column_letter(SH['CPO'])
sl.conditional_formatting.add(f"{cp}{sf}:{cp}{slast}",
    CellIsRule(operator='equal',formula=['"Yes"'],fill=PatternFill('solid',fgColor='C6EFCE'),
               font=Font(name=FONT,size=10,bold=True,color='006100')))
sm=get_column_letter(SH['Still Meets'])
sl.conditional_formatting.add(f"{sm}{sf}:{sm}{slast}",
    CellIsRule(operator='equal',formula=['"no"'],fill=PatternFill('solid',fgColor='FFC7CE'),
               font=Font(name=FONT,size=10,color='9C0006')))
sl.cell(row=slast+2,column=1,value='"Still Meets" recalculates from the Search Criteria tab. If you tighten a filter, rows that no longer qualify turn red here rather than disappearing - check the All Cars tab for anything new that qualifies.').font=Font(name=FONT,size=9,italic=True,color='595959')

# ============================== TOP 10 COMPARE ==============================
tc=wb.create_sheet('Top 10 Compare',3)
title(tc,'Top 10 Head-to-Head','The ten highest-scoring cars from the Shortlist, side by side. Everything here is live-linked to All Cars.')
top=passing[:10]
ATTRS=[('Score','0.0'),('Year','0'),('Make',None),('Model',None),('Trim',None),('Price','$#,##0'),
 ('Mileage','#,##0'),('Miles/Yr','#,##0'),('Est. Monthly','$#,##0'),('CPO',None),('Value Badge',None),
 ('Est. Market','$#,##0'),('vs Market','0.0%;(0.0%);-'),('Reliability',None),('Repair $/yr','$#,##0'),
 ('No Accidents',None),('Owners','0'),('Personal Use',None),('Service Recs','0'),('Drivetrain',None),
 ('Fuel',None),('MPG','0'),('Heated Seats',None),('Heated Steering',None),('Sunroof/Moonroof',None),
 ('Leather',None),('Power Liftgate',None),('Blind Spot',None),('Remote Start',None),('Navigation',None),
 ('Wireless Charging',None),('Ext Color',None),('Int Color',None),('Dealer',None),('City',None),
 ('Miles Away','0.0'),('Dealer Rating','0.0'),('Days Listed','0'),('Price Drop','$#,##0')]
tc.cell(row=4,column=1,value='Attribute')
for j,d in enumerate(top):
    tc.cell(row=4,column=2+j,value=f"#{j+1}  {d['year']} {d['make']} {d['model']} {d['trim']}".strip())
    tc.column_dimensions[get_column_letter(2+j)].width=19
style_header(tc,4,1+len(top),h=46)
tc.column_dimensions['A'].width=20
tc.freeze_panes='B5'
for i,(attr,fmt) in enumerate(ATTRS):
    R=5+i
    a=tc.cell(row=R,column=1,value=attr)
    a.font=Font(name=FONT,size=10,bold=True); a.fill=PatternFill('solid',fgColor='F2F2F2'); a.border=BOX
    for j,d in enumerate(top):
        c=tc.cell(row=R,column=2+j,value=ref(d['vin'],attr))
        c.font=Font(name=FONT,size=10); c.alignment=Alignment(horizontal='center'); c.border=BOX
        if fmt: c.number_format=fmt
R=5+len(ATTRS)
tc.cell(row=R,column=1,value='CarFax').font=Font(name=FONT,size=10,bold=True)
tc.cell(row=R,column=1).fill=PatternFill('solid',fgColor='F2F2F2'); tc.cell(row=R,column=1).border=BOX
for j,d in enumerate(top):
    c=tc.cell(row=R,column=2+j,value='Open CARFAX listing')
    c.hyperlink=d['url']; c.font=Font(name=FONT,size=10,color='0563C1',underline='single')
    c.alignment=Alignment(horizontal='center'); c.border=BOX
R2=R+1
tc.cell(row=R2,column=1,value='Listing URL').font=Font(name=FONT,size=10,bold=True)
tc.cell(row=R2,column=1).fill=PatternFill('solid',fgColor='F2F2F2'); tc.cell(row=R2,column=1).border=BOX
for j,d in enumerate(top):
    c=tc.cell(row=R2,column=2+j,value=d['url'])
    c.font=Font(name=FONT,size=8,color='0563C1'); c.border=BOX
    c.alignment=Alignment(horizontal='center',wrap_text=True)
tc.conditional_formatting.add(f"B5:{get_column_letter(1+len(top))}5",
    ColorScaleRule(start_type='min',start_color='F8696B',mid_type='percentile',mid_value=50,
                   mid_color='FFEB84',end_type='max',end_color='63BE7B'))

# ============================== MODEL BENCHMARKS ==============================
mb=wb.create_sheet('Model Benchmarks',4)
title(mb,'Model Benchmarks','What each model and year actually goes for in this market. Everything is a live formula over the All Cars tab - it updates if you edit that data.')
cols=[('Make',12),('Model',12),('Year',8),('Cars Listed',11),('Avg Price',12),('Min Price',12),
      ('Max Price',12),('Avg Mileage',12),('Avg Score',11),('CPO Count',11),('Meets-All Count',14)]
MH={n:i+1 for i,(n,_) in enumerate(cols)}
for n,w in cols:
    mb.cell(row=4,column=MH[n],value=n); mb.column_dimensions[get_column_letter(MH[n])].width=w
style_header(mb,4,len(cols))
combos=sorted({(d['make'],d['model'],d['year']) for d in D})
LAST=LAY['last']
rngs={n:f"{AC}!${L(n)}${FIRST}:${L(n)}${LAST}" for n in
      ['Make','Model','Year','Price','Mileage','Score','CPO','Meets All']}
for i,(mk,mo,yr) in enumerate(combos):
    R=5+i
    crit=f'{rngs["Make"]},$A{R},{rngs["Model"]},$B{R},{rngs["Year"]},$C{R}'
    vals=[('Make',mk,None),('Model',mo,None),('Year',yr,'0'),
     ('Cars Listed',f'=COUNTIFS({crit})','0'),
     ('Avg Price',f'=IFERROR(AVERAGEIFS({rngs["Price"]},{crit}),"")','$#,##0'),
     ('Min Price',f'=IFERROR(_xlfn.MINIFS({rngs["Price"]},{crit}),"")','$#,##0'),
     ('Max Price',f'=IFERROR(_xlfn.MAXIFS({rngs["Price"]},{crit}),"")','$#,##0'),
     ('Avg Mileage',f'=IFERROR(AVERAGEIFS({rngs["Mileage"]},{crit}),"")','#,##0'),
     ('Avg Score',f'=IFERROR(AVERAGEIFS({rngs["Score"]},{crit}),"")','0.0'),
     ('CPO Count',f'=COUNTIFS({crit},{rngs["CPO"]},"Yes")','0'),
     ('Meets-All Count',f'=COUNTIFS({crit},{rngs["Meets All"]},"YES")','0')]
    for n,v,fmt in vals:
        c=mb.cell(row=R,column=MH[n],value=v)
        c.font=Font(name=FONT,size=10); c.border=BOX; c.alignment=Alignment(horizontal='center')
        if fmt: c.number_format=fmt
mlast=4+len(combos)
mb.auto_filter.ref=f"A4:{get_column_letter(len(cols))}{mlast}"
wb.save(OUT)
print('shortlist',len(passing),'combos',len(combos),'saved')
