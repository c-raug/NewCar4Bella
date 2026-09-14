import json, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule

ROOT=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA=os.path.join(ROOT,'data')
BASE=os.path.join(ROOT,'output')
os.makedirs(BASE,exist_ok=True)
D=json.load(open(os.path.join(DATA,'scored_latest.json')))
OUT=os.path.join(BASE,'Bella_Car_Search.xlsx')

FONT='Arial'
NAVY='1F3864'; ACC='2E5C8A'; LIGHT='DCE6F1'; YEL='FFF2CC'; GREY='F2F2F2'
HDR=PatternFill('solid',fgColor=NAVY)
SUBHDR=PatternFill('solid',fgColor=ACC)
INPUT=PatternFill('solid',fgColor=YEL)
BANDA=PatternFill('solid',fgColor='FFFFFF')
thin=Side(style='thin',color='BFBFBF')
BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
BLUE=Font(name=FONT,size=10,color='0000FF')

wb=Workbook()

def style_header(ws,row,ncols,fill=HDR):
    for c in range(1,ncols+1):
        cell=ws.cell(row=row,column=c)
        cell.font=Font(name=FONT,size=10,bold=True,color='FFFFFF')
        cell.fill=fill
        cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
        cell.border=BOX
    ws.row_dimensions[row].height=30

def title(ws,text,sub=None):
    ws['A1']=text
    ws['A1'].font=Font(name=FONT,size=16,bold=True,color=NAVY)
    if sub:
        ws['A2']=sub
        ws['A2'].font=Font(name=FONT,size=10,italic=True,color='595959')

# ============================ SEARCH CRITERIA ============================
cr=wb.active; cr.title='Search Criteria'
title(cr,'Search Criteria','Edit the yellow cells. Every YES/NO in the "Meets All" column recalculates from these.')
cr['A4']='Filter'; cr['B4']='Value'; cr['C4']='Notes'
style_header(cr,4,3)
crit=[
 ('Max Price ($)',40000,'Listing price, before tax/title/doc fee.'),
 ('Min Year',2022,'Model year.'),
 ('Max Distance (mi)',75,'Straight-line distance from Elkhart, IN 46514.'),
 ('Require AWD / 4WD',True,'TRUE or FALSE.'),
 ('Require No Accidents',True,'No accident or damage reported to CARFAX.'),
 ('Require Service History',True,'At least one service record on file.'),
 ('Require Heated Seats',True,'From the trim equipment list.'),
 ('Require Good/Great Value',True,"CARFAX's own value badge must be GOOD or GREAT."),
 ('Require Certified Pre-Owned',False,'Set TRUE to see only CPO cars (that leaves 6).'),
]
r=5
for lab,val,note in crit:
    cr.cell(row=r,column=1,value=lab).font=Font(name=FONT,size=10,bold=True)
    c=cr.cell(row=r,column=2,value=val); c.fill=INPUT; c.font=BLUE; c.border=BOX
    c.alignment=Alignment(horizontal='center')
    if 'Price' in lab: c.number_format='$#,##0'
    cr.cell(row=r,column=3,value=note).font=Font(name=FONT,size=9,color='595959')
    cr.cell(row=r,column=1).border=BOX
    r+=1
CRIT_ROW={lab:5+i for i,(lab,_,_) in enumerate(crit)}

cr['A16']='Mileage Cap by Make'; cr['A16'].font=Font(name=FONT,size=12,bold=True,color=NAVY)
cr['A17']='Make'; cr['B17']='Max Miles'; cr['C17']='Notes'
style_header(cr,17,3)
caps=[('Toyota',100000,'Your stated cap.'),('Honda',70000,'Your stated cap.'),('Subaru',50000,'Your stated cap.')]
for i,(mk,v,note) in enumerate(caps):
    rr=18+i
    cr.cell(row=rr,column=1,value=mk).font=Font(name=FONT,size=10,bold=True)
    c=cr.cell(row=rr,column=2,value=v); c.fill=INPUT; c.font=BLUE; c.number_format='#,##0'
    c.alignment=Alignment(horizontal='center'); c.border=BOX
    cr.cell(row=rr,column=1).border=BOX
    cr.cell(row=rr,column=3,value=note).font=Font(name=FONT,size=9,color='595959')
cr.column_dimensions['A'].width=30; cr.column_dimensions['B'].width=14; cr.column_dimensions['C'].width=62
cr['A23']='Models searched: Toyota 4Runner, Toyota RAV4, Honda HR-V, Honda CR-V, Subaru Outback.'
cr['A23'].font=Font(name=FONT,size=9,italic=True,color='595959')
cr['A24']='Apple CarPlay is standard on every 2022+ car in all five of these model lines, so it never eliminates anyone. It is shown as a column for confirmation.'
cr['A24'].font=Font(name=FONT,size=9,italic=True,color='595959')

# ============================ SCORING WEIGHTS ============================
sw=wb.create_sheet('Scoring Weights')
title(sw,'Scoring Weights','Edit the yellow cells to re-rank every car instantly. They should add to 100.')
sw['A4']='Component'; sw['B4']='Weight'; sw['C4']='What it measures'
style_header(sw,4,3)
weights=[
 ('Reliability & History',26,"CARFAX model reliability rating, projected repair cost and risk (60%) blended with this car's own history - accident-free, one owner, personal use, number of service records (40%). Your #1 priority."),
 ('Price / Value',22,"CARFAX Great/Good/Fair value badge (55%) blended with how far the asking price sits below what a same-model car of that year and mileage typically lists for (45%). Your #2 priority."),
 ('Mileage',18,'Miles against your cap for that make (70%) blended with miles-per-year against a 12,000/yr normal (30%). Your #3 priority.'),
 ('Certified Pre-Owned',15,'100 if factory certified, 0 if not. Weighted heavily so CPO cars rise, without excluding good non-CPO cars.'),
 ('Features',14,'Heated steering, sunroof, blind spot, power liftgate, leather, remote start, nav, wireless charging, parking sensors, power seat, premium audio. Your #4 priority.'),
 ('Distance',5,'100 at your doorstep down to 0 at 75 miles. Tiebreaker only.'),
]
for i,(lab,w,desc) in enumerate(weights):
    rr=5+i
    sw.cell(row=rr,column=1,value=lab).font=Font(name=FONT,size=10,bold=True)
    c=sw.cell(row=rr,column=2,value=w); c.fill=INPUT; c.font=BLUE; c.border=BOX
    c.alignment=Alignment(horizontal='center'); c.number_format='0'
    sw.cell(row=rr,column=1).border=BOX
    sw.cell(row=rr,column=3,value=desc).font=Font(name=FONT,size=9,color='595959')
    sw.cell(row=rr,column=3).alignment=Alignment(wrap_text=True,vertical='top')
    sw.row_dimensions[rr].height=42
sw['A11']='TOTAL'; sw['A11'].font=Font(name=FONT,size=10,bold=True)
sw['B11']='=SUM(B5:B10)'; sw['B11'].font=Font(name=FONT,size=10,bold=True)
sw['B11'].alignment=Alignment(horizontal='center'); sw['B11'].border=BOX
sw['C11']='=IF(B11=100,"OK - weights add to 100","ADJUST - weights must add to 100, currently "&B11)'
sw['C11'].font=Font(name=FONT,size=10,bold=True,color='008000')
sw.column_dimensions['A'].width=26; sw.column_dimensions['B'].width=12; sw.column_dimensions['C'].width=90
W={'rel':'$B$5','val':'$B$6','mile':'$B$7','cpo':'$B$8','feat':'$B$9','dist':'$B$10','tot':'$B$11'}

# ============================ ALL CARS ============================
FEATCOLS=['Heated Steering','Sunroof/Moonroof','Blind Spot','Power Liftgate','Leather',
          'Remote Start','Navigation','Wireless Charging','Parking Sensors','Power Driver Seat','Premium Audio']
COLS=[('Rank',7),('Score',9),('Meets All',11),('Year',7),('Make',9),('Model',11),('Trim',18),
 ('Price',11),('Mileage',10),('Miles/Yr',10),('CPO',7),('Drivetrain',11),('Fuel',10),('MPG',7),
 ('Value Badge',12),('Est. Market',12),('vs Market',10),('Reliability',11),('Repair $/yr',11),
 ('No Accidents',12),('Owners',8),('Personal Use',12),('Service Recs',12),
 ('Heated Seats',12),('Apple CarPlay',13)]+[(f,13) for f in FEATCOLS]+[
 ('Ext Color',12),('Int Color',11),('Engine',12),('Dealer',30),('City',15),('ST',6),
 ('Miles Away',11),('Dealer Rating',12),('Days Listed',11),('Price Drop',11),
 ('Est. Monthly',12),('Doc Fee',9),
 ('Sub: Reliability',13),('Sub: Value',11),('Sub: Mileage',12),('Sub: Features',12),
 ('Sub: CPO',10),('Sub: Distance',12),('VIN',20),('CarFax Link',30),('Listing URL',46)]
H={name:i+1 for i,(name,_) in enumerate(COLS)}
NC=len(COLS)

ac=wb.create_sheet('All Cars')
title(ac,'All Cars - Full Inventory','Every 2022+ 4Runner / RAV4 / HR-V / CR-V / Outback under $40,000 within 75 miles of Elkhart. Click the filter arrows to narrow. Data pulled from CARFAX 2026-09-14.')
HROW=4
for name,w in COLS:
    ac.cell(row=HROW,column=H[name],value=name)
    ac.column_dimensions[get_column_letter(H[name])].width=w
style_header(ac,HROW,NC)
ac.freeze_panes='D5'

def yn(b): return 'Yes' if b else 'No'
D.sort(key=lambda r:-(0.26*r['s_rel']+0.22*r['s_val']+0.18*r['s_mile']+0.15*r['s_cpo']+0.14*r['s_feat']+0.05*r['s_dist']))

CRS="'Search Criteria'"
SWS="'Scoring Weights'"
first=HROW+1
for i,d in enumerate(D):
    R=first+i
    def put(name,val,fmt=None,bold=False,align=None):
        c=ac.cell(row=R,column=H[name],value=val)
        c.font=Font(name=FONT,size=10,bold=bold)
        if fmt: c.number_format=fmt
        if align: c.alignment=Alignment(horizontal=align)
        c.border=BOX
        return c
    L=lambda n: get_column_letter(H[n])
    put('Score',f"=ROUND(({L('Sub: Reliability')}{R}*{SWS}!{W['rel']}+{L('Sub: Value')}{R}*{SWS}!{W['val']}"
                f"+{L('Sub: Mileage')}{R}*{SWS}!{W['mile']}+{L('Sub: CPO')}{R}*{SWS}!{W['cpo']}"
                f"+{L('Sub: Features')}{R}*{SWS}!{W['feat']}+{L('Sub: Distance')}{R}*{SWS}!{W['dist']})"
                f"/MAX({SWS}!{W['tot']},1),1)",'0.0',bold=True,align='center')
    put('Rank',f"=RANK({L('Score')}{R},${L('Score')}${first}:${L('Score')}${first+len(D)-1})",'0',align='center')
    put('Meets All',
        f'=IF(AND({L("Price")}{R}<={CRS}!$B${CRIT_ROW["Max Price ($)"]},'
        f'{L("Year")}{R}>={CRS}!$B${CRIT_ROW["Min Year"]},'
        f'{L("Miles Away")}{R}<={CRS}!$B${CRIT_ROW["Max Distance (mi)"]},'
        f'{L("Mileage")}{R}<=INDEX({CRS}!$B$18:$B$20,MATCH({L("Make")}{R},{CRS}!$A$18:$A$20,0)),'
        f'OR(NOT({CRS}!$B${CRIT_ROW["Require AWD / 4WD"]}),{L("Drivetrain")}{R}<>"FWD"),'
        f'OR(NOT({CRS}!$B${CRIT_ROW["Require No Accidents"]}),{L("No Accidents")}{R}="Yes"),'
        f'OR(NOT({CRS}!$B${CRIT_ROW["Require Service History"]}),{L("Service Recs")}{R}>0),'
        f'OR(NOT({CRS}!$B${CRIT_ROW["Require Heated Seats"]}),{L("Heated Seats")}{R}="Yes"),'
        f'OR(NOT({CRS}!$B${CRIT_ROW["Require Good/Great Value"]}),{L("Value Badge")}{R}="GREAT",{L("Value Badge")}{R}="GOOD"),'
        f'OR(NOT({CRS}!$B${CRIT_ROW["Require Certified Pre-Owned"]}),{L("CPO")}{R}="Yes")),"YES","no")',
        align='center')
    put('Year',d['year'],'0',align='center'); put('Make',d['make']); put('Model',d['model']); put('Trim',d['trim'])
    put('Price',d['price'],'$#,##0'); put('Mileage',d['miles'],'#,##0')
    put('Miles/Yr',f"=IFERROR(ROUND({L('Mileage')}{R}/MAX(2026-{L('Year')}{R},0.5),0),\"\")",'#,##0',align='center')
    put('CPO',yn(d['certified']),align='center'); put('Drivetrain',d['drivetype'],align='center')
    put('Fuel',d['fuel'],align='center'); put('MPG',d['mpg'],'0',align='center')
    put('Value Badge',d['value_badge'],align='center')
    put('Est. Market',d['pred_price'],'$#,##0')
    put('vs Market',(d['delta_pct']/100.0) if d['delta_pct'] is not None else None,'0.0%;(0.0%);-',align='center')
    put('Reliability',d['rel_badge'],align='center')
    put('Repair $/yr',d['repair_cost'],'$#,##0')
    put('No Accidents',yn(d['no_accidents']),align='center')
    put('Owners',d['owners'],'0',align='center'); put('Personal Use',yn(d['personal']),align='center')
    put('Service Recs',d['svc_records'],'0',align='center')
    put('Heated Seats',yn(d['feats']['Heated Seats']),align='center')
    put('Apple CarPlay','Yes',align='center')
    for f in FEATCOLS: put(f,yn(d['feats'][f]),align='center')
    put('Ext Color',d['ext']); put('Int Color',d['int_']); put('Engine',d['engine'].strip())
    put('Dealer',d['dealer']); put('City',d['city']); put('ST',d['state'],align='center')
    put('Miles Away',d['distance'],'0.0',align='center')
    put('Dealer Rating',d['dealer_rating'],'0.0',align='center')
    put('Days Listed',d['days_listed'],'0',align='center')
    put('Price Drop',d['price_drop'] or None,'$#,##0')
    put('Est. Monthly',d['monthly'],'$#,##0')
    put('Doc Fee',d['doc_fee'] or None,'$#,##0')
    for k,n in (('s_rel','Sub: Reliability'),('s_val','Sub: Value'),('s_mile','Sub: Mileage'),
                ('s_feat','Sub: Features'),('s_cpo','Sub: CPO'),('s_dist','Sub: Distance')):
        put(n,d[k],'0.0',align='center')
    put('VIN',d['vin'])
    c=ac.cell(row=R,column=H['CarFax Link'],value=f"{d['year']} {d['make']} {d['model']}")
    c.hyperlink=d['url']; c.font=Font(name=FONT,size=10,color='0563C1',underline='single'); c.border=BOX
    u=ac.cell(row=R,column=H['Listing URL'],value=d['url'])
    u.font=Font(name=FONT,size=9,color='0563C1'); u.border=BOX
last=first+len(D)-1
ac.auto_filter.ref=f"A{HROW}:{get_column_letter(NC)}{last}"
sc=get_column_letter(H['Score'])
ac.conditional_formatting.add(f"{sc}{first}:{sc}{last}",
    ColorScaleRule(start_type='min',start_color='F8696B',mid_type='percentile',mid_value=50,
                   mid_color='FFEB84',end_type='max',end_color='63BE7B'))
mc=get_column_letter(H['Meets All'])
ac.conditional_formatting.add(f"{mc}{first}:{mc}{last}",
    CellIsRule(operator='equal',formula=['"YES"'],fill=PatternFill('solid',fgColor='C6EFCE'),
               font=Font(name=FONT,size=10,bold=True,color='006100')))
json.dump({'first':first,'last':last,'H':H,'NC':NC,'FEATCOLS':FEATCOLS},open(os.path.join(DATA,'layout.json'),'w'))
wb.save(OUT)
print('saved',OUT,'rows',len(D))
