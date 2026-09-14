import json, time, urllib.parse, urllib.request, os, sys

ZIP, RADIUS = "46514", 75
TARGETS = [("Toyota","4Runner"),("Toyota","RAV4"),("Honda","HR-V"),("Honda","CR-V"),("Subaru","Outback")]
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "raw")
os.makedirs(OUT, exist_ok=True)

def fetch(params, tries=4):
    url = "https://helix.carfax.com/search/v2/vehicles?" + urllib.parse.urlencode(params)
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers={"User-Agent":"Mozilla/5.0","Accept":"application/json"})
            with urllib.request.urlopen(req, timeout=60) as r:
                return json.loads(r.read().decode())
        except Exception as e:
            if i == tries-1: raise
            time.sleep(2**i)

all_rows, seen = [], set()
for make, model in TARGETS:
    page, total_pages = 1, 1
    while page <= total_pages and page <= 20:
        p = {"zip":ZIP,"radius":RADIUS,"make":make,"model":model,"yearMin":2022,
             "priceMax":40000,"vehicleCondition":"USED","sort":"BEST","page":page}
        d = fetch(p)
        total_pages = d.get("totalPageCount",1)
        for l in d.get("listings",[]):
            vin = l.get("vin")
            if not vin or vin in seen: continue
            seen.add(vin); all_rows.append(l)
        print(f"{make} {model} page {page}/{total_pages} total={d.get('totalListingCount')} kept={len(all_rows)}", flush=True)
        page += 1
        time.sleep(0.4)

json.dump(all_rows, open(os.path.join(OUT,"listings.json"),"w"))
print("TOTAL UNIQUE:", len(all_rows))
