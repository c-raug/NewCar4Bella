# Bella's Car Search

A filterable, scored spreadsheet of live CARFAX inventory for a used SUV search around
Elkhart, Indiana.

**Deliverable:** [`output/Bella_Car_Search.xlsx`](output/Bella_Car_Search.xlsx)

To use it as a Google Sheet: upload the file to Google Drive, then open it with Google
Sheets (Drive converts it on open). Formulas, filters, conditional formatting, and the
per-car CARFAX links all carry over.

## The search

| Parameter | Value |
|---|---|
| Location | Elkhart, IN 46514, 75 mile radius |
| Models | Toyota 4Runner, Toyota RAV4, Honda HR-V, Honda CR-V, Subaru Outback |
| Max price | $40,000 |
| Min year | 2022 |
| Mileage cap | Toyota 100,000 · Honda 70,000 · Subaru 50,000 |
| Required | AWD/4WD · no accidents · service history · GOOD or GREAT value · heated seats · Apple CarPlay |
| Scored bonus | Certified Pre-Owned |

262 listings matched the model/year/price/radius search. 36 meet every requirement;
6 of those are also Certified Pre-Owned.

## Tabs

| Tab | What it is |
|---|---|
| Start Here | Overview, the filter funnel, how scoring works, what to check before buying |
| Shortlist | The 36 cars meeting every criterion, ranked best first |
| Top 10 Compare | The ten best side by side |
| All Cars | All 262 listings, every data point, filterable |
| Search Criteria | Editable filters that drive the YES/no "Meets All" column |
| Scoring Weights | Editable weights that drive the Score |
| Model Benchmarks | What each model-year actually lists for in this market |

Both editable tabs are live: change a yellow cell and every score, rank and flag
recalculates.

## Scoring

Six sub-scores from 0-100, weighted. The weighting follows the stated priority order --
reliability first, then price, then mileage, then features.

| Component | Weight | Basis |
|---|---|---|
| Reliability & History | 26% | CARFAX model reliability, repair cost and risk (60%) + this car's own history: accident-free, one owner, personal use, service record count (40%) |
| Price / Value | 22% | CARFAX value badge (55%) + how far under the fitted market price it is listed (45%) |
| Mileage | 18% | Miles vs. the cap for that make (70%) + miles/year vs. a 12,000/yr normal (30%) |
| Certified Pre-Owned | 15% | 100 if factory certified, else 0 |
| Features | 14% | Heated steering, sunroof, blind spot, power liftgate, leather, remote start, nav, wireless charging, parking sensors, power seat, premium audio |
| Distance | 5% | 100 at the door, 0 at 75 miles |

"Est. Market" is fitted per model from the 262 listings in this dataset (price regressed
on year and mileage). It is a sanity check on asking price, not an appraisal.

## Keeping it up to date

Two ways, same scoring logic:

**In Google Sheets (no local setup).** Paste `apps_script/Code.gs` into the
spreadsheet once; a **Car Search → Refresh from CARFAX** menu then re-pulls and
rescores in place, preserving your criteria, weights and notes. Setup steps and
caveats: [`apps_script/README.md`](apps_script/README.md).

**Locally.**

```bash
pip install openpyxl
sudo apt-get install -y libreoffice-calc   # only needed for the recalc step
./scripts/run_all.sh
```

Listings change daily. `scripts/1_fetch_listings.py` re-pulls from the CARFAX search API;
edit `TARGETS`, `ZIP` or `RADIUS` at the top of that file to change the search.

## Caveats

- Certification comes from the dealer's feed. Confirm it is manufacturer certified
  (Toyota TCUV, HondaTrue, Subaru Certified) and not an in-house dealer warranty.
- Feature data is the factory equipment list for that trim, so it is reliable for
  standard equipment but can miss optional packages. Confirm on the window sticker.
- Data snapshot: September 14, 2026.
