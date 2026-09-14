#!/usr/bin/env bash
# Refresh the whole search end to end. Run from the repo root.
set -euo pipefail
cd "$(dirname "$0")/.."
python3 scripts/1_fetch_listings.py     # pull live CARFAX inventory -> data/carfax_listings_latest.json
python3 scripts/2_score_listings.py     # compute sub-scores          -> data/scored_latest.json
python3 scripts/3a_build_allcars.py     # All Cars tab
python3 scripts/3b_build_views.py       # Shortlist, Top 10, Benchmarks
python3 scripts/3c_build_guide.py       # Start Here
python3 scripts/recalc.py output/Bella_Car_Search.xlsx 400
