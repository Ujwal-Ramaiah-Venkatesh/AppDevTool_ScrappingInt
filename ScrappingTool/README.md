# ScrappingTool — Website-Lead Finder

Finds local businesses that likely **need a website** and exports them to an
**Excel (.xlsx)** sheet for outreach.

Two data sources:

- **`osm` (default) — FREE, no API key, no billing.** Uses OpenStreetMap via
  the Overpass API (business listings) + Nominatim (geocoding). OSM data is
  open-licensed, so it's fully legal.
- **`google` — richer data, but needs a billed Google Places API key.**

Searches a **radius around each of your localities** (default: 5 North-East
Bangalore areas, 10 km each), merges and deduplicates the results.

## Output columns

`Business Name, Category, Link to Open, Website Status, Contact Number,
Contact Email, Address / Area, Source, Notes`

Leads are ranked so the most useful are on top: **contactable (has a phone)**
first, then **needs a website** (`Website Status` = `None` / `Social-only`).
The Excel sheet color-codes prospects (green = no site, yellow = social-only)
with a frozen header and filters.

## Setup

```powershell
pip install -r requirements.txt
```

That's it for the free `osm` source — no key needed.

(Optional, only for `google`) Get a Places API key, enable "Places API (New)",
then `Copy-Item .env.example .env` and paste your key into `.env`.

## Run

```powershell
# Uses config.yaml (edit areas, radius_km, target_count there)
python scraper.py

# Override on the fly
python scraper.py --area "Koramangala, Bangalore" --radius 10 --count 300
python scraper.py --source osm        # free (default)
python scraper.py --source google     # needs API key
python scraper.py --no-emails         # faster; skip email scan
```

The sheet is written to `leads.xlsx` (opens directly in Excel). Use a `.csv`
name in `--out` or `output_csv` for plain CSV.

## Notes

- **Never fabricates.** Missing phone/email cells stay blank.
- **Website Status = OK** only means a site exists; judging "outdated /
  not mobile-friendly" needs a manual look.
- **OSM coverage** varies — many small shops have a name + phone but no email.
  The ranking pushes the contactable ones to the top.
- **Overpass** is a shared free service; if an area returns a 429/504, the tool
  retries other mirrors. Re-run later if a busy area is skipped.
- **Outreach law:** Contacting businesses on publicly listed numbers is
  generally low-risk, but cold email/calls are regulated (India DPDP Act,
  GDPR, CAN-SPAM). Always offer an opt-out and don't spam.
