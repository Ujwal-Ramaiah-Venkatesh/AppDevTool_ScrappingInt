# ScrappingTool — Website-Lead Finder

Finds local businesses that likely **need a website** and exports them to an
**Excel (.xlsx)** sheet you can use for outreach. Uses the **official Google
Places API (New)** — Google's supported, Terms-compliant way to query business
listings (no HTML scraping, so you won't get IP-banned).

Searches within a **radius around your locality** (default: Bangalore,
Indiranagar, 10 km) and can widen later.

## Output columns

`Business Name, Category, Link to Open, Website Status, Contact Number,
Contact Email, Address / Area, Source, Notes`

Rows are sorted with the **best prospects on top** (`Website Status` = `None`
or `Social-only`), and the Excel sheet color-codes them (green = no site,
yellow = social-only) with a frozen header and filters.

## Setup

1. **Get a Google API key** and enable **"Places API (New)"**:
   https://developers.google.com/maps/documentation/places/web-service/get-api-key
2. Copy the key into a `.env` file:
   ```powershell
   Copy-Item .env.example .env
   # then edit .env and paste your key
   ```
3. Install dependencies:
   ```powershell
   pip install -r requirements.txt
   ```

## Run

```powershell
# Uses config.yaml (edit area, radius_km, categories, target_count there)
python scraper.py

# Or override on the fly — center the 10 km radius on your locality
python scraper.py --area "Koramangala, Bangalore" --radius 10 --count 200
python scraper.py --out leads.xlsx
python scraper.py --no-emails
```

The sheet is written to `leads.xlsx` (opens directly in Excel). Use a `.csv`
name in `--out` or `output_csv` if you prefer plain CSV.

## Important notes

- **Emails:** The Places API does **not** return emails. When `enrich_emails`
  is on, the tool fetches each business's own public website and scans for a
  visible email. It **never invents** an address — blanks stay blank.
- **Website Status = OK** just means a site exists. Detecting "outdated /
  not mobile-friendly" reliably needs a manual look — that's your judgment call.
- **Cost:** Places API is pay-as-you-go with a monthly free tier. 200 leads is
  well within typical free credits, but check your Google Cloud billing.
- **Outreach law:** Contacting businesses on publicly listed numbers is
  generally low-risk, but cold email/calls are regulated (India DPDP Act,
  GDPR, CAN-SPAM). Always offer an opt-out and don't spam.
