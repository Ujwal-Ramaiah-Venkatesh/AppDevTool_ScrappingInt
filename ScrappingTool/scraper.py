"""
Lead-generation scraper for local businesses that may need a website.

Uses the official Google Places API (New) — this is Google's supported,
Terms-compliant way to query business listings. It does NOT scrape Google
Maps HTML (which violates Google's ToS and gets you blocked).

Output: a CSV with the exact columns requested, best prospects on top.

Usage:
    python scraper.py                       # uses config.yaml
    python scraper.py --location "Kochi"    # override location
    python scraper.py --area "Indiranagar, Bangalore" --radius 10
    python scraper.py --no-emails           # skip email enrichment
"""

from __future__ import annotations

import argparse
import csv
import os
import re
import sys
import time
from dataclasses import dataclass, field
from typing import Iterable

import requests
import yaml
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

PLACES_SEARCH_URL = "https://places.googleapis.com/v1/places:searchText"

# Max radius the Places API allows for a circular restriction (meters).
MAX_RADIUS_M = 50000

# Fields we ask Google to return. Keeping this tight keeps the API cheap.
FIELD_MASK = ",".join(
    [
        "places.id",
        "places.displayName",
        "places.formattedAddress",
        "places.websiteUri",
        "places.nationalPhoneNumber",
        "places.internationalPhoneNumber",
        "places.googleMapsUri",
        "places.primaryTypeDisplayName",
        "nextPageToken",
    ]
)

# Hosts that indicate the business has only a social page, not a real website.
SOCIAL_HOSTS = (
    "facebook.com",
    "fb.com",
    "instagram.com",
    "linktr.ee",
    "wa.me",
    "whatsapp.com",
    "linkedin.com",
    "twitter.com",
    "x.com",
    "youtube.com",
    "t.me",
)

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

# Emails we never want (tracking pixels, asset filenames, sample text).
EMAIL_BLOCKLIST = ("example.com", "sentry.io", "wixpress.com", ".png", ".jpg", ".gif")

CSV_COLUMNS = [
    "Business Name",
    "Category",
    "Link to Open",
    "Website Status",
    "Contact Number",
    "Contact Email",
    "Address / Area",
    "Source",
    "Notes",
]


@dataclass
class Lead:
    place_id: str
    name: str
    category: str
    link: str
    website_status: str
    phone: str
    email: str
    address: str
    source: str
    website: str = ""  # kept for enrichment; not written to CSV
    notes: str = ""

    def as_row(self) -> dict[str, str]:
        return {
            "Business Name": self.name,
            "Category": self.category,
            "Link to Open": self.link,
            "Website Status": self.website_status,
            "Contact Number": self.phone,
            "Contact Email": self.email,
            "Address / Area": self.address,
            "Source": self.source,
            "Notes": self.notes,
        }


@dataclass
class Config:
    location: str
    target_count: int
    categories: list[str]
    enrich_emails: bool
    output_csv: str
    area: str = ""
    areas: list[str] = field(default_factory=list)
    center_lat: float | None = None
    center_lng: float | None = None
    radius_km: float = 10.0
    api_key: str = field(repr=False, default="")


def load_config(path: str) -> Config:
    load_dotenv()
    api_key = os.getenv("GOOGLE_MAPS_API_KEY", "").strip()

    with open(path, "r", encoding="utf-8") as fh:
        raw = yaml.safe_load(fh) or {}

    return Config(
        location=str(raw.get("location", "")).strip(),
        target_count=int(raw.get("target_count", 200)),
        categories=list(raw.get("categories", [])),
        enrich_emails=bool(raw.get("enrich_emails", True)),
        output_csv=str(raw.get("output_csv", "leads.csv")).strip(),
        area=str(raw.get("area", "")).strip(),
        areas=[str(a).strip() for a in (raw.get("areas") or []) if str(a).strip()],
        center_lat=raw.get("center_lat"),
        center_lng=raw.get("center_lng"),
        radius_km=float(raw.get("radius_km", 10.0)),
        api_key=api_key,
    )


def classify_website(website: str) -> tuple[str, str]:
    """Return (status, note) for a business's website URL."""
    if not website:
        return "None", "No website listed — strong candidate."
    host = website.lower()
    if any(social in host for social in SOCIAL_HOSTS):
        return "Social-only", "Only a social page — strong candidate."
    return "OK", "Has a website — verify if outdated/mobile-friendly."


def geocode_area(area: str, cfg: Config) -> tuple[float, float] | None:
    """Geocode one locality/landmark string to (lat, lng) via Places search."""
    headers = {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": cfg.api_key,
        "X-Goog-FieldMask": "places.location,places.formattedAddress",
    }
    resp = requests.post(
        PLACES_SEARCH_URL,
        headers=headers,
        json={"textQuery": area, "pageSize": 1},
        timeout=30,
    )
    if resp.status_code != 200:
        print(f"  ! Could not geocode '{area}': {resp.status_code} {resp.text[:150]}", file=sys.stderr)
        return None
    places = resp.json().get("places", [])
    if not places:
        return None
    loc = places[0].get("location", {})
    lat, lng = loc.get("latitude"), loc.get("longitude")
    if lat is None or lng is None:
        return None
    return float(lat), float(lng)


def resolve_centers(cfg: Config) -> list[tuple[str, float, float]]:
    """Resolve all search centers as (label, lat, lng).

    Priority: explicit center_lat/center_lng > `areas` list > single `area`.
    Returns [] when nothing is configured (falls back to text-only search).
    """
    if cfg.center_lat is not None and cfg.center_lng is not None:
        return [("custom", float(cfg.center_lat), float(cfg.center_lng))]

    area_list = cfg.areas or ([cfg.area] if cfg.area else [])
    centers: list[tuple[str, float, float]] = []
    for area in area_list:
        coords = geocode_area(area, cfg)
        if coords is None:
            print(f"  ! Skipping '{area}' — could not geocode.", file=sys.stderr)
            continue
        print(f"Center: {area} ({coords[0]:.5f}, {coords[1]:.5f})")
        centers.append((area, coords[0], coords[1]))
    return centers


def search_places(query: str, cfg: Config, remaining: int, center: tuple[float, float] | None) -> Iterable[dict]:
    """Yield place dicts for one text query, following pagination."""
    headers = {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": cfg.api_key,
        "X-Goog-FieldMask": FIELD_MASK,
    }
    radius_m = min(cfg.radius_km * 1000, MAX_RADIUS_M)
    page_token = None
    fetched = 0

    while fetched < remaining:
        body: dict[str, object] = {"textQuery": query, "pageSize": 20}
        if center is not None:
            # Restrict results to a circle around the user's location.
            body["locationRestriction"] = {
                "circle": {
                    "center": {"latitude": center[0], "longitude": center[1]},
                    "radius": radius_m,
                }
            }
        if page_token:
            body["pageToken"] = page_token

        resp = requests.post(PLACES_SEARCH_URL, headers=headers, json=body, timeout=30)
        if resp.status_code != 200:
            print(f"  ! API error {resp.status_code} for '{query}': {resp.text[:200]}", file=sys.stderr)
            return

        data = resp.json()
        for place in data.get("places", []):
            yield place
            fetched += 1
            if fetched >= remaining:
                return

        page_token = data.get("nextPageToken")
        if not page_token:
            return
        # New page tokens need a brief delay before they become valid.
        time.sleep(2)


def extract_email(website: str) -> str:
    """Best-effort public email from a business homepage. Never fabricates."""
    try:
        resp = requests.get(
            website,
            timeout=12,
            headers={"User-Agent": "Mozilla/5.0 (lead-research; contact-scan)"},
        )
    except requests.RequestException:
        return ""
    if resp.status_code != 200:
        return ""

    for match in EMAIL_RE.findall(resp.text):
        candidate = match.strip().lower()
        if not any(bad in candidate for bad in EMAIL_BLOCKLIST):
            return candidate
    return ""


def place_to_lead(place: dict, category: str) -> Lead:
    name = (place.get("displayName") or {}).get("text", "").strip()
    website = (place.get("websiteUri") or "").strip()
    status, note = classify_website(website)
    phone = (place.get("nationalPhoneNumber") or place.get("internationalPhoneNumber") or "").strip()
    api_category = (place.get("primaryTypeDisplayName") or {}).get("text", "").strip()

    return Lead(
        place_id=place.get("id", ""),
        name=name,
        category=api_category or category,
        link=place.get("googleMapsUri", ""),
        website_status=status,
        phone=phone,
        email="",  # filled during enrichment
        address=place.get("formattedAddress", "").strip(),
        source="Google Places API",
        website=website,
        notes=note,
    )


def collect_leads(cfg: Config) -> list[Lead]:
    seen_ids: set[str] = set()
    seen_phones: set[str] = set()
    leads: list[Lead] = []

    centers = resolve_centers(cfg)
    if not centers:
        print("No center/area configured — searching by text only (no radius filter).")
        centers = [("", None, None)]  # sentinel: text-only search
    else:
        print(f"Searching {len(centers)} area(s) within {cfg.radius_km:g} km each.")

    # Spread the target evenly across every area × category combination.
    combos = max(1, len(centers) * len(cfg.categories))
    per_query = max(5, cfg.target_count // combos + 5)

    for label, lat, lng in centers:
        if len(leads) >= cfg.target_count:
            break
        center = (lat, lng) if lat is not None and lng is not None else None
        for category in cfg.categories:
            if len(leads) >= cfg.target_count:
                break
            where = label or cfg.location
            query = f"{category} in {where}"
            print(f"Searching: {query}")
            for place in search_places(query, cfg, per_query, center):
                pid = place.get("id", "")
                if not pid or pid in seen_ids:
                    continue
                lead = place_to_lead(place, category)
                # Dedupe on phone as well as place id.
                if lead.phone and lead.phone in seen_phones:
                    continue
                seen_ids.add(pid)
                if lead.phone:
                    seen_phones.add(lead.phone)
                leads.append(lead)

    return leads


def sort_leads(leads: list[Lead]) -> list[Lead]:
    priority = {"None": 0, "Social-only": 1, "Outdated": 2, "OK": 3}
    return sorted(leads, key=lambda l: priority.get(l.website_status, 9))


def write_csv(leads: list[Lead], path: str) -> None:
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=CSV_COLUMNS)
        writer.writeheader()
        for lead in leads:
            writer.writerow(lead.as_row())


# Row highlight colors by how strong the lead is.
STATUS_FILL = {
    "None": "C6EFCE",         # green  — best prospect
    "Social-only": "FFEB9C",  # yellow — good prospect
    "Outdated": "FCE4D6",     # orange
}


def write_xlsx(leads: list[Lead], path: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Leads"

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4472C4")
    for col_idx, name in enumerate(CSV_COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=name)
        cell.font = header_font
        cell.fill = header_fill

    for row in leads:
        data = row.as_row()
        ws.append([data[c] for c in CSV_COLUMNS])
        fill_hex = STATUS_FILL.get(row.website_status)
        if fill_hex:
            ws.cell(row=ws.max_row, column=4).fill = PatternFill("solid", fgColor=fill_hex)

    # Freeze header, enable autofilter, and set sensible column widths.
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(CSV_COLUMNS))}{ws.max_row}"
    widths = [28, 18, 40, 14, 18, 30, 45, 18, 40]
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width

    wb.save(path)


def write_output(leads: list[Lead], path: str) -> None:
    if path.lower().endswith((".xlsx", ".xls")):
        write_xlsx(leads, path)
    else:
        write_csv(leads, path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Scrape local-business leads via Google Places API.")
    parser.add_argument("--config", default="config.yaml")
    parser.add_argument("--location", help="Override the location from config.")
    parser.add_argument("--area", help="Locality/landmark to center the radius on (geocoded).")
    parser.add_argument("--radius", type=float, help="Radius in km around the center.")
    parser.add_argument("--count", type=int, help="Override target lead count.")
    parser.add_argument("--out", help="Output file (.xlsx or .csv).")
    parser.add_argument("--no-emails", action="store_true", help="Skip email enrichment.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    cfg = load_config(args.config)

    if args.location:
        cfg.location = args.location
    if args.area:
        cfg.areas = [args.area]
    if args.radius:
        cfg.radius_km = args.radius
    if args.count:
        cfg.target_count = args.count
    if args.out:
        cfg.output_csv = args.out
    if args.no_emails:
        cfg.enrich_emails = False

    if not cfg.api_key:
        print("ERROR: GOOGLE_MAPS_API_KEY is not set. Copy .env.example to .env and add your key.", file=sys.stderr)
        return 1
    if not cfg.location:
        print("ERROR: No location set in config.yaml.", file=sys.stderr)
        return 1

    leads = collect_leads(cfg)
    print(f"\nCollected {len(leads)} unique businesses.")

    # Email enrichment: fetch each real website once and scan for a public email.
    if cfg.enrich_emails:
        print("Enriching emails from business websites (best-effort)...")
        for lead in leads:
            if lead.website and lead.website_status == "OK":
                lead.email = extract_email(lead.website)

    leads = sort_leads(leads)
    write_output(leads, cfg.output_csv)
    print(f"Wrote {len(leads)} leads to {cfg.output_csv}")

    # Quick summary so you know how many are strong prospects.
    strong = sum(1 for l in leads if l.website_status in ("None", "Social-only"))
    print(f"Strong prospects (no site / social-only): {strong}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
