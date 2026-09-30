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

# --- Free (no-key) data sources: OpenStreetMap -------------------------------
# Overpass = business POIs; Nominatim = geocoding area names to lat/lng.
OVERPASS_URL = "https://overpass-api.de/api/interpreter"
# Mirrors tried in order when one is overloaded (504) or times out.
OVERPASS_MIRRORS = (
    "https://overpass-api.de/api/interpreter",
    "https://overpass.kumi.systems/api/interpreter",
    "https://maps.mail.ru/osm/tools/overpass/api/interpreter",
)
NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
# Nominatim requires a real User-Agent identifying the app.
USER_AGENT = "ScrappingTool-LeadFinder/1.0 (local business research)"

# OSM tags that identify a "business" worth prospecting, mapped to a label.
OSM_AMENITIES = (
    "restaurant|cafe|fast_food|bar|pub|food_court|ice_cream|"
    "clinic|doctors|dentist|pharmacy|veterinary|"
    "car_rental|car_wash|driving_school|"
    "cinema|nightclub"
)
OSM_TOURISM = "hotel|guest_house|motel|hostel|apartment|resort"

# Amenity/office values that are NOT website-pitch prospects (chains, govt, etc.).
OSM_EXCLUDE_AMENITY = {
    "bank", "atm", "bureau_de_change", "post_office", "townhall", "courthouse",
    "police", "fire_station", "prison", "public_building", "community_centre",
    "school", "college", "university", "kindergarten", "hospital", "fuel",
    "place_of_worship", "toilets", "parking", "bus_station", "fountain",
}
OSM_EXCLUDE_OFFICE = {"government", "administrative", "diplomatic"}

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
    data_source: str = "osm"  # "osm" (free) or "google" (needs API key)
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
        data_source=str(raw.get("data_source", "osm")).strip().lower(),
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
    """Geocode one locality/landmark to (lat, lng) using free Nominatim (OSM)."""
    try:
        resp = requests.get(
            NOMINATIM_URL,
            params={
                "q": area,
                "format": "json",
                "limit": 1,
                "countrycodes": "in",
                # Bias toward the Bangalore metro to avoid same-name mismatches.
                "viewbox": "77.35,13.25,77.85,12.75",
                "bounded": 1,
            },
            headers={"User-Agent": USER_AGENT},
            timeout=30,
        )
    except requests.RequestException as exc:
        print(f"  ! Geocode request failed for '{area}': {exc}", file=sys.stderr)
        return None
    if resp.status_code != 200 or not resp.json():
        print(f"  ! Could not geocode '{area}': {resp.status_code}", file=sys.stderr)
        return None
    top = resp.json()[0]
    # Nominatim asks for max ~1 request/second.
    time.sleep(1)
    return float(top["lat"]), float(top["lon"])


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


def collect_leads_google(cfg: Config) -> list[Lead]:
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


def _osm_category(tags: dict) -> str:
    """Human label for an OSM element from its primary classifying tag."""
    for key in ("shop", "amenity", "tourism", "office", "craft", "healthcare", "leisure"):
        if tags.get(key):
            return str(tags[key]).replace("_", " ")
    return "business"


def _osm_address(tags: dict) -> str:
    parts = [
        tags.get("addr:housenumber", ""),
        tags.get("addr:street", ""),
        tags.get("addr:suburb", ""),
        tags.get("addr:city", ""),
        tags.get("addr:postcode", ""),
    ]
    return ", ".join(p for p in parts if p).strip(", ")


def element_to_lead(el: dict) -> Lead | None:
    tags = el.get("tags", {})
    name = (tags.get("name") or "").strip()
    if not name:
        return None  # unnamed POIs aren't usable leads

    # Skip non-prospects (banks, government, schools, fuel, chains-ish).
    if tags.get("amenity") in OSM_EXCLUDE_AMENITY:
        return None
    if tags.get("office") in OSM_EXCLUDE_OFFICE:
        return None

    lat = el.get("lat") or (el.get("center") or {}).get("lat")
    lon = el.get("lon") or (el.get("center") or {}).get("lon")
    website = (tags.get("website") or tags.get("contact:website") or "").strip()
    phone = (tags.get("phone") or tags.get("contact:phone") or tags.get("contact:mobile") or "").strip()
    email = (tags.get("email") or tags.get("contact:email") or "").strip()
    status, note = classify_website(website)

    link = f"https://www.google.com/maps/search/?api=1&query={lat},{lon}" if lat and lon else ""
    osm_id = f"{el.get('type', 'n')}/{el.get('id', '')}"

    return Lead(
        place_id=osm_id,
        name=name,
        category=_osm_category(tags),
        link=link,
        website_status=status,
        phone=phone,
        email=email,
        address=_osm_address(tags),
        source="OpenStreetMap (Overpass)",
        website=website,
        notes=note,
    )


def overpass_fetch(lat: float, lng: float, radius_m: float) -> list[dict]:
    """Query the free Overpass API for business POIs within a radius.

    Tries each mirror in turn on timeout/504 so one busy server doesn't
    lose the whole area.
    """
    r = int(radius_m)
    query = f"""
    [out:json][timeout:90];
    (
      nwr(around:{r},{lat},{lng})[shop];
      nwr(around:{r},{lat},{lng})[amenity~"^({OSM_AMENITIES})$"];
      nwr(around:{r},{lat},{lng})[tourism~"^({OSM_TOURISM})$"];
      nwr(around:{r},{lat},{lng})[office];
      nwr(around:{r},{lat},{lng})[craft];
      nwr(around:{r},{lat},{lng})[leisure=fitness_centre];
    );
    out center tags;
    """
    for mirror in OVERPASS_MIRRORS:
        try:
            resp = requests.post(
                mirror,
                data={"data": query},
                headers={"User-Agent": USER_AGENT},
                timeout=180,
            )
        except requests.RequestException as exc:
            print(f"  ! {mirror} failed ({exc}); trying next mirror...", file=sys.stderr)
            continue
        if resp.status_code == 200:
            return resp.json().get("elements", [])
        print(f"  ! {mirror} returned {resp.status_code}; trying next mirror...", file=sys.stderr)
        time.sleep(3)
    print("  ! All Overpass mirrors failed for this area.", file=sys.stderr)
    return []


def collect_leads_osm(cfg: Config) -> list[Lead]:
    seen_ids: set[str] = set()
    seen_phones: set[str] = set()
    leads: list[Lead] = []

    centers = resolve_centers(cfg)
    if not centers:
        print("ERROR: OSM mode needs at least one area to center on.", file=sys.stderr)
        return []

    radius_m = min(cfg.radius_km * 1000, MAX_RADIUS_M)
    print(f"Searching {len(centers)} area(s) within {cfg.radius_km:g} km each (OpenStreetMap).")

    for label, lat, lng in centers:
        print(f"Querying Overpass around {label}...")
        elements = overpass_fetch(lat, lng, radius_m)
        print(f"  {len(elements)} POIs returned.")
        for el in elements:
            lead = element_to_lead(el)
            if lead is None or lead.place_id in seen_ids:
                continue
            if lead.phone and lead.phone in seen_phones:
                continue
            seen_ids.add(lead.place_id)
            if lead.phone:
                seen_phones.add(lead.phone)
            leads.append(lead)
        # Be polite to the shared public Overpass instance.
        time.sleep(2)

    return leads


def collect_leads(cfg: Config) -> list[Lead]:
    if cfg.data_source == "google":
        return collect_leads_google(cfg)
    return collect_leads_osm(cfg)


def sort_leads(leads: list[Lead]) -> list[Lead]:
    """Rank actionable, high-intent leads first.

    Priority: has a phone (contactable) > needs a website > has an email.
    """
    status_rank = {"None": 0, "Social-only": 1, "Outdated": 2, "OK": 3}
    return sorted(
        leads,
        key=lambda l: (
            0 if l.phone else 1,
            status_rank.get(l.website_status, 9),
            0 if l.email else 1,
            l.name.lower(),
        ),
    )


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
    parser.add_argument("--source", choices=["osm", "google"], help="Data source (default from config).")
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
    if args.source:
        cfg.data_source = args.source
    if args.no_emails:
        cfg.enrich_emails = False

    # Only the Google source needs an API key; OSM is free and keyless.
    if cfg.data_source == "google" and not cfg.api_key:
        print("ERROR: GOOGLE_MAPS_API_KEY is not set. Copy .env.example to .env and add your key.", file=sys.stderr)
        return 1
    if not cfg.location:
        print("ERROR: No location set in config.yaml.", file=sys.stderr)
        return 1

    leads = collect_leads(cfg)
    print(f"\nCollected {len(leads)} unique businesses (before ranking/cap).")

    # Rank best leads first, then keep only the target count.
    leads = sort_leads(leads)[: cfg.target_count]

    # Email enrichment: fetch each real website once and scan for a public email.
    if cfg.enrich_emails:
        print("Enriching emails from business websites (best-effort)...")
        for lead in leads:
            if lead.website and not lead.email and lead.website_status == "OK":
                lead.email = extract_email(lead.website)

    write_output(leads, cfg.output_csv)
    print(f"Wrote {len(leads)} leads to {cfg.output_csv}")

    # Quick summary so you know how many are strong prospects.
    strong = sum(1 for l in leads if l.website_status in ("None", "Social-only"))
    with_phone = sum(1 for l in leads if l.phone)
    print(f"Strong prospects (no site / social-only): {strong}")
    print(f"Contactable (has phone): {with_phone}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
