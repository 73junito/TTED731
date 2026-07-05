import csv
import json
import time
import urllib.parse
import urllib.request
from pathlib import Path

ROOT = Path(r"A:\TTED731")
A3 = ROOT / "03_Assignment3_Instructional_Design"
INPUT = A3 / "Reference Verification" / "assignment3_reference_targets.csv"
OUTPUT = A3 / "Reference Verification" / "crossref_verification_results.csv"
JSON_OUT = A3 / "Reference Verification" / "crossref_verification_results.json"

MAILTO = "73junito@gmail.com"

def get_json(url):
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": f"TTED731ReferenceVerifier/1.0 (mailto:{MAILTO})"
        }
    )
    with urllib.request.urlopen(req, timeout=30) as response:
        return json.loads(response.read().decode("utf-8"))

def lookup_by_doi(doi):
    url = f"https://api.crossref.org/works/{urllib.parse.quote(doi)}"
    data = get_json(url)
    return data.get("message", {})

def lookup_by_title(title):
    query = urllib.parse.urlencode({
        "query.title": title,
        "rows": 1
    })
    url = f"https://api.crossref.org/works?{query}"
    data = get_json(url)
    items = data.get("message", {}).get("items", [])
    return items[0] if items else {}

def authors_to_text(authors):
    names = []
    for a in authors or []:
        given = a.get("given", "")
        family = a.get("family", "")
        full = f"{family}, {given}".strip(", ")
        if full:
            names.append(full)
    return "; ".join(names)

def get_year(item):
    for field in ["published-print", "published-online", "published"]:
        parts = item.get(field, {}).get("date-parts", [])
        if parts and parts[0]:
            return parts[0][0]
    return ""

rows = []
all_items = {}

with INPUT.open(newline="", encoding="utf-8-sig") as f:
    reader = csv.DictReader(f)
    for row in reader:
        key = row["key"]
        target_title = row["title"]
        target_year = row["year"]
        target_doi = row["doi"].strip()

        try:
            if target_doi:
                item = lookup_by_doi(target_doi)
                method = "doi"
            else:
                item = lookup_by_title(target_title)
                method = "title"

            found_title = (item.get("title") or [""])[0]
            found_doi = item.get("DOI", "")
            found_year = get_year(item)
            container = (item.get("container-title") or [""])[0]
            volume = item.get("volume", "")
            issue = item.get("issue", "")
            page = item.get("page", "")
            url = item.get("URL", "")
            authors = authors_to_text(item.get("author", []))

            title_match = target_title.lower().replace(":", "").replace("-", " ")[:40] in found_title.lower().replace(":", "").replace("-", " ")
            year_match = str(target_year) == str(found_year)
            doi_match = target_doi.lower() == found_doi.lower() if target_doi else ""

            status = "VERIFIED" if (doi_match or title_match) and (not target_year or year_match) else "CHECK"

            rows.append({
                "key": key,
                "status": status,
                "method": method,
                "target_title": target_title,
                "found_title": found_title,
                "target_year": target_year,
                "found_year": found_year,
                "target_doi": target_doi,
                "found_doi": found_doi,
                "journal": container,
                "volume": volume,
                "issue": issue,
                "pages": page,
                "authors": authors,
                "url": url
            })

            all_items[key] = item
            print(f"{status}: {key} -> {found_title}")

        except Exception as e:
            rows.append({
                "key": key,
                "status": "ERROR",
                "method": "doi" if target_doi else "title",
                "target_title": target_title,
                "found_title": "",
                "target_year": target_year,
                "found_year": "",
                "target_doi": target_doi,
                "found_doi": "",
                "journal": "",
                "volume": "",
                "issue": "",
                "pages": "",
                "authors": "",
                "url": "",
                "error": str(e)
            })
            print(f"ERROR: {key} -> {e}")

        time.sleep(1)

fieldnames = [
    "key","status","method","target_title","found_title","target_year","found_year",
    "target_doi","found_doi","journal","volume","issue","pages","authors","url"
]

with OUTPUT.open("w", newline="", encoding="utf-8") as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
    writer.writeheader()
    writer.writerows(rows)

with JSON_OUT.open("w", encoding="utf-8") as f:
    json.dump(all_items, f, indent=2)

print(f"\nSaved CSV: {OUTPUT}")
print(f"Saved JSON: {JSON_OUT}")
