#!/usr/bin/env python3
"""
find_emails.py  —  bulk work-email finder for an organisation list.

Reads one Excel workbook that holds everything (the companies to look up, your
API keys, and a few settings), finds the most likely work email for each
organisation (or a named person at it), and writes the results back into the
same workbook. It uses Hunter.io and/or Apollo.io, and when one API key runs
out of credits it automatically rotates to the next key, then to the other
provider.

    Companies sheet   -> what to look up + where results are written
    API Keys sheet    -> one row per key; add as many as you like
    Settings sheet    -> a few knobs (provider order, thresholds, delay)

Typical use
-----------
    pip install openpyxl requests
    python find_emails.py --workbook Aurora_Email_Finder.xlsx

    # try just the first 5 rows:
    python find_emails.py --workbook Aurora_Email_Finder.xlsx --limit 5

    # no network, fills fake data so you can check the plumbing:
    python find_emails.py --workbook Aurora_Email_Finder.xlsx --dry-run

Nothing is hard-coded: every key and setting lives in the workbook, so the
same script serves the funder list, a corporate-sponsor list, or any other
"organisation -> contact email" job.
"""

import argparse
import datetime as dt
import sys
import time

try:
    import requests
except ImportError:
    requests = None

from openpyxl import load_workbook

HUNTER_BASE = "https://api.hunter.io/v2"
APOLLO_BASE = "https://api.apollo.io/api/v1"

# Sheet names (change here if you rename tabs in the workbook)
SHEET_COMPANIES = "Companies"
SHEET_KEYS = "API Keys"
SHEET_SETTINGS = "Settings"

# Column map for the Companies sheet (1-based). Keep in sync with the workbook.
COL = {
    "id": 1, "org": 2, "domain": 3, "first": 4, "last": 5, "role": 6, "mode": 7,
    "out_domain": 8, "out_email": 9, "out_name": 10, "out_position": 11,
    "out_confidence": 12, "out_provider": 13, "out_status": 14,
    "out_other": 15, "out_run": 16,
}
FIRST_DATA_ROW = 3  # row 1 = headers, row 2 = example row


# --------------------------------------------------------------------------- #
#  API key pool with automatic rotation
# --------------------------------------------------------------------------- #
class ApiKey:
    def __init__(self, provider, key, priority, notes=""):
        self.provider = str(provider).strip().lower()
        self.key = str(key).strip()
        self.priority = priority if priority is not None else 999
        self.notes = notes or ""
        self.dead = False
        self.calls = 0

    @property
    def tail(self):
        return self.key[-4:] if len(self.key) >= 4 else self.key


class KeyExhausted(Exception):
    """No usable key left for a provider (all out of credits / invalid)."""


class ProviderError(Exception):
    """A single call failed for a reason that is not the key's fault."""


class KeyPool:
    def __init__(self, keys):
        self.keys = sorted(keys, key=lambda k: k.priority)

    def next_key(self, provider):
        for k in self.keys:
            if k.provider == provider and k.key and not k.dead:
                return k
        return None

    def kill(self, key, reason):
        key.dead = True
        print(f"      key ...{key.tail} ({key.provider}) retired: {reason}")

    def summary(self):
        out = {}
        for k in self.keys:
            out.setdefault(k.provider, []).append(k)
        return out


# --------------------------------------------------------------------------- #
#  Low-level provider calls (each rotates keys on a credit / limit error)
# --------------------------------------------------------------------------- #
def _json(resp):
    try:
        return resp.json()
    except Exception:
        return {}


def _hunter_err(body):
    try:
        return (body.get("errors") or [{}])[0].get("id", "")
    except Exception:
        return ""


def hunter_get(pool, path, params):
    """GET a Hunter endpoint, rotating keys when one is out of credits."""
    last = None
    while True:
        key = pool.next_key("hunter")
        if not key:
            raise KeyExhausted(last or "no active Hunter keys")
        p = dict(params)
        p["api_key"] = key.key
        try:
            r = requests.get(f"{HUNTER_BASE}{path}", params=p, timeout=30)
        except requests.RequestException as e:
            raise ProviderError(f"network error: {e}")
        key.calls += 1
        if r.status_code == 200:
            return _json(r)

        eid = _hunter_err(_json(r))
        # 429 = rate limit (short) OR monthly usage cap. Back off once for a
        # plain rate limit, then retry the same key before giving up on it.
        if r.status_code == 429 and eid != "usage_limit_reached":
            time.sleep(2.0)
            try:
                r = requests.get(f"{HUNTER_BASE}{path}", params=p, timeout=30)
                if r.status_code == 200:
                    return _json(r)
                eid = _hunter_err(_json(r))
            except requests.RequestException as e:
                raise ProviderError(f"network error: {e}")
        last = f"HTTP {r.status_code} {eid}".strip()
        # 401 invalid, 402 payment required, 403 forbidden, 429 usage cap -> next key
        pool.kill(key, last)


def apollo_post(pool, path, payload):
    """POST an Apollo endpoint, rotating keys when one is out of credits."""
    last = None
    while True:
        key = pool.next_key("apollo")
        if not key:
            raise KeyExhausted(last or "no active Apollo keys")
        headers = {
            "X-Api-Key": key.key,
            "Content-Type": "application/json",
            "Cache-Control": "no-cache",
        }
        try:
            r = requests.post(f"{APOLLO_BASE}{path}", json=payload, headers=headers, timeout=30)
        except requests.RequestException as e:
            raise ProviderError(f"network error: {e}")
        key.calls += 1
        if r.status_code == 200:
            return _json(r)
        if r.status_code == 429:  # rate limited: back off once, retry same key
            time.sleep(2.0)
            try:
                r = requests.post(f"{APOLLO_BASE}{path}", json=payload, headers=headers, timeout=30)
                if r.status_code == 200:
                    return _json(r)
            except requests.RequestException as e:
                raise ProviderError(f"network error: {e}")
        last = f"HTTP {r.status_code}"
        pool.kill(key, last)  # 401/402/403/429 cap -> next key


# --------------------------------------------------------------------------- #
#  Normalised lookups
# --------------------------------------------------------------------------- #
def resolve_domain(pool, org, provider_order):
    """Turn a company NAME into a domain. Hunter can do this in one call and
    also hand back the emails, so we prefer it and cache that payload."""
    if "hunter" in provider_order and pool.next_key("hunter"):
        data = hunter_get(pool, "/domain-search", {"company": org, "limit": 25}).get("data", {})
        if data.get("domain"):
            return data["domain"], data  # (domain, hunter payload we can reuse)
    if "apollo" in provider_order and pool.next_key("apollo"):
        res = apollo_post(pool, "/organizations/search", {"q_organization_name": org, "per_page": 1})
        orgs = res.get("organizations") or res.get("accounts") or []
        if orgs:
            dom = orgs[0].get("primary_domain") or orgs[0].get("website_url") or ""
            dom = dom.replace("https://", "").replace("http://", "").replace("www.", "").strip("/")
            if dom:
                return dom, None
    return "", None


def pick_from_hunter_emails(emails, role_hint, min_conf):
    """Choose the best email from a Hunter domain-search 'emails' list."""
    if not emails:
        return None, []
    hint = (role_hint or "").lower().strip()

    def score(e):
        s = e.get("confidence") or 0
        if hint:
            hay = " ".join(str(e.get(f) or "") for f in ("position", "department", "seniority")).lower()
            if hint in hay:
                s += 1000  # force hinted contacts to the top
        if e.get("first_name") or e.get("last_name"):
            s += 5  # a named person beats a generic inbox at equal confidence
        return s

    ordered = sorted(emails, key=score, reverse=True)
    best = ordered[0]
    others = [e.get("value") for e in ordered[1:] if e.get("value")]
    return best, others


def lookup_row(pool, row, settings):
    """Return a result dict for one input row."""
    org = (row.get("org") or "").strip()
    domain = (row.get("domain") or "").strip().replace("https://", "").replace("http://", "").replace("www.", "").strip("/")
    first = (row.get("first") or "").strip()
    last = (row.get("last") or "").strip()
    role = (row.get("role") or "").strip()
    mode = (row.get("mode") or settings["default_mode"]).strip().lower()
    order = settings["provider_order"]
    min_conf = settings["min_confidence"]

    if mode == "auto":
        mode = "person" if (first and last) else "domain"

    cached_hunter = None
    if not domain and org:
        domain, cached_hunter = resolve_domain(pool, org, order)
    if not domain:
        return {"status": "no_domain", "provider": "", "confidence": "",
                "email": "", "name": "", "position": "", "domain": "", "other": ""}

    # ---- person mode: find one specific person's email ----
    if mode == "person" and first and last:
        for prov in order:
            try:
                if prov == "hunter" and pool.next_key("hunter"):
                    d = hunter_get(pool, "/email-finder",
                                   {"domain": domain, "first_name": first, "last_name": last}).get("data", {})
                    if d.get("email"):
                        conf = d.get("score") or 0
                        return {"status": "found" if conf >= min_conf else "found_low_confidence",
                                "provider": "hunter", "confidence": conf, "email": d["email"],
                                "name": f"{first} {last}".strip(),
                                "position": d.get("position") or role, "domain": domain, "other": ""}
                if prov == "apollo" and pool.next_key("apollo"):
                    payload = {"first_name": first, "last_name": last,
                               "domain": domain, "organization_name": org,
                               "reveal_personal_emails": False}
                    person = (apollo_post(pool, "/people/match", payload) or {}).get("person", {}) or {}
                    email = person.get("email") or ""
                    if email and "not_unlocked" not in email and "email_not" not in email:
                        return {"status": "found", "provider": "apollo", "confidence": "",
                                "email": email, "name": f"{first} {last}".strip(),
                                "position": person.get("title") or role, "domain": domain, "other": ""}
            except KeyExhausted:
                continue
            except ProviderError as e:
                print(f"      {prov} error: {e}")
                continue
        # fall through to domain mode if the person could not be found
        mode = "domain"

    # ---- domain mode: pull the contacts at the company ----
    for prov in order:
        try:
            if prov == "hunter" and (cached_hunter or pool.next_key("hunter")):
                data = cached_hunter or hunter_get(pool, "/domain-search", {"domain": domain, "limit": 25}).get("data", {})
                best, others = pick_from_hunter_emails(data.get("emails", []), role, min_conf)
                if best and best.get("value"):
                    conf = best.get("confidence") or 0
                    name = (str(best.get("first_name") or "") + " " + str(best.get("last_name") or "")).strip()
                    return {"status": "found" if conf >= min_conf else "found_low_confidence",
                            "provider": "hunter", "confidence": conf, "email": best["value"],
                            "name": name, "position": best.get("position") or "",
                            "domain": domain, "other": "; ".join(others[: settings["max_emails"]])}
            if prov == "apollo" and pool.next_key("apollo"):
                payload = {"q_organization_domains": domain, "page": 1, "per_page": settings["max_emails"]}
                people = (apollo_post(pool, "/mixed_people/search", payload) or {}).get("people", []) or []
                hits = [p for p in people if p.get("email") and "not_unlocked" not in (p.get("email") or "")]
                if hits:
                    top = hits[0]
                    return {"status": "found", "provider": "apollo", "confidence": "",
                            "email": top["email"],
                            "name": (str(top.get("first_name") or "") + " " + str(top.get("last_name") or "")).strip(),
                            "position": top.get("title") or "", "domain": domain,
                            "other": "; ".join(p["email"] for p in hits[1:])}
        except KeyExhausted:
            continue
        except ProviderError as e:
            print(f"      {prov} error: {e}")
            continue

    return {"status": "not_found", "provider": "", "confidence": "",
            "email": "", "name": "", "position": "", "domain": domain, "other": ""}


def dry_lookup(row, settings):
    """Offline stand-in so you can test the workbook wiring without keys."""
    org = (row.get("org") or "").strip()
    domain = (row.get("domain") or "").strip() or (org.lower().replace(" ", "") + ".com.au" if org else "")
    first, last = (row.get("first") or "").strip(), (row.get("last") or "").strip()
    if first and last:
        email = f"{first}.{last}@{domain}".lower()
        name = f"{first} {last}"
    else:
        email, name = f"info@{domain}", ""
    return {"status": "found", "provider": "dry-run", "confidence": 88, "email": email,
            "name": name, "position": row.get("role") or "", "domain": domain, "other": ""}


# --------------------------------------------------------------------------- #
#  Workbook I/O
# --------------------------------------------------------------------------- #
def read_settings(ws):
    s = {"provider_order": ["hunter", "apollo"], "default_mode": "auto",
         "max_emails": 5, "min_confidence": 50, "delay": 0.3, "only_empty": True}
    for r in range(1, ws.max_row + 1):
        k = ws.cell(r, 1).value
        v = ws.cell(r, 2).value
        if not k:
            continue
        k = str(k).strip().lower()
        if k == "provider_order" and v:
            s["provider_order"] = [x.strip().lower() for x in str(v).split(",") if x.strip()]
        elif k == "default_mode" and v:
            s["default_mode"] = str(v).strip().lower()
        elif k == "max_emails_per_company" and v is not None:
            s["max_emails"] = int(v)
        elif k == "min_confidence" and v is not None:
            s["min_confidence"] = float(v)
        elif k == "rate_limit_delay_seconds" and v is not None:
            s["delay"] = float(v)
        elif k == "only_empty_rows" and v is not None:
            s["only_empty"] = str(v).strip().lower() in ("yes", "true", "1", "y")
    return s


def read_keys(ws):
    keys = []
    for r in range(3, ws.max_row + 1):  # row 1 = title, row 2 = headers, data from row 3
        provider = ws.cell(r, 1).value
        key = ws.cell(r, 2).value
        priority = ws.cell(r, 3).value
        enabled = ws.cell(r, 4).value
        notes = ws.cell(r, 5).value
        if not provider or not key:
            continue
        if str(enabled).strip().lower() in ("no", "false", "0", "n"):
            continue
        if str(key).strip().upper().startswith("PASTE"):
            continue  # placeholder row, skip
        keys.append(ApiKey(provider, key, priority, notes))
    return keys


def main():
    ap = argparse.ArgumentParser(description="Bulk work-email finder (Hunter + Apollo) driven by an Excel workbook.")
    ap.add_argument("--workbook", required=True, help="path to the .xlsx workbook")
    ap.add_argument("--limit", type=int, default=0, help="only process the first N pending rows")
    ap.add_argument("--dry-run", action="store_true", help="no network calls; fill plausible fake data")
    args = ap.parse_args()

    if not args.dry_run and requests is None:
        sys.exit("The 'requests' package is needed for live runs. Install it with:  pip install requests")

    wb = load_workbook(args.workbook)
    for name in (SHEET_COMPANIES, SHEET_KEYS, SHEET_SETTINGS):
        if name not in wb.sheetnames:
            sys.exit(f"Workbook is missing the '{name}' sheet.")
    ws = wb[SHEET_COMPANIES]
    settings = read_settings(wb[SHEET_SETTINGS])

    pool = None
    if not args.dry_run:
        pool = KeyPool(read_keys(wb[SHEET_KEYS]))
        if not pool.keys:
            sys.exit("No usable API keys found on the 'API Keys' sheet. Add at least one, or use --dry-run.")
        have = {p: len(v) for p, v in pool.summary().items()}
        print(f"Loaded keys: {have}")

    processed = found = 0
    for r in range(FIRST_DATA_ROW, ws.max_row + 1):
        org = ws.cell(r, COL["org"]).value
        idv = ws.cell(r, COL["id"]).value
        if not org:
            continue
        if "EXAMPLE" in (str(idv) + " " + str(org)).upper():
            continue  # template/example row, never processed
        if settings["only_empty"] and ws.cell(r, COL["out_email"]).value:
            continue  # already done on a previous run

        row = {"id": ws.cell(r, COL["id"]).value, "org": org,
               "domain": ws.cell(r, COL["domain"]).value, "first": ws.cell(r, COL["first"]).value,
               "last": ws.cell(r, COL["last"]).value, "role": ws.cell(r, COL["role"]).value,
               "mode": ws.cell(r, COL["mode"]).value}
        label = str(org).strip()
        print(f"[{r}] {label} ...", end=" ")

        try:
            res = dry_lookup(row, settings) if args.dry_run else lookup_row(pool, row, settings)
        except KeyExhausted as e:
            print(f"\nStopped: every API key is out of credits ({e}). Progress saved.")
            break

        ws.cell(r, COL["out_domain"]).value = res["domain"]
        ws.cell(r, COL["out_email"]).value = res["email"]
        ws.cell(r, COL["out_name"]).value = res["name"]
        ws.cell(r, COL["out_position"]).value = res["position"]
        ws.cell(r, COL["out_confidence"]).value = res["confidence"]
        ws.cell(r, COL["out_provider"]).value = res["provider"]
        ws.cell(r, COL["out_status"]).value = res["status"]
        ws.cell(r, COL["out_other"]).value = res["other"]
        ws.cell(r, COL["out_run"]).value = dt.datetime.utcnow().strftime("%Y-%m-%d %H:%M")
        print(f"{res['status']:>20}  {res['email']}")

        processed += 1
        if res["email"]:
            found += 1
        wb.save(args.workbook)  # save after every row so nothing is lost
        if args.limit and processed >= args.limit:
            break
        if not args.dry_run:
            time.sleep(settings["delay"])

    wb.save(args.workbook)
    print(f"\nDone. {processed} rows processed, {found} emails found. Saved to {args.workbook}")
    if pool:
        for prov, ks in pool.summary().items():
            live = sum(1 for k in ks if not k.dead)
            print(f"  {prov}: {live}/{len(ks)} keys still live, {sum(k.calls for k in ks)} calls made")


if __name__ == "__main__":
    main()
