# Email finder — Hunter.io + Apollo, driven by one Excel workbook

Finds the most likely work email for each organisation on a list, using Hunter.io and/or
Apollo.io. When one API key runs out of credits it rotates to the next key, then to the
other provider. Everything lives in **`Aurora_Email_Finder.xlsx`** — keys, inputs, settings,
and results — so the script itself never needs editing.

## Setup (once)

```
pip install openpyxl requests
```

## Run it

```
# 1. Safe first: fake data, no keys used, proves the wiring works
python find_emails.py --workbook Aurora_Email_Finder.xlsx --dry-run

# 2. Live, but just the first 5 rows while you sanity-check
python find_emails.py --workbook Aurora_Email_Finder.xlsx --limit 5

# 3. The whole list
python find_emails.py --workbook Aurora_Email_Finder.xlsx
```

Results write into the green columns and save after **every row**, so you can stop and
re-run any time — it skips rows that already have an email.

## The workbook, three tabs

**API Keys** — one row per key. Paste as many as you have. `Priority 1` is used first; when it
hits its credit cap the script retires it and moves to `Priority 2`, then Apollo. Set
`Enabled` to `no` to park a key.

| Provider | API Key | Priority | Enabled |
|---|---|---|---|
| hunter | your_key_1 | 1 | yes |
| hunter | your_key_2 | 2 | yes |
| apollo | your_key_3 | 3 | yes |

**Companies** — paste your list into the yellow columns. `Organisation` is the only required one.

- Already know the website? Put it in `Website / Domain` — skips the name→domain lookup and saves a credit.
- Chasing a named person? Add `First name` + `Last name`, set `Lookup mode` to `person`.
- Just want the best inbox at a company? Leave the name blank (`domain` mode).
- `Role hint` (e.g. `marketing`, `CEO`) pushes domain mode toward the right contact.

**Settings** — provider order, confidence threshold, delay between calls, and whether to skip
already-filled rows. Sensible defaults are in place.

## How key rotation works

Each call checks the HTTP response. `401` (bad key), `402` (payment/credits), `403`, and a
`429` monthly usage cap all retire that key and the script moves to the next one. A plain
`429` rate-limit just waits two seconds and retries the same key. If every key for both
providers is spent, it stops cleanly with progress saved.

## Which input list to use

The **funder / organisation list** is the right input — paste its `Organisation` column into
column B. These APIs find work emails from a **company domain**, so the **wealthy-locals
letterbox list won't work here** (home addresses, no company) — that list is for the physical
flyer drop. If any of those locals are business owners, add their name plus their company and
the script can look them up by name.

## Please use it responsibly

Respect Hunter's and Apollo's terms, only approach organisations that welcome contact, and
follow Australia's Spam Act 2003: identify yourself and give an easy unsubscribe in every
email. This finds business contacts for legitimate outreach — keep it to that.
