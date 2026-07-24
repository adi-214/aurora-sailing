# Email finder — Hunter.io + Apollo, driven by one Excel workbook

Finds the most likely work email for each organisation on a list. Everything lives in
**`Aurora_Email_Finder.xlsx`** — your keys, your inputs, your settings, and the results — so
you never have to edit the Python.

Download both files, paste your API key into the workbook, run one command.

## Setup (once)

```
pip install openpyxl requests
```

Python 3.8 or newer. A free [Hunter.io](https://hunter.io/users/sign_up) account is enough to
get started — it gives you 50 search credits a month and full API access.

## Run it

**1. Check your key first.** This is free, uses zero credits, and tells you your plan and
exactly how many searches you have left:

```
python find_emails.py --workbook Aurora_Email_Finder.xlsx --check-keys
```

**2. Test the wiring with fake data.** No network, no credits:

```
python find_emails.py --workbook Aurora_Email_Finder.xlsx --dry-run
```

**3. Go live**, a few rows at a time while you sanity-check:

```
python find_emails.py --workbook Aurora_Email_Finder.xlsx --limit 5
python find_emails.py --workbook Aurora_Email_Finder.xlsx
```

Results write into the green columns and save after **every row**, so you can stop and re-run
any time — it skips rows that already have an email. Add `--redo` to run them again anyway,
or `--verbose` to see the raw error body when a call fails.

## The workbook, three tabs

**API Keys** — one row per key, replacing the `PASTE_…` placeholders. `Priority 1` is used
first; when it genuinely runs out the script retires it and moves to the next, then to the
other provider. Set `Enabled` to `no` to park a key without deleting it.

| Provider | API Key | Priority | Enabled |
|---|---|---|---|
| hunter | your_key_1 | 1 | yes |
| hunter | your_key_2 | 2 | yes |
| apollo | your_key_3 | 3 | yes |

**Companies** — paste your list into the yellow columns. `Organisation` is the only required
field.

- **Know the website? Put it in `Website / Domain`.** This is the single most useful thing you
  can do: it skips name resolution, and that is the step most likely to fail on small clubs
  and charities.
- Chasing a named person? Add `First name` + `Last name`. Leave `Lookup mode` on `auto`.
- Just want the best inbox at an organisation? Leave the name blank.
- `Role hint` (`marketing`, `Commodore`, `CEO`) steers both providers toward the right
  contact instead of whoever happens to rank first.

**Settings** — confidence threshold, delay between calls, how many extra addresses to keep,
and whether to skip already-filled rows. The defaults are sensible; you can ignore this tab.

## What a row costs

Hunter's free plan gives you 50 credits a month.

| Call | Credits | What it does |
|---|---|---|
| Discover | **0** | Company name → domain. Tried first, always. |
| Domain Search | 1 | Up to 10 addresses at a domain. |
| Email Finder | 1 | One named person at a domain. |

So a row with the website filled in costs **1 credit**. A row without one costs 1 if Discover
recognises the name, 2 if it has to fall through to a paid name search. Filling in column C is
free and roughly doubles how far 50 credits goes.

Apollo works differently: its people search is free but returns **no email addresses**, so the
script enriches the single best match afterwards — and only that step costs Apollo credits.
Apollo also gates API access behind its paid tiers and its search endpoint wants a master key,
so expect `403` on a free Apollo account. That is normal; the script says so plainly and
carries on with Hunter.

## How key rotation works

A key is retired **only when the provider says the key itself is the problem** — `401` (bad
key), `402` (payment required), `403` (forbidden), or a `429` that specifically reports a
monthly usage cap. A plain `429` rate-limit waits two seconds and retries the same key.

Anything else — a malformed request, a `404`, a server hiccup — is reported with the
provider's own error message and the run carries on. A bad request never costs you a key, and
one bad row never ends the run.

## Reading the Status column

| Status | Meaning |
|---|---|
| `found` | Usable address at or above your confidence threshold. |
| `found_low_confidence` | Real address, but Hunter is unsure. Verify before sending. |
| `not_found` | Nobody in the database. Check the organisation's own contact page. |
| `no_domain` | No provider recognised the name. Paste the website into column C. |

The script will not write a domain it cannot match back to the name you typed. A blank costs
you nothing; a wrong address gets emailed to a stranger.

## Troubleshooting

**"every API key is out of credits" but the dashboard says otherwise.** That was a bug in an
earlier version, which asked Hunter for 25 results per page. Hunter's free plan rejects
`limit + offset > 10` with `400 pagination_error`, and the old error handler read every
non-200 as an exhausted key. Fixed — the page size is now 10, and `--check-keys` will always
give you the provider's own number rather than a guess.

**Nothing found for a small club or charity.** Expected. Hunter indexes corporate domains well
and volunteer-run organisations poorly. When a row comes back empty, that organisation's own
website usually lists the address you want — often faster than any API.

**`PermissionError` on save.** The workbook is open in Excel. The script writes
`…_results.xlsx` alongside it instead, so nothing is lost. Close Excel and re-run.

## What kind of list this works on

These APIs find **work emails from an organisation's domain**. A list of people and their home
addresses will not work — there is no company to look up. The right input is a list of
organisations: paste their names into column B, and their websites into column C wherever you
already know them.

## Please use it responsibly

Respect Hunter's and Apollo's terms of service, only approach organisations that welcome
contact, and follow Australia's Spam Act 2003: identify yourself and give an easy unsubscribe
in every email you send. This finds business contacts for legitimate outreach — keep it there.
