# Mist MPSK Generator

Creates multiple pre-shared keys (Mist cloud PSKs) at **site level** for a list of sites, then produces:

- an **Excel record** of every site / PSK name / SSID / VLAN / passphrase
- a **PDF per site** with the site name as the title and, for each PSK, its name, SSID, password and a Wi-Fi **QR code** that joins the network when scanned. There are three per page, with a "Page x of y" footer.

Every site in the sites CSV gets every PSK in the PSKs CSV.

## Setup

```bash
pip install -r requirements.txt
cp mist_mpsk_generator.ini.example mist_mpsk_generator.ini
cp sites.csv.example sites.csv
cp psks.csv.example psks.csv
```

Edit the ini: `cloud`, `api_token` (needs write access to the sites), `org_id`, the allowed key characters and the key length (8-63).

### sites.csv

```csv
site_name
London Office
Manchester Office
```

### psks.csv

```csv
name,ssid,vlan
Finance,Corp-PSK,110
Visitors,Guest-PSK,{{guest_vlan}}
```

`vlan` can be a number (1-4094), a Mist site variable such as `{{guest_vlan}}`, or blank (no VLAN override).

## Usage

```bash
python3 mist_mpsk_generator.py              # create PSKs, write Excel + PDFs
python3 mist_mpsk_generator.py --dry-run    # read-only checks, write Excel + PDFs, create nothing
python3 mist_mpsk_generator.py --rotate     # also give existing PSKs (same name + SSID) a new key
python3 mist_mpsk_generator.py --delete-unlisted        # also delete site PSKs not in psks.csv (asks first)
python3 mist_mpsk_generator.py --delete-unlisted --yes  # same, no prompts (unattended runs)
python3 mist_mpsk_generator.py -c customer.ini
```

Output goes to `output/<timestamp>/` (or `<timestamp>_dryrun`).

## Start-up summary and confirmation

After a few API calls, the script reads the org and matches your sites. Before scanning each site, it shows a summary:

```
================================================================
  Org        : Example-Org  (00000000-0000-...)
  Cloud      : api.eu.mist.com
  Sites      : 3 from sites.csv (of 3 in the org)
  PSKs       : 2 per site from psks.csv, on SSID(s) Corp-PSK, IoT-PSK
  Affected   : up to 6 site PSK(s) created (existing ones are kept)
  DELETE     : ON - any PSK on these sites not in the PSKs CSV will be deleted
  Mode       : LIVE - changes will be written to Mist
  API calls  : about 12 reads + up to 6 writes + deletes (estimate)
  API token  : 4,855 of 5,000 calls left this hour
================================================================
Continue with org 'Example-Org'? [y/N]:
```

- A live run asks **Continue?** The default is No.
- With `--delete-unlisted`, it then asks a second time whether to **keep delete mode on**. Later it lists the exact PSKs and asks you to type `DELETE`.
- A dry run shows the summary without prompting.
- `--yes` skips every prompt. A non-interactive run without `--yes` refuses to make changes.

## What it checks before writing anything

If any check fails, the script lists every problem and stops before anything is created:

- the ini and both CSVs are valid (no duplicate names, VLANs in range, SSIDs no longer than 32 bytes)
- every site name exists in the org, and exactly once
- every SSID is an enabled WLAN at each site, using `/sites/{id}/wlans/derived` so template WLANs count, with PSK auth and local keys (multiple passphrases). WLANs using dynamic PSK, either RADIUS or org-level Cloud PSKs, are rejected because they ignore site PSKs

## Behaviour

- **Unique keys**: every key generated in a run is unique, and none matches a passphrase that already exists on those sites. Keys use Python's `secrets` module.
- **Re-runs**: a PSK that already exists at the site (same name and SSID) is **kept**, and its current key goes into the Excel file and PDF. Use `--rotate` to replace it.
- PSKs are created as `usage: multi` with no expiry.
- The **QR code** uses the standard `WIFI:T:WPA;S:…;P:…;;` format, escaping `\ ; , : "`. `H:true` is added when the WLAN is hidden.
- The script **warns** if the allowed characters include QR-escaped characters, a space, or look-alikes (`0/O`, `1/l/I`).
- **Partial failure**: if an API write fails midway, the run stops and writes `mpsk_keys_<ts>_INCOMPLETE.xlsx`, which records what was written. In that case no PDFs are produced. Re-run after fixing the problem: PSKs already created are kept.

## Deleting PSKs that aren't in the file

`--delete-unlisted` deletes every PSK on the **listed sites** whose name + SSID pair is not in `psks.csv`. This includes PSKs on SSIDs that aren't in the file at all. Sites missing from `sites.csv` are never touched.

- Every normal run reports how many unlisted PSKs exist, but leaves them alone.
- `--dry-run --delete-unlisted` lists exactly what would be deleted. **Always run this first**, especially if `psks.csv` has been cut down to only some of the PSKs.
- A live run lists the PSKs and asks you to type `DELETE`. A non-interactive run refuses unless `--yes` is given.
- Deletes run after creates and rotates, so a failed create never leaves a site short of keys.
- The Excel file has a **Deleted** sheet with each deleted PSK's name, SSID, VLAN and old passphrase, so a PSK deleted by mistake can be recreated with the same key.

## API token usage

Every run prints and records its API usage:

- the token's usage from `GET /self/usage` at the start, just before writing, and at the end: calls used this hour, the hourly limit (normally 5,000) and calls left
- the number of API calls this run made, by method (GET/POST/PUT/DELETE)

These figures are also saved on an **API Usage** sheet in the Excel file. If the token doesn't have enough calls left this hour for the planned writes, the run stops before writing anything. The token figures include any other script or person using the same token. Each usage check costs one call.

A run uses about 2 calls per site to read, plus 1 call per PSK created, rotated or deleted, plus a few fixed calls.

## Proxy / Zscaler

Add the optional `[network]` section (see the example ini):

- `ca_bundle`: the proxy's root CA, as a PEM file
- `verify_ssl`: TLS certificate checking; `false` turns it off, as a last resort only
- `http_proxy` / `https_proxy`: proxy addresses

Without the section, the script uses the normal certifi CA store and any proxy environment variables.

## Security

The Excel file and PDFs contain **plaintext Wi-Fi keys**. `*.ini`, `*.csv` (except the examples), `*.xlsx`, `*.pdf` and `output/` are git-ignored.
