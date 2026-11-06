#!/usr/bin/env python3
"""
Mist MPSK Generator

Creates multiple pre-shared keys (Mist cloud PSKs) at site level for a list of
sites, then writes:
  * an Excel record of every site / PSK name / SSID / VLAN / passphrase
  * one PDF per site with each PSK's name, SSID and passphrase plus a Wi-Fi
    QR code (three per A4 page, "Page x of y" footer)

Inputs:
  * mist_mpsk_generator.ini  - API token, org ID, cloud, key alphabet and length
  * sites CSV                - one site name per row       (column: site_name)
  * PSKs CSV                 - PSKs every site should get  (columns: name, ssid, vlan)

Every site gets every PSK. All checks (sites exist, SSIDs exist at each site
with PSK auth, CSVs are valid) run before anything is written - if any check
fails the script stops and lists every problem, nothing is created.

Usage:
    python3 mist_mpsk_generator.py                 # create PSKs + outputs
    python3 mist_mpsk_generator.py --dry-run       # read-only checks + outputs, no writes
    python3 mist_mpsk_generator.py --rotate        # replace keys for PSKs that already exist
    python3 mist_mpsk_generator.py --delete-unlisted   # also delete site PSKs not in the PSKs CSV
    python3 mist_mpsk_generator.py -c other.ini
"""

from __future__ import annotations

import argparse
import configparser
import csv
import math
import re
import secrets
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

try:
    import requests
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from reportlab.graphics import renderPDF
    from reportlab.graphics.barcode.qr import QrCodeWidget
    from reportlab.graphics.shapes import Drawing
    from reportlab.lib.colors import HexColor
    from reportlab.lib.pagesizes import A4, LETTER
    from reportlab.lib.units import cm
    from reportlab.pdfbase.pdfmetrics import stringWidth
    from reportlab.pdfgen import canvas
except ImportError as e:  # pragma: no cover
    sys.exit(f"Missing dependency: {e.name}. Run: pip install -r requirements.txt")

SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_INI = SCRIPT_DIR / "mist_mpsk_generator.ini"
REQUEST_TIMEOUT = 30
PSK_AUTH_TYPES = {"psk", "psk-tkip", "psk-wpa2-tkip"}

# Characters that must be escaped inside a WIFI: QR payload. The script escapes
# them, but some older scanners still get them wrong.
QR_SPECIAL_CHARS = set('\\;,:"')
LOOKALIKE_GROUPS = ["0O", "1lI"]


class MpskError(Exception):
    """Raised for any problem that should stop the run with a clear message."""


def info(msg: str) -> None:
    print(msg, flush=True)


def warn(msg: str) -> None:
    print(f"WARNING: {msg}", flush=True)


# --------------------------------------------------------------------------
# Config
# --------------------------------------------------------------------------

@dataclass
class Config:
    base_url: str
    api_token: str
    org_id: str
    alphabet: str
    length: int
    sites_csv: Path
    psks_csv: Path
    output_dir: Path
    page_size: tuple
    footer_text: str
    verify: object = True
    proxies: dict = field(default_factory=dict)


def cloud_to_base_url(cloud: str) -> str:
    cloud = cloud.strip().lower()
    cloud = re.sub(r"^https?://", "", cloud).split("/")[0]
    if not cloud:
        raise MpskError("[mist] cloud is required (e.g. api.mist.com, api.eu.mist.com, api.gc3.mist.com).")
    if "." not in cloud:
        cloud = {"global": "api.mist.com", "us": "api.mist.com", "eu": "api.eu.mist.com"}.get(
            cloud, f"api.{cloud}.mist.com")
    if cloud.startswith("manage."):
        cloud = "api." + cloud[len("manage."):]
    return f"https://{cloud}/api/v1"


def read_network(parser: configparser.ConfigParser):
    """Optional [network] section for TLS-inspecting proxies (e.g. Zscaler).
    With no section, behaviour is normal (certifi CA store, env-var proxies)."""
    if "network" not in parser:
        return True, {}
    section = parser["network"]
    ca_bundle = section.get("ca_bundle", "").strip()
    verify_ssl = section.getboolean("verify_ssl", fallback=True)
    if ca_bundle:
        ca_path = Path(ca_bundle).expanduser()
        if not ca_path.is_file():
            raise MpskError(
                f"[network] ca_bundle does not exist: {ca_path}\n"
                "This should be a PEM file containing your proxy's root CA certificate "
                "(e.g. exported from Zscaler).")
        verify = str(ca_path)
    else:
        verify = verify_ssl
    proxies = {}
    for scheme in ("http", "https"):
        value = section.get(f"{scheme}_proxy", "").strip()
        if value:
            proxies[scheme] = value
    return verify, proxies


def resolve_path(value: str, ini_dir: Path) -> Path:
    p = Path(value).expanduser()
    return p if p.is_absolute() else (ini_dir / p)


def load_config(ini_path: Path) -> Config:
    if not ini_path.is_file():
        raise MpskError(f"Config file not found: {ini_path}\n"
                        f"Copy mist_mpsk_generator.ini.example to {ini_path.name} and fill it in.")
    # Interpolation off so '%' can appear in the allowed character list.
    parser = configparser.ConfigParser(interpolation=None)
    parser.read(ini_path, encoding="utf-8")
    problems = []

    def get(section, key, default=None):
        if parser.has_option(section, key):
            return parser.get(section, key).strip()
        if default is None:
            problems.append(f"[{section}] {key} is missing")
            return ""
        return default

    api_token = get("mist", "api_token")
    org_id = get("mist", "org_id")
    cloud = get("mist", "cloud")

    # allowed_characters is read raw (not stripped) so a space can be allowed
    # deliberately - but surrounding whitespace from the ini line is removed.
    alphabet_raw = get("psk", "allowed_characters")
    length_raw = get("psk", "length")

    ini_dir = ini_path.parent
    sites_csv = resolve_path(get("files", "sites_csv", "sites.csv"), ini_dir)
    psks_csv = resolve_path(get("files", "psks_csv", "psks.csv"), ini_dir)
    output_dir = resolve_path(get("files", "output_dir", "output"), ini_dir)

    page_name = get("pdf", "page_size", "A4").upper()
    footer_text = get("pdf", "footer_text", "")

    if problems:
        raise MpskError("Config problems in " + str(ini_path) + ":\n  - " + "\n  - ".join(problems))

    # --- alphabet ---------------------------------------------------------
    seen = []
    for ch in alphabet_raw:
        if ch not in seen:
            seen.append(ch)
    alphabet = "".join(seen)
    bad = [ch for ch in alphabet if not (0x20 <= ord(ch) <= 0x7E)]
    if bad:
        problems.append("[psk] allowed_characters may only contain printable ASCII "
                        f"(WPA passphrase rule); found: {' '.join(repr(c) for c in bad)}")
    if len(alphabet) < 10:
        problems.append(f"[psk] allowed_characters has only {len(alphabet)} distinct characters; use at least 10")

    try:
        length = int(length_raw)
        if not 8 <= length <= 63:
            problems.append("[psk] length must be between 8 and 63 (WPA passphrase limit)")
    except ValueError:
        length = 0
        problems.append(f"[psk] length must be a whole number, got {length_raw!r}")

    if page_name not in ("A4", "LETTER"):
        problems.append(f"[pdf] page_size must be A4 or Letter, got {page_name!r}")

    try:
        base_url = cloud_to_base_url(cloud)
    except MpskError as e:
        problems.append(str(e))
        base_url = ""

    if problems:
        raise MpskError("Config problems in " + str(ini_path) + ":\n  - " + "\n  - ".join(problems))

    # --- warnings (don't stop the run) -----------------------------------
    special = sorted(QR_SPECIAL_CHARS & set(alphabet))
    if special:
        warn("allowed_characters includes " + " ".join(special) +
             " - these are escaped in the QR code, but some older phone scanners mishandle them "
             "and they are awkward to type.")
    if " " in alphabet:
        warn("allowed_characters includes a space - keys may be hard to read from the PDF.")
    for group in LOOKALIKE_GROUPS:
        present = [c for c in group if c in alphabet]
        if len(present) > 1:
            warn(f"allowed_characters includes look-alike characters {' '.join(present)} - "
                 "consider removing some to make keys easier to type.")

    verify, proxies = read_network(parser)
    return Config(base_url=base_url, api_token=api_token, org_id=org_id, alphabet=alphabet,
                  length=length, sites_csv=sites_csv, psks_csv=psks_csv, output_dir=output_dir,
                  page_size=A4 if page_name == "A4" else LETTER, footer_text=footer_text,
                  verify=verify, proxies=proxies)


# --------------------------------------------------------------------------
# CSV input
# --------------------------------------------------------------------------

@dataclass
class PskSpec:
    name: str
    ssid: str
    vlan: object  # int or "{{var}}" string, or None


def read_csv_rows(path: Path, required: list[str]) -> list[dict]:
    if not path.is_file():
        raise MpskError(f"CSV file not found: {path}")
    with path.open(newline="", encoding="utf-8-sig") as fh:
        reader = csv.DictReader(fh)
        if not reader.fieldnames:
            raise MpskError(f"{path.name} is empty")
        headers = {h.strip().lower(): h for h in reader.fieldnames if h}
        missing = [r for r in required if r not in headers]
        if missing:
            raise MpskError(f"{path.name} is missing column(s): {', '.join(missing)} "
                            f"(found: {', '.join(reader.fieldnames)})")
        rows = []
        for line_no, row in enumerate(reader, start=2):
            clean = {k: (row.get(headers[k]) or "").strip() for k in headers}
            if not any(clean.values()):
                continue
            clean["_line"] = line_no
            rows.append(clean)
    return rows


def load_sites(path: Path) -> list[str]:
    rows = read_csv_rows(path, ["site_name"])
    problems, names, seen = [], [], set()
    for row in rows:
        name = row["site_name"]
        if not name:
            problems.append(f"line {row['_line']}: site_name is blank")
        elif name.lower() in seen:
            problems.append(f"line {row['_line']}: duplicate site {name!r}")
        else:
            seen.add(name.lower())
            names.append(name)
    if problems:
        raise MpskError(f"Problems in {path.name}:\n  - " + "\n  - ".join(problems))
    if not names:
        raise MpskError(f"{path.name} has no sites")
    return names


def parse_vlan(raw: str):
    if raw == "":
        return None
    if re.fullmatch(r"\d+", raw):
        v = int(raw)
        if 1 <= v <= 4094:
            return v
        raise ValueError(f"VLAN {raw} is outside 1-4094")
    if re.fullmatch(r"\{\{\s*[A-Za-z0-9_]+\s*\}\}", raw):
        return raw.replace(" ", "")
    raise ValueError(f"VLAN {raw!r} must be a number 1-4094 or a site variable like {{{{guest_vlan}}}}")


def load_psks(path: Path) -> list[PskSpec]:
    rows = read_csv_rows(path, ["name", "ssid", "vlan"])
    problems, specs, seen = [], [], set()
    for row in rows:
        ln = row["_line"]
        if not row["name"]:
            problems.append(f"line {ln}: name is blank")
            continue
        if not row["ssid"]:
            problems.append(f"line {ln}: ssid is blank")
            continue
        if len(row["ssid"].encode()) > 32:
            problems.append(f"line {ln}: SSID {row['ssid']!r} is longer than 32 bytes")
        try:
            vlan = parse_vlan(row["vlan"])
        except ValueError as e:
            problems.append(f"line {ln}: {e}")
            continue
        key = (row["name"].lower(), row["ssid"])
        if key in seen:
            problems.append(f"line {ln}: duplicate PSK name {row['name']!r} for SSID {row['ssid']!r}")
            continue
        seen.add(key)
        specs.append(PskSpec(row["name"], row["ssid"], vlan))
    if problems:
        raise MpskError(f"Problems in {path.name}:\n  - " + "\n  - ".join(problems))
    if not specs:
        raise MpskError(f"{path.name} has no PSKs")
    return specs


# --------------------------------------------------------------------------
# Mist API
# --------------------------------------------------------------------------

class Mist:
    def __init__(self, cfg: Config):
        self.base = cfg.base_url
        self.s = requests.Session()
        self.s.headers.update({"Authorization": f"Token {cfg.api_token}",
                               "Content-Type": "application/json"})
        self.s.verify = cfg.verify
        self.calls: dict[str, int] = {}   # API calls made by this run, by method
        if cfg.proxies:
            self.s.proxies.update(cfg.proxies)
        if cfg.verify is False:
            warn("TLS certificate verification is DISABLED ([network] verify_ssl = false). "
                 "Only use this as a last resort.")
            requests.packages.urllib3.disable_warnings(
                requests.packages.urllib3.exceptions.InsecureRequestWarning)

    def request(self, method: str, path: str, **kwargs):
        url = self.base + path
        for attempt in range(5):
            self.calls[method] = self.calls.get(method, 0) + 1
            try:
                r = self.s.request(method, url, timeout=REQUEST_TIMEOUT, **kwargs)
            except requests.exceptions.SSLError as e:
                raise MpskError(f"TLS certificate verification failed for {url}: {e}\n"
                                "If you're behind a TLS-inspecting proxy (e.g. Zscaler), set "
                                "[network] ca_bundle to your proxy's root CA (PEM).")
            except requests.exceptions.ProxyError as e:
                raise MpskError(f"Proxy error calling {url}: {e}\nCheck [network] http_proxy / https_proxy.")
            except requests.exceptions.RequestException as e:
                raise MpskError(f"Network error calling {url}: {e}")
            if r.status_code == 429 and attempt < 4:
                wait = int(r.headers.get("Retry-After", "0") or 0) or 2 ** (attempt + 2)
                warn(f"Rate limited by Mist API, waiting {wait}s...")
                time.sleep(wait)
                continue
            if r.status_code >= 400:
                detail = r.text.strip()[:500]
                hint = ""
                if r.status_code == 401:
                    hint = " (check api_token and cloud)"
                elif r.status_code == 403:
                    hint = " (token lacks permission for this org/site - needs write access to create PSKs)"
                elif r.status_code == 404:
                    hint = " (check org_id and cloud)"
                raise MpskError(f"{method} {path} -> HTTP {r.status_code}{hint}: {detail}")
            return r.json() if r.content else {}
        raise MpskError(f"{method} {path} kept hitting the API rate limit")

    @property
    def total_calls(self) -> int:
        return sum(self.calls.values())

    def usage(self) -> dict | None:
        """GET /self/usage - the token's API calls this hour vs its hourly limit.
        Costs one call itself. Returns None if it can't be read (never fatal)."""
        try:
            u = self.request("GET", "/self/usage")
            limit, used = u.get("request_limit"), u.get("requests")
            if isinstance(limit, int) and isinstance(used, int):
                return {"limit": limit, "used": used, "left": limit - used, "seconds": u.get("seconds")}
        except MpskError as e:
            warn(f"Could not read API token usage (/self/usage): {e}")
        return None

    def get_all(self, path: str, params: dict | None = None) -> list:
        """Page-based listing (sites, PSKs)."""
        out, page = [], 1
        params = dict(params or {})
        while True:
            params.update({"limit": 1000, "page": page})
            batch = self.request("GET", path, params=params)
            if not isinstance(batch, list):
                raise MpskError(f"Unexpected response from {path}: {str(batch)[:200]}")
            out.extend(batch)
            if len(batch) < 1000:
                return out
            page += 1


# --------------------------------------------------------------------------
# Planning / validation
# --------------------------------------------------------------------------

@dataclass
class PlanItem:
    site_name: str
    site_id: str
    spec: PskSpec
    hidden: bool
    existing: dict | None = None      # existing PSK object at site, if any
    action: str = ""                  # create / rotate / existing
    passphrase: str = ""
    status: str = "pending"


@dataclass
class DeleteItem:
    site_name: str
    site_id: str
    psk: dict                         # the existing PSK object, as returned by Mist
    status: str = "pending"


def resolve_sites(api: Mist, cfg: Config, site_names: list[str]) -> tuple[dict, list, int]:
    """Reads the org and matches the CSV site names to site IDs.
    Returns (org, [(name, site_id)], number of sites in the org)."""
    info(f"Reading org {cfg.org_id} ...")
    org = api.request("GET", f"/orgs/{cfg.org_id}")
    org_sites = api.get_all(f"/orgs/{cfg.org_id}/sites")

    by_name: dict[str, list[dict]] = {}
    for s in org_sites:
        by_name.setdefault((s.get("name") or "").strip().lower(), []).append(s)

    problems, resolved = [], []
    for name in site_names:
        matches = by_name.get(name.lower(), [])
        exact = [m for m in matches if m.get("name", "").strip() == name]
        matches = exact or matches
        if not matches:
            problems.append(f"Site {name!r} not found in org")
        elif len(matches) > 1:
            problems.append(f"Site name {name!r} matches {len(matches)} sites in the org - make it unique")
        else:
            resolved.append((name, matches[0]["id"]))
    if problems:
        raise MpskError("Site problems:\n  - " + "\n  - ".join(problems))
    return org, resolved, len(org_sites)


def build_plan(api: Mist, resolved: list, specs: list[PskSpec]) -> tuple[list[PlanItem], list[DeleteItem]]:
    """Returns the PSKs to create/rotate/keep and the site PSKs that are NOT in
    the PSKs CSV (candidates for --delete-unlisted)."""
    problems = []
    listed = {(spec.name.lower(), spec.ssid) for spec in specs}
    plan: list[PlanItem] = []
    unlisted: list[DeleteItem] = []
    for i, (name, site_id) in enumerate(resolved, 1):
        info(f"  [{i}/{len(resolved)}] Checking {name} ...")
        wlans = api.request("GET", f"/sites/{site_id}/wlans/derived")
        ssid_wlans: dict[str, list[dict]] = {}
        for w in wlans:
            if w.get("enabled", True) is False:
                continue
            ssid_wlans.setdefault(w.get("ssid", ""), []).append(w)

        existing = {}
        for p in api.get_all(f"/sites/{site_id}/psks"):
            key = ((p.get("name") or "").lower(), p.get("ssid", ""))
            existing[key] = p
            if key not in listed:
                unlisted.append(DeleteItem(name, site_id, p))

        for spec in specs:
            candidates = ssid_wlans.get(spec.ssid, [])
            if not candidates:
                problems.append(f"{name}: SSID {spec.ssid!r} is not an enabled WLAN at this site")
                continue
            psk_wlans = [w for w in candidates if (w.get("auth") or {}).get("type") in PSK_AUTH_TYPES]
            if not psk_wlans:
                types = ", ".join(sorted({(w.get('auth') or {}).get('type', '?') for w in candidates}))
                problems.append(f"{name}: SSID {spec.ssid!r} uses auth type {types}, not PSK")
                continue
            if any((w.get("dynamic_psk") or {}).get("enabled") for w in psk_wlans):
                problems.append(f"{name}: SSID {spec.ssid!r} uses RADIUS dynamic PSK - cloud PSKs are not used")
                continue
            hidden = any(w.get("hide_ssid") for w in psk_wlans)
            plan.append(PlanItem(name, site_id, spec, hidden,
                                 existing=existing.get((spec.name.lower(), spec.ssid))))

    if problems:
        raise MpskError("SSID problems (nothing has been created):\n  - " + "\n  - ".join(problems))
    return plan, unlisted


def generate_passphrases(plan: list[PlanItem], cfg: Config, rotate: bool) -> None:
    space = len(cfg.alphabet) ** cfg.length
    needed = len(plan)
    if space < needed * 1000:
        raise MpskError(f"Only {space:,} possible keys from {len(cfg.alphabet)} characters at length "
                        f"{cfg.length} - too few to safely make {needed} unique keys. "
                        "Increase [psk] length or allowed_characters.")

    used = {item.existing["passphrase"] for item in plan
            if item.existing and item.existing.get("passphrase")}
    for item in plan:
        if item.existing and not rotate:
            item.action = "existing"
            item.passphrase = item.existing.get("passphrase", "")
            continue
        item.action = "rotate" if item.existing else "create"
        while True:
            key = "".join(secrets.choice(cfg.alphabet) for _ in range(cfg.length))
            # A passphrase can't start or end with a space (stripped by many clients).
            if key not in used and key[0] != " " and key[-1] != " ":
                break
        used.add(key)
        item.passphrase = key


def apply_plan(api: Mist, plan: list[PlanItem]) -> None:
    todo = [p for p in plan if p.action in ("create", "rotate")]
    for i, item in enumerate(todo, 1):
        body = {"name": item.spec.name, "ssid": item.spec.ssid,
                "passphrase": item.passphrase, "usage": "multi"}
        if item.spec.vlan is not None:
            body["vlan_id"] = item.spec.vlan
        if item.action == "create":
            api.request("POST", f"/sites/{item.site_id}/psks", json=body)
        else:
            api.request("PUT", f"/sites/{item.site_id}/psks/{item.existing['id']}", json=body)
        item.status = "created" if item.action == "create" else "rotated"
        info(f"  [{i}/{len(todo)}] {item.action:6} {item.site_name} / {item.spec.name} ({item.spec.ssid})")


def show_usage(label: str, u: dict | None) -> None:
    if u:
        info(f"API token usage {label}: {u['used']:,} of {u['limit']:,} calls used this hour, "
             f"{u['left']:,} left")


def check_usage_headroom(api: Mist, writes: int) -> dict | None:
    """Stop before writing anything if the token hasn't enough calls left this hour."""
    u = api.usage()
    show_usage("before writing", u)
    if u and writes + 5 > u["left"]:
        reset = f" (resets in about {int(u['seconds']) // 60} min)" if isinstance(u.get("seconds"), (int, float)) else ""
        raise MpskError(f"This run needs {writes} write call(s) but the API token only has {u['left']} "
                        f"call(s) left this hour{reset}. Nothing has been written - try again later.")
    return u


def ask_yes_no(question: str) -> bool:
    try:
        return input(f"{question} [y/N]: ").strip().lower() in ("y", "yes")
    except EOFError:
        return False


def confirm_start(org: dict, cfg: Config, resolved: list, org_site_count: int, specs: list[PskSpec],
                  args, usage: dict | None) -> None:
    """Start-up summary of what the run will touch, and the go/no-go prompt."""
    n_sites, n_psks = len(resolved), len(specs)
    ssids = sorted({s.ssid for s in specs})
    reads = 2 * n_sites + 6
    line = "=" * 64
    info("")
    info(line)
    info(f"  Org        : {org.get('name', '?')}  ({cfg.org_id})")
    info(f"  Cloud      : {cfg.base_url.split('/')[2]}")
    info(f"  Sites      : {n_sites} from {cfg.sites_csv.name} (of {org_site_count} in the org)")
    info(f"  PSKs       : {n_psks} per site from {cfg.psks_csv.name}, on SSID(s) {', '.join(ssids)}")
    info(f"  Affected   : up to {n_sites * n_psks} site PSK(s) "
         + ("given NEW keys where they already exist (--rotate)" if args.rotate
            else "created (existing ones are kept)"))
    if args.delete_unlisted:
        info("  DELETE     : ON - any PSK on these sites not in the PSKs CSV will be deleted")
    info(f"  Mode       : {'DRY RUN - nothing written to Mist' if args.dry_run else 'LIVE - changes will be written to Mist'}")
    est = f"about {reads} reads" + ("" if args.dry_run else f" + up to {n_sites * n_psks} writes"
                                   + (" + deletes" if args.delete_unlisted else ""))
    info(f"  API calls  : {est} (estimate)")
    if usage:
        info(f"  API token  : {usage['left']:,} of {usage['limit']:,} calls left this hour")
    info(line)

    if args.dry_run or args.yes:
        return
    if not sys.stdin.isatty():
        raise MpskError("Refusing to make live changes without confirmation in a non-interactive run - add --yes.")
    if not ask_yes_no(f"Continue with org '{org.get('name', '?')}'?"):
        raise MpskError("Cancelled - nothing has been written to Mist.")
    if args.delete_unlisted and not ask_yes_no(
            f"Delete mode is ON: PSKs on these {n_sites} site(s) that are not in {cfg.psks_csv.name} "
            "will be deleted (you'll see the exact list before anything is deleted). Keep delete mode on?"):
        raise MpskError("Cancelled - run again without --delete-unlisted if you don't want deletions. "
                        "Nothing has been written to Mist.")


def describe_psk(d: DeleteItem) -> str:
    vlan = d.psk.get("vlan_id")
    return (f"{d.site_name} / {d.psk.get('name', '?')} ({d.psk.get('ssid', '?')})"
            + (f" VLAN {vlan}" if vlan not in (None, "") else ""))


def confirm_deletions(deletions: list[DeleteItem], assume_yes: bool) -> None:
    info(f"The following {len(deletions)} PSK(s) are not in the PSKs CSV and will be DELETED:")
    for d in deletions:
        info(f"  - {describe_psk(d)}")
    if assume_yes:
        return
    if not sys.stdin.isatty():
        raise MpskError("Refusing to delete without confirmation in a non-interactive run - add --yes.")
    answer = input(f"Type DELETE to delete these {len(deletions)} PSK(s), anything else to cancel: ")
    if answer.strip() != "DELETE":
        raise MpskError("Deletion not confirmed - nothing has been written to Mist.")


def apply_deletions(api: Mist, deletions: list[DeleteItem]) -> None:
    for i, d in enumerate(deletions, 1):
        api.request("DELETE", f"/sites/{d.site_id}/psks/{d.psk['id']}")
        d.status = "deleted"
        info(f"  [{i}/{len(deletions)}] delete {describe_psk(d)}")


# --------------------------------------------------------------------------
# Excel record
# --------------------------------------------------------------------------

def _style_sheet(ws, widths) -> None:
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(vertical="center")
    for col, width in zip("ABCDEFG", widths):
        ws.column_dimensions[col].width = width
    for row in ws.iter_rows(min_row=2, min_col=5, max_col=5):
        for c in row:
            c.font = Font(name="Courier New")
            c.number_format = "@"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def write_excel(plan: list[PlanItem], path: Path, dry_run: bool,
                deletions: list[DeleteItem] | None = None, usage: list | None = None) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "PSKs"
    ws.append(["Site", "PSK Name", "SSID", "VLAN", "Passphrase", "Action", "Status"])
    for item in plan:
        ws.append([item.site_name, item.spec.name, item.spec.ssid,
                   "" if item.spec.vlan is None else item.spec.vlan,
                   item.passphrase, item.action, "dry run - not written" if dry_run else item.status])
    _style_sheet(ws, (28, 24, 24, 16, 40, 10, 22))

    if deletions:
        # Keep the old passphrases so a deleted PSK can be recreated by hand if needed.
        wd = wb.create_sheet("Deleted")
        wd.append(["Site", "PSK Name", "SSID", "VLAN", "Passphrase", "Action", "Status"])
        for d in deletions:
            vlan = d.psk.get("vlan_id")
            wd.append([d.site_name, d.psk.get("name", ""), d.psk.get("ssid", ""),
                       "" if vlan is None else vlan, d.psk.get("passphrase", ""), "delete",
                       "dry run - not deleted" if dry_run else d.status])
        _style_sheet(wd, (28, 24, 24, 16, 40, 10, 22))

    if usage:
        wu = wb.create_sheet("API Usage")
        wu.append(["Item", "Value"])
        for row in usage:
            wu.append(list(row))
        for c in wu[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F4E78")
        wu.column_dimensions["A"].width = 44
        wu.column_dimensions["B"].width = 18
    wb.save(path)


# --------------------------------------------------------------------------
# PDF
# --------------------------------------------------------------------------

PER_PAGE = 3
INK = HexColor("#1A1A1A")
MUTED = HexColor("#5F6368")
RULE = HexColor("#C8CCD0")


def wifi_qr_payload(ssid: str, passphrase: str, hidden: bool) -> str:
    def esc(s: str) -> str:
        return "".join("\\" + ch if ch in QR_SPECIAL_CHARS else ch for ch in s)
    return f"WIFI:T:WPA;S:{esc(ssid)};P:{esc(passphrase)};{'H:true;' if hidden else ''};"


def draw_qr(c: canvas.Canvas, payload: str, x: float, y: float, size: float) -> None:
    widget = QrCodeWidget(payload, barLevel="M", barBorder=4)
    x0, y0, x1, y1 = widget.getBounds()
    w, h = x1 - x0, y1 - y0
    d = Drawing(size, size, transform=[size / w, 0, 0, size / h, 0, 0])
    d.add(widget)
    renderPDF.draw(d, c, x, y)


def fit_font(text: str, font: str, max_size: float, min_size: float, width: float) -> float:
    size = max_size
    while size > min_size and stringWidth(text, font, size) > width:
        size -= 0.5
    return size


def wrap_mono(text: str, font: str, size: float, width: float) -> list[str]:
    per_line = max(1, int(width // stringWidth("M", font, size)))
    return [text[i:i + per_line] for i in range(0, len(text), per_line)]


def write_site_pdf(site_name: str, items: list[PlanItem], path: Path, cfg: Config) -> None:
    page_w, page_h = cfg.page_size
    margin = 1.8 * cm
    header_h = 3.5 * cm
    footer_h = 1.6 * cm
    qr_size = 5.6 * cm
    pages = max(1, math.ceil(len(items) / PER_PAGE))

    c = canvas.Canvas(str(path), pagesize=cfg.page_size)
    c.setTitle(f"{site_name} - Wi-Fi Access")
    c.setAuthor("Mist MPSK Generator")

    slot_h = (page_h - header_h - footer_h - margin) / PER_PAGE

    for page in range(pages):
        # Header
        top = page_h - margin
        c.setFillColor(INK)
        title_size = fit_font(site_name, "Helvetica-Bold", 24, 12, page_w - 2 * margin)
        c.setFont("Helvetica-Bold", title_size)
        c.drawString(margin, top - title_size, site_name)
        c.setFont("Helvetica", 10.5)
        c.setFillColor(MUTED)
        c.drawString(margin, top - title_size - 16, "Wi-Fi access - scan a QR code with your phone camera to connect")
        c.setStrokeColor(INK)
        c.setLineWidth(1.2)
        rule_y = page_h - header_h
        c.line(margin, rule_y, page_w - margin, rule_y)

        # Slots
        for slot, item in enumerate(items[page * PER_PAGE:(page + 1) * PER_PAGE]):
            slot_top = rule_y - slot * slot_h
            if slot > 0:
                c.setStrokeColor(RULE)
                c.setLineWidth(0.6)
                c.setDash(3, 3)
                c.line(margin, slot_top, page_w - margin, slot_top)
                c.setDash()

            qr_x = margin
            qr_y = slot_top - (slot_h + qr_size) / 2
            draw_qr(c, wifi_qr_payload(item.spec.ssid, item.passphrase, item.hidden), qr_x, qr_y, qr_size)

            tx = margin + qr_size + 1.2 * cm
            tw = page_w - margin - tx
            y = slot_top - (slot_h - qr_size) / 2 - 4

            name_size = fit_font(item.spec.name, "Helvetica-Bold", 17, 10, tw)
            c.setFillColor(INK)
            c.setFont("Helvetica-Bold", name_size)
            c.drawString(tx, y - name_size, item.spec.name)
            y -= name_size + 22

            def label(text, yy):
                c.setFillColor(MUTED)
                c.setFont("Helvetica", 8.5)
                c.drawString(tx, yy, text.upper())

            label("Network (SSID)", y)
            y -= 16
            c.setFillColor(INK)
            ssid_size = fit_font(item.spec.ssid, "Helvetica", 13, 8, tw)
            c.setFont("Helvetica", ssid_size)
            c.drawString(tx, y, item.spec.ssid)
            y -= 24

            label("Password", y)
            y -= 18
            pw_size = 14 if len(item.passphrase) <= 24 else 11
            c.setFont("Courier-Bold", pw_size)
            c.setFillColor(INK)
            for line in wrap_mono(item.passphrase, "Courier-Bold", pw_size, tw):
                c.drawString(tx, y, line)
                y -= pw_size + 4
            y -= 10

            extra = []
            if item.spec.vlan is not None:
                extra.append(f"VLAN {item.spec.vlan}")
            if item.hidden:
                extra.append("Hidden network")
            if extra:
                c.setFillColor(MUTED)
                c.setFont("Helvetica", 8.5)
                c.drawString(tx, y, "  ·  ".join(extra))

        # Footer
        c.setStrokeColor(RULE)
        c.setLineWidth(0.6)
        c.line(margin, footer_h, page_w - margin, footer_h)
        c.setFont("Helvetica", 8.5)
        c.setFillColor(MUTED)
        c.drawString(margin, footer_h - 14, cfg.footer_text or site_name)
        c.drawRightString(page_w - margin, footer_h - 14, f"Page {page + 1} of {pages}")
        c.showPage()
    c.save()


def safe_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|\x00-\x1f]+', "_", name).strip(" .") or "site"


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def main() -> int:
    ap = argparse.ArgumentParser(description="Create Mist site-level multi-PSKs and per-site QR code PDFs.")
    ap.add_argument("-c", "--config", type=Path, default=DEFAULT_INI, help="ini file (default: %(default)s)")
    ap.add_argument("--dry-run", action="store_true",
                    help="run all read-only checks and produce Excel/PDFs, but create nothing in Mist")
    ap.add_argument("--rotate", action="store_true",
                    help="give PSKs that already exist (same name + SSID) a new key; default is to keep them")
    ap.add_argument("--delete-unlisted", action="store_true",
                    help="delete PSKs on the listed sites whose name + SSID are not in the PSKs CSV "
                         "(asks you to type DELETE first; sites not in the sites CSV are never touched)")
    ap.add_argument("--yes", action="store_true",
                    help="skip all confirmation prompts, including DELETE (for unattended runs)")
    args = ap.parse_args()

    try:
        cfg = load_config(args.config.expanduser().resolve())
        site_names = load_sites(cfg.sites_csv)
        specs = load_psks(cfg.psks_csv)
        api = Mist(cfg)
        usage_start = api.usage()
        org, resolved, org_site_count = resolve_sites(api, cfg, site_names)
        confirm_start(org, cfg, resolved, org_site_count, specs, args, usage_start)
        info("Checking SSIDs and existing PSKs on each site ...")
        plan, unlisted = build_plan(api, resolved, specs)
        generate_passphrases(plan, cfg, args.rotate)
        deletions = unlisted if args.delete_unlisted else []

        counts = {a: sum(1 for p in plan if p.action == a) for a in ("create", "rotate", "existing")}
        info(f"Plan: {counts['create']} to create, {counts['rotate']} to rotate, "
             f"{counts['existing']} already exist (kept)"
             + (f", {len(deletions)} to delete" if args.delete_unlisted else ""))
        if unlisted and not args.delete_unlisted:
            info(f"  Note: {len(unlisted)} PSK(s) on these sites are not in the PSKs CSV "
                 "(left alone - use --delete-unlisted to remove them)")
        if deletions and args.dry_run:
            info("Would delete (dry run):")
            for d in deletions:
                info(f"  - {describe_psk(d)}")
        elif deletions:
            confirm_deletions(deletions, args.yes)
        for p in plan:
            if p.action == "existing":
                p.status = "already existed - kept"

        writes = counts["create"] + counts["rotate"] + len(deletions)
        if not args.dry_run and writes:
            check_usage_headroom(api, writes)

        def usage_rows() -> list:
            u_end = api.usage()
            show_usage("at end", u_end)
            rows = [("API calls made by this run", api.total_calls)]
            rows += [(f"  {m}", n) for m, n in sorted(api.calls.items())]
            if usage_start:
                rows += [("Token calls used this hour - at start", usage_start["used"])]
            if u_end:
                rows += [("Token calls used this hour - at end", u_end["used"]),
                         ("Token hourly limit", u_end["limit"]),
                         ("Token calls left this hour - at end", u_end["left"])]
            rows += [("Note", "Token usage covers every script/user of this token, not only this run")]
            info(f"API calls made by this run: {api.total_calls} ("
                 + ", ".join(f"{m} {n}" for m, n in sorted(api.calls.items())) + ")")
            return rows

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_dir = cfg.output_dir / (f"{stamp}_dryrun" if args.dry_run else stamp)
        out_dir.mkdir(parents=True, exist_ok=True)
        xlsx = out_dir / f"mpsk_keys_{stamp}.xlsx"

        if not args.dry_run and (counts["create"] or counts["rotate"]):
            info("Writing PSKs to Mist ...")
            try:
                apply_plan(api, plan)
            except MpskError:
                # Keep a record of what did get written, so no key is lost.
                for p in plan:
                    if p.status == "pending":
                        p.status = "NOT written (run stopped)"
                partial = out_dir / f"mpsk_keys_{stamp}_INCOMPLETE.xlsx"
                write_excel(plan, partial, dry_run=False, deletions=deletions, usage=usage_rows())
                print(f"Partial record of written keys saved to {partial}", file=sys.stderr)
                raise

        # Deletions run after creates/rotates, so a failed create never leaves a site short.
        if not args.dry_run and deletions:
            info("Deleting unlisted PSKs ...")
            try:
                apply_deletions(api, deletions)
            except MpskError:
                for d in deletions:
                    if d.status == "pending":
                        d.status = "NOT deleted (run stopped)"
                partial = out_dir / f"mpsk_keys_{stamp}_INCOMPLETE.xlsx"
                write_excel(plan, partial, dry_run=False, deletions=deletions, usage=usage_rows())
                print(f"Partial record saved to {partial}", file=sys.stderr)
                raise

        write_excel(plan, xlsx, args.dry_run, deletions, usage_rows())

        by_site: dict[str, list[PlanItem]] = {}
        for p in plan:
            by_site.setdefault(p.site_name, []).append(p)
        for site, items in by_site.items():
            write_site_pdf(site, items, out_dir / f"{safe_filename(site)}.pdf", cfg)

        info(f"Done. Output in {out_dir}")
        info(f"  {xlsx.name}  +  {len(by_site)} site PDF(s)")
        info("  These files contain plaintext Wi-Fi keys - store and share them carefully.")
        return 0
    except MpskError as e:
        print(f"\nERROR: {e}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("\nCancelled.", file=sys.stderr)
        return 130


if __name__ == "__main__":
    sys.exit(main())
