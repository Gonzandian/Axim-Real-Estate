"""
AXIM Real Estate Partners — Single-Property Due-Diligence Pull
===============================================================
Reusable tool: given ONE NYC property (address or BBL), retrieve everything on
Jake's list and write it into a formatted Excel workbook.

ALL dataset IDs, field names, and query filters below were verified against the
live NYC Open Data / DOF endpoints on 2026-06-26 using a real Brooklyn property
(925 Lenox Road, BBL 3046420025). See README for the verification notes.

Sources (verified)
-------------------
  DOF   Market value & assessment (current, 5+ yrs) .. 8y4t-faws  (parid filter)
        Exemptions ................................... muvi-b6kx  (parid filter)
        Abatements ................................... rgyu-ii48  (parid filter)
        Tax bill (Statement of Account) + balance .... a836-edms.nyc.gov  (SOA PDF)
  DOB   Certificate of Occupancy ..................... bs8b-p36w  (bin / bbl)
        DOB violations (open) ........................ 3h2n-5cm9  (bin)
        ECB / environmental violations (open) ........ 6bgk-3dad  (bin)
  HPD   Housing maintenance violations (open) ........ wvxf-dwi5  (boroid/block/lot)
  ZoLa  Zoning, lot dims, units, age, size ........... 64uk-42ks  (borocode/block/lot)
  ACRIS Mortgage / deed / owner ...................... bnx9-e6tj, 8h5j-fqxa, 636b-3b5g
  Geocode (address -> BBL/BIN) ....................... geosearch.planninglabs.nyc

NOTE: the older assessment dataset yjxr-fw8i is FROZEN at tax year 2018/19 — do
not use it. 8y4t-faws is the current roll (verified through tax year 2027).

Design
------
* Every source is wrapped so one failure never kills the run; status is recorded
  on the workbook's Run_Log tab.
* Requires normal outbound network — run in Google Colab / locally.
"""

import re
import sys
import time
import datetime as _dt
from collections import OrderedDict

import requests

# ----------------------------------------------------------------------------
# Config
# ----------------------------------------------------------------------------
SOCRATA_HOST = "https://data.cityofnewyork.us"
GEOSEARCH = "https://geosearch.planninglabs.nyc/v2/search"
EDMS = "https://a836-edms.nyc.gov/dctm-rest/repositories/dofedmspts/StatementSearch"

APP_TOKEN = ""  # optional Socrata token; not required for single-property pulls

DATASETS = {
    "pluto":         "64uk-42ks",
    "assessment":    "8y4t-faws",   # current roll (NOT yjxr-fw8i, which is frozen)
    "exemption":     "muvi-b6kx",
    "abatement":     "rgyu-ii48",
    "cofo":          "bs8b-p36w",
    "dob_viol":      "3h2n-5cm9",
    "ecb_viol":      "6bgk-3dad",
    "hpd_viol":      "wvxf-dwi5",
    "acris_master":  "bnx9-e6tj",
    "acris_legals":  "8h5j-fqxa",
    "acris_parties": "636b-3b5g",
}

BOROUGH_NAME = {1: "Manhattan", 2: "Bronx", 3: "Brooklyn", 4: "Queens", 5: "Staten Island"}

PARTY_GRANTOR = "1"   # seller / mortgagor
PARTY_GRANTEE = "2"   # buyer / lender (mortgagee)


# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def _headers():
    h = {"User-Agent": "AXIM-PropertyDD/1.1"}
    if APP_TOKEN:
        h["X-App-Token"] = APP_TOKEN
    return h


def soda_get(resource, where=None, select=None, order=None, limit=5000, extra=None):
    """Query a Socrata resource -> list[dict]."""
    params = {"$limit": limit}
    if where:
        params["$where"] = where
    if select:
        params["$select"] = select
    if order:
        params["$order"] = order
    if extra:
        params.update(extra)
    r = requests.get(f"{SOCRATA_HOST}/resource/{resource}.json",
                     params=params, headers=_headers(), timeout=60)
    r.raise_for_status()
    return r.json()


def pick(row, *names, default=None):
    if not row:
        return default
    low = {k.lower(): v for k, v in row.items()}
    for n in names:
        v = low.get(n.lower())
        if v not in (None, "", " "):
            return v
    return default


def _num(x):
    if x in (None, "", " "):
        return None
    try:
        return float(str(x).replace(",", "").replace("$", ""))
    except ValueError:
        return None


def _money(x):
    v = _num(x)
    return f"${v:,.0f}" if v is not None else ""


def _clean(x):
    """PLUTO stores numbers as strings with trailing zeros ('28.0000000'). Tidy them."""
    v = _num(x)
    if v is None:
        return x
    return int(v) if v == int(v) else round(v, 2)


# ----------------------------------------------------------------------------
# Geocode: address / BBL -> identifiers
# ----------------------------------------------------------------------------
def normalize_bbl(bbl):
    s = re.sub(r"\D", "", str(bbl))
    if len(s) != 10:
        raise ValueError(f"BBL must be 10 digits, got {bbl!r}")
    return int(s[0]), int(s[1:6]), int(s[6:10]), s  # boro, block, lot, bbl10


def resolve_property(query):
    """
    Accept a 10-digit BBL or a free-text address. Returns an identifier dict.
    Surfaces the geocoder's match quality so a fallback/approximate hit is visible.
    """
    q = str(query).strip()
    digits = re.sub(r"\D", "", q)
    ident = {"input": q, "bin": None, "address": None, "lat": None, "lon": None,
             "geocode_match": "exact (BBL supplied)", "geocode_confidence": ""}

    if len(digits) == 10 and len(q) <= 14:                      # looks like a BBL
        boro, block, lot, bbl10 = normalize_bbl(digits)
        ident.update(boro=boro, block=block, lot=lot, bbl=bbl10)
        return ident

    r = requests.get(GEOSEARCH, params={"text": q, "size": 1},
                     headers=_headers(), timeout=30)
    r.raise_for_status()
    feats = r.json().get("features", [])
    if not feats:
        raise ValueError(f"GeoSearch found no match for {q!r}")
    f = feats[0]
    props = f.get("properties", {})
    pad = (props.get("addendum", {}) or {}).get("pad", {}) or {}
    bbl10 = re.sub(r"\D", "", str(pad.get("bbl", "")))
    if len(bbl10) != 10:
        raise ValueError(f"GeoSearch returned no usable BBL for {q!r}: {pad}")
    boro, block, lot, bbl10 = normalize_bbl(bbl10)
    geom = f.get("geometry", {}).get("coordinates", [None, None])
    ident.update(
        boro=boro, block=block, lot=lot, bbl=bbl10,
        bin=pad.get("bin"),
        address=props.get("label"),
        lon=geom[0], lat=geom[1],
        geocode_match=props.get("match_type", "?"),
        geocode_confidence=props.get("confidence", ""),
    )
    return ident


# ----------------------------------------------------------------------------
# ZoLa / PLUTO  -> zoning, lot, units, age, size, owner
# ----------------------------------------------------------------------------
def get_pluto(ident):
    # PLUTO stores bbl as a decimal string ('3046420025.00000000'), so filter on
    # borocode + block + lot instead (all verified plain-integer text fields).
    rows = soda_get(DATASETS["pluto"],
                    where=(f"borocode='{ident['boro']}' AND block='{ident['block']}' "
                           f"AND lot='{ident['lot']}'"), limit=2)
    if not rows:
        return {}, {}
    r = rows[0]
    zoning = ", ".join([z for z in [pick(r, "zonedist1"), pick(r, "zonedist2"),
                                    pick(r, "zonedist3"), pick(r, "zonedist4")] if z])
    overlays = ", ".join([z for z in [pick(r, "overlay1"), pick(r, "overlay2")] if z])
    summary = OrderedDict([
        ("Zoning district", zoning),
        ("Commercial overlay", overlays),
        ("Lot frontage (ft)", _clean(pick(r, "lotfront"))),
        ("Lot depth (ft)", _clean(pick(r, "lotdepth"))),
        ("Lot area (sf)", _clean(pick(r, "lotarea"))),
        ("Residential units", _clean(pick(r, "unitsres"))),
        ("Total units", _clean(pick(r, "unitstotal"))),
        ("Year built", pick(r, "yearbuilt")),
        ("Year altered", pick(r, "yearalter1")),
        ("Building area (gsf)", _clean(pick(r, "bldgarea"))),
        ("Residential area (sf)", _clean(pick(r, "resarea"))),
        ("Number of floors", _clean(pick(r, "numfloors"))),
        ("Number of buildings", _clean(pick(r, "numbldgs"))),
        ("Building class", pick(r, "bldgclass")),
        ("Land use", pick(r, "landuse")),
        ("Built FAR", _clean(pick(r, "builtfar"))),
        ("Max residential FAR", _clean(pick(r, "residfar"))),
        ("Owner (PLUTO)", pick(r, "ownername")),
        ("PLUTO address", pick(r, "address")),
    ])
    return summary, r


# ----------------------------------------------------------------------------
# DOF assessments (current roll, 5+ yrs) and benefits
# ----------------------------------------------------------------------------
def get_assessments(ident, years=6):
    rows = soda_get(DATASETS["assessment"], where=f"parid='{ident['bbl']}'",
                    order="year DESC", limit=200)
    table = []
    for r in rows:
        table.append(OrderedDict([
            ("Tax year", pick(r, "year")),
            ("Tax class", pick(r, "curtaxclass", "pytaxclass")),
            ("Market value (est.)", _num(pick(r, "curmkttot", "finmkttot", "tenmkttot"))),
            ("Assessed land", _num(pick(r, "curactland", "finactland"))),
            ("Assessed total", _num(pick(r, "curacttot", "finacttot", "tenacttot"))),
            ("Exempt total", _num(pick(r, "curactextot", "finactextot"))),
            ("Bldg class", pick(r, "bldg_class")),
        ]))
    seen, dedup = set(), []          # the dataset has duplicate rows per year
    for row in table:
        if row["Tax year"] not in seen:
            seen.add(row["Tax year"])
            dedup.append(row)
    return dedup[:years]


def get_benefits(ident):
    out = []
    # Exemptions (Property Exemption Detail) — keyed on parid
    try:
        for r in soda_get(DATASETS["exemption"], where=f"parid='{ident['bbl']}'",
                          order="year DESC", limit=500):
            out.append(OrderedDict([
                ("Type", "Exemption"),
                ("Benefit / holder", pick(r, "exname", default="")),
                ("Code", pick(r, "exmp_code", default="")),
                ("Tax year", pick(r, "year", default="")),
                ("Amount", _num(pick(r, "curexmptot", "tenexmptot"))),
            ]))
    except Exception:
        pass
    # Abatements (Property Abatement Detail) — keyed on parid; tccode = benefit type
    try:
        for r in soda_get(DATASETS["abatement"], where=f"parid='{ident['bbl']}'",
                          order="taxyr DESC", limit=500):
            out.append(OrderedDict([
                ("Type", "Abatement"),
                ("Benefit / holder", " / ".join([x for x in [pick(r, "tccode"),
                                                              pick(r, "tcsubcode")] if x])),
                ("Code", pick(r, "tccode", default="")),
                ("Tax year", pick(r, "taxyr", default="")),
                ("Amount", _num(pick(r, "appliedabt", "actabtpyr"))),
            ]))
    except Exception:
        pass
    return out


# ----------------------------------------------------------------------------
# DOB: Certificate of Occupancy + violations; HPD violations
# ----------------------------------------------------------------------------
def get_cofo(ident):
    rows = []
    if ident.get("bin"):
        rows = soda_get(DATASETS["cofo"], where=f"bin='{ident['bin']}'",
                        order="c_o_issue_date DESC", limit=50)
    if not rows:
        rows = soda_get(DATASETS["cofo"], where=f"bbl='{ident['bbl']}'",
                        order="c_o_issue_date DESC", limit=50)
    if not rows:
        return {}, []
    cur = rows[0]
    summary = OrderedDict([
        ("C of O issue date", (pick(cur, "c_o_issue_date", default="") or "")[:10]),
        ("Job number", pick(cur, "job_number")),
        ("Job type", pick(cur, "job_type")),
        ("Issue type", pick(cur, "issue_type")),
        ("Status", pick(cur, "application_status_raw", "filing_status_raw")),
    ])
    return summary, rows


def get_dob_violations(ident):
    if not ident.get("bin"):
        return []                       # DOB violations are keyed on BIN
    rows = soda_get(DATASETS["dob_viol"], where=f"bin='{ident['bin']}'", limit=2000)
    out = []
    for r in rows:
        cat = (pick(r, "violation_category", default="") or "").upper()
        if "RESOLVE" in cat or "DISMISS" in cat:
            continue
        out.append(OrderedDict([
            ("Agency", "DOB"),
            ("Type", pick(r, "violation_type", "violation_type_code", default="")),
            ("Issued", pick(r, "issue_date", default="")),
            ("Status", pick(r, "violation_category", default="")),
            ("Description", (pick(r, "description", default="") or "")[:300]),
            ("Number", pick(r, "violation_number", "isn_dob_bis_viol", default="")),
        ]))
    return out


def get_ecb_violations(ident):
    if not ident.get("bin"):
        return []
    rows = soda_get(DATASETS["ecb_viol"], where=f"bin='{ident['bin']}'", limit=2000)
    out = []
    for r in rows:
        st = (pick(r, "ecb_violation_status", default="") or "").upper()
        if st and st != "ACTIVE":
            continue
        out.append(OrderedDict([
            ("Agency", "DOB/ECB"),
            ("Type", pick(r, "violation_type", default="")),
            ("Issued", pick(r, "issue_date", default="")),
            ("Status", pick(r, "ecb_violation_status", default="")),
            ("Description", (pick(r, "violation_description",
                                  "section_law_description1", default="") or "")[:300]),
            ("Number", pick(r, "ecb_violation_number", default="")),
            ("Penalty due", _money(pick(r, "balance_due", "penality_imposed"))),
        ]))
    return out


def get_hpd_violations(ident):
    rows = soda_get(DATASETS["hpd_viol"],
                    where=(f"boroid='{ident['boro']}' AND block='{ident['block']}' "
                           f"AND lot='{ident['lot']}'"), limit=5000)
    out = []
    for r in rows:
        st = (pick(r, "violationstatus", "currentstatus", default="") or "").upper()
        if "OPEN" not in st:
            continue
        out.append(OrderedDict([
            ("Agency", "HPD"),
            ("Class", pick(r, "class", default="")),
            ("Issued", (pick(r, "novissueddate", "inspectiondate", default="") or "")[:10]),
            ("Status", pick(r, "violationstatus", default="")),
            ("Description", (pick(r, "novdescription", default="") or "")[:300]),
            ("Number", pick(r, "violationid", default="")),
        ]))
    return out


# ----------------------------------------------------------------------------
# ACRIS: current mortgage, last deed, current owner entity
# ----------------------------------------------------------------------------
def _acris_doc_ids(ident):
    legals = soda_get(DATASETS["acris_legals"],
                      where=(f"borough={ident['boro']} AND block={ident['block']} "
                             f"AND lot={ident['lot']}"),
                      select="document_id", limit=10000)
    return list({r.get("document_id") for r in legals if r.get("document_id")})


def _acris_masters(doc_ids):
    masters = []
    for i in range(0, len(doc_ids), 75):
        quoted = ",".join(f"'{d}'" for d in doc_ids[i:i + 75])
        masters += soda_get(DATASETS["acris_master"],
                            where=f"document_id in ({quoted})", limit=5000)
    return masters


def _acris_parties(doc_id):
    return soda_get(DATASETS["acris_parties"],
                    where=f"document_id='{doc_id}'", limit=200)


def _latest(masters, doc_types):
    cand = [m for m in masters if (m.get("doc_type") or "").upper() in doc_types]
    cand.sort(key=lambda m: m.get("recorded_datetime") or m.get("document_date") or "",
              reverse=True)
    return cand[0] if cand else None


def get_acris(ident):
    summary, txns = OrderedDict(), []
    doc_ids = _acris_doc_ids(ident)
    if not doc_ids:
        return summary, txns
    masters = _acris_masters(doc_ids)

    for m in sorted(masters, key=lambda m: m.get("recorded_datetime") or "", reverse=True):
        txns.append(OrderedDict([
            ("Doc type", pick(m, "doc_type")),
            ("Doc date", (pick(m, "document_date", default="") or "")[:10]),
            ("Recorded", (pick(m, "recorded_datetime", default="") or "")[:10]),
            ("Amount", _num(pick(m, "document_amt"))),
            ("Doc ID", pick(m, "document_id")),
        ]))

    mtg = _latest(masters, {"MTGE", "AGMT", "MODA", "CEMA"})
    if mtg:
        parties = _acris_parties(mtg["document_id"])
        lender = next((p.get("name") for p in parties
                       if str(p.get("party_type")) == PARTY_GRANTEE), None)
        summary["Mortgage lender"] = lender or ""
        summary["Mortgage amount"] = _money(pick(mtg, "document_amt"))
        summary["Mortgage doc type"] = pick(mtg, "doc_type")
        summary["Mortgage recorded"] = (pick(mtg, "recorded_datetime", default="") or "")[:10]
        summary["Mortgage doc ID"] = pick(mtg, "document_id")

    deed = _latest(masters, {"DEED", "DEEDO", "RPTT&RET"})
    if deed:
        parties = _acris_parties(deed["document_id"])
        grantee = next((p.get("name") for p in parties
                        if str(p.get("party_type")) == PARTY_GRANTEE), None)
        grantor = next((p.get("name") for p in parties
                        if str(p.get("party_type")) == PARTY_GRANTOR), None)
        summary["Current owner entity (last grantee)"] = grantee or ""
        summary["Prior owner (grantor)"] = grantor or ""
        summary["Last deed price"] = _money(pick(deed, "document_amt"))
        summary["Last deed recorded"] = (pick(deed, "recorded_datetime", default="") or "")[:10]
        summary["Last deed doc ID"] = pick(deed, "document_id")
    return summary, txns


# ----------------------------------------------------------------------------
# DOF EDMS: Statement of Account (tax bill) + account balance
# ----------------------------------------------------------------------------
def _stmt_candidates():
    """
    Real DOF quarterly SOA statement dates cluster near mid-Feb, early-Jun,
    mid-Aug and mid-Nov (verified examples: 20250215, 20241116, 20230603).
    Probe a window around each, most-recent-first, current + prior 2 years.
    """
    bases = [(2, 15), (6, 1), (8, 15), (11, 15)]
    offsets = [0, 1, 2, 3, -1, 4, 5, 6]
    today = _dt.date.today()
    dates = []
    for yr in (today.year, today.year - 1, today.year - 2):
        for mo, dy in bases:
            for off in offsets:
                try:
                    d = _dt.date(yr, mo, dy) + _dt.timedelta(days=off)
                except ValueError:
                    continue
                if d <= today:
                    dates.append(d)
    dates = sorted(set(dates), reverse=True)
    return [d.strftime("%Y%m%d") for d in dates]


def get_tax_statement(ident, save_path=None, max_probes=64):
    """
    Fetch the most recent DOF Statement of Account PDF for the BBL and parse the
    amount due. Best-effort: probes real statement-date windows until one returns
    a genuine PDF (the endpoint returns 'No data found.' for non-statement dates).
    """
    info = OrderedDict([("Statement found", "No"), ("Statement date", ""),
                        ("Statement URL", ""), ("Amount due", ""),
                        ("Est. market value (from bill)", "")])
    for n, stmt_date in enumerate(_stmt_candidates()):
        if n >= max_probes:
            break
        url = f"{EDMS}?bbl={ident['bbl']}&stmtDate={stmt_date}&stmtType=SOA"
        try:
            r = requests.get(url, headers=_headers(), timeout=45)
        except Exception:
            continue
        body = r.content or b""
        if r.status_code == 200 and len(body) > 800 and b"No data found" not in body:
            info["Statement found"] = "Yes"
            info["Statement date"] = f"{stmt_date[:4]}-{stmt_date[4:6]}-{stmt_date[6:]}"
            info["Statement URL"] = url
            if save_path:
                try:
                    with open(save_path, "wb") as fh:
                        fh.write(body)
                except Exception:
                    pass
            text = _pdf_text(body)
            due = _find_amount(text, r"total amount due[^$]*\$\s*([\d,]+\.\d{2})") \
                or _find_amount(text, r"amount due[^$]*\$\s*([\d,]+\.\d{2})")
            if due is not None:
                info["Amount due"] = _money(due)
            mkt = _find_amount(text, r"market value[:\s]*\$\s*([\d,]+)")
            if mkt is not None:
                info["Est. market value (from bill)"] = _money(mkt)
            return info, body
    return info, None


def _pdf_text(pdf_bytes):
    try:
        import pdfplumber, io
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            return "\n".join((p.extract_text() or "") for p in pdf.pages)
    except Exception:
        # Some EDMS responses are already plain text rather than binary PDF
        try:
            return pdf_bytes.decode("utf-8", errors="ignore")
        except Exception:
            return ""


def _find_amount(text, pattern):
    m = re.search(pattern, text or "", re.I)
    return _num(m.group(1)) if m else None


# ----------------------------------------------------------------------------
# Orchestrator
# ----------------------------------------------------------------------------
def run_property(query, out_xlsx=None, save_pdf=True):
    log = []

    def step(name, fn):
        t0 = time.time()
        try:
            res = fn()
            if isinstance(res, tuple):
                n = len(res[0]) if isinstance(res[0], (list, dict)) else 1
            elif isinstance(res, (list, dict)):
                n = len(res)
            else:
                n = 1
            log.append((name, "OK", n, round(time.time() - t0, 1), ""))
            return res
        except Exception as e:
            log.append((name, "FAILED", 0, round(time.time() - t0, 1), str(e)[:200]))
            return None if name == "Geocode" else ({}, [])

    ident = step("Geocode", lambda: resolve_property(query))
    if not ident:
        raise RuntimeError("Geocoding failed; cannot continue.")

    def first(x):  # normalize ({},[]) fallback to the right shape
        return x

    pluto_sum, pluto_raw = step("PLUTO/ZoLa", lambda: get_pluto(ident)) or ({}, {})
    assess = step("DOF assessments", lambda: get_assessments(ident))
    assess = assess if isinstance(assess, list) else []
    benefits = step("DOF benefits", lambda: get_benefits(ident))
    benefits = benefits if isinstance(benefits, list) else []
    cofo_sum, cofo_rows = step("DOB C of O", lambda: get_cofo(ident)) or ({}, [])
    dobv = step("DOB violations", lambda: get_dob_violations(ident))
    dobv = dobv if isinstance(dobv, list) else []
    ecbv = step("ECB/environmental", lambda: get_ecb_violations(ident))
    ecbv = ecbv if isinstance(ecbv, list) else []
    hpdv = step("HPD violations", lambda: get_hpd_violations(ident))
    hpdv = hpdv if isinstance(hpdv, list) else []
    acris_sum, acris_txns = step("ACRIS", lambda: get_acris(ident)) or ({}, [])

    pdf_path = out_xlsx.replace(".xlsx", "_SOA.pdf") if (save_pdf and out_xlsx) else None
    tax_info, _ = step("DOF tax statement",
                       lambda: get_tax_statement(ident, pdf_path)) or ({}, None)

    result = dict(ident=ident, pluto_sum=pluto_sum or {}, pluto_raw=pluto_raw or {},
                  assess=assess, benefits=benefits, cofo_sum=cofo_sum or {},
                  cofo_rows=cofo_rows or [], dobv=dobv, ecbv=ecbv, hpdv=hpdv,
                  acris_sum=acris_sum or {}, acris_txns=acris_txns or [],
                  tax_info=tax_info or {}, log=log)

    if out_xlsx:
        from axim_xlsx_writer import write_workbook
        write_workbook(result, out_xlsx)
        result["xlsx"] = out_xlsx
    return result


if __name__ == "__main__":
    q = sys.argv[1] if len(sys.argv) > 1 else "925 Lenox Road, Brooklyn"
    out = f"AXIM_PropertyDD_{re.sub(r'[^A-Za-z0-9]+','_', q)[:40]}.xlsx"
    res = run_property(q, out_xlsx=out)
    print("\nRun log:")
    for row in res["log"]:
        print(f"  {row[0]:<22} {row[1]:<8} rows={row[2]:<4} {row[3]}s  {row[4]}")
    print(f"\nWrote {out}")
