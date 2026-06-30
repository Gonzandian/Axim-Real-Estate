"""
AXIM Real Estate Partners — Property Due-Diligence Tool
Streamlit web app: enter an address or BBL, get an Excel workbook.
"""

import io
import re
import tempfile
import traceback

import streamlit as st

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="AXIM Real Estate Partners — Property Due-Diligence",
    page_icon="A",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ── Styling (investment-bank aesthetic) ────────────────────────────────────────
st.markdown(
    """
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@500;600;700&family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <style>
        /* ── Global ── */
        html, body, [class*="st-"], .stMarkdown, p, span, div, label, input, button {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
            color: #0B1F3A;
        }
        .stApp { background: #FBFAF7; }
        .block-container { max-width: 780px; padding-top: 1rem; padding-bottom: 4rem; }
        #MainMenu, footer, header[data-testid="stHeader"] { visibility: hidden; }

        /* ── Brand bar ── */
        .axim-bar {
            border-bottom: 1px solid #1A1A1A;
            padding: 14px 0 12px 0;
            margin-bottom: 36px;
            display: flex;
            justify-content: space-between;
            align-items: baseline;
        }
        .axim-wordmark {
            font-family: 'Inter', sans-serif;
            font-weight: 600;
            font-size: 13px;
            letter-spacing: 0.22em;
            color: #0B1F3A;
            text-transform: uppercase;
        }
        .axim-bar-meta {
            font-family: 'Inter', sans-serif;
            font-weight: 400;
            font-size: 11px;
            letter-spacing: 0.16em;
            color: #6B6B6B;
            text-transform: uppercase;
        }

        /* ── Headlines ── */
        h1, .axim-h1 {
            font-family: 'Playfair Display', 'Georgia', serif !important;
            font-weight: 600 !important;
            color: #0B1F3A !important;
            font-size: 2.5rem !important;
            line-height: 1.15 !important;
            letter-spacing: -0.01em !important;
            margin: 0 0 0.5rem 0 !important;
            padding: 0 !important;
        }
        .axim-eyebrow {
            font-family: 'Inter', sans-serif;
            font-weight: 600;
            font-size: 11px;
            letter-spacing: 0.22em;
            color: #B89968;
            text-transform: uppercase;
            margin-bottom: 0.6rem;
        }
        .axim-lede {
            font-family: 'Inter', sans-serif;
            font-weight: 400;
            font-size: 0.95rem;
            color: #4A5568;
            line-height: 1.55;
            margin-bottom: 2rem;
            max-width: 640px;
        }
        h2, h3 {
            font-family: 'Playfair Display', 'Georgia', serif !important;
            font-weight: 600 !important;
            color: #0B1F3A !important;
            letter-spacing: -0.005em !important;
        }

        /* ── Input ── */
        .stTextInput > div > div > input {
            font-family: 'Inter', sans-serif;
            font-size: 0.95rem;
            color: #0B1F3A;
            background: #FFFFFF;
            border: 1px solid #D1D5DB;
            border-radius: 2px;
            padding: 0.75rem 1rem;
        }
        .stTextInput > div > div > input:focus {
            border-color: #0B1F3A;
            box-shadow: 0 0 0 1px #0B1F3A;
            outline: none;
        }

        /* ── Button ── */
        .stButton > button {
            background-color: #0B1F3A;
            color: #FBFAF7;
            font-family: 'Inter', sans-serif;
            font-weight: 500;
            font-size: 0.85rem;
            letter-spacing: 0.14em;
            text-transform: uppercase;
            width: 100%;
            padding: 0.85rem 1rem;
            border-radius: 2px;
            border: 1px solid #0B1F3A;
            transition: background 0.15s ease;
        }
        .stButton > button:hover:not([disabled]) {
            background-color: #B89968;
            border-color: #B89968;
            color: #0B1F3A;
        }
        .stButton > button[disabled] {
            background-color: #E5E7EB;
            border-color: #E5E7EB;
            color: #9CA3AF;
        }

        /* ── Download button (secondary style) ── */
        .stDownloadButton > button {
            background-color: #B89968;
            color: #0B1F3A;
            font-family: 'Inter', sans-serif;
            font-weight: 600;
            font-size: 0.85rem;
            letter-spacing: 0.14em;
            text-transform: uppercase;
            width: 100%;
            padding: 0.85rem 1rem;
            border-radius: 2px;
            border: 1px solid #B89968;
        }
        .stDownloadButton > button:hover {
            background-color: #0B1F3A;
            color: #FBFAF7;
            border-color: #0B1F3A;
        }

        /* ── Metrics ── */
        [data-testid="stMetric"] {
            background: #FFFFFF;
            border: 1px solid #E5E7EB;
            border-radius: 2px;
            padding: 1rem 1.1rem;
        }
        [data-testid="stMetricLabel"] {
            font-family: 'Inter', sans-serif !important;
            font-size: 10px !important;
            letter-spacing: 0.18em !important;
            text-transform: uppercase !important;
            color: #6B6B6B !important;
            font-weight: 500 !important;
        }
        [data-testid="stMetricValue"] {
            font-family: 'Playfair Display', 'Georgia', serif !important;
            font-size: 1.55rem !important;
            font-weight: 600 !important;
            color: #0B1F3A !important;
            letter-spacing: -0.01em !important;
        }

        /* ── Success/info alerts ── */
        [data-testid="stAlert"] {
            border-radius: 2px;
            border-left-width: 3px;
        }

        /* ── Divider ── */
        hr {
            border: none;
            border-top: 1px solid #E5E7EB;
            margin: 2rem 0;
        }

        /* ── Progress bar ── */
        [data-testid="stProgressBar"] > div > div > div {
            background-color: #0B1F3A;
        }

        /* ── Expander ── */
        [data-testid="stExpander"] summary {
            font-family: 'Inter', sans-serif;
            font-size: 12px;
            letter-spacing: 0.14em;
            text-transform: uppercase;
            color: #6B6B6B;
            font-weight: 500;
        }
    </style>

    <div class="axim-bar">
        <div class="axim-wordmark">AXIM Real Estate Partners</div>
        <div class="axim-bar-meta">New York · Property Intelligence</div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="axim-eyebrow">Property Due-Diligence</div>', unsafe_allow_html=True)
st.markdown('<h1 class="axim-h1">NYC Property Intelligence Workbook</h1>', unsafe_allow_html=True)
st.markdown(
    '<p class="axim-lede">Enter a New York City address or 10-digit BBL to compile a full '
    'due-diligence package — title, zoning, assessments, tax benefits, violations, '
    'and ownership history — delivered as a single Excel workbook.</p>',
    unsafe_allow_html=True,
)

# ── Input ──────────────────────────────────────────────────────────────────────
query = st.text_input(
    "Address or BBL",
    placeholder="925 Lenox Road, Brooklyn   ·   or   ·   3046420025",
    label_visibility="collapsed",
)

run = st.button("Compile Due-Diligence Package", disabled=not query.strip())

# ── Run ────────────────────────────────────────────────────────────────────────
if run and query.strip():
    from axim_property_dd import run_property

    steps = [
        "Geocode",
        "PLUTO/ZoLa",
        "DOF assessments",
        "DOF benefits",
        "DOB C of O",
        "DOB violations",
        "ECB/environmental",
        "HPD violations",
        "ACRIS",
        "DOF tax statement",
    ]
    n_steps = len(steps)

    progress = st.progress(0, text="Starting…")
    status_box = st.empty()

    # Monkey-patch the step logger so we can update the progress bar live
    import axim_property_dd as _dd

    _step_idx = [0]
    _orig_run = _dd.run_property  # keep reference (not used after override)

    # We'll run step-by-step by calling the underlying fetchers directly,
    # updating progress between each one, then write the workbook.
    result = {}

    def _tick(label):
        _step_idx[0] += 1
        pct = min(_step_idx[0] / n_steps, 1.0)
        progress.progress(pct, text=f"Fetching: {label}…")

    try:
        # 1. Geocode
        status_box.info("Resolving address…")
        ident = _dd.resolve_property(query.strip())
        _tick("Geocode")
        log = [("Geocode", "OK", 1, 0, "")]

        def _step(name, fn):
            _tick(name)
            status_box.info(f"{name}…")
            t0 = __import__("time").time()
            try:
                res = fn()
                n = len(res[0]) if isinstance(res, tuple) and isinstance(res[0], (list, dict)) \
                    else (len(res) if isinstance(res, (list, dict)) else 1)
                log.append((name, "OK", n, round(__import__("time").time() - t0, 1), ""))
                return res
            except Exception as exc:
                log.append((name, "FAILED", 0,
                            round(__import__("time").time() - t0, 1), str(exc)[:200]))
                return ({}, []) if name != "Geocode" else {}

        pluto_sum, pluto_raw = _step("PLUTO/ZoLa",   lambda: _dd.get_pluto(ident))
        assess              = _step("DOF assessments",lambda: _dd.get_assessments(ident))
        benefits            = _step("DOF benefits",   lambda: _dd.get_benefits(ident))
        cofo_sum, cofo_rows = _step("DOB C of O",    lambda: _dd.get_cofo(ident))
        dobv                = _step("DOB violations", lambda: _dd.get_dob_violations(ident))
        ecbv                = _step("ECB/environmental", lambda: _dd.get_ecb_violations(ident))
        hpdv                = _step("HPD violations", lambda: _dd.get_hpd_violations(ident))
        acris_sum, acris_txns = _step("ACRIS",       lambda: _dd.get_acris(ident))

        # Tax statement (slow — up to 64 probes)
        _tick("DOF tax statement")
        status_box.info("Fetching DOF tax statement (may take ~30 seconds)…")
        import time as _time
        t0 = _time.time()
        try:
            tax_info, _ = _dd.get_tax_statement(ident)
            log.append(("DOF tax statement", "OK", 1, round(_time.time() - t0, 1), ""))
        except Exception as exc:
            tax_info = {}
            log.append(("DOF tax statement", "FAILED", 0, round(_time.time() - t0, 1), str(exc)[:200]))

        progress.progress(1.0, text="Building workbook…")
        status_box.info("Writing Excel workbook…")

        res = dict(
            ident=ident,
            pluto_sum=pluto_sum or {},
            pluto_raw=pluto_raw or {},
            assess=assess if isinstance(assess, list) else [],
            benefits=benefits if isinstance(benefits, list) else [],
            cofo_sum=cofo_sum or {},
            cofo_rows=cofo_rows or [],
            dobv=dobv if isinstance(dobv, list) else [],
            ecbv=ecbv if isinstance(ecbv, list) else [],
            hpdv=hpdv if isinstance(hpdv, list) else [],
            acris_sum=acris_sum or {},
            acris_txns=acris_txns or [],
            tax_info=tax_info or {},
            log=log,
        )

        # Write workbook to a temp file, then read back into BytesIO
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp_path = tmp.name
        from axim_xlsx_writer import write_workbook
        write_workbook(res, tmp_path)
        with open(tmp_path, "rb") as f:
            xlsx_bytes = f.read()

        progress.empty()
        status_box.empty()

        # ── Results summary ────────────────────────────────────────────────────
        st.success("Due-diligence package compiled.")

        address_label = (
            ident.get("address")
            or res["pluto_sum"].get("PLUTO address")
            or ident.get("input")
        )
        st.markdown(
            '<div class="axim-eyebrow" style="margin-top:1.5rem;">Subject Property</div>',
            unsafe_allow_html=True,
        )
        st.markdown(
            f'<h2 style="font-family:Playfair Display,Georgia,serif;'
            f'font-weight:600;color:#0B1F3A;margin:0 0 1.2rem 0;'
            f'font-size:1.7rem;letter-spacing:-0.005em;">{address_label or ""}</h2>',
            unsafe_allow_html=True,
        )

        from axim_property_dd import BOROUGH_NAME
        col1, col2, col3 = st.columns(3)
        col1.metric("Borough", BOROUGH_NAME.get(ident.get("boro"), "—"))
        col2.metric("BBL", ident.get("bbl") or "—")
        col3.metric("BIN", ident.get("bin") or "—")

        st.divider()

        # Valuation
        latest = res["assess"][0] if res["assess"] else {}
        mv = latest.get("Market value (est.)")
        c1, c2, c3 = st.columns(3)
        c1.metric("Est. market value", f"${mv:,.0f}" if mv else "—")
        c2.metric("Tax year", latest.get("Tax year") or "—")
        c3.metric(
            "Assessed total",
            (f"${latest.get('Assessed total'):,.0f}"
             if latest.get("Assessed total") else "—"),
        )

        # Violations
        total_v = len(res["dobv"]) + len(res["ecbv"]) + len(res["hpdv"])
        v1, v2, v3, v4 = st.columns(4)
        v1.metric("Open violations (total)", total_v)
        v2.metric("DOB", len(res["dobv"]))
        v3.metric("ECB", len(res["ecbv"]))
        v4.metric("HPD", len(res["hpdv"]))

        # Owner / mortgage
        if res["acris_sum"]:
            st.markdown("**Ownership & financing**")
            for k, v in res["acris_sum"].items():
                if v:
                    st.markdown(f"- **{k}:** {v}")

        st.divider()

        # ── Download ───────────────────────────────────────────────────────────
        safe_name = re.sub(r"[^A-Za-z0-9]+", "_", query.strip())[:40]
        filename = f"AXIM_PropertyDD_{safe_name}.xlsx"

        st.download_button(
            label="Download Excel workbook",
            data=xlsx_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Run log (collapsed)
        with st.expander("Run log"):
            for name, status, n, secs, err in log:
                marker = "OK " if status == "OK" else "FAIL"
                st.markdown(f"`{marker}` **{name}** — {n} rows, {secs}s  {err}")

        # ── Footer ────────────────────────────────────────────────────────────
        st.markdown(
            '<div style="margin-top:3rem;padding-top:1.5rem;border-top:1px solid #E5E7EB;'
            'font-family:Inter,sans-serif;font-size:11px;letter-spacing:0.14em;'
            'text-transform:uppercase;color:#9CA3AF;text-align:center;">'
            'AXIM Real Estate Partners · Confidential · Internal use only'
            '</div>',
            unsafe_allow_html=True,
        )

    except Exception:
        progress.empty()
        status_box.empty()
        st.error("Something went wrong. See details below.")
        st.code(traceback.format_exc())
