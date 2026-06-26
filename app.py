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
    page_title="AXIM Property DD",
    page_icon="🏢",
    layout="centered",
)

# ── Header ─────────────────────────────────────────────────────────────────────
st.markdown(
    """
    <style>
        .block-container { max-width: 720px; padding-top: 2rem; }
        h1 { color: #1F3864; }
        .stButton > button {
            background-color: #1F3864;
            color: white;
            font-weight: 600;
            width: 100%;
            padding: 0.6rem 1rem;
            border-radius: 6px;
            border: none;
        }
        .stButton > button:hover { background-color: #2c4f8c; }
        .metric-box {
            background: #f0f4fc;
            border-radius: 8px;
            padding: 0.8rem 1rem;
            margin-bottom: 0.5rem;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🏢 AXIM Property Due-Diligence")
st.caption("Enter a NYC address or 10-digit BBL to pull the full due-diligence package.")

# ── Input ──────────────────────────────────────────────────────────────────────
query = st.text_input(
    "Address or BBL",
    placeholder="e.g.  925 Lenox Road, Brooklyn  —or—  3046420025",
    label_visibility="collapsed",
)

run = st.button("Run due-diligence →", disabled=not query.strip())

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
        status_box.info("📍 Resolving address…")
        ident = _dd.resolve_property(query.strip())
        _tick("Geocode")
        log = [("Geocode", "OK", 1, 0, "")]

        def _step(name, fn):
            _tick(name)
            status_box.info(f"⏳ {name}…")
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
        status_box.info("⏳ Fetching DOF tax statement (may take ~30 seconds)…")
        import time as _time
        t0 = _time.time()
        try:
            tax_info, _ = _dd.get_tax_statement(ident)
            log.append(("DOF tax statement", "OK", 1, round(_time.time() - t0, 1), ""))
        except Exception as exc:
            tax_info = {}
            log.append(("DOF tax statement", "FAILED", 0, round(_time.time() - t0, 1), str(exc)[:200]))

        progress.progress(1.0, text="Building workbook…")
        status_box.info("📊 Writing Excel workbook…")

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
        st.success("✅ Done!")

        address_label = (
            ident.get("address")
            or res["pluto_sum"].get("PLUTO address")
            or ident.get("input")
        )
        st.subheader(address_label or "")

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
            label="⬇️  Download Excel workbook",
            data=xlsx_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Run log (collapsed)
        with st.expander("Run log"):
            for name, status, n, secs, err in log:
                icon = "✅" if status == "OK" else "❌"
                st.markdown(f"{icon} **{name}** — {n} rows, {secs}s  {err}")

    except Exception:
        progress.empty()
        status_box.empty()
        st.error("Something went wrong. See details below.")
        st.code(traceback.format_exc())
