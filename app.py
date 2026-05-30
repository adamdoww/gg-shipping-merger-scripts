#!/usr/bin/env python3
"""Streamlit web UI for the orders + shipping-charges merge.

A zero-install front end for non-developers: upload the two Shopify exports,
click Merge, download the styled .xlsx. Reuses merge.py's logic verbatim.

Run locally:
    streamlit run app.py

Deploy: push to GitHub, then connect the repo at https://share.streamlit.io
and point it at app.py.
"""
from __future__ import annotations

import tempfile
from datetime import date
from pathlib import Path

import streamlit as st

from merge import merge_orders_and_charges

st.set_page_config(page_title="GG Shipping Merger", page_icon="📦", layout="centered")


# --- Optional password gate ---------------------------------------------------
# Set an "app_password" secret in Streamlit Cloud (Settings → Secrets) to enable.
# If no secret is configured, the gate is skipped (handy for local testing).

def check_password() -> bool:
    try:
        expected = st.secrets["app_password"]
    except Exception:
        # No secrets file / key configured at all → open access (e.g. local runs).
        expected = None
    if not expected:
        return True  # no password configured → open access

    if st.session_state.get("password_ok"):
        return True

    st.title("📦 GG Shipping Merger")
    entered = st.text_input("Enter the access password", type="password")
    if entered:
        if entered == expected:
            st.session_state["password_ok"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False


if not check_password():
    st.stop()


# --- Main UI ------------------------------------------------------------------

st.title("📦 GG Shipping Merger")
st.write(
    "Upload your Shopify **orders** export and **charges** export. "
    "This tool adds each order's company shipping cost and the shipping "
    "difference, then gives you a polished Excel file to download."
)

col1, col2 = st.columns(2)
with col1:
    orders_file = st.file_uploader(
        "1. Orders CSV", type="csv", key="orders",
        help="The orders_export.csv from Shopify.",
    )
with col2:
    charges_file = st.file_uploader(
        "2. Charges CSV", type="csv", key="charges",
        help="The charges/billing export with shipping_fee rows.",
    )

st.caption("Order matters: the **orders** file goes on the left, **charges** on the right.")

run = st.button("Merge", type="primary", disabled=not (orders_file and charges_file))

if run and orders_file and charges_file:
    with st.spinner("Merging…"):
        # Streamlit hands us in-memory uploads; merge.py reads from disk, so we
        # stage both files (and the output) in a throwaway temp directory.
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            orders_path = tmp_path / "orders.csv"
            charges_path = tmp_path / "charges.csv"
            output_path = tmp_path / f"orders_with_shipping_cost_{date.today().isoformat()}.xlsx"

            orders_path.write_bytes(orders_file.getvalue())
            charges_path.write_bytes(charges_file.getvalue())

            try:
                stats = merge_orders_and_charges(orders_path, charges_path, output_path)
                xlsx_bytes = output_path.read_bytes()
            except SystemExit as exc:
                # merge.py raises SystemExit with a friendly message on bad input.
                st.error(f"Could not merge: {exc}")
                st.stop()
            except Exception as exc:  # noqa: BLE001 — surface anything else to the user
                st.error(f"Something went wrong: {exc}")
                st.stop()

    st.success("Done! Your merged file is ready.")

    st.download_button(
        "⬇️ Download merged Excel file",
        data=xlsx_bytes,
        file_name=output_path.name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    if stats["error_count"]:
        st.warning(f"{stats['error_count']} formula error(s) detected — double-check the output.")

    st.subheader("Summary")
    m1, m2, m3 = st.columns(3)
    m1.metric("Data rows", f"{stats['data_rows']:,}")
    m2.metric("Orders matched", f"{stats['matched_orders']:,}")
    m3.metric("Unbilled orders", f"{stats['unbilled_count']:,}")

    t1, t2, t3 = st.columns(3)
    t1.metric("Shipping (customer paid)", f"${stats['total_shipping_paid']:,.2f}")
    t2.metric("Company shipping cost", f"${stats['total_company_cost']:,.2f}")
    t3.metric("Shipping difference", f"${stats['total_difference']:,.2f}")

    st.caption(
        f"{stats['unique_orders_in_charges']:,} unique orders in the charges file · "
        f"{stats['num_columns']} columns (A..{stats['last_column_letter']})"
    )
