import io
import tempfile
from pathlib import Path
from datetime import datetime

import streamlit as st
from modules.apc1_converter import convert_uploaded_file
from modules.qa_compare import write_result
from modules.oit_inline_diff import compare_oit_workbooks

st.set_page_config(page_title="Engineering Data Tools", page_icon="🧩", layout="wide", initial_sidebar_state="collapsed")
st.title("Engineering Data Tools")
st.caption("APC1 → OIT Converter | OIT vs webDM QA Comparison")

tab1, tab2, tab3 = st.tabs(["Tab 1 · APC1 → OIT", "Tab 2 · QA PSV Compare", "Tab 3 · OIT Excel Diff"])

def temp_file(suffix, data):
    f = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    f.write(data); f.flush(); f.close()
    return Path(f.name)

with tab1:
    st.header("APC1 → OIT Converter")
    st.write("Upload APC1 Excel workbook and convert it using the v9 mapping and embedded OIT template.")
    apc = st.file_uploader("Upload APC1 Excel", type=["xlsx", "xlsm"], key="apc_upload")
    if apc is not None:
        st.info(f"Selected: {apc.name} · {apc.size:,} bytes")
        if st.button("Convert APC1 to OIT", type="primary", key="apc_convert"):
            tmp_out = None
            progress = st.progress(0, text="Preparing APC1 conversion...")
            status = st.empty()
            def update_apc_progress(value, message):
                progress.progress(min(100, max(0, int(value))), text=message)
                status.caption(message)
            try:
                update_apc_progress(5, "Reading uploaded APC1 workbook...")
                suffix = Path(apc.name).suffix or ".xlsx"
                m = __import__("re").search(r"PT\d+", apc.name, flags=__import__("re").I)
                pt = m.group(0).upper() if m else "PT_UNKNOWN"
                out_name = f"{datetime.now():%m%d}_OIT_result_{pt}.xlsx"
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
                    tmp_out = Path(f.name)
                update_apc_progress(12, "Validating APC1 headers and preparing OIT template...")
                count = convert_uploaded_file(apc.getvalue(), apc.name, tmp_out, progress_callback=update_apc_progress)
                st.session_state["apc_result"] = (tmp_out.read_bytes(), out_name, count)
                update_apc_progress(100, "Conversion completed.")
                st.success(f"Conversion completed. Rows converted: {count}")
            except Exception as e:
                st.error(f"APC1 → OIT conversion failed: {e}")
            finally:
                if tmp_out and tmp_out.exists():
                    tmp_out.unlink(missing_ok=True)
    if "apc_result" in st.session_state:
        data, name, count = st.session_state["apc_result"]
        st.metric("Rows converted", count)
        st.download_button("Download OIT Result", data, file_name=name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="apc_download")

with tab2:
    st.header("QA PSV Comparison")
    st.write("Upload OIT1 Excel and webDM2 raw PSV, PSV-export Excel, or normal Excel Device Listing.")
    c1, c2 = st.columns(2)
    with c1:
        oit = st.file_uploader("Upload OIT1 Excel", type=["xlsx", "xlsm"], key="qa_oit")
    with c2:
        webdm = st.file_uploader("Upload webDM2 (.psv / .xlsx / .xlsm)", type=["psv", "xlsx", "xlsm"], key="qa_webdm")
    if oit is not None and webdm is not None:
        st.caption(f"OIT1: {oit.name} · webDM2: {webdm.name}")
        if st.button("Run QA Comparison", type="primary", key="qa_run"):
            progress = st.progress(0, text="Preparing QA comparison...")
            status = st.empty()
            def update_qa_progress(value, message):
                progress.progress(min(100, max(0, int(value))), text=message)
                status.caption(message)
            paths = []
            try:
                update_qa_progress(3, "Preparing uploaded files...")
                oit_path = temp_file(Path(oit.name).suffix or ".xlsx", oit.getvalue()); paths.append(oit_path)
                webdm_path = temp_file(Path(webdm.name).suffix or ".psv", webdm.getvalue()); paths.append(webdm_path)
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
                    out_path = Path(f.name)
                try:
                    stats = write_result(oit_path, webdm_path, out_path, progress_callback=update_qa_progress)
                    out_name = f"{datetime.now():%m%d}_QApsv_result_{__import__('re').search(r'PT\d+', oit.name, flags=__import__('re').I).group(0).upper() if __import__('re').search(r'PT\d+', oit.name, flags=__import__('re').I) else 'PTxxxx'}.xlsx"
                    st.session_state["qa_result"] = (out_path.read_bytes(), out_name, stats)
                    update_qa_progress(100, "QA comparison completed.")
                finally:
                    out_path.unlink(missing_ok=True)
            except Exception as e:
                st.error(f"QA PSV comparison failed: {e}")
            finally:
                for p in paths:
                    p.unlink(missing_ok=True)
    if "qa_result" in st.session_state:
        data, name, stats = st.session_state["qa_result"]
        st.success("QA comparison completed.")
        cols = st.columns(4)
        cols[0].metric("Matched IDs", stats["matched"])
        cols[1].metric("Different fields", stats["differences"])
        cols[2].metric("OIT-only IDs", stats["oit_only"])
        cols[3].metric("webDM-only IDs", stats["webdm_only"])
        st.metric("Devices with differences", stats["devices_with_differences"])
        st.download_button("Download QA Report", data, file_name=name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="qa_download")


with tab3:
    st.header("OIT Excel Inline Diff")
    st.write("Upload the NEW and OLD OIT Excel workbooks. The report highlights changed, added and removed cell values.")
    dcol1, dcol2 = st.columns(2)
    with dcol1:
        diff_new = st.file_uploader("Upload NEW OIT workbook", type=["xlsx", "xlsm"], key="diff_new")
    with dcol2:
        diff_old = st.file_uploader("Upload OLD OIT workbook", type=["xlsx", "xlsm"], key="diff_old")

    if diff_new is not None and diff_old is not None:
        st.caption(f"NEW: {diff_new.name} · OLD: {diff_old.name}")
        if diff_new.name == diff_old.name:
            st.warning("NEW and OLD file names are identical. Please confirm that you selected the correct versions.")
        if st.button("Run OIT Diff", type="primary", key="oit_diff_run"):
            diff_progress = st.progress(0, text="Preparing OIT comparison...")
            diff_status = st.empty()

            def update_diff_progress(value, message):
                diff_progress.progress(min(100, max(0, int(value))), text=message)
                diff_status.caption(message)

            try:
                excel_bytes, html_report, records, counts = compare_oit_workbooks(
                    diff_new.getvalue(),
                    diff_old.getvalue(),
                    diff_new.name,
                    diff_old.name,
                    progress_callback=update_diff_progress,
                )
                stamp = datetime.now().strftime("%m%d_%H%M%S")
                stem = Path(diff_new.name).stem
                st.session_state["oit_diff_result"] = {
                    "excel": excel_bytes,
                    "html": html_report.encode("utf-8"),
                    "records": records,
                    "counts": counts,
                    "excel_name": f"{stamp}_{stem}_color_diff.xlsx",
                    "html_name": f"{stamp}_{stem}_diff_report.html",
                    "new_name": diff_new.name,
                    "old_name": diff_old.name,
                }
                st.success("OIT diff completed.")
            except Exception as e:
                st.error(f"OIT Excel diff failed: {e}")

    if "oit_diff_result" in st.session_state:
        result = st.session_state["oit_diff_result"]
        st.subheader("Diff Summary")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Changed", result["counts"]["CHANGED"])
        m2.metric("Added", result["counts"]["ADDED"])
        m3.metric("Removed", result["counts"]["REMOVED"])
        m4.metric("Total Differences", len(result["records"]))

        b1, b2 = st.columns(2)
        with b1:
            st.download_button(
                "Download Diff Excel",
                result["excel"],
                file_name=result["excel_name"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="oit_diff_excel_download",
            )
        with b2:
            st.download_button(
                "Download Diff HTML",
                result["html"],
                file_name=result["html_name"],
                mime="text/html",
                key="oit_diff_html_download",
            )

        st.subheader("Diff Preview")
        st.caption("Inline preview of the HTML report. Yellow = changed, green = added, red = removed; text-level insertions/deletions are highlighted.")
        try:
            import streamlit.components.v1 as components
            components.html(result["html"].decode("utf-8"), height=650, scrolling=True)
        except Exception:
            st.warning("HTML preview cannot be rendered in this environment. Please download the HTML report.")

        with st.expander("Diff Records Preview", expanded=False):
            import pandas as pd
            if result["records"]:
                preview_df = pd.DataFrame(result["records"])
                st.dataframe(preview_df.head(500), use_container_width=True, height=350)
                if len(preview_df) > 500:
                    st.caption("Preview limited to the first 500 differences. Download the full Excel or HTML report for all differences.")
            else:
                st.success("No differences found between the selected workbooks.")


st.divider()
st.caption("Engineering Data Tools · Browser upload/download edition. Use only company-approved hosting for company data.")
