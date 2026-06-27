import streamlit as st
import os
import sys
import json
import datetime as dt
from pathlib import Path
import requests

from report_generator import build_data_dict, build_report

# ----------------------------------------------------------------------------
# Streamlit Web App Master Dashboard Layout
# ----------------------------------------------------------------------------
st.set_page_config(page_title="LAMCO Daily Site Report Master Dashboard", layout="wide")

st.title("📋 LAMCO Daily Site Report Master Dashboard")
st.caption("Pulls real-time submission records from KoboToolbox. Fully integrated version formatting engine.")

SERVER = st.secrets.get("KOBO_SERVER", "kf.kobotoolbox.org")
ASSET_UID = st.secrets.get("KOBO_ASSET_UID", "")
TOKEN = st.secrets.get("KOBO_TOKEN", "")
FORM_URL = st.secrets.get("KOBO_FORM_URL", "")

btn_cols = st.columns([3.2, 2.5, 3.8, 10.5])
col_idx = 0
if FORM_URL:
    with btn_cols[col_idx]:
        st.link_button("📝 Open Kobo Online Form", FORM_URL, use_container_width=True)
    col_idx += 1
with btn_cols[col_idx]:
    if st.button("🔄 Refresh List", use_container_width=True):
        st.rerun()
col_idx += 1
with btn_cols[col_idx]:
    if st.button("⚡ Force Refresh (Clear Cache)", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
st.markdown("") # Spacer

if not ASSET_UID or not TOKEN:
    st.warning("⚠️ Configuration Keys Missing inside Advanced Cloud Secrets panel. Please type tokens manually:")
    col_creds1, col_creds2, col_creds3 = st.columns([2, 2, 3])
    with col_creds1: ASSET_UID = st.text_input("Kobo Asset Form UID ID", value=ASSET_UID)
    with col_creds2: TOKEN = st.text_input("Kobo Account API Token", type="password", value=TOKEN)
    with col_creds3: FORM_URL = st.text_input("Kobo Form Online URL", value=FORM_URL)

@st.cache_data(ttl=120)
def fetch_all_submissions(server, asset_uid, token):
    base_url = f"https://{server}/api/v2/assets/{asset_uid}/"
    session = requests.Session()
    session.headers.update({"Authorization": f"Token {token}"})
    
    results = []
    url = f"{base_url}data/"
    params = {"sort": json.dumps({"_id": -1}), "limit": 1000}
    
    while url:
        res = session.get(url, params=params, timeout=45)
        res.raise_for_status()
        data = res.json()
        results.extend(data.get("results", []))
        url = data.get("next")
        params = None
        
    return results

if ASSET_UID and TOKEN:
    try:
        with st.spinner("Synchronizing data stream logs from Kobo Link (Fetching full dataset up to 1000 items)..."):
            submissions = fetch_all_submissions(SERVER, ASSET_UID, TOKEN)
            
        st.subheader(f"📊 Live Submissions Log Ledger ({len(submissions)} records found)")
        
        # Table Header Row
        header_cols = st.columns([1.2, 1.8, 2, 3, 2, 2.5])
        header_cols[0].markdown("**ID**")
        header_cols[1].markdown("**Report Date**")
        header_cols[2].markdown("**Reporter**")
        header_cols[3].markdown("**Project Name**")
        header_cols[4].markdown("**Progress Status**")
        header_cols[5].markdown("**Action Options**")
        st.markdown("---")
        
        # Print Grid Items Loops
        for sub in submissions:
            sub_id = sub.get("_id")
            rep_date = sub.get("grp_daily/report_date") or sub.get("today_date") or "—"
            reporter = sub.get("grp_reporter/reporter_name") or "—"
            p_name = sub.get("grp_project/project_name") or "—"
            status = sub.get("grp_summary/progress_status") or "—"
            
            row_cols = st.columns([1.2, 1.8, 2, 3, 2, 2.5])
            row_cols[0].write(f"#{sub_id}")
            row_cols[1].write(str(rep_date))
            row_cols[2].write(str(reporter))
            row_cols[3].write(str(p_name))
            row_cols[4].write(str(status).upper())
            
            if row_cols[5].button(f"📄 Compile Report", key=f"btn_{sub_id}"):
                with st.spinner(f"Acquiring attachments and formatting Document for Entry #{sub_id}..."):
                    base_url = f"https://{SERVER}/api/v2/assets/{ASSET_UID}/"
                    session = requests.Session()
                    session.headers.update({"Authorization": f"Token {TOKEN}"})
                    
                    # Establish central output directory
                    output_dir = Path("./output")
                    output_dir.mkdir(parents=True, exist_ok=True)
                    
                    photo_dir = output_dir / f"photos_{sub_id}"
                    photo_dir.mkdir(parents=True, exist_ok=True)
                    
                    attachment_map = {}
                    for att in sub.get("_attachments", []):
                        fn = att.get("filename", "")
                        dl_url = att.get("download_url") or att.get("download_large_url")
                        if dl_url and fn:
                            p_name_att = Path(fn).name
                            local_p = photo_dir / p_name_att
                            if not local_p.exists():
                                try:
                                    r_file = session.get(dl_url, timeout=30)
                                    if r_file.status_code == 200:
                                        local_p.write_bytes(r_file.content)
                                        attachment_map[p_name_att] = str(local_p)
                                except Exception:
                                    pass
                            else:
                                attachment_map[p_name_att] = str(local_p)
                    
                    data = build_data_dict(sub, attachment_map)
                    data["_generated_at"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
                    
                    # Auto-Versioning to prevent [Errno 13] File Lock errors inside the output folder
                    base_filename = f"LAMCO_Daily_Report_ID{sub_id}_{str(rep_date).replace('/', '-')}"
                    out_filepath = output_dir / f"{base_filename}.docx"
                    version = 1
                    
                    while out_filepath.exists():
                        version += 1
                        out_filepath = output_dir / f"{base_filename}_v{version}.docx"
                        
                    data["_version"] = version
                    
                    build_report(data, photo_dir, str(out_filepath))
                    
                    with open(out_filepath, "rb") as f:
                        st.download_button(
                            label="📥 Save Word File (.docx)",
                            data=f,
                            file_name=out_filepath.name,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"dl_{sub_id}"
                        )
                    st.success(f"Report for #{sub_id} successfully compiled (Saved as v{version})!")
            st.markdown("<h5 style='margin:0; opacity:0.1; border-bottom:1px solid gray;'></h5>", unsafe_allow_html=True)
            
    except Exception as e:
        st.error(f"Failed connection execution loop error: {e}")