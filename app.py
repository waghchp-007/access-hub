import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime, date
import urllib.parse
import io
import shutil

DATA_DIR = Path('data')
DATA_FILE = DATA_DIR / 'alllinks.xlsx'
TRACKER_FILE = DATA_DIR / 'daily_tracker.xlsx'
SHEET_NAME = 'Links'
TRACKER_SHEET = 'Tracker'
COLUMNS = ["Category","Application Name","Description","Link Type","URL","Owner","Uploaded By","Uploaded Date","Last Modified By","Last Modified Date","Status","Update Available","Version","Department","Priority","Remarks"]
TRACKER_COLUMNS = ["Task Name","Owner","Start Date","Expected End Date","ETA","Progress","Status","Priority","Remarks","Created By","Created Date","Last Modified By","Last Modified Date"]
STATUS_COLORS = {"Active":"🟢","In Progress":"🔵","Pending":"🟠","Retired":"⚪","Issue":"🔴"}
PRIORITY_COLORS = {"High":"🔴","Medium":"🟡","Low":"🟢"}
TYPE_ICONS = {"Power BI":"📊","SharePoint":"🔷","Excel":"📗","Web Page":"🌐","Application":"⚙️","Document":"📄","PPT / Deck":"📽️","Ops Resource":"🔧","Other":"🔗"}
PROGRESS_COLORS = {"0%":"#EF4444","10%":"#F97316","20%":"#F97316","30%":"#F59E0B","40%":"#F59E0B","50%":"#EAB308","60%":"#84CC16","70%":"#84CC16","80%":"#22C55E","90%":"#22C55E","100%":"#059669"}
ADMIN_PASSWORD = "admin@hub2026"

st.set_page_config(page_title="Centralized Application Access Hub", page_icon="📊", layout="wide")

CSS_LINES = [
    "#MainMenu, footer, header {visibility: hidden;}",
    ".block-container {padding-top: 1rem; padding-bottom: 1rem;}",
    ".top-header {background: linear-gradient(135deg, #1E3A5F, #2563EB, #3B82F6); padding: 16px 28px; border-radius: 14px; margin-bottom: 14px; display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; box-shadow: 0 4px 20px rgba(37,99,235,0.2);}",
    ".top-header .title-area {text-align: center; flex: 1;}",
    ".top-header h1 {font-size: 22px; font-weight: 800; color: #fff; margin: 0;}",
    ".top-header .subtitle {color: #BFDBFE; font-size: 12px; margin: 2px 0 0;}",
    ".top-header .user-info {text-align: right; color: #BFDBFE; font-size: 11px; min-width: 180px;}",
    ".top-header .user-info b {color: #fff; font-size: 13px;}",
    ".top-header .logo {width: 42px; height: 42px; border-radius: 10px; min-width: 42px; background: rgba(255,255,255,0.15); display: flex; align-items: center; justify-content: center; font-size: 22px;}",
    ".stat-row {display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 14px;}",
    ".stat-card {flex: 1; min-width: 130px; padding: 14px 16px; border-radius: 12px; text-align: center; border: 1px solid rgba(0,0,0,0.06); box-shadow: 0 2px 8px rgba(0,0,0,0.04);}",
    ".stat-card .num {font-size: 28px; font-weight: 800; line-height: 1.1;}",
    ".stat-card .lbl {font-size: 10px; color: #475569; text-transform: uppercase; letter-spacing: 0.06em; margin-top: 2px;}",
    ".sc-blue {background: #EFF6FF;} .sc-blue .num {color: #2563EB;}",
    ".sc-green {background: #ECFDF5;} .sc-green .num {color: #059669;}",
    ".sc-red {background: #FEF2F2;} .sc-red .num {color: #DC2626;}",
    ".sc-purple {background: #F5F3FF;} .sc-purple .num {color: #7C3AED;}",
    ".sc-yellow {background: #FEFCE8;} .sc-yellow .num {color: #CA8A04;}",
    ".sc-cyan {background: #ECFEFF;} .sc-cyan .num {color: #0891B2;}",
    ".badge {display: inline-block; padding: 2px 10px; border-radius: 999px; font-size: 11px; font-weight: 600;}",
    ".badge-active {background: #D1FAE5; color: #059669;}",
    ".badge-inprogress {background: #DBEAFE; color: #3B82F6;}",
    ".badge-pending {background: #FEF3C7; color: #D97706;}",
    ".badge-retired {background: #F1F5F9; color: #64748B;}",
    ".badge-issue {background: #FEE2E2; color: #DC2626;}",
    ".badge-yes {background: #FEE2E2; color: #DC2626;}",
    ".badge-no {background: #D1FAE5; color: #059669;}",
    ".badge-done {background: #D1FAE5; color: #059669;}",
    ".badge-todo {background: #DBEAFE; color: #3B82F6;}",
    ".badge-blocked {background: #FEE2E2; color: #DC2626;}",
    ".cat-tag {display: inline-block; padding: 2px 8px; border-radius: 6px; font-size: 11px; font-weight: 600;}",
    ".cat-dashboard {background: #DBEAFE; color: #1E40AF;}",
    ".cat-application {background: #FEF3C7; color: #92400E;}",
    ".cat-document {background: #FCE7F3; color: #9D174D;}",
    ".cat-report {background: #E0E7FF; color: #3730A3;}",
    ".cat-webpage {background: #CCFBF1; color: #065F46;}",
    ".cat-other {background: #F1F5F9; color: #64748B;}",
    ".htable {width: 100%; border-collapse: collapse; font-size: 12.5px;}",
    ".htable th {background: linear-gradient(135deg, #1E3A5F, #2563EB); color: white; padding: 10px; text-align: left; font-size: 10px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.04em; white-space: nowrap;}",
    ".htable td {padding: 8px 10px; border-bottom: 1px solid #E2E8F0; vertical-align: middle;}",
    ".htable tr:nth-child(even) {background: #F8FAFC;}",
    ".htable tr:hover {background: #EFF6FF;}",
    ".htable .nm {font-weight: 600; color: #0F172A;}",
    ".htable .ds {color: #64748B; font-size: 11.5px; max-width: 170px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;}",
    ".htable .mt {font-size: 11px; color: #64748B;}",
    ".htable .sr {text-align: center; color: #94A3B8; font-weight: 600; font-size: 11px;}",
    ".abtn {display: inline-flex; align-items: center; justify-content: center; width: 26px; height: 26px; border-radius: 6px; border: 1px solid #E2E8F0; background: white; font-size: 12px; text-decoration: none; transition: all 0.12s;}",
    ".abtn:hover {background: #DBEAFE; border-color: #3B82F6; transform: translateY(-1px);}",
    ".pbar {width: 100%; height: 8px; background: #E2E8F0; border-radius: 4px; overflow: hidden;}",
    ".pbar-fill {height: 100%; border-radius: 4px;}",
    ".stTabs [data-baseweb='tab-list'] {gap: 8px;}",
    ".stTabs [data-baseweb='tab'] {border-radius: 8px 8px 0 0; padding: 8px 20px; font-weight: 600; font-size: 13px;}",
    ".stButton > button {border-radius: 8px; font-weight: 600;}",
    ".stDownloadButton > button {border-radius: 8px; background: #059669; color: white; border: none;}",
    ".main .block-container {background: #FAFBFE;}",
    "div[data-testid='stSidebar'] {display: none;}",
    ".sec-title {font-size: 13px; font-weight: 700; color: #1E3A5F; margin: 12px 0 8px; padding-bottom: 4px; border-bottom: 2px solid #3B82F6;}",
]
CSS = "<style>" + " ".join(CSS_LINES) + "</style>"
st.markdown(CSS, unsafe_allow_html=True)

def ensure_files():
    DATA_DIR.mkdir(exist_ok=True)
    if not DATA_FILE.exists():
        pd.DataFrame(columns=COLUMNS).to_excel(DATA_FILE, sheet_name=SHEET_NAME, index=False)
    if not TRACKER_FILE.exists():
        pd.DataFrame(columns=TRACKER_COLUMNS).to_excel(TRACKER_FILE, sheet_name=TRACKER_SHEET, index=False)

def load_links():
    ensure_files()
    try:
        df = pd.read_excel(DATA_FILE, sheet_name=SHEET_NAME).fillna("")
        for c in COLUMNS:
            if c not in df.columns: df[c] = ""
        return df
    except: return pd.DataFrame(columns=COLUMNS)

def save_links(df):
    ensure_files()
    try:
        if DATA_FILE.exists():
            bk = DATA_DIR / ("links_bk_" + datetime.now().strftime("%H%M%S") + ".xlsx")
            shutil.copy2(DATA_FILE, bk)
            for f in sorted(DATA_DIR.glob("links_bk_*.xlsx"))[:-3]: f.unlink()
        df.to_excel(DATA_FILE, sheet_name=SHEET_NAME, index=False)
        st.session_state.links = df.copy()
    except Exception as e: st.warning("Save error: " + str(e))

def get_links():
    if "links" not in st.session_state: st.session_state.links = load_links()
    return st.session_state.links.copy()

def load_tracker():
    ensure_files()
    try:
        df = pd.read_excel(TRACKER_FILE, sheet_name=TRACKER_SHEET).fillna("")
        for c in TRACKER_COLUMNS:
            if c not in df.columns: df[c] = ""
        return df
    except: return pd.DataFrame(columns=TRACKER_COLUMNS)

def save_tracker(df):
    ensure_files()
    try:
        if TRACKER_FILE.exists():
            bk = DATA_DIR / ("tracker_bk_" + datetime.now().strftime("%H%M%S") + ".xlsx")
            shutil.copy2(TRACKER_FILE, bk)
            for f in sorted(DATA_DIR.glob("tracker_bk_*.xlsx"))[:-3]: f.unlink()
        df.to_excel(TRACKER_FILE, sheet_name=TRACKER_SHEET, index=False)
        st.session_state.tracker = df.copy()
    except Exception as e: st.warning("Save error: " + str(e))

def get_tracker():
    if "tracker" not in st.session_state: st.session_state.tracker = load_tracker()
    return st.session_state.tracker.copy()

def cat_class(c):
    c = str(c).lower()
    if "dashboard" in c or "power bi" in c: return "cat-dashboard"
    if "app" in c: return "cat-application"
    if "doc" in c: return "cat-document"
    if "report" in c or "excel" in c: return "cat-report"
    if "web" in c: return "cat-webpage"
    return "cat-other"

def status_class(s):
    k = str(s).lower().replace(" ","")
    m = {"active":"badge-active","inprogress":"badge-inprogress","pending":"badge-pending","retired":"badge-retired","issue":"badge-issue","done":"badge-done","todo":"badge-todo","blocked":"badge-blocked"}
    return m.get(k, "badge-active")

def esc(t):
    return str(t).replace("<","&lt;").replace(">","&gt;").replace('"',"&quot;")

def render_links_table(df):
    if df.empty:
        st.info("No records. Go to Data Management tab to upload or add entries.")
        return
    html = '<div style="overflow-x:auto;max-height:550px;overflow-y:auto;border-radius:10px;border:1px solid #E2E8F0;"><table class="htable"><thead><tr>'
    for h in ["#","Category","Name","Type","Description","Owner","Status","Update","Ver.","Modified By","Modified Date","Actions"]:
        html += "<th>" + h + "</th>"
    html += "</tr></thead><tbody>"
    for i, (_, r) in enumerate(df.iterrows()):
        url = str(r.get("URL",""))
        hu = url.startswith("http")
        cc = cat_class(r.get("Category",""))
        sc = status_class(r.get("Status",""))
        u = str(r.get("Update Available","")).lower().strip()
        uc = "badge-yes" if u=="yes" else "badge-no"
        ut = "Yes" if u=="yes" else "No"
        pd2 = PRIORITY_COLORS.get(str(r.get("Priority","")),"")
        ti = TYPE_ICONS.get(str(r.get("Link Type","")),"🔗")
        md = str(r.get("Last Modified Date",""))[:20]
        if md == "nan": md = "-"
        ob = '<a href="' + url + '" target="_blank" class="abtn" title="Open">🔗</a>' if hu else '<span class="abtn" style="opacity:.3">-</span>'
        desc = esc(r.get("Description","-"))
        nm = esc(r.get("Application Name","-"))
        stv = str(r.get("Status","-"))
        sicon = STATUS_COLORS.get(stv,"")
        html += "<tr>"
        html += '<td class="sr">' + str(i+1) + "</td>"
        html += '<td><span class="cat-tag ' + cc + '">' + pd2 + " " + str(r.get("Category","-")) + "</span></td>"
        html += '<td class="nm">' + nm + "</td>"
        html += '<td style="font-size:12px">' + ti + " " + str(r.get("Link Type","")) + "</td>"
        html += '<td class="ds" title="' + desc + '">' + desc + "</td>"
        html += '<td class="mt">' + str(r.get("Owner","-")) + "</td>"
        html += '<td><span class="badge ' + sc + '">' + sicon + stv + "</span></td>"
        html += '<td><span class="badge ' + uc + '">' + ut + "</span></td>"
        html += '<td class="mt">' + str(r.get("Version","-")) + "</td>"
        html += '<td class="mt">' + str(r.get("Last Modified By","-")) + "</td>"
        html += '<td class="mt">' + md + "</td>"
        html += '<td style="white-space:nowrap">' + ob + "</td>"
        html += "</tr>"
    html += "</tbody></table></div>"
    st.markdown(html, unsafe_allow_html=True)

def render_tracker_table(df):
    if df.empty:
        st.info("No tasks yet. Click Add Task below to start tracking.")
        return
    html = '<div style="overflow-x:auto;max-height:500px;overflow-y:auto;border-radius:10px;border:1px solid #E2E8F0;"><table class="htable"><thead><tr>'
    for h in ["#","Task Name","Owner","Start","End","ETA","Progress","Status","Priority","Remarks","Modified By","Modified Date"]:
        html += "<th>" + h + "</th>"
    html += "</tr></thead><tbody>"
    for i, (_, r) in enumerate(df.iterrows()):
        sc = status_class(r.get("Status",""))
        pr = str(r.get("Progress","0%")).replace("%","")
        try: pv = int(pr)
        except: pv = 0
        pc = PROGRESS_COLORS.get(str((pv//10)*10) + "%", "#94A3B8")
        pd2 = PRIORITY_COLORS.get(str(r.get("Priority","")),"")
        md = str(r.get("Last Modified Date",""))[:20]
        if md == "nan": md = "-"
        sd = str(r.get("Start Date",""))[:10]
        ed = str(r.get("Expected End Date",""))[:10]
        if sd == "nan": sd = "-"
        if ed == "nan": ed = "-"
        html += "<tr>"
        html += '<td class="sr">' + str(i+1) + "</td>"
        html += '<td class="nm">' + str(r.get("Task Name","-")) + "</td>"
        html += '<td class="mt">' + str(r.get("Owner","-")) + "</td>"
        html += '<td class="mt">' + sd + "</td>"
        html += '<td class="mt">' + ed + "</td>"
        html += '<td class="mt">' + str(r.get("ETA","-")) + "</td>"
        html += '<td style="min-width:100px"><div class="pbar"><div class="pbar-fill" style="width:' + str(pv) + '%;background:' + pc + '"></div></div><span style="font-size:10px;color:' + pc + ';font-weight:700">' + str(pv) + '%</span></td>'
        html += '<td><span class="badge ' + sc + '">' + str(r.get("Status","-")) + "</span></td>"
        html += "<td>" + pd2 + " " + str(r.get("Priority","-")) + "</td>"
        html += '<td class="ds">' + str(r.get("Remarks","-")) + "</td>"
        html += '<td class="mt">' + str(r.get("Last Modified By","-")) + "</td>"
        html += '<td class="mt">' + md + "</td>"
        html += "</tr>"
    html += "</tbody></table></div>"
    st.markdown(html, unsafe_allow_html=True)


def main():
    if "user_name" not in st.session_state: st.session_state.user_name = ""
    if "last_access" not in st.session_state: st.session_state.last_access = datetime.now().strftime("%d-%b-%Y %I:%M %p")
    if "t_edit_idx" not in st.session_state: st.session_state.t_edit_idx = None
    if "edit_idx" not in st.session_state: st.session_state.edit_idx = None

    df = get_links()
    tdf = get_tracker()
    total = len(df)
    active = len(df[df["Status"].str.lower()=="active"]) if total else 0
    updates = len(df[df["Update Available"].str.lower()=="yes"]) if total else 0
    cats = df["Category"].nunique() if total else 0
    tasks = len(tdf)
    tasks_done = len(tdf[tdf["Status"].str.lower()=="done"]) if tasks else 0
    un = st.session_state.user_name or "Guest"

    header_html = '<div class="top-header"><div class="logo">📊</div><div class="title-area"><h1>Centralized Application Access Hub</h1><div class="subtitle">All dashboards, applications, documents and resources in one place</div></div><div class="user-info"><b>' + un + '</b><br>🕐 ' + st.session_state.last_access + '</div></div>'
    st.markdown(header_html, unsafe_allow_html=True)

    stats_html = '<div class="stat-row">'
    stats_html += '<div class="stat-card sc-blue"><div class="num">' + str(total) + '</div><div class="lbl">Total Links</div></div>'
    stats_html += '<div class="stat-card sc-green"><div class="num">' + str(active) + '</div><div class="lbl">Active</div></div>'
    stats_html += '<div class="stat-card sc-red"><div class="num">' + str(updates) + '</div><div class="lbl">Updates</div></div>'
    stats_html += '<div class="stat-card sc-purple"><div class="num">' + str(cats) + '</div><div class="lbl">Categories</div></div>'
    stats_html += '<div class="stat-card sc-yellow"><div class="num">' + str(tasks) + '</div><div class="lbl">Tasks</div></div>'
    stats_html += '<div class="stat-card sc-cyan"><div class="num">' + str(tasks_done) + '</div><div class="lbl">Completed</div></div>'
    stats_html += '</div>'
    st.markdown(stats_html, unsafe_allow_html=True)

    c1, c2, c3 = st.columns([2,2,4])
    with c1:
        uinp = st.text_input("Your Name", value=st.session_state.user_name, placeholder="Enter name", label_visibility="collapsed")
        st.session_state.user_name = uinp

    tab1, tab2, tab3 = st.tabs(["📊 Access Hub", "📂 Data Management", "📋 Daily Status Tracker"])

    with tab1:
        fc1, fc2, fc3, fc4 = st.columns(4)
        with fc1: search = st.text_input("🔍 Search", placeholder="Name, category...", key="s1")
        with fc2:
            copts = ["All"] + sorted(df["Category"].unique().tolist()) if total else ["All"]
            fcat = st.selectbox("Category", copts, key="fc1")
        with fc3:
            topts = ["All"] + sorted(df["Link Type"].unique().tolist()) if total else ["All"]
            ftype = st.selectbox("Type", topts, key="ft1")
        with fc4:
            sopts = ["All"] + sorted(df["Status"].unique().tolist()) if total else ["All"]
            fstat = st.selectbox("Status", sopts, key="fs1")
        filt = df.copy()
        if search: filt = filt[filt.apply(lambda r: search.lower() in " ".join(r.astype(str).values).lower(), axis=1)]
        if fcat != "All": filt = filt[filt["Category"]==fcat]
        if ftype != "All": filt = filt[filt["Link Type"]==ftype]
        if fstat != "All": filt = filt[filt["Status"]==fstat]
        st.markdown("**" + str(len(filt)) + "** of **" + str(total) + "** records")
        render_links_table(filt)

    with tab2:
        dm1, dm2 = st.columns(2)
        with dm1:
            st.markdown('<div class="sec-title">Upload Excel</div>', unsafe_allow_html=True)
            umode = st.radio("Mode", ["Append new rows", "Replace all (admin)"], key="um", horizontal=True)
            apw = ""
            if "Replace" in umode:
                apw = st.text_input("Admin password", type="password", key="ap")
            upf = st.file_uploader("Upload", type=["xlsx","xls"], key="uf")
            if upf:
                try:
                    ndf = pd.read_excel(upf, sheet_name=0).fillna("")
                    for c in COLUMNS:
                        if c not in ndf.columns: ndf[c] = ""
                    st.info("**" + upf.name + "** - **" + str(len(ndf)) + " rows**")
                    if "Append" in umode:
                        ex = get_links()
                        eu = set(ex["URL"].str.strip().str.lower())
                        nr = ndf[~ndf["URL"].str.strip().str.lower().isin(eu)]
                        d = len(ndf) - len(nr)
                        if d > 0: st.warning(str(d) + " duplicates skipped")
                        if len(nr) > 0:
                            if st.button("Append " + str(len(nr)) + " records", type="primary", key="ab"):
                                save_links(pd.concat([ex, nr], ignore_index=True))
                                st.success("Added " + str(len(nr)) + " records")
                                st.rerun()
                        else: st.info("All URLs already exist.")
                    else:
                        if apw == ADMIN_PASSWORD:
                            st.warning("Replace " + str(len(get_links())) + " records?")
                            if st.button("Replace All", type="primary", key="rb"):
                                save_links(ndf)
                                st.success("Replaced")
                                st.rerun()
                        else: st.error("Enter admin password.")
                except Exception as e: st.error("Error: " + str(e))
            st.markdown("---")
            if total > 0:
                buf = io.BytesIO()
                df.to_excel(buf, sheet_name=SHEET_NAME, index=False)
                st.download_button("📥 Export Excel", buf.getvalue(), file_name="AccessHub_" + datetime.now().strftime("%Y%m%d") + ".xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.markdown("---")
            hurl = st.text_input("Hub URL (for emails)", placeholder="Your Streamlit URL", key="hu")
            if hurl and total > 0:
                subj = urllib.parse.quote("Access Hub - Bookmark This")
                body = urllib.parse.quote("Hi Team,\n\nAccess all resources:\n" + hurl + "\n\nBookmark this link.\n\nTotal: " + str(total) + "\n\nRegards,\n" + un)
                st.markdown('<a href="mailto:?subject=' + subj + '&body=' + body + '" style="display:inline-block;padding:10px 20px;background:linear-gradient(135deg,#2563EB,#3B82F6);color:white;border-radius:8px;text-decoration:none;font-size:13px;font-weight:600;">📧 Notify All</a>', unsafe_allow_html=True)

        with dm2:
            st.markdown('<div class="sec-title">Add / Edit Entry</div>', unsafe_allow_html=True)
            editing = st.session_state.get("edit_idx") is not None
            edit_row = df.iloc[st.session_state.edit_idx].to_dict() if editing and st.session_state.edit_idx < len(df) else {}
            if total > 0:
                eo = {"-- New Entry --": None}
                eo.update({str(i+1) + ". " + r["Application Name"]: i for i, (_, r) in enumerate(df.iterrows())})
                sel = st.selectbox("Select to edit", list(eo.keys()), key="se")
                if eo[sel] is not None:
                    st.session_state.edit_idx = eo[sel]
                    edit_row = df.iloc[eo[sel]].to_dict()
                    editing = True
                else:
                    st.session_state.edit_idx = None
                    editing = False
                    edit_row = {}
            with st.form("entry_form", clear_on_submit=True):
                r1, r2 = st.columns(2)
                with r1:
                    fcat2 = st.text_input("Category", value=str(edit_row.get("Category","")), key="fc2")
                    ftype2 = st.selectbox("Link Type", ["Power BI","SharePoint","Excel","Web Page","Application","Document","PPT / Deck","Ops Resource","Other"], key="ft2")
                    fown2 = st.text_input("Owner", value=str(edit_row.get("Owner","")), key="fo2")
                    fupd2 = st.selectbox("Update Available", ["No","Yes"], key="fu2")
                with r2:
                    fnm2 = st.text_input("Application Name *", value=str(edit_row.get("Application Name","")), key="fn2")
                    furl2 = st.text_input("URL", value=str(edit_row.get("URL","")), key="furl2")
                    fst2 = st.selectbox("Status", ["Active","In Progress","Pending","Retired","Issue"], key="fs2")
                    fver2 = st.text_input("Version", value=str(edit_row.get("Version","")), key="fv2")
                fdesc2 = st.text_area("Description", value=str(edit_row.get("Description","")), key="fd2", height=60)
                fdept2 = st.text_input("Department", value=str(edit_row.get("Department","")), key="fdp2")
                fpri2 = st.selectbox("Priority", ["Medium","High","Low"], key="fp2")
                frem2 = st.text_area("Remarks", value=str(edit_row.get("Remarks","")), key="fr2", height=60)
                sb1, sb2, sb3 = st.columns([1,1,2])
                with sb1: save_btn = st.form_submit_button("💾 Save", type="primary")
                with sb2: del_btn = st.form_submit_button("🗑️ Delete")
                if save_btn:
                    if not fnm2.strip(): st.error("Name required.")
                    else:
                        now = datetime.now().strftime("%d-%b-%Y %I:%M:%S %p")
                        user = st.session_state.user_name or "Unknown"
                        nr = {"Category":fcat2,"Application Name":fnm2,"Description":fdesc2,"Link Type":ftype2,"URL":furl2,"Owner":fown2,"Status":fst2,"Update Available":fupd2,"Version":fver2,"Department":fdept2,"Priority":fpri2,"Remarks":frem2,"Last Modified By":user,"Last Modified Date":now}
                        if editing:
                            nr["Uploaded By"] = edit_row.get("Uploaded By", user)
                            nr["Uploaded Date"] = edit_row.get("Uploaded Date", now)
                            for c in COLUMNS: df.at[st.session_state.edit_idx, c] = nr.get(c, "")
                            save_links(df)
                            st.success("Updated")
                            st.session_state.edit_idx = None
                            st.rerun()
                        else:
                            nr["Uploaded By"] = user
                            nr["Uploaded Date"] = now
                            save_links(pd.concat([df, pd.DataFrame([nr])], ignore_index=True))
                            st.success("Added")
                            st.rerun()
                if del_btn and editing:
                    df = df.drop(index=st.session_state.edit_idx).reset_index(drop=True)
                    save_links(df)
                    st.success("Deleted")
                    st.session_state.edit_idx = None
                    st.rerun()

    with tab3:
        st.markdown('<div class="sec-title">Daily Status Tracker</div>', unsafe_allow_html=True)
        tc1, tc2 = st.columns([3,1])
        with tc2: tsearch = st.text_input("🔍 Search tasks", key="ts")
        tfilt = tdf.copy()
        if tsearch: tfilt = tfilt[tfilt.apply(lambda r: tsearch.lower() in " ".join(r.astype(str).values).lower(), axis=1)]
        render_tracker_table(tfilt)
        st.markdown("---")
        st.markdown('<div class="sec-title">Add / Edit Task</div>', unsafe_allow_html=True)
        t_editing = st.session_state.get("t_edit_idx") is not None
        t_edit_row = tdf.iloc[st.session_state.t_edit_idx].to_dict() if t_editing and st.session_state.t_edit_idx < len(tdf) else {}
        if tasks > 0:
            to2 = {"-- New Task --": None}
            to2.update({str(i+1) + ". " + r["Task Name"]: i for i, (_, r) in enumerate(tdf.iterrows())})
            tsel = st.selectbox("Select task", list(to2.keys()), key="tse")
            if to2[tsel] is not None:
                st.session_state.t_edit_idx = to2[tsel]
                t_edit_row = tdf.iloc[to2[tsel]].to_dict()
                t_editing = True
            else:
                st.session_state.t_edit_idx = None
                t_editing = False
                t_edit_row = {}
        with st.form("task_form", clear_on_submit=True):
            tr1, tr2 = st.columns(2)
            with tr1:
                tname = st.text_input("Task Name *", value=str(t_edit_row.get("Task Name","")), key="tn")
                towner = st.text_input("Owner", value=str(t_edit_row.get("Owner","")), key="to2")
                try: sdv = pd.to_datetime(t_edit_row.get("Start Date")).date() if t_editing and t_edit_row.get("Start Date") else date.today()
                except: sdv = date.today()
                tstart = st.date_input("Start Date", value=sdv, key="tsd")
                try: edv = pd.to_datetime(t_edit_row.get("Expected End Date")).date() if t_editing and t_edit_row.get("Expected End Date") else date.today()
                except: edv = date.today()
                tend = st.date_input("Expected End", value=edv, key="ted")
            with tr2:
                teta = st.text_input("ETA", value=str(t_edit_row.get("ETA","")), key="teta", placeholder="e.g. 3 days")
                tprog = st.select_slider("Progress", options=["0%","10%","20%","30%","40%","50%","60%","70%","80%","90%","100%"], value=str(t_edit_row.get("Progress","0%")) if t_editing and t_edit_row.get("Progress") else "0%", key="tp")
                tstatus = st.selectbox("Status", ["To Do","In Progress","Done","Blocked","Pending"], key="tst")
                tpri = st.selectbox("Priority", ["Medium","High","Low"], key="tpr")
            trem = st.text_area("Remarks", value=str(t_edit_row.get("Remarks","")), key="trm", height=60)
            tb1, tb2, tb3 = st.columns([1,1,2])
            with tb1: tsave = st.form_submit_button("💾 Save Task", type="primary")
            with tb2: tdel = st.form_submit_button("🗑️ Delete Task")
            if tsave:
                if not tname.strip(): st.error("Task Name required.")
                else:
                    now = datetime.now().strftime("%d-%b-%Y %I:%M:%S %p")
                    user = st.session_state.user_name or "Unknown"
                    nt = {"Task Name":tname,"Owner":towner,"Start Date":str(tstart),"Expected End Date":str(tend),"ETA":teta,"Progress":tprog,"Status":tstatus,"Priority":tpri,"Remarks":trem,"Last Modified By":user,"Last Modified Date":now}
                    if t_editing:
                        nt["Created By"] = t_edit_row.get("Created By", user)
                        nt["Created Date"] = t_edit_row.get("Created Date", now)
                        for c in TRACKER_COLUMNS: tdf.at[st.session_state.t_edit_idx, c] = nt.get(c, "")
                        save_tracker(tdf)
                        st.success("Task updated")
                        st.session_state.t_edit_idx = None
                        st.rerun()
                    else:
                        nt["Created By"] = user
                        nt["Created Date"] = now
                        save_tracker(pd.concat([tdf, pd.DataFrame([nt])], ignore_index=True))
                        st.success("Task added")
                        st.rerun()
            if tdel and t_editing:
                tdf = tdf.drop(index=st.session_state.t_edit_idx).reset_index(drop=True)
                save_tracker(tdf)
                st.success("Deleted")
                st.session_state.t_edit_idx = None
                st.rerun()
        if tasks > 0:
            st.markdown("---")
            buf2 = io.BytesIO()
            tdf.to_excel(buf2, sheet_name=TRACKER_SHEET, index=False)
            st.download_button("📥 Export Tracker", buf2.getvalue(), file_name="Tracker_" + datetime.now().strftime("%Y%m%d") + ".xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.markdown('<div style="text-align:center;color:#94A3B8;font-size:10px;margin-top:20px;">CENTRALIZED APPLICATION ACCESS HUB</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
