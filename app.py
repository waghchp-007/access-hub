import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime
import urllib.parse
import io
import shutil
import os

# ============================================================
# CONFIG
# ============================================================
DATA_DIR = Path("data")
DATA_FILE = DATA_DIR / "alllinks.xlsx"
SHEET_NAME = "Links"
COLUMNS = [
    "Category", "Application Name", "Description", "Link Type", "URL", "Owner",
    "Uploaded By", "Uploaded Date", "Last Modified By", "Last Modified Date",
    "Status", "Update Available", "Version", "Department", "Priority", "Remarks"
]
STATUS_COLORS = {"Active": "🟢", "In Progress": "🔵", "Pending": "🟠", "Retired": "⚪", "Issue": "🔴"}
PRIORITY_COLORS = {"High": "🔴", "Medium": "🟡", "Low": "🟢"}
TYPE_ICONS = {
    "Power BI": "📊", "SharePoint": "🔷", "Excel": "📗", "Web Page": "🌐",
    "Application": "⚙️", "Document": "📄", "PPT / Deck": "📽️", "Ops Resource": "🔧", "Other": "🔗"
}

# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="Centralized Access Hub",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# CUSTOM CSS
# ============================================================
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .hub-header {
        background: linear-gradient(135deg, #0F1B33, #1B2A4A, #243B65);
        padding: 20px 28px; border-radius: 12px; margin-bottom: 16px; color: white;
    }
    .hub-header h1 { font-size: 24px; font-weight: 800; margin: 0; }
    .hub-header p { color: #94A3B8; font-size: 13px; margin: 4px 0 0; }
    .stat-row { display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 16px; }
    .stat-card {
        flex: 1; min-width: 120px; background: white;
        border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 16px; text-align: center;
    }
    .stat-card .num { font-size: 26px; font-weight: 800; line-height: 1.1; }
    .stat-card .lbl { font-size: 11px; color: #64748B; text-transform: uppercase; letter-spacing: 0.06em; }
    .stat-card.blue .num { color: #3B82F6; }
    .stat-card.green .num { color: #059669; }
    .stat-card.red .num { color: #DC2626; }
    .stat-card.purple .num { color: #7C3AED; }
    .badge { display: inline-block; padding: 2px 10px; border-radius: 999px; font-size: 12px; font-weight: 600; }
    .badge-active { background: #D1FAE5; color: #059669; }
    .badge-inprogress { background: #DBEAFE; color: #3B82F6; }
    .badge-pending { background: #FEF3C7; color: #D97706; }
    .badge-retired { background: #F1F5F9; color: #64748B; }
    .badge-issue { background: #FEE2E2; color: #DC2626; }
    .badge-yes { background: #FEE2E2; color: #DC2626; }
    .badge-no { background: #D1FAE5; color: #059669; }
    .cat-tag { display: inline-block; padding: 2px 8px; border-radius: 6px; font-size: 11px; font-weight: 600; }
    .cat-dashboard { background: #DBEAFE; color: #1E40AF; }
    .cat-application { background: #FEF3C7; color: #92400E; }
    .cat-document { background: #FCE7F3; color: #9D174D; }
    .cat-report { background: #E0E7FF; color: #3730A3; }
    .cat-webpage { background: #CCFBF1; color: #065F46; }
    .cat-other { background: #F1F5F9; color: #64748B; }
    .data-table { width: 100%; border-collapse: collapse; font-size: 13px; }
    .data-table th {
        background: #1B2A4A; color: white; padding: 10px 10px; text-align: left;
        font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.04em;
        white-space: nowrap; position: sticky; top: 0; z-index: 10;
    }
    .data-table td { padding: 8px 10px; border-bottom: 1px solid #E2E8F0; vertical-align: middle; max-width: 200px; }
    .data-table tr:nth-child(even) { background: #FAFBFD; }
    .data-table tr:hover { background: #EFF6FF; }
    .data-table .name-cell { font-weight: 600; color: #0F172A; }
    .data-table .desc-cell { color: #64748B; font-size: 12px; max-width: 180px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .data-table .meta-cell { font-size: 11px; color: #64748B; }
    .data-table .sr-cell { text-align: center; color: #94A3B8; font-weight: 600; font-size: 11px; }
    .act-btn {
        display: inline-flex; align-items: center; justify-content: center;
        width: 28px; height: 28px; border-radius: 6px;
        border: 1px solid #E2E8F0; background: white;
        cursor: pointer; font-size: 13px; text-decoration: none; transition: all 0.15s;
    }
    .act-btn:hover { background: #DBEAFE; border-color: #3B82F6; }
    .section-title { font-size: 14px; font-weight: 700; color: #1B2A4A; margin: 18px 0 10px; border-bottom: 2px solid #3B82F6; padding-bottom: 6px; }
    div[data-testid="stSidebar"] { background: #F8FAFC; }
    .stButton > button { border-radius: 8px; }
</style>
""", unsafe_allow_html=True)


# ============================================================
# DATA FUNCTIONS — uses session_state as live store + Excel file as base
# ============================================================
def ensure_data_file():
    DATA_DIR.mkdir(exist_ok=True)
    if not DATA_FILE.exists():
        df = pd.DataFrame(columns=COLUMNS)
        df.to_excel(DATA_FILE, sheet_name=SHEET_NAME, index=False)


def load_data_from_file():
    ensure_data_file()
    try:
        df = pd.read_excel(DATA_FILE, sheet_name=SHEET_NAME)
        for col in COLUMNS:
            if col not in df.columns:
                df[col] = ""
        df = df.fillna("")
        return df
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return pd.DataFrame(columns=COLUMNS)


def get_data():
    """Get data from session state (live edits) or fall back to file."""
    if "data" not in st.session_state:
        st.session_state.data = load_data_from_file()
    return st.session_state.data.copy()


def save_data(df):
    """Save to both session state AND Excel file."""
    st.session_state.data = df.copy()
    try:
        ensure_data_file()
        # Backup before writing
        if DATA_FILE.exists():
            backup = DATA_DIR / f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            shutil.copy2(DATA_FILE, backup)
            # Keep only last 3 backups
            backups = sorted(DATA_DIR.glob("backup_*.xlsx"))
            for old in backups[:-3]:
                old.unlink()
        df.to_excel(DATA_FILE, sheet_name=SHEET_NAME, index=False)
    except Exception as e:
        st.warning(f"Could not write to Excel file: {e}. Data is saved in session.")


def get_cat_class(cat):
    c = str(cat).lower()
    if "dashboard" in c or "power bi" in c: return "cat-dashboard"
    if "app" in c: return "cat-application"
    if "doc" in c: return "cat-document"
    if "report" in c or "excel" in c: return "cat-report"
    if "web" in c: return "cat-webpage"
    return "cat-other"


def get_status_class(s):
    sl = str(s).lower().replace(" ", "")
    return {"active": "badge-active", "inprogress": "badge-inprogress",
            "pending": "badge-pending", "retired": "badge-retired",
            "issue": "badge-issue"}.get(sl, "badge-active")


def render_table(df):
    if df.empty:
        st.info("No records found. Use **Add Entry** or **Upload Excel** to get started.")
        return

    html = '<div style="overflow-x:auto;max-height:600px;overflow-y:auto;"><table class="data-table"><thead><tr>'
    headers = ["#", "Category", "Name", "Type", "Description", "Owner", "Status",
               "Update", "Ver.", "Uploaded By", "Modified By", "Modified Date", "Actions"]
    for h in headers:
        html += f"<th>{h}</th>"
    html += "</tr></thead><tbody>"

    for idx, (_, row) in enumerate(df.iterrows()):
        url = str(row.get("URL", ""))
        has_url = url.startswith("http")
        cat_cls = get_cat_class(row.get("Category", ""))
        status_cls = get_status_class(row.get("Status", ""))
        upd = str(row.get("Update Available", "")).strip().lower()
        upd_cls = "badge-yes" if upd == "yes" else "badge-no"
        upd_txt = "▲ Yes" if upd == "yes" else "✓ No"
        pri = str(row.get("Priority", ""))
        pri_dot = PRIORITY_COLORS.get(pri, "")
        type_icon = TYPE_ICONS.get(str(row.get("Link Type", "")), "🔗")
        mod_date = str(row.get("Last Modified Date", ""))[:20]
        if mod_date == "nan": mod_date = "—"

        open_btn = f'<a href="{url}" target="_blank" class="act-btn" title="Open">🔗</a>' if has_url else '<span class="act-btn" style="opacity:0.3;">✖</span>'

        subj = urllib.parse.quote(f"Review: {row.get('Application Name', 'Resource')}")
        body_text = f"Hi {row.get('Owner', 'Team')},\n\nPlease review: {row.get('Application Name', '')}\nLink: {url}\n\nRegards"
        body = urllib.parse.quote(body_text)
        notify_btn = f'<a href="mailto:?subject={subj}&body={body}" class="act-btn" title="Notify">🔔</a>'

        # Escape HTML in description
        desc = str(row.get('Description', '—')).replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')
        name = str(row.get('Application Name', '—')).replace('<', '&lt;').replace('>', '&gt;')

        html += f"""<tr>
            <td class="sr-cell">{idx+1}</td>
            <td><span class="cat-tag {cat_cls}">{pri_dot} {row.get('Category','—')}</span></td>
            <td class="name-cell">{name}</td>
            <td><span style="font-size:12px;">{type_icon} {row.get('Link Type','')}</span></td>
            <td class="desc-cell" title="{desc}">{desc}</td>
            <td class="meta-cell">{row.get('Owner','—')}</td>
            <td><span class="badge {status_cls}">{STATUS_COLORS.get(str(row.get('Status','')), '')} {row.get('Status','—')}</span></td>
            <td><span class="badge {upd_cls}">{upd_txt}</span></td>
            <td class="meta-cell">{row.get('Version','—')}</td>
            <td class="meta-cell">{row.get('Uploaded By','—')}</td>
            <td class="meta-cell">{row.get('Last Modified By','—')}</td>
            <td class="meta-cell">{mod_date}</td>
            <td style="white-space:nowrap;">{open_btn} {notify_btn}</td>
        </tr>"""

    html += "</tbody></table></div>"
    st.markdown(html, unsafe_allow_html=True)


# ============================================================
# MAIN APP
# ============================================================
def main():
    if "user_name" not in st.session_state:
        st.session_state.user_name = ""
    if "show_add_form" not in st.session_state:
        st.session_state.show_add_form = False
    if "edit_idx" not in st.session_state:
        st.session_state.edit_idx = None

    df = get_data()

    # ---- HEADER ----
    st.markdown("""
    <div class="hub-header">
        <h1>📊 Centralized Access Hub</h1>
        <p>All dashboards, applications, documents, and resources — one place.</p>
    </div>
    """, unsafe_allow_html=True)

    # ---- STATS ----
    total = len(df)
    active = len(df[df["Status"].str.lower() == "active"]) if total else 0
    updates = len(df[df["Update Available"].str.lower() == "yes"]) if total else 0
    cats = df["Category"].nunique() if total else 0

    st.markdown(f"""
    <div class="stat-row">
        <div class="stat-card blue"><div class="num">{total}</div><div class="lbl">Total Links</div></div>
        <div class="stat-card green"><div class="num">{active}</div><div class="lbl">Active</div></div>
        <div class="stat-card red"><div class="num">{updates}</div><div class="lbl">Updates Pending</div></div>
        <div class="stat-card purple"><div class="num">{cats}</div><div class="lbl">Categories</div></div>
    </div>
    """, unsafe_allow_html=True)

    # ---- SIDEBAR ----
    with st.sidebar:
        st.markdown("### ⚙️ Settings")
        user_name = st.text_input("Your Name", value=st.session_state.user_name, placeholder="Enter your name")
        st.session_state.user_name = user_name

        st.markdown("---")
        st.markdown("### 🔍 Filters")
        search = st.text_input("Search", placeholder="Name, category, owner...")

        categories = ["All"] + sorted(df["Category"].unique().tolist()) if total else ["All"]
        f_cat = st.selectbox("Category", categories)

        types = ["All"] + sorted(df["Link Type"].unique().tolist()) if total else ["All"]
        f_type = st.selectbox("Link Type", types)

        statuses = ["All"] + sorted(df["Status"].unique().tolist()) if total else ["All"]
        f_status = st.selectbox("Status", statuses)

        st.markdown("---")
        st.markdown("### 📂 Data Management")

        uploaded = st.file_uploader("📤 Upload Excel", type=["xlsx", "xls"], label_visibility="collapsed")
        if uploaded:
            try:
                new_df = pd.read_excel(uploaded, sheet_name=0)
                new_df = new_df.fillna("")
                for col in COLUMNS:
                    if col not in new_df.columns:
                        new_df[col] = ""
                save_data(new_df)
                st.success(f"✅ Loaded {len(new_df)} records from {uploaded.name}")
                st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")

        if total > 0:
            buf = io.BytesIO()
            df.to_excel(buf, sheet_name=SHEET_NAME, index=False)
            st.download_button("📥 Export Excel", buf.getvalue(),
                               file_name=f"AccessHub_Export_{datetime.now().strftime('%Y%m%d')}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.markdown("---")
        hub_url = st.text_input("🔗 Hub URL (for emails)",
                                value=st.session_state.get("hub_url", ""),
                                placeholder="Your Streamlit Cloud URL",
                                key="hub_url_input")
        if hub_url:
            st.session_state.hub_url = hub_url
            subj = urllib.parse.quote("Centralized Access Hub – Bookmark This Link")
            body = urllib.parse.quote(f"Hi Team,\n\nAccess all resources:\n{hub_url}\n\nPlease bookmark this link.\n\nTotal: {total} | Active: {active}\n\nRegards,\n{user_name}")
            st.markdown(f'<a href="mailto:?subject={subj}&body={body}" style="display:inline-block;padding:8px 16px;background:#3B82F6;color:white;border-radius:8px;text-decoration:none;font-size:13px;font-weight:600;width:100%;text-align:center;">📧 Notify All Team</a>', unsafe_allow_html=True)

    # ---- FILTER ----
    filtered = df.copy()
    if search:
        mask = filtered.apply(lambda row: search.lower() in " ".join(row.astype(str).values).lower(), axis=1)
        filtered = filtered[mask]
    if f_cat != "All":
        filtered = filtered[filtered["Category"] == f_cat]
    if f_type != "All":
        filtered = filtered[filtered["Link Type"] == f_type]
    if f_status != "All":
        filtered = filtered[filtered["Status"] == f_status]

    # ---- ACTION BUTTONS ----
    col1, col2, col3, col4 = st.columns([1, 1, 1, 4])
    with col1:
        if st.button("➕ Add Entry", use_container_width=True):
            st.session_state.show_add_form = True
            st.session_state.edit_idx = None
    with col2:
        if st.button("🔄 Refresh", use_container_width=True):
            st.session_state.data = load_data_from_file()
            st.rerun()
    with col3:
        st.markdown(f"**{len(filtered)}** of **{total}** shown")

    # ---- ADD / EDIT FORM ----
    if st.session_state.show_add_form or st.session_state.edit_idx is not None:
        editing = st.session_state.edit_idx is not None
        title = "✏️ Edit Entry" if editing else "➕ Add New Entry"

        with st.expander(title, expanded=True):
            edit_row = df.iloc[st.session_state.edit_idx].to_dict() if editing else {}

            c1, c2 = st.columns(2)
            with c1:
                f_category = st.text_input("Category", value=str(edit_row.get("Category", "")), key="fc")
                type_opts = ["Power BI", "SharePoint", "Excel", "Web Page", "Application", "Document", "PPT / Deck", "Ops Resource", "Other"]
                type_default = type_opts.index(edit_row["Link Type"]) if editing and edit_row.get("Link Type") in type_opts else 0
                f_link_type = st.selectbox("Link Type", type_opts, index=type_default, key="ft")
                f_owner = st.text_input("Owner", value=str(edit_row.get("Owner", "")), key="fo")
                f_update = st.selectbox("Update Available", ["No", "Yes"], index=1 if str(edit_row.get("Update Available", "")).lower() == "yes" else 0, key="fu")
                f_dept = st.text_input("Department", value=str(edit_row.get("Department", "")), key="fd")
            with c2:
                f_name = st.text_input("Application Name *", value=str(edit_row.get("Application Name", "")), key="fn")
                f_url = st.text_input("URL", value=str(edit_row.get("URL", "")), key="furl")
                status_opts = ["Active", "In Progress", "Pending", "Retired", "Issue"]
                status_default = status_opts.index(edit_row["Status"]) if editing and edit_row.get("Status") in status_opts else 0
                f_status_input = st.selectbox("Status", status_opts, index=status_default, key="fs")
                f_version = st.text_input("Version", value=str(edit_row.get("Version", "")), key="fv")
                pri_opts = ["Medium", "High", "Low"]
                pri_default = pri_opts.index(edit_row["Priority"]) if editing and edit_row.get("Priority") in pri_opts else 0
                f_priority = st.selectbox("Priority", pri_opts, index=pri_default, key="fp")

            f_desc = st.text_area("Description", value=str(edit_row.get("Description", "")), key="fdesc", height=68)
            f_remarks = st.text_area("Remarks", value=str(edit_row.get("Remarks", "")), key="fr", height=68)

            bc1, bc2, bc3 = st.columns([1, 1, 4])
            with bc1:
                if st.button("💾 Save", type="primary", use_container_width=True):
                    if not f_name.strip():
                        st.error("Application Name is required.")
                    else:
                        now = datetime.now().strftime("%d-%b-%Y %I:%M:%S %p")
                        user = st.session_state.user_name or "Unknown"
                        new_row = {
                            "Category": f_category, "Application Name": f_name,
                            "Description": f_desc, "Link Type": f_link_type,
                            "URL": f_url, "Owner": f_owner,
                            "Status": f_status_input, "Update Available": f_update,
                            "Version": f_version, "Department": f_dept,
                            "Priority": f_priority, "Remarks": f_remarks,
                            "Last Modified By": user, "Last Modified Date": now,
                        }
                        if editing:
                            new_row["Uploaded By"] = edit_row.get("Uploaded By", user)
                            new_row["Uploaded Date"] = edit_row.get("Uploaded Date", now)
                            for col in COLUMNS:
                                df.at[st.session_state.edit_idx, col] = new_row.get(col, "")
                            save_data(df)
                            st.success(f"✅ Updated by {user}")
                        else:
                            new_row["Uploaded By"] = user
                            new_row["Uploaded Date"] = now
                            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                            save_data(df)
                            st.success(f"✅ Added by {user}")
                        st.session_state.show_add_form = False
                        st.session_state.edit_idx = None
                        st.rerun()
            with bc2:
                if st.button("Cancel", use_container_width=True):
                    st.session_state.show_add_form = False
                    st.session_state.edit_idx = None
                    st.rerun()

    # ---- TABLE ----
    st.markdown('<div class="section-title">📋 All Resources</div>', unsafe_allow_html=True)
    render_table(filtered)

    # ---- EDIT / DELETE ----
    if total > 0:
        st.markdown("---")
        st.markdown("##### Quick Edit")
        edit_options = {f"{i+1}. {row['Application Name']} ({row['Category']})": i for i, (_, row) in enumerate(df.iterrows())}
        selected = st.selectbox("Select entry to edit", [""] + list(edit_options.keys()), key="es")
        if selected and selected in edit_options:
            ec1, ec2 = st.columns([1, 1])
            with ec1:
                if st.button("✏️ Edit Selected", use_container_width=True):
                    st.session_state.edit_idx = edit_options[selected]
                    st.session_state.show_add_form = False
                    st.rerun()
            with ec2:
                if st.button("🗑️ Delete Selected", use_container_width=True):
                    idx = edit_options[selected]
                    df = df.drop(index=idx).reset_index(drop=True)
                    save_data(df)
                    st.success("🗑️ Deleted.")
                    st.rerun()

    # ---- FOOTER ----
    st.markdown("---")
    st.markdown(
        f'<div style="text-align:center;color:#94A3B8;font-size:11px;">'
        f'CENTRALIZED ACCESS HUB · '
        f'Last refresh: {datetime.now().strftime("%d-%b-%Y %I:%M %p")}'
        f'</div>', unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
