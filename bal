import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Therapy Clients Dashboard", layout="wide")

# ────────────────────────────────────────────────
#   TITLE & UPLOAD
# ────────────────────────────────────────────────
st.title("📊 Therapy Clients Dashboard")
st.caption("Upload → view summary → edit attendance → see replacement hours & module progress live")

uploaded = st.file_uploader("Upload Clients Main Data.xlsx", type=["xlsx"])

if not uploaded:
    st.info("Please upload your Excel file to begin.")
    st.stop()

try:
    xls = pd.ExcelFile(uploaded)
except Exception as e:
    st.error(f"Could not read the Excel file.\n{e}")
    st.stop()

# ────────────────────────────────────────────────
#   Identify client sheets (exclude known non-client sheets)
# ────────────────────────────────────────────────
all_sheets = xls.sheet_names
exclude = {"📅 Main Schedule", "Schedule Puasa", "📨 Schedule Template"}
client_sheets = [s for s in all_sheets if s not in exclude and len(s.strip()) > 0]

if not client_sheets:
    st.error("No client sheets found in the workbook.")
    st.stop()

# ────────────────────────────────────────────────
#   Helper – read one client sheet & extract important parts
# ────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def parse_client_sheet(sheet_name: str) -> tuple:
    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

    target_hours = 8.0
    attendance_start_row = None
    module_start_row = None

    for i, row in df.iterrows():
        txt = str(row[0]).strip().lower()
        if "monthly target" in txt:
            try:
                target_hours = float(txt.split(":")[-1].strip().split()[0])
            except:
                pass
        if txt == "date":
            attendance_start_row = i + 1
        if "module tracker" in txt:
            module_start_row = i + 2

    # ─── Attendance table ───────────────────────────────
    attendance_df = pd.DataFrame()
    if attendance_start_row is not None:
        # Try to find where the table ends (before next big section)
        end_row = len(df)
        for j in range(attendance_start_row, len(df)):
            if "monthly summary" in str(df.iloc[j,0]).lower():
                end_row = j
                break
        attendance_df = df.iloc[attendance_start_row:end_row, :6].copy()
        attendance_df.columns = ["DATE", "DAY", "TIME", "HOURS", "ATTENDANCE", "NOTES"]
        attendance_df = attendance_df[attendance_df["DATE"].notna()].reset_index(drop=True)
        attendance_df["HOURS"]   = pd.to_numeric(attendance_df["HOURS"],   errors="coerce").fillna(0)
        attendance_df["ATTENDANCE"] = attendance_df["ATTENDANCE"].astype(str).str.strip()

    # ─── Module table ───────────────────────────────────
    modules_df = pd.DataFrame(columns=["MODULE","STATUS","Start","End"])
    if module_start_row is not None:
        mod_end = len(df)
        for j in range(module_start_row, len(df)):
            txt = str(df.iloc[j,0]).strip().lower()
            if "replacement" in txt or "fee" in txt or "achievement" in txt:
                mod_end = j
                break
        modules_df = df.iloc[module_start_row:mod_end, :4].copy()
        modules_df.columns = ["MODULE","STATUS","Start","End"]
        modules_df = modules_df[modules_df["MODULE"].notna()].reset_index(drop=True)
        modules_df["Start"] = pd.to_datetime(modules_df["Start"], errors="coerce")
        modules_df["End"]   = pd.to_datetime(modules_df["End"],   errors="coerce")

    return target_hours, attendance_df, modules_df, df  # last one = full sheet for possible export


# ────────────────────────────────────────────────
#   GLOBAL SUMMARY
# ────────────────────────────────────────────────
summary_list = []

for name in client_sheets:
    target, att_df, _, _ = parse_client_sheet(name)
    present = att_df["ATTENDANCE"].str.lower().str.contains("present")
    completed = att_df.loc[present, "HOURS"].sum()
    remaining = max(0.0, target - completed)

    summary_list.append({
        "Client": name,
        "Target": target,
        "Completed": round(completed, 1),
        "Remaining": round(remaining, 1),
        "Sessions": len(att_df),
        "Present": present.sum(),
        "Replace Needed": "Yes" if remaining > 0 else "No"
    })

summary = pd.DataFrame(summary_list)

st.subheader("Overview – All Clients")
st.dataframe(
    summary.style.format({
        "Target": "{:.1f}",
        "Completed": "{:.1f}",
        "Remaining": "{:.1f}"
    }),
    use_container_width=True,
    hide_index=True
)

need_replace = summary[summary["Replace Needed"] == "Yes"]
if not need_replace.empty:
    st.warning(f"**{len(need_replace)} clients** need replacement classes (total ≈ {need_replace['Remaining'].sum():.1f} hours)")
else:
    st.success("All clients have met or exceeded their monthly target 🎉")


# ────────────────────────────────────────────────
#   SINGLE CLIENT DETAIL
# ────────────────────────────────────────────────
st.divider()
st.subheader("Detail & Live Editing")

chosen = st.selectbox("Select client", client_sheets, index=0)

target, att_df, mod_df, full_sheet = parse_client_sheet(chosen)

edited_att = st.data_editor(
    att_df,
    column_config={
        "DATE":         st.column_config.TextColumn("Date"),
        "DAY":          st.column_config.TextColumn("Day"),
        "TIME":         st.column_config.TextColumn("Time"),
        "HOURS":        st.column_config.NumberColumn("Hours", min_value=0.0, step=0.5, format="%.1f"),
        "ATTENDANCE":   st.column_config.SelectboxColumn("Status", options=["", "Present", "Absent"]),
        "NOTES":        st.column_config.TextColumn("Notes")
    },
    num_rows="dynamic",
    use_container_width=True,
    key=f"attendance_{chosen}"
)

# Recalculate
present = edited_att["ATTENDANCE"].str.lower().str.contains("present")
completed_h = edited_att.loc[present, "HOURS"].sum()
remaining_h = max(0.0, target - completed_h)

colA, colB, colC = st.columns([2,2,3])
colA.metric("Target", f"{target:.1f} hrs")
colB.metric("Completed", f"{completed_h:.1f} hrs", delta_color="normal")
colC.metric("Still needed", f"{remaining_h:.1f} hrs", delta_color="inverse" if remaining_h > 0 else "normal")

if remaining_h > 0:
    st.warning(f"→ **Replacement classes needed: {remaining_h:.1f} hours**")
else:
    st.success("→ Monthly target already met or exceeded")


# ────────────────────────────────────────────────
#   MODULE PROGRESS
# ────────────────────────────────────────────────
st.subheader(f"Module Progress – {chosen}")

edited_mod = st.data_editor(
    mod_df,
    column_config={
        "MODULE": st.column_config.TextColumn("Module"),
        "STATUS": st.column_config.SelectboxColumn("Status", options=["Not Started", "In Progress", "✅ Done"]),
        "Start":  st.column_config.DateColumn("Start Date"),
        "End":    st.column_config.DateColumn("End / Completed Date")
    },
    num_rows="dynamic",
    use_container_width=True,
    key=f"modules_{chosen}"
)

# ────────────────────────────────────────────────
#   Simple export placeholder (you can expand later)
# ────────────────────────────────────────────────
if st.button("Download current view (CSV)"):
    csv = edited_att.to_csv(index=False).encode('utf-8')
    st.download_button(
        "Download attendance table (CSV)",
        csv,
        f"{chosen}_attendance_{datetime.now().strftime('%Y%m%d')}.csv",
        "text/csv"
    )

st.caption("Tip: changes are live in this session only. Save/export when you're done.")
