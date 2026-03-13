import streamlit as st
import pandas as pd
from datetime import datetime, date
import io
import calendar

st.set_page_config(page_title="Therapy Clients Dashboard", layout="wide", page_icon="📋")

# ─────────────────────────────────────────
#  CUSTOM CSS
# ─────────────────────────────────────────
st.markdown("""
<style>
    .metric-card {
        background: #f8f9fa;
        border-radius: 12px;
        padding: 16px 20px;
        border-left: 5px solid #4e8df5;
        margin-bottom: 10px;
    }
    .metric-card.warn { border-left-color: #f5a623; }
    .metric-card.good { border-left-color: #27ae60; }
    .metric-card.danger { border-left-color: #e74c3c; }
    .metric-card h4 { margin: 0; font-size: 13px; color: #666; }
    .metric-card h2 { margin: 4px 0 0; font-size: 26px; }
    .section-header {
        font-size: 18px; font-weight: 700;
        border-bottom: 2px solid #4e8df5;
        padding-bottom: 6px; margin: 20px 0 12px;
        color: #1a1a2e;
    }
    .badge {
        display:inline-block; padding:3px 10px; border-radius:20px;
        font-size:12px; font-weight:600;
    }
    .badge-done { background:#d4edda; color:#155724; }
    .badge-progress { background:#fff3cd; color:#856404; }
    .badge-pending { background:#f8d7da; color:#721c24; }
    .schedule-msg {
        background:#eef6ff; border:1px solid #bee3f8;
        border-radius:10px; padding:16px 20px;
        white-space:pre-wrap; font-family:monospace; font-size:13px;
    }
    .stTabs [data-baseweb="tab"] { font-size:15px; font-weight:600; }
    .client-header {
        background: linear-gradient(135deg,#4e8df5,#764ef5);
        color:white; border-radius:12px; padding:20px 24px; margin-bottom:20px;
    }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────
CLIENT_SHEETS = ['Jannah','Ariq','Aisy Affan','Anas','Hanis','Adeeb',
                 'Aleya','Aisha Sofea','Aleena','Megat','Aaisyah Sophea',
                 'Afeef Zayyan','Orkid']

DAY_MAP = {0:'Isnin',1:'Selasa',2:'Rabu',3:'Khamis',4:'Jumaat',5:'Sabtu',6:'Ahad'}
DAY_MAP_EN = {0:'Monday',1:'Tuesday',2:'Wednesday',3:'Thursday',4:'Friday',5:'Saturday',6:'Sunday'}

def parse_client_sheet(df):
    """Parse a client individual sheet into structured data."""
    raw = df.iloc[:, 0].astype(str)
    name = raw.iloc[0].strip()
    target_str = raw.iloc[1]
    try:
        target = float(''.join(filter(lambda c: c.isdigit() or c=='.', target_str)))
    except:
        target = 8.0

    # Sessions: rows where col0 looks like a date
    sessions = []
    for _, row in df.iterrows():
        val = row.iloc[0]
        if pd.isna(val):
            continue
        try:
            d = pd.to_datetime(val)
            day_    = str(row.iloc[1]) if not pd.isna(row.iloc[1]) else ''
            time_   = str(row.iloc[2]) if not pd.isna(row.iloc[2]) else ''
            hours_  = row.iloc[3]
            attend_ = str(row.iloc[4]) if not pd.isna(row.iloc[4]) else ''
            notes_  = str(row.iloc[5]) if not pd.isna(row.iloc[5]) else ''
            if day_ not in ('nan','') or time_ not in ('nan',''):
                sessions.append({
                    'date': d,
                    'day': day_,
                    'time': time_,
                    'hours': pd.to_numeric(hours_, errors='coerce'),
                    'attendance': attend_,
                    'notes': notes_
                })
        except:
            pass

    # Monthly summary row values
    hours_done = 0.0
    for _, row in df.iterrows():
        if str(row.iloc[0]).strip() == 'Hours Completed This Month':
            try:
                hours_done = float(row.iloc[3])
            except:
                pass

    # Modules
    modules = []
    in_module = False
    for _, row in df.iterrows():
        label = str(row.iloc[0]).strip()
        if '📚 MODULE' in label or 'MODULE TRACKER' in label:
            in_module = True
            continue
        if in_module and label not in ('nan','') and 'MODULE' not in label:
            status = str(row.iloc[3]).strip() if not pd.isna(row.iloc[3]) else ''
            if status not in ('nan','','STATUS') and label not in ('🔄 REPLACEMENT','💳 FEE','🏆 ACHIEVE'):
                modules.append({'module': label, 'status': status})
        if in_module and ('🔄' in label or '💳' in label or '🏆' in label):
            in_module = False

    # Replacement sessions
    replacements = []
    in_repl = False
    for _, row in df.iterrows():
        label = str(row.iloc[0]).strip()
        if '🔄 REPLACEMENT' in label:
            in_repl = True
            continue
        if in_repl:
            if '💳' in label or '🏆' in label:
                in_repl = False
                continue
            val = row.iloc[0]
            try:
                d = pd.to_datetime(val)
                replacements.append({
                    'date': d,
                    'day': str(row.iloc[1]) if not pd.isna(row.iloc[1]) else '',
                    'time': str(row.iloc[2]) if not pd.isna(row.iloc[2]) else '',
                    'hours': pd.to_numeric(row.iloc[3], errors='coerce'),
                    'status': str(row.iloc[4]) if not pd.isna(row.iloc[4]) else ''
                })
            except:
                pass

    # Fee tracking
    fees = []
    in_fee = False
    for _, row in df.iterrows():
        label = str(row.iloc[0]).strip()
        if '💳 FEE' in label:
            in_fee = True
            continue
        if in_fee:
            if '🏆' in label or label == 'TOTAL':
                in_fee = False
                continue
            try:
                d = pd.to_datetime(row.iloc[0])
                fees.append({
                    'month': d,
                    'amount': pd.to_numeric(row.iloc[1], errors='coerce'),
                    'status': str(row.iloc[2]) if not pd.isna(row.iloc[2]) else 'Pending',
                    'notes': str(row.iloc[5]) if not pd.isna(row.iloc[5]) else ''
                })
            except:
                pass

    return {
        'name': name,
        'target': target,
        'sessions': sessions,
        'hours_done': hours_done,
        'modules': modules,
        'replacements': replacements,
        'fees': fees
    }


def load_all_clients(xl):
    clients = {}
    for sheet in CLIENT_SHEETS:
        if sheet in xl.sheet_names:
            df = xl.parse(sheet, header=None)
            data = parse_client_sheet(df)
            clients[sheet] = data
    return clients


def hours_balance(client_data):
    done = client_data['hours_done']
    target = client_data['target']
    remaining_sessions = [s for s in client_data['sessions']
                          if pd.isna(s['hours']) and s['date'] >= pd.Timestamp(datetime.today())]
    scheduled_future = sum(
        s.get('hours', 0) or 0
        for s in client_data['sessions']
        if not pd.isna(s.get('hours')) and str(s.get('attendance','')).strip() == ''
    )
    balance = target - done
    return done, target, balance


def generate_schedule_message(client_name, sessions, month_year):
    """Generate WhatsApp-style schedule message."""
    valid = [s for s in sessions if not pd.isna(s['date'])]
    valid.sort(key=lambda x: x['date'])

    lines = [
        f"Assalamualaikum Tuan/Puan ,",
        f"",
        f"Berikut jadual sesi kelas fizikal untuk bulan {month_year}. Sila rujuk jadual di bawah :",
        f""
    ]
    for i, s in enumerate(valid, 1):
        d = s['date']
        date_str = f"{d.day}/{d.month}/{str(d.year)[2:]}"
        lines.append(f"📒 Kelas {i} : {date_str} ({s['day']}) , {s['time']}")

    lines += [
        "",
        "Keprihatinan tuan/puan amat kami hargai. Kepada ibubapa yang telah",
        "menjelaskan yuran, diucapkan ribuan terima kasih."
    ]
    return "\n".join(lines)


# ─────────────────────────────────────────
#  MAIN APP
# ─────────────────────────────────────────
st.title("📋 Therapy Clients Dashboard")
st.caption("Upload → View Summary → Edit Attendance → Schedule Generator → Track Progress")

uploaded = st.file_uploader("📁 Upload Clients_Main_Data.xlsx", type=["xlsx"])

if not uploaded:
    st.info("Please upload your Excel file to begin.")
    st.stop()

xl = pd.ExcelFile(uploaded)
clients = load_all_clients(xl)

# ═══════════════════════════════════════
#  TABS
# ═══════════════════════════════════════
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Overview",
    "📅 Schedule View",
    "⏱️ Hours & Balance",
    "📚 Module Progress",
    "📨 Schedule Generator"
])

# ══════════════════════════════
#  TAB 1 – OVERVIEW
# ══════════════════════════════
with tab1:
    st.markdown('<div class="section-header">🏠 All Clients Summary</div>', unsafe_allow_html=True)

    total_clients = len(clients)
    need_replacement = sum(1 for c in clients.values()
                           if c['target'] - c['hours_done'] > 0 and c['hours_done'] > 0)
    paid_count = sum(1 for c in clients.values()
                     if any(f['status'] == 'Paid' for f in c['fees']))

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Clients", total_clients)
    with col2:
        total_hrs = sum(c['hours_done'] for c in clients.values())
        st.metric("Total Hours Done (March)", f"{total_hrs:.1f} hrs")
    with col3:
        st.metric("Need Replacement", need_replacement, delta="clients" if need_replacement else "✓ All good")
    with col4:
        st.metric("Fees Paid", f"{paid_count}/{total_clients}")

    st.markdown("---")

    # Summary table
    rows = []
    for sheet, c in clients.items():
        done, target, balance = hours_balance(c)
        attended = sum(1 for s in c['sessions'] if str(s['attendance']).strip() == 'Present')
        absent   = sum(1 for s in c['sessions'] if str(s['attendance']).strip() == 'Absent')
        fee_status = next((f['status'] for f in c['fees']
                           if f['month'].month == datetime.today().month), 'Pending')
        module_statuses = [m['status'] for m in c['modules']]
        current_module = next((m['module'] for m in c['modules']
                               if 'progress' in m['status'].lower() or 'Progress' in m['status']), '—')
        rows.append({
            'Client': c['name'],
            'Target (hrs)': target,
            'Done (hrs)': done,
            'Balance': balance,
            'Attended': attended,
            'Absent': absent,
            'Needs Replacement': '⚠️ Yes' if balance > 0 else '✅ No',
            'Fee (March)': fee_status,
            'Current Module': current_module
        })

    df_summary = pd.DataFrame(rows)

    def highlight_balance(val):
        if isinstance(val, str) and '⚠️' in val:
            return 'background-color: #fff3cd; color:#856404'
        if isinstance(val, str) and '✅' in val:
            return 'background-color: #d4edda; color:#155724'
        return ''

    def highlight_fee(val):
        if val == 'Paid':
            return 'background-color: #d4edda; color:#155724'
        if val == 'Pending':
            return 'background-color: #f8d7da; color:#721c24'
        return ''

    styled = df_summary.style \
        .applymap(highlight_balance, subset=['Needs Replacement']) \
        .applymap(highlight_fee, subset=['Fee (March)'])

    st.dataframe(styled, use_container_width=True, hide_index=True)


# ══════════════════════════════
#  TAB 2 – SCHEDULE VIEW
# ══════════════════════════════
with tab2:
    st.markdown('<div class="section-header">📅 Schedule View</div>', unsafe_allow_html=True)

    view_mode = st.radio("View Mode", ["Weekly Recurring", "By Client", "Calendar View"],
                         horizontal=True)

    if view_mode == "Weekly Recurring":
        if '📅 Main Schedule' in xl.sheet_names:
            df_sched = xl.parse('📅 Main Schedule', header=None)
            st.markdown("**Regular Weekly Schedule**")
            day_order = ['SELASA','RABU','KHAMIS','JUMAAT','SABTU']
            day_emoji = {'SELASA':'🔵','RABU':'🟢','KHAMIS':'🟡','JUMAAT':'🟠','SABTU':'🔴'}
            current_day = None
            day_data = {}
            for _, row in df_sched.iterrows():
                val = str(row.iloc[0]).strip()
                if val.upper() in day_order:
                    current_day = val.upper()
                    day_data[current_day] = []
                elif current_day and str(row.iloc[2]) not in ('nan','NaN','') and not pd.isna(row.iloc[2]):
                    day_data[current_day].append({
                        'Time': str(row.iloc[1]),
                        'Client': str(row.iloc[2]),
                        'Therapy': str(row.iloc[3]),
                        'Package/Age': str(row.iloc[4]),
                        'Target hrs/month': row.iloc[5],
                        'Notes': str(row.iloc[6]) if not pd.isna(row.iloc[6]) else ''
                    })

            cols = st.columns(len(day_data))
            for i, (day, sessions) in enumerate(day_data.items()):
                with cols[i]:
                    emoji = day_emoji.get(day, '📌')
                    st.markdown(f"**{emoji} {day}**")
                    if sessions:
                        df_d = pd.DataFrame(sessions)
                        st.dataframe(df_d[['Time','Client','Therapy','Package/Age']],
                                     hide_index=True, use_container_width=True)
                    else:
                        st.caption("No sessions")

    elif view_mode == "By Client":
        selected = st.selectbox("Select Client", list(clients.keys()))
        c = clients[selected]
        st.markdown(f"### {c['name']}")

        sessions_df = pd.DataFrame(c['sessions'])
        if not sessions_df.empty:
            sessions_df['date'] = pd.to_datetime(sessions_df['date']).dt.strftime('%d/%m/%Y')

            def color_attendance(val):
                if val == 'Present': return 'color: green; font-weight:600'
                if val == 'Absent': return 'color: red; font-weight:600'
                return 'color: gray'

            styled_s = sessions_df[['date','day','time','hours','attendance','notes']].style \
                .applymap(color_attendance, subset=['attendance'])
            st.dataframe(styled_s, use_container_width=True, hide_index=True)
        else:
            st.info("No sessions recorded.")

    elif view_mode == "Calendar View":
        sel_month = st.selectbox("Select Month", [
            "March 2026", "April 2026", "May 2026", "June 2026"])
        month_num = {"March 2026":3,"April 2026":4,"May 2026":5,"June 2026":6}[sel_month]

        # Collect all sessions for that month
        all_sessions = []
        for sheet, c in clients.items():
            for s in c['sessions']:
                if pd.to_datetime(s['date']).month == month_num:
                    all_sessions.append({
                        'date': pd.to_datetime(s['date']),
                        'client': c['name'],
                        'time': s['time'],
                        'attendance': s['attendance']
                    })

        if all_sessions:
            df_cal = pd.DataFrame(all_sessions)
            df_cal['day_num'] = df_cal['date'].dt.day

            st.markdown(f"**{sel_month} — All Sessions**")
            pivot = df_cal.groupby(['day_num','client']).agg(
                time=('time','first'),
                attendance=('attendance','first')
            ).reset_index()
            pivot.columns = ['Day','Client','Time','Attendance']
            st.dataframe(pivot, use_container_width=True, hide_index=True)
        else:
            st.info("No sessions found for this month.")


# ══════════════════════════════
#  TAB 3 – HOURS & BALANCE
# ══════════════════════════════
with tab3:
    st.markdown('<div class="section-header">⏱️ Hours Done & Balance to Replace</div>',
                unsafe_allow_html=True)

    # Filter
    filter_opt = st.radio("Filter", ["All Clients", "Needs Replacement Only", "On Track"],
                          horizontal=True)

    for sheet, c in clients.items():
        done, target, balance = hours_balance(c)

        if filter_opt == "Needs Replacement Only" and balance <= 0:
            continue
        if filter_opt == "On Track" and balance > 0:
            continue

        status_class = "danger" if balance > 2 else ("warn" if balance > 0 else "good")
        status_icon  = "⚠️" if balance > 0 else "✅"

        with st.expander(f"{status_icon} {c['name']}  — Done: {done}h / Target: {target}h / Balance: {balance}h"):
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Hours Done", f"{done} hrs")
            with col2:
                st.metric("Target", f"{target} hrs")
            with col3:
                st.metric("Balance / To Replace", f"{balance} hrs",
                          delta=f"{'⚠️ Need replacement' if balance > 0 else '✅ Met'}")

            # Progress bar
            pct = min(done / target, 1.0) if target > 0 else 0
            st.progress(pct, text=f"{pct*100:.0f}% complete")

            # Sessions table
            sessions_df = pd.DataFrame(c['sessions'])
            if not sessions_df.empty:
                sessions_df['date'] = pd.to_datetime(sessions_df['date']).dt.strftime('%d/%m/%Y')
                st.dataframe(sessions_df[['date','day','time','hours','attendance','notes']],
                             hide_index=True, use_container_width=True)

            # Replacement sessions
            if c['replacements']:
                st.markdown("**🔄 Replacement Sessions**")
                repl_df = pd.DataFrame(c['replacements'])
                repl_df['date'] = pd.to_datetime(repl_df['date']).dt.strftime('%d/%m/%Y')
                st.dataframe(repl_df, hide_index=True, use_container_width=True)
            elif balance > 0:
                st.warning(f"⚠️ {balance} hour(s) need to be replaced — no replacement sessions scheduled yet.")

    st.markdown("---")
    st.markdown('<div class="section-header">📊 Summary Chart</div>', unsafe_allow_html=True)

    chart_data = pd.DataFrame([
        {'Client': c['name'], 'Done': c['hours_done'], 'Target': c['target']}
        for c in clients.values()
    ]).set_index('Client')

    st.bar_chart(chart_data)


# ══════════════════════════════
#  TAB 4 – MODULE PROGRESS
# ══════════════════════════════
with tab4:
    st.markdown('<div class="section-header">📚 Module Progress Tracker</div>',
                unsafe_allow_html=True)

    # Overview table
    mod_rows = []
    for sheet, c in clients.items():
        for m in c['modules']:
            mod_rows.append({'Client': c['name'], 'Module': m['module'], 'Status': m['status']})

    if mod_rows:
        df_mod = pd.DataFrame(mod_rows)

        def badge_status(val):
            v = str(val).lower()
            if '✅' in val or 'done' in v:
                return 'background-color:#d4edda; color:#155724; font-weight:600'
            if 'progress' in v:
                return 'background-color:#fff3cd; color:#856404; font-weight:600'
            return ''

        styled_mod = df_mod.style.applymap(badge_status, subset=['Status'])
        st.dataframe(styled_mod, use_container_width=True, hide_index=True)

    st.markdown("---")

    # Per-client detail
    selected_mod = st.selectbox("Deep-dive: Select Client", list(clients.keys()), key='mod_sel')
    c = clients[selected_mod]
    st.markdown(f"#### {c['name']} — Modules")
    if c['modules']:
        for m in c['modules']:
            status = m['status']
            icon = "✅" if ('✅' in status or status.lower() in ('done','✓')) else \
                   ("🔄" if 'progress' in status.lower() else "⬜")
            col1, col2 = st.columns([3,1])
            with col1:
                st.markdown(f"{icon} **{m['module']}**")
            with col2:
                color = "green" if icon=="✅" else ("orange" if icon=="🔄" else "gray")
                st.markdown(f"<span style='color:{color}'>{status}</span>", unsafe_allow_html=True)
    else:
        st.info("No modules recorded.")


# ══════════════════════════════
#  TAB 5 – SCHEDULE GENERATOR
# ══════════════════════════════
with tab5:
    st.markdown('<div class="section-header">📨 Monthly Schedule Generator</div>',
                unsafe_allow_html=True)
    st.markdown("Generate WhatsApp-ready schedule messages and plan next month's classes.")

    sub1, sub2 = st.tabs(["📋 Generate from Existing Data", "➕ Plan Next Month"])

    # ── Sub-tab 1: existing sessions → message
    with sub1:
        sel_client = st.selectbox("Select Client", list(clients.keys()), key='gen_sel')
        c = clients[sel_client]
        sessions = [s for s in c['sessions'] if not pd.isna(s['date'])]
        sessions.sort(key=lambda x: x['date'])

        if sessions:
            first_date = sessions[0]['date']
            month_year = first_date.strftime('%B %Y')
        else:
            month_year = datetime.today().strftime('%B %Y')

        month_override = st.text_input("Month label in message (e.g. March 2026)", value=month_year)

        msg = generate_schedule_message(c['name'], sessions, month_override)
        st.markdown("**Preview:**")
        st.markdown(f'<div class="schedule-msg">{msg}</div>', unsafe_allow_html=True)

        if st.button("📋 Copy-ready text", key='copy1'):
            st.code(msg, language=None)

    # ── Sub-tab 2: build next month manually
    with sub2:
        st.markdown("#### Plan Next Month's Classes")

        col_a, col_b = st.columns(2)
        with col_a:
            sel_client2 = st.selectbox("Client", list(clients.keys()), key='plan_sel')
        with col_b:
            next_month_label = st.text_input("Month (for message)", value="April 2026")

        c2 = clients[sel_client2]
        st.info(f"**{c2['name']}** — Target: {c2['target']} hrs/month")

        st.markdown("**Add class dates & times:**")

        if 'planned_sessions' not in st.session_state:
            st.session_state.planned_sessions = []

        with st.form("add_session_form", clear_on_submit=True):
            fc1, fc2, fc3 = st.columns(3)
            with fc1:
                new_date = st.date_input("Date", value=date.today())
            with fc2:
                new_time_start = st.time_input("Start time", value=datetime.strptime("08:00", "%H:%M").time())
            with fc3:
                new_time_end   = st.time_input("End time", value=datetime.strptime("09:00", "%H:%M").time())
            submitted = st.form_submit_button("➕ Add Session")
            if submitted:
                day_name = DAY_MAP[new_date.weekday()]
                time_str = f"{new_time_start.strftime('%I:%M').lstrip('0')} – {new_time_end.strftime('%I:%M%p').lower()}"
                st.session_state.planned_sessions.append({
                    'date': pd.Timestamp(new_date),
                    'day': day_name,
                    'time': f"{new_time_start.strftime('%H:%M')} – {new_time_end.strftime('%H:%M')}",
                    'time_display': f"{new_time_start.strftime('%-I:%M')} – {new_time_end.strftime('%-I:%M%p')}"
                })
                st.success(f"Added: {new_date.strftime('%d/%m/%Y')} ({day_name})")

        if st.session_state.planned_sessions:
            st.markdown("**Planned Sessions:**")
            planned_df = pd.DataFrame(st.session_state.planned_sessions)
            planned_df['date_fmt'] = planned_df['date'].dt.strftime('%d/%m/%Y')
            st.dataframe(planned_df[['date_fmt','day','time']], hide_index=True, use_container_width=True)

            # Generate message from planned
            plan_msg_sessions = []
            for s in st.session_state.planned_sessions:
                plan_msg_sessions.append({
                    'date': s['date'],
                    'day': s['day'],
                    'time': s['time'],
                    'hours': None,
                    'attendance': ''
                })

            plan_msg = generate_schedule_message(c2['name'], plan_msg_sessions, next_month_label)
            st.markdown("**📨 WhatsApp Message Preview:**")
            st.markdown(f'<div class="schedule-msg">{plan_msg}</div>', unsafe_allow_html=True)
            st.code(plan_msg, language=None)

            if st.button("🗑️ Clear Planned Sessions"):
                st.session_state.planned_sessions = []
                st.rerun()

            # Bulk generate all clients
            st.markdown("---")
            st.markdown("#### 📦 Bulk: Generate for ALL Clients (copy same dates)")
            if st.button("Generate All Clients from Same Dates"):
                dates_used = [s['date'] for s in st.session_state.planned_sessions]
                days_used  = [s['day'] for s in st.session_state.planned_sessions]
                times_used = [s['time'] for s in st.session_state.planned_sessions]

                all_msgs = []
                for sheet, cc in clients.items():
                    # Use existing time from their current sessions if available, else use planned time
                    msgs_sessions = []
                    for i, d in enumerate(dates_used):
                        # Try to find matching weekday in existing sessions
                        weekday = d.weekday()
                        existing_time = next(
                            (s['time'] for s in cc['sessions']
                             if pd.Timestamp(s['date']).weekday() == weekday),
                            times_used[i] if i < len(times_used) else "TBD"
                        )
                        msgs_sessions.append({
                            'date': d,
                            'day': days_used[i],
                            'time': existing_time,
                            'hours': None,
                            'attendance': ''
                        })
                    msg_bulk = generate_schedule_message(cc['name'], msgs_sessions, next_month_label)
                    all_msgs.append(f"{'='*50}\n{msg_bulk}\n")

                full_bulk = "\n".join(all_msgs)
                st.text_area("All Client Messages (copy from here)", full_bulk, height=400)
        else:
            st.caption("No sessions planned yet. Add sessions using the form above.")

    # ── Export section
    st.markdown("---")
    st.markdown('<div class="section-header">💾 Export</div>', unsafe_allow_html=True)

    if st.button("📥 Export Summary to Excel"):
        rows = []
        for sheet, c in clients.items():
            done, target, balance = hours_balance(c)
            attended = sum(1 for s in c['sessions'] if str(s['attendance']).strip() == 'Present')
            absent   = sum(1 for s in c['sessions'] if str(s['attendance']).strip() == 'Absent')
            fee_status = next((f['status'] for f in c['fees']
                               if f['month'].month == datetime.today().month), 'Pending')
            rows.append({
                'Client': c['name'],
                'Target (hrs)': target,
                'Hours Done': done,
                'Balance': balance,
                'Sessions Attended': attended,
                'Sessions Absent': absent,
                'Fee Status': fee_status,
                'Needs Replacement': 'Yes' if balance > 0 else 'No'
            })

        out_df = pd.DataFrame(rows)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            out_df.to_excel(writer, index=False, sheet_name='Summary')
        buf.seek(0)
        st.download_button(
            label="⬇️ Download Summary Excel",
            data=buf,
            file_name=f"therapy_summary_{datetime.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
