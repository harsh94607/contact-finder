"""
Apollo Contact Finder - WEB version (Streamlit)   |   Developed by Harsh Goyal
------------------------------------------------------------------------------
Same flow as the desktop app, in the browser:
  1. Type your Apollo API key (kept only in your browser session - never saved)
  2. Upload an Excel file with two columns:  Company Name | Domain
  3. Pick a company -> see the people Apollo finds (FREE) -> tick the ones you want
     -> "Load more" for broader results -> "Reveal selected & save" (credits used only here)
  4. Download the Excel file with every company block you saved

Deploy: see README.md (GitHub + Streamlit Community Cloud).
"""

import io
import re
import time

import pandas as pd
import requests
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# ----------------------------- CONFIG ---------------------------------
BASE_URL = "https://api.apollo.io/api/v1"
OUTPUT_NAME = "apollo_contacts.xlsx"

MAX_CONTACTS_PER_COMPANY = 20    # most people you can tick per company
PAGE_SIZE = 10                   # people added each time you click "Load more"

TARGET_TITLES = [
    "head of merchandising", "ceo", "president", "owner",
    "chief operating officer", "coo",
    "buyer", "head of buying", "head of sourcing", "sourcing manager",
    "procurement manager", "merchandising director",
    "head of product", "product development", "production manager",
    "vp operations", "head of production",
]
TARGET_SENIORITIES = ["owner", "founder", "c_suite", "vp", "head", "director", "manager"]

BROAD_TITLES = [
    "founder", "co-founder", "creative director", "design director", "head of design",
    "designer", "product manager", "merchandiser", "merchandising", "buying",
    "sourcing", "supply chain", "operations", "production", "manufacturing",
    "procurement", "purchasing", "category manager", "general manager",
    "managing director", "brand manager", "vice president", "director",
]

SEARCH_LEVELS = [
    ("Your target roles", {"person_titles": TARGET_TITLES, "include_similar_titles": True,
                           "person_seniorities": TARGET_SENIORITIES}),
    ("Target roles, any seniority", {"person_titles": TARGET_TITLES,
                                     "include_similar_titles": True}),
    ("Broader roles", {"person_titles": BROAD_TITLES, "include_similar_titles": True}),
    ("Everyone else at the company", {}),
]
# -----------------------------------------------------------------------


# =========================== LOGIC (no UI) ============================
class AuthError(Exception):
    pass


def clean_domain(raw):
    d = str(raw or "").strip().lower()
    d = re.sub(r"^https?://", "", d)
    d = re.sub(r"^www\.", "", d)
    return d.split("/")[0].strip()


def read_input(file_bytes):
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [str(h or "").strip().lower() for h in rows[0]]
    name_col = next((i for i, h in enumerate(header) if "company" in h or "name" in h), 0)
    dom_col = next((i for i, h in enumerate(header)
                    if "domain" in h or "website" in h or "url" in h), 1)
    companies = []
    for row in rows[1:]:
        if not row or len(row) <= max(name_col, dom_col):
            continue
        name = str(row[name_col] or "").strip()
        domain = clean_domain(row[dom_col])
        if name and domain:
            companies.append((name, domain))
    return companies


def apollo_request(key, method, endpoint, payload=None, params=None, retries=3):
    url = f"{BASE_URL}/{endpoint}"
    headers = {"Content-Type": "application/json", "Cache-Control": "no-cache",
               "x-api-key": key}
    for attempt in range(1, retries + 1):
        try:
            r = requests.request(method, url, headers=headers, json=payload,
                                 params=params, timeout=60)
        except requests.RequestException as e:
            print(f"Network error: {e}")
            time.sleep(2 * attempt)
            continue
        if r.status_code == 200:
            return r.json()
        if r.status_code == 429:
            time.sleep(10 * attempt)
            continue
        if r.status_code == 401:
            raise AuthError(r.text[:300])
        print(f"{r.status_code} on '{endpoint}': {r.text[:300]}")
        return None
    return None


def get_company(key, domain):
    data = apollo_request(key, "GET", "organizations/enrich", params={"domain": domain})
    org = (data or {}).get("organization") or {}
    org.setdefault("primary_domain", domain)
    org.setdefault("website_url", f"http://www.{domain}")
    return org


def new_state(levels):
    return {"levels": levels, "level": 0, "page": 1}


def exhausted(state):
    return state["level"] >= len(state["levels"])


def fetch_batch(key, org, domain, state, shown, session):
    """Next batch of people not already shown (FREE). Moves to broader levels
    automatically. Records Apollo's total for the current search in session."""
    known = {p["id"] for p in shown if p.get("id")}
    batch = []
    while not exhausted(state) and not batch:
        label, filters = state["levels"][state["level"]]
        payload = dict(filters)
        payload.update({"page": state["page"], "per_page": PAGE_SIZE})
        if org.get("id"):
            payload["organization_ids"] = [org["id"]]
        else:
            payload["q_organization_domains_list"] = [domain]

        data = apollo_request(key, "POST", "mixed_people/api_search", payload)
        people = (data or {}).get("people") or []
        total = (data or {}).get("total_entries") or \
            ((data or {}).get("pagination") or {}).get("total_entries")
        if total is not None:
            session["last_total"] = total
        session["last_label"] = label

        for p in people:
            pid = p.get("id")
            if pid and pid not in known:
                known.add(pid)
                p["_via"] = label
                batch.append(p)

        if data is None or len(people) < PAGE_SIZE:
            state["level"] += 1
            state["page"] = 1
        else:
            state["page"] += 1
    return batch


def person_name(p):
    first = p.get("first_name") or ""
    last = p.get("last_name") or p.get("last_name_obfuscated") or ""
    return f"{first} {last}".strip() or "(no name)"


def reveal_people(key, people):
    revealed = {}
    for i in range(0, len(people), 10):
        chunk = people[i:i + 10]
        payload = {"details": [{"id": p["id"]} for p in chunk if p.get("id")]}
        data = apollo_request(key, "POST", "people/bulk_match", payload,
                              params={"reveal_personal_emails": "false"})
        for m in (data or {}).get("matches") or []:
            if m and m.get("id"):
                revealed[m["id"]] = m
    return revealed


def build_contacts(chosen, revealed):
    rows = []
    for p in chosen:
        r = revealed.get(p.get("id")) or p
        rows.append((person_name(r), r.get("title") or p.get("title"),
                     r.get("email"), r.get("linkedin_url")))
    return rows


# ------------------------------ EXCEL ----------------------------------
GREEN = PatternFill("solid", fgColor="C6E0B4")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
LINK_FONT = Font(color="0563C1", underline="single")
WRAP_TOP = Alignment(wrap_text=True, vertical="top")


def style_header(ws, row, labels):
    for col, text in enumerate(labels, 1):
        c = ws.cell(row=row, column=col, value=text)
        c.font = Font(bold=True)
        c.fill = GREEN
        c.border = BORDER
        c.alignment = WRAP_TOP


def put(ws, row, col, value, link=False, bold=False):
    c = ws.cell(row=row, column=col, value=value or None)
    c.border = BORDER
    c.alignment = WRAP_TOP
    if bold:
        c.font = Font(bold=True)
    if link and value:
        c.hyperlink = value
        c.font = LINK_FONT
    return c


def open_workbook(prev_bytes=None):
    """New workbook, or the previous output file the user uploaded (so work can continue)."""
    if prev_bytes:
        wb = load_workbook(io.BytesIO(prev_bytes))
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Contacts"
        for col, w in zip("ABCD", (26, 60, 40, 55)):
            ws.column_dimensions[col].width = w
    return wb, ws


def done_companies(ws):
    done = set()
    for r in range(1, ws.max_row):
        if ws.cell(r, 1).value == "Company Name":
            nxt = ws.cell(r + 1, 1).value
            if nxt:
                done.add(str(nxt).strip().lower())
    return done


def write_block(ws, company, overview, website, linkedin, contacts):
    empty_sheet = ws.max_row == 1 and ws["A1"].value is None
    r = 1 if empty_sheet else ws.max_row + 2

    style_header(ws, r, ["Company Name", "Company Overview", "Website", "Company LinkedIn"])
    r += 1
    put(ws, r, 1, company, bold=True)
    put(ws, r, 2, overview)
    put(ws, r, 3, website, link=True)
    put(ws, r, 4, linkedin, link=True)
    ws.row_dimensions[r].height = max(30, 15 * (len(overview or "") // 55 + 1))
    r += 1

    style_header(ws, r, ["Contact Person", "Role", "Email Address", "LinkedIn URL"])
    r += 1
    for name, role, email, li in contacts:
        put(ws, r, 1, name)
        put(ws, r, 2, role)
        put(ws, r, 3, email)
        put(ws, r, 4, li, link=True)
        r += 1


def workbook_bytes(wb):
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def sample_input_bytes():
    wb = Workbook()
    ws = wb.active
    ws.append(["Company Name", "Domain"])
    ws.append(["Example Jewellers", "example.com"])
    return workbook_bytes(wb)



# =============================== WEB UI ================================
def stretch(fn, *args, **kwargs):
    """Full-width widget that works on both older and newer Streamlit versions."""
    try:
        return fn(*args, width="stretch", **kwargs)
    except Exception:           # noqa  (older Streamlit: 'width' must be a number)
        return fn(*args, use_container_width=True, **kwargs)


def flash(msg, kind="success"):
    st.session_state["flash"] = (kind, msg)


def secret(name):
    try:
        return st.secrets.get(name)
    except Exception:           # noqa  (no secrets configured)
        return None


def init_state():
    defaults = {"companies": [], "status": [], "sessions": {}, "cur": None, "nonce": 0,
                "wb": None, "ws": None, "saved_contacts": 0, "saved_companies": 0,
                "confirm_reveal": None, "authed": False}
    for k, v in defaults.items():
        st.session_state.setdefault(k, v)


def next_pending(after):
    n = len(st.session_state.status)
    for step in range(1, n + 1):
        i = (after + step) % n
        if st.session_state.status[i] == "pending":
            return i
    return None


def go_to(idx):
    st.session_state.cur = idx
    st.session_state.nonce += 1          # forces the company dropdown to show the new company


def advance():
    nxt = next_pending(st.session_state.cur)
    if nxt is not None:
        go_to(nxt)
    st.session_state.confirm_reveal = None


def ensure_session(key, idx):
    sessions = st.session_state.sessions
    if idx in sessions:
        return sessions[idx]
    name, domain = st.session_state.companies[idx]
    s = {"org": None, "shown": [], "selected": [], "default": new_state(SEARCH_LEVELS),
         "custom": None, "last_total": None, "last_label": "", "ver": 0}
    with st.spinner(f"Searching Apollo for people at {name} (free) ..."):
        s["org"] = get_company(key, domain)
        s["shown"].extend(fetch_batch(key, s["org"], domain, s["default"], s["shown"], s))
    sessions[idx] = s
    return s


def load_companies(file_bytes, prev_bytes):
    companies = read_input(file_bytes)
    wb, ws = open_workbook(prev_bytes)
    if not companies:
        raise ValueError("No companies with a domain were found. Columns needed: Company Name | Domain")
    done = done_companies(ws) if prev_bytes else set()
    ss = st.session_state
    ss.companies = companies
    ss.status = ["done" if n.lower() in done else "pending" for n, _ in companies]
    ss.sessions = {}
    ss.wb, ss.ws = wb, ws
    ss.saved_contacts = ss.saved_companies = 0
    ss.confirm_reveal = None
    first = next((i for i, s in enumerate(ss.status) if s == "pending"), 0)
    go_to(first)


def main():
    st.set_page_config(page_title="Apollo Contact Finder", page_icon="\u25C6", layout="wide")
    st.markdown("""
        <style>
        .block-container {padding-top: 3.2rem; max-width: 1250px;}
        div[data-testid="stMetric"] {border: 1px solid rgba(128,128,128,.28);
              border-radius: 10px; padding: 10px 16px;}
        .by {font-size: .8rem; opacity: .8; text-align: right; padding-top: 18px;}
        .by b {background: #2bd4ff; color: #02141a; padding: 3px 12px; border-radius: 14px;
              font-size: .9rem; margin-left: 4px;}
        </style>""", unsafe_allow_html=True)
    init_state()
    ss = st.session_state

    c1, c2 = st.columns([4, 2])
    c1.title("Apollo Contact Finder")
    c1.caption("Find the right people at each company (free) \u2192 pick them \u2192 reveal emails (uses credits) \u2192 download Excel")
    c2.markdown("<div class='by'>Developed by <b>Harsh Goyal</b></div>", unsafe_allow_html=True)

    # ---------------- optional password for the whole app (set APP_PASSWORD in secrets) ----------------
    app_pw = secret("APP_PASSWORD")
    if app_pw and not ss.authed:
        pw = st.text_input("App password", type="password")
        if pw:
            if pw == app_pw:
                ss.authed = True
                st.rerun()
            else:
                st.error("Wrong password.")
        st.stop()

    # ---------------- sidebar ----------------
    sb = st.sidebar
    sb.header("1. API key")
    server_key = secret("APOLLO_API_KEY")
    if server_key:
        key = str(server_key).strip().strip("'\"").strip()
        sb.success("Apollo key provided by the server.")
    else:
        key = (sb.text_input("Apollo API key", type="password",
                             help="Used only for this browser session. It is never saved or shown.") or "")
        key = key.strip().strip("'\"").strip()

    sb.header("2. Companies file")
    up = sb.file_uploader("Excel with Company Name | Domain", type=["xlsx", "xlsm"])
    prev = sb.file_uploader("Previous output file (optional)", type=["xlsx"],
                            help="Upload apollo_contacts.xlsx from an earlier session to continue it: "
                                 "companies already in it show as done and new blocks are added to it.")
    sb.download_button("Download a sample input file", sample_input_bytes(), "companies_sample.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    ask_confirm = sb.toggle("Ask before using credits", value=True)

    if sb.button("Load companies", type="primary", disabled=not (key and up)):
        try:
            load_companies(up.getvalue(), prev.getvalue() if prev else None)
            st.rerun()
        except Exception as e:                          # noqa
            sb.error(f"Could not read the file: {e}")

    if ss.wb is not None and ss.saved_contacts:
        sb.header("3. Download")
        sb.download_button(f"\u2B07 Download Excel ({ss.saved_contacts} contacts)",
                           workbook_bytes(ss.wb), OUTPUT_NAME,
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary")
        sb.caption("Results live only in this browser session. Download before closing the tab.")

    # ---------------- gate: key + file ----------------
    if not key:
        st.info("Start in the sidebar: **type your Apollo API key**, then upload your companies file.")
        st.stop()
    if not ss.companies:
        st.info("Upload your Excel file in the sidebar and click **Load companies**. "
                "It needs two columns: **Company Name** and **Domain**.")
        st.stop()

    try:
        render_workspace(key, ask_confirm)
    except AuthError:
        st.error("Apollo says the API key is invalid. Check it in the sidebar and try again.")


def render_workspace(key, ask_confirm):
    ss = st.session_state
    n_all = len(ss.companies)
    cur = ss.cur
    mark = {"done": "OK", "skipped": "--", "pending": "  "}

    # ---------------- company picker + progress ----------------
    done = ss.status.count("done")
    m = st.columns(4)
    m[0].metric("Companies done", f"{done}/{n_all}")
    m[1].metric("Contacts saved", ss.saved_contacts)
    pick_key = f"pick_{ss.nonce}"

    def on_pick():
        ss.cur = ss[pick_key]
        ss.confirm_reveal = None

    st.selectbox("Company", range(n_all), index=cur, key=pick_key, on_change=on_pick,
                 format_func=lambda i: f"{i + 1:03d}  [{mark[ss.status[i]]}]  {ss.companies[i][0]}")
    cur = ss.cur
    name, domain = ss.companies[cur]
    note = ss.pop("flash", None)
    if note:
        (st.success if note[0] == "success" else st.info)(note[1])
    s = ensure_session(key, cur)
    org = s["org"] or {}

    m[2].metric("People listed", len(s["shown"]))
    m[3].metric("Selected", f"{len(s['selected'])}/{MAX_CONTACTS_PER_COMPANY}")

    st.subheader(f"{name}  \u00B7  {domain}")
    overview = org.get("short_description") or org.get("seo_description") or ""
    if overview:
        st.caption(overview[:300] + ("..." if len(overview) > 300 else ""))
    total = s["last_total"]
    st.caption(f"Search level: {s['last_label'] or '-'}"
               + (f"   \u00B7   Apollo total for last search: {total:,}" if total else ""))
    if ss.status[cur] == "done":
        st.success("This company is already saved.")

    # ---------------- search other roles ----------------
    with st.form("roles"):
        r1, r2 = st.columns([4, 1])
        roles_txt = r1.text_input("Search other roles (comma separated)", placeholder="designer, buyer, production",
                                  label_visibility="collapsed")
        go = r2.form_submit_button("Search roles")
    if go:
        roles = [r.strip() for r in roles_txt.split(",") if r.strip()]
        if not roles:
            st.warning("Type one or more roles, e.g. designer, buyer")
        else:
            s["custom"] = new_state([(f"Roles: {', '.join(roles)}",
                                      {"person_titles": roles, "include_similar_titles": True})])
            with st.spinner("Searching those roles (free) ..."):
                batch = fetch_batch(key, s["org"], domain, s["custom"], s["shown"], s)
            s["shown"].extend(batch)
            s["ver"] += 1
            flash(f"Added {len(batch)} people for those roles." if batch
                  else "Nothing new for those roles (they may already be listed).", "info")
            st.rerun()

    # ---------------- people table (tick to select) ----------------
    if not s["shown"]:
        st.warning("No people found for this company. Try 'Search roles' or skip it.")
    else:
        rows = []
        for i, p in enumerate(s["shown"], 1):
            has = p.get("has_email")
            rows.append({"Pick": p["id"] in s["selected"], "#": i, "Name": person_name(p),
                         "Role": p.get("title") or "-",
                         "Email?": "yes" if has is True else ("no" if has is False else "?"),
                         "Found via": p.get("_via", "")})
        edited = stretch(
            st.data_editor, pd.DataFrame(rows), hide_index=True,
            disabled=["#", "Name", "Role", "Email?", "Found via"],
            column_config={"Pick": st.column_config.CheckboxColumn("Pick", width="small")},
            key=f"editor_{cur}_{s['ver']}")
        s["selected"] = [p["id"] for p, pick in zip(s["shown"], edited["Pick"]) if pick]
        st.caption("Tick the people you want. Last names stay masked until you reveal.")

    n_sel = len(s["selected"])
    too_many = n_sel > MAX_CONTACTS_PER_COMPANY
    if too_many:
        st.error(f"You can select up to {MAX_CONTACTS_PER_COMPANY} people per company "
                 f"(you ticked {n_sel}). Untick a few.")

    # ---------------- actions ----------------
    b = st.columns([1, 1, 1, 1, 2])
    if stretch(b[0].button, "\u25BC Load more"):
        state = s["custom"] if (s["custom"] and not exhausted(s["custom"])) else s["default"]
        with st.spinner("Loading more people (free) ..."):
            batch = fetch_batch(key, s["org"], domain, state, s["shown"], s)
            if not batch and state is s["custom"]:
                batch = fetch_batch(key, s["org"], domain, s["default"], s["shown"], s)
        s["shown"].extend(batch)
        s["ver"] += 1
        flash(f"Added {len(batch)} more people." if batch else "No more people found at this company.", "info")
        st.rerun()
    if stretch(b[1].button, "Select all"):
        s["selected"] = [p["id"] for p in s["shown"]][:MAX_CONTACTS_PER_COMPANY]
        s["ver"] += 1
        st.rerun()
    if stretch(b[2].button, "Clear"):
        s["selected"] = []
        s["ver"] += 1
        st.rerun()
    if stretch(b[3].button, "Skip company"):
        ss.status[cur] = "skipped"
        advance()
        st.rerun()
    if stretch(b[4].button, "Reveal selected & save  \u25B6", type="primary",
               disabled=(n_sel == 0 or too_many)):
        if ask_confirm:
            ss.confirm_reveal = cur
        else:
            do_reveal(key, cur, s)
        st.rerun()

    if ss.confirm_reveal == cur and n_sel:
        st.warning(f"Reveal emails for {n_sel} {'person' if n_sel == 1 else 'people'}? "
                   f"This uses up to {n_sel} Apollo credit{'s' if n_sel != 1 else ''}.")
        y, c = st.columns(2)
        if stretch(y.button, "Yes, reveal and save", type="primary"):
            do_reveal(key, cur, s)
            st.rerun()
        if stretch(c.button, "Cancel"):
            ss.confirm_reveal = None
            st.rerun()

    # ---------------- all companies ----------------
    with st.expander("All companies"):
        stretch(st.dataframe, pd.DataFrame({"#": range(1, n_all + 1),
                                   "Company": [c[0] for c in ss.companies],
                                   "Domain": [c[1] for c in ss.companies],
                                   "Status": ss.status}), hide_index=True)

    if ss.status.count("pending") == 0:
        st.success("All companies are finished. Download your Excel file from the sidebar.")


def do_reveal(key, idx, s):
    ss = st.session_state
    name, domain = ss.companies[idx]
    org = s["org"] or {}
    overview = org.get("short_description") or org.get("seo_description") or ""
    chosen = [p for p in s["shown"] if p["id"] in s["selected"]]
    with st.spinner(f"Revealing {len(chosen)} emails and saving ..."):
        revealed = reveal_people(key, chosen)
        contacts = build_contacts(chosen, revealed)
        write_block(ss.ws, name, overview, org.get("website_url"), org.get("linkedin_url"), contacts)
    ss.status[idx] = "done"
    ss.saved_contacts += len(contacts)
    ss.saved_companies += 1
    flash(f"Saved {len(contacts)} contacts for {name}. Download the Excel from the sidebar any time.")
    advance()


if __name__ == "__main__":
    main()
