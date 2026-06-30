"""
app.py
Streamlit wizard for the מיכפל salary import template automation.

Run with:
    streamlit run app.py

Flow: pick company -> extract & preview -> generate (no write) -> insert
(explicit, separate step, with automatic backup).
"""

import datetime
import os

import pandas as pd
import streamlit as st

import config
import engine

SRGL_FILENAME = "Q8SRGL26.000"


st.set_page_config(page_title="מיכפל Template Automation", layout="wide")
st.title("מיכפל — Salary Import Template Automation")

# ---------------------------------------------------------------------------
# Sidebar settings -- persisted to the config file, survive between runs
# ---------------------------------------------------------------------------
cfg = config.load()

with st.sidebar:
    st.header("Settings")
    data_dir = st.text_input("Data folder", value=cfg["data_dir"])
    out_dir = st.text_input("Output folder (generated files)", value=cfg["out_dir"])

# Persist any change as soon as it happens, so the new value is the default
# next run (and for the CLI, which reads the same config).
if data_dir != cfg["data_dir"] or out_dir != cfg["out_dir"]:
    config.save({**cfg, "data_dir": data_dir, "out_dir": out_dir})
    st.sidebar.caption("Settings saved.")

srgl_path = os.path.join(data_dir, SRGL_FILENAME)

# ---------------------------------------------------------------------------
# Session state init
# ---------------------------------------------------------------------------
_defaults = {
    "last_company": None,
    "extracted": False,
    "components": None,
    "skipped_components": 0,
    "employees": None,
    "included_inactive": [],
    "invalid_employees": 0,
    "generated": False,
    "template_text": None,
    "col_map": None,
    "stats_cols": None,
    "excel_path": None,
    "carry_new_hires": [],
    "carry_count": 0,
    "inserted": False,
    "backup_path": None,
    # batch mode: persist generated templates across reruns so the insert
    # step (a separate button click) can act on them.
    "batch_results": None,
    "batch_inserted": False,
    "batch_insert_summary": None,
}
for k, v in _defaults.items():
    st.session_state.setdefault(k, v)


def _reset_downstream():
    st.session_state.extracted = False
    st.session_state.generated = False
    st.session_state.inserted = False
    st.session_state.included_inactive = []
    st.session_state.carry_new_hires = []
    st.session_state.carry_count = 0


@st.cache_data(show_spinner=False)
def load_employees_with_status(path, mtime):
    """
    Cached wrapper around engine.extract_employees_with_status. `mtime`
    (file modification time) is part of the cache key so the cache busts
    automatically when the underlying file changes; it isn't used in the
    body. Avoids re-reading the (potentially large) Q8OVDM26 file on every
    Streamlit rerun.
    """
    return engine.extract_employees_with_status(path)


# ---------------------------------------------------------------------------
# Step 1 -- select company
# ---------------------------------------------------------------------------
st.header("Step 1 — Select company")

if not os.path.isdir(data_dir):
    st.error(f"Data folder not found: {data_dir}")
    st.stop()

companies = engine.discover_companies(data_dir)
if not companies:
    st.warning(
        "No companies found (need matching Q8MIFL26.* and Q8OVDM26.* pairs in the data folder)."
    )
    st.stop()

# ---------------------------------------------------------------------------
# Mode: one company (full wizard incl. insert) vs. batch generate (many at
# once, generate-only). Insert into Q8SRGL26.000 stays single-company on
# purpose -- it's the one destructive, per-template, confirmed step, and is
# not safe to loop unattended.
# ---------------------------------------------------------------------------
SINGLE_MODE = "חברה אחת (כולל הוספה לקובץ התבניות)"
BATCH_MODE = "מספר חברות יחד (יצירת קבצים בלבד)"
run_mode = st.radio(
    "מצב הרצה",
    [SINGLE_MODE, BATCH_MODE],
    index=0,
    horizontal=True,
    help=(
        "חברה אחת: התהליך המלא כולל סקירה והוספה ל-Q8SRGL26.000. "
        "מספר חברות: יצירת ה-Excel וה-txt לכמה חברות בבת אחת (ללא הוספה לקובץ "
        "התבניות -- שלב ההוספה נשאר פר-חברה מטעמי בטיחות)."
    ),
)

# ===========================================================================
# BATCH MODE -- generate files for several companies at once, then OPTIONALLY
# insert their templates into Q8SRGL26.000 in one confirmed step (mirrors the
# single-company Step 3, looped with per-company backup + isolation).
# Rendered instead of the single-company wizard; st.stop() at the end keeps
# the wizard below from also running.
# ===========================================================================
if run_mode == BATCH_MODE:
    st.header("Batch — generate & insert for multiple companies")

    selected = st.multiselect(
        "בחר חברות",
        companies,
        default=companies,
        help="כל החברות נבחרות כברירת מחדל; אפשר לצמצם.",
    )

    today = datetime.date.today()
    bc1, bc2 = st.columns(2)
    with bc1:
        b_year = st.number_input(
            "שנת מס (year)",
            min_value=2000,
            max_value=2100,
            value=today.year,
            step=1,
            key="batch_year",
        )
    with bc2:
        b_month = st.number_input(
            "חודש דיווח (month)",
            min_value=1,
            max_value=12,
            value=today.month,
            step=1,
            key="batch_month",
        )

    b_carry = st.checkbox(
        "מלא אוטומטית ערכי כמות/מחיר מהחודש הקודם (carry-forward)",
        value=False,
        key="batch_carry",
        help=(
            "בהרצת אצווה לא מוצג בורר רכיבים פר-חברה; רכיבי חופשה/מחלה/חג/לידה "
            "מזוהים אוטומטית לפי שם ואינם מועברים. סמן 'העבר הכל' כדי להעביר גם אותם."
        ),
    )
    b_carry_all = False
    if b_carry:
        b_carry_all = st.checkbox(
            "העבר הכל (כולל חופשה/מחלה)", value=False, key="batch_carry_all"
        )

    if not selected:
        st.info("בחר לפחות חברה אחת.")
    elif st.button(f"Generate {len(selected)} companies", type="primary"):
        os.makedirs(out_dir, exist_ok=True)
        rows = []
        progress = st.progress(0.0)
        for idx, comp in enumerate(selected, 1):
            mifl_path = os.path.join(data_dir, f"Q8MIFL26.{comp}")
            ovdm_path = os.path.join(data_dir, f"Q8OVDM26.{comp}")
            tname = f"תבנית משכורות אוטומציה {comp}"
            try:
                comps, _skip = engine.extract_components(mifl_path)
                deds = engine.extract_deductions(mifl_path)
                emps, _inv = engine.extract_employees(ovdm_path)
                ttext, cmap, scols, dmap = engine.build_template(tname, comps, deds)

                prior = None
                prior_ded = None
                no_carry = set()
                if b_carry:
                    prior = engine.read_prior_month(ovdm_path)
                    prior_ded = engine.read_prior_month_deductions(ovdm_path)
                    no_carry = (
                        set()
                        if b_carry_all
                        else engine.default_no_carry_components(comps)
                    )

                xpath = os.path.join(out_dir, f"template_{comp}.xlsx")
                rep = engine.build_excel(
                    comp,
                    comps,
                    cmap,
                    scols,
                    emps,
                    xpath,
                    int(b_year),
                    int(b_month),
                    prior_month=prior,
                    no_carry=no_carry,
                    ded_map=dmap,
                    prior_deductions=prior_ded,
                )
                rows.append(
                    {
                        "company": comp,
                        "template_name": tname,
                        # keep the template block IN MEMORY for the insert
                        # step -- no loose .txt is written to the output folder
                        "template_text": ttext,
                        "ok": True,
                        "employees": len(emps),
                        "components": len(comps),
                        "carried": rep["carried"],
                        "new_hires": len(rep["new_hires"]),
                        "xlsx_path": xpath,
                        "error": None,
                    }
                )
            except Exception as e:
                # Isolate per company -- one failure must not abort the batch.
                rows.append(
                    {
                        "company": comp,
                        "template_name": tname,
                        "template_text": None,
                        "ok": False,
                        "employees": 0,
                        "components": 0,
                        "carried": 0,
                        "new_hires": 0,
                        "xlsx_path": None,
                        "error": str(e),
                    }
                )
            progress.progress(idx / len(selected))

        st.session_state.batch_results = rows
        st.session_state.batch_inserted = False
        st.session_state.batch_insert_summary = None

    # ---- persistent results + insert step (survives reruns) --------------
    results = st.session_state.batch_results
    if results:
        ok = [r for r in results if r["ok"]]
        st.success(f"נוצרו {len(ok)} מתוך {len(results)} חברות (קובצי Excel).")
        st.dataframe(
            [
                {
                    "חברה": r["company"],
                    "סטטוס": "✓" if r["ok"] else f"✗ {r['error']}",
                    "עובדים": r["employees"],
                    "רכיבים": r["components"],
                    "הועברו": r["carried"],
                    "עובדים חדשים": r["new_hires"],
                }
                for r in results
            ],
            use_container_width=True,
        )

        # ZIP of the generated EXCEL files (the deliverable to fill & import).
        import io
        import zipfile

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for r in ok:
                if r["xlsx_path"]:
                    zf.write(r["xlsx_path"], arcname=os.path.basename(r["xlsx_path"]))
        buf.seek(0)
        st.download_button(
            f"⬇ Download all Excel files ({len(ok)}) as ZIP",
            buf,
            file_name=f"templates_batch_{int(b_year)}_{int(b_month):02d}.zip",
            mime="application/zip",
        )
        if b_carry:
            st.caption(
                "תזכורת: carry-forward קורא את הערכים מקובץ Q8OVDM26 הנוכחי, "
                "שמכיל את נתוני החודש הקודם עד לייבוא החדש. יש להפיק לפני הייבוא."
            )

        # ---- batch template insert into Q8SRGL26.000 --------------------
        st.divider()
        st.subheader("עדכון התבניות בקובץ Q8SRGL26.000")
        insertable = [r for r in ok if r["template_text"]]
        if not os.path.exists(srgl_path):
            st.warning(
                f"{SRGL_FILENAME} not found at {srgl_path} — cannot insert "
                f"templates until it exists."
            )
        elif not insertable:
            st.info("אין תבניות להוספה.")
        else:
            st.warning(
                f"This will modify `{srgl_path}` for **{len(insertable)} "
                f"companies**.\n\nEach company's template is inserted "
                f"separately, and a backup of {SRGL_FILENAME} is created "
                f"before each insert (existing same-named templates are "
                f"replaced)."
            )
            confirm_batch = st.checkbox(
                "אני מאשר עדכון התבניות לכל החברות שנוצרו.",
                key="batch_insert_confirm",
            )
            if st.button(
                f"Insert {len(insertable)} templates",
                disabled=not confirm_batch,
                type="primary",
            ):
                summary = []
                prog = st.progress(0.0)
                for i, r in enumerate(insertable, 1):
                    try:
                        res = engine.regenerate_and_insert_template(
                            srgl_path,
                            r["template_name"],
                            r["template_text"],
                            r["company"],
                        )
                        summary.append(
                            {
                                "company": r["company"],
                                "ok": True,
                                "backup": res.get("insert_backup"),
                                "error": None,
                            }
                        )
                    except PermissionError:
                        summary.append(
                            {
                                "company": r["company"],
                                "ok": False,
                                "backup": None,
                                "error": "Permission denied (write access to data folder needed)",
                            }
                        )
                    except Exception as e:
                        # Isolate: a bad company must not abort remaining inserts.
                        summary.append(
                            {
                                "company": r["company"],
                                "ok": False,
                                "backup": None,
                                "error": str(e),
                            }
                        )
                    prog.progress(i / len(insertable))
                st.session_state.batch_inserted = True
                st.session_state.batch_insert_summary = summary

        summary = st.session_state.batch_insert_summary
        if summary:
            ins_ok = [s for s in summary if s["ok"]]
            if len(ins_ok) == len(summary):
                st.success(f"הוכנסו {len(ins_ok)} תבניות אל {SRGL_FILENAME}.")
            else:
                st.error(
                    f"הוכנסו {len(ins_ok)} מתוך {len(summary)} — חלק נכשלו (ראה טבלה)."
                )
            st.dataframe(
                [
                    {
                        "חברה": s["company"],
                        "סטטוס": "✓" if s["ok"] else f"✗ {s['error']}",
                        "גיבוי": os.path.basename(s["backup"]) if s["backup"] else "",
                    }
                    for s in summary
                ],
                use_container_width=True,
            )
            st.caption("לשחזור ידני: העתק את קובץ הגיבוי על Q8SRGL26.000.")

    st.stop()  # batch mode handled; do not fall through to the single wizard

existing_templates = []
if os.path.exists(srgl_path):
    existing_templates = engine.list_existing_templates(srgl_path)
else:
    st.warning(
        f"{SRGL_FILENAME} not found at {srgl_path} -- insertion (Step 3) will not be possible until it exists."
    )

company = st.selectbox("Company number", companies, key="company_select")
template_name = f"תבנית משכורות אוטומציה {company}"

if template_name in existing_templates:
    st.info(
        f"ℹ️ A template named '{template_name}' already exists in {SRGL_FILENAME}. "
        f"Inserting again will replace it (the old block is removed first)."
    )

if company != st.session_state.last_company:
    st.session_state.last_company = company
    _reset_downstream()

today = datetime.date.today()
c1, c2 = st.columns(2)
with c1:
    year = st.number_input(
        "שנת מס (year)", min_value=2000, max_value=2100, value=today.year, step=1
    )
with c2:
    month = st.number_input(
        "חודש דיווח (month)", min_value=1, max_value=12, value=today.month, step=1
    )

st.divider()

# ---------------------------------------------------------------------------
# Step 2 -- extract, preview & generate (no write to Q8SRGL26.000 yet)
# ---------------------------------------------------------------------------
st.header("Step 2 — Extract & generate (no write yet)")

# --- inactive-employee handling -------------------------------------------
ACTIVE_ONLY = "עובדים פעילים בלבד"
ALL_INACTIVE = "כל העובדים (כולל לא פעילים)"
SELECT_INACTIVE = "בחירת עובדים לא פעילים לכלול"

inactive_mode = st.radio(
    "עובדים לא פעילים (קוד הפסקה 1)",
    [ACTIVE_ONLY, ALL_INACTIVE, SELECT_INACTIVE],
    index=0,
    help=(
        "כברירת מחדל נכללים רק עובדים פעילים. אפשר לכלול את כל הלא-פעילים, "
        "או לבחור ידנית אילו לכלול."
    ),
)

selected_inactive = []
if inactive_mode == SELECT_INACTIVE:
    ovdm_path = os.path.join(data_dir, f"Q8OVDM26.{company}")
    if not os.path.exists(ovdm_path):
        st.error(f"Q8OVDM26.{company} not found at {ovdm_path}")
    else:
        try:
            _, inactive_list, _ = load_employees_with_status(
                ovdm_path, os.path.getmtime(ovdm_path)
            )
        except Exception as e:
            st.error(f"Could not read employees: {e}")
            inactive_list = []
        if not inactive_list:
            st.caption("No inactive employees found for this company.")
        else:
            st.caption('סמן בעמודת "כלול" את העובדים הלא-פעילים להוספה לתבנית:')
            editor_rows = pd.DataFrame(
                {
                    "כלול": [False] * len(inactive_list),
                    "תעודת זהות": [tz for tz, _ in inactive_list],
                    "שם": [name for _, name in inactive_list],
                }
            )
            edited = st.data_editor(
                editor_rows,
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                disabled=["תעודת זהות", "שם"],
                column_config={
                    "כלול": st.column_config.CheckboxColumn(
                        "כלול", help="סמן כדי לכלול עובד זה בתבנית", default=False
                    ),
                    "תעודת זהות": st.column_config.TextColumn("תעודת זהות"),
                    "שם": st.column_config.TextColumn("שם"),
                },
                key=f"inactive_editor_{company}",
            )
            # Reconstruct (tz, name) straight from the checked rows so the
            # selection is robust to any visual re-sorting of the table.
            selected_inactive = [
                (row["תעודת זהות"], row["שם"])
                for _, row in edited.iterrows()
                if row["כלול"]
            ]
            st.caption(f"נבחרו {len(selected_inactive)} מתוך {len(inactive_list)}.")

# --- carry-forward (prior-month values) -----------------------------------
st.subheader("מילוי נתוני חודש קודם (carry-forward)")
carry_forward = st.checkbox(
    "מלא אוטומטית ערכי כמות/מחיר מהחודש הקודם",
    value=False,
    help=(
        "קורא את הערכים הקיימים בקובץ העובדים (Q8OVDM26) -- שמכיל את נתוני "
        "החודש הקודם עד לייבוא החדש -- וממלא אותם מראש בגיליון. עובדים חדשים "
        '(שלא היו בחודש הקודם) יקבלו שורה ריקה עם ת"ז ושם מסומנים בצהוב.'
    ),
)

no_carry_numbers = set()
if carry_forward:
    mifl_path_cf = os.path.join(data_dir, f"Q8MIFL26.{company}")
    try:
        _cf_components, _ = engine.extract_components(mifl_path_cf)
    except Exception as e:
        st.error(f"Could not read components for carry-forward options: {e}")
        _cf_components = []

    if _cf_components:
        default_nc = engine.default_no_carry_components(_cf_components)
        st.caption(
            "בחר רכיבים שלא יועברו מהחודש הקודם. כברירת מחדל מסומנים רכיבי "
            "חופשה/מחלה/חג/לידה -- אפשר לשנות:"
        )
        # The carry-forward join key is the slot real number, which equals
        # `actual` (= rechiv_extracted - 1), with the משכורת block = 1. Build
        # the hidden _real column with that, so the no_carry set matches the
        # keys read_prior_month() / build_excel() use.
        nc_rows = pd.DataFrame(
            {
                "אל תעביר": [
                    (1 if rc == 2 else rc - 1) in default_nc
                    for rc, _n, _k in _cf_components
                ],
                "קוד": [(1 if rc == 2 else rc - 1) for rc, _n, _k in _cf_components],
                "שם": [n for _rc, n, _k in _cf_components],
                "_real": [(1 if rc == 2 else rc - 1) for rc, _n, _k in _cf_components],
            }
        )
        nc_edited = st.data_editor(
            nc_rows,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            disabled=["קוד", "שם", "_real"],
            column_config={
                "אל תעביר": st.column_config.CheckboxColumn(
                    "אל תעביר", help="סמן רכיבים שלא יועברו מהחודש הקודם"
                ),
                "קוד": st.column_config.NumberColumn("קוד"),
                "שם": st.column_config.TextColumn("שם"),
                "_real": None,  # hidden helper column
            },
            key=f"nocarry_editor_{company}",
        )
        no_carry_numbers = {
            int(row["_real"]) for _, row in nc_edited.iterrows() if row["אל תעביר"]
        }
        st.caption(f"לא יועברו {len(no_carry_numbers)} רכיבים.")

if st.button("Extract & generate", type="primary"):
    mifl_path = os.path.join(data_dir, f"Q8MIFL26.{company}")
    ovdm_path = os.path.join(data_dir, f"Q8OVDM26.{company}")
    try:
        components, skipped = engine.extract_components(mifl_path)
        deductions = engine.extract_deductions(mifl_path)
        active_emps, inactive_emps, _inv = load_employees_with_status(
            ovdm_path, os.path.getmtime(ovdm_path)
        )
    except Exception as e:
        st.error(f"Extraction failed: {e}")
    else:
        # Decide which inactive employees to include, per the chosen mode.
        if inactive_mode == ALL_INACTIVE:
            included_inactive = list(inactive_emps)
        elif inactive_mode == SELECT_INACTIVE:
            chosen_set = set(selected_inactive)
            included_inactive = [e for e in inactive_emps if e in chosen_set]
        else:
            included_inactive = []

        employees = sorted(active_emps + included_inactive, key=lambda x: x[0])
        invalid = sum(1 for tz, _name in employees if not tz)

        st.session_state.components = components
        st.session_state.skipped_components = skipped
        st.session_state.employees = employees
        st.session_state.included_inactive = included_inactive
        st.session_state.invalid_employees = invalid
        st.session_state.extracted = True
        st.session_state.generated = False
        st.session_state.inserted = False

        os.makedirs(out_dir, exist_ok=True)
        template_text, col_map, stats_cols, ded_map = engine.build_template(
            template_name, components, deductions
        )

        # carry-forward: read previous month's values from the same Q8OVDM26
        # (it holds last month's data until the new sheet is imported).
        prior_month = None
        prior_deductions = None
        if carry_forward:
            try:
                prior_month = engine.read_prior_month(ovdm_path)
                prior_deductions = engine.read_prior_month_deductions(ovdm_path)
            except Exception as e:
                st.error(f"Could not read previous month's values: {e}")
                prior_month = None
                prior_deductions = None

        excel_path = os.path.join(out_dir, f"template_{company}.xlsx")
        try:
            cf_result = engine.build_excel(
                company,
                components,
                col_map,
                stats_cols,
                employees,
                excel_path,
                int(year),
                int(month),
                prior_month=prior_month,
                no_carry=no_carry_numbers,
                ded_map=ded_map,
                prior_deductions=prior_deductions,
            )
        except Exception as e:
            st.error(f"Excel generation failed: {e}")
        else:
            st.session_state.template_text = template_text
            st.session_state.col_map = col_map
            st.session_state.stats_cols = stats_cols
            st.session_state.excel_path = excel_path
            st.session_state.carry_new_hires = cf_result["new_hires"]
            st.session_state.carry_count = cf_result["carried"]
            st.session_state.generated = True

if st.session_state.extracted:
    components = st.session_state.components
    employees = st.session_state.employees

    included_inactive = st.session_state.get("included_inactive", [])
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Components found", len(components))
    m2.metric("Components filtered out", st.session_state.skipped_components)
    m3.metric("Employees (total)", len(employees))
    m4.metric("Inactive included", len(included_inactive))
    m5.metric("Invalid תז blocks", st.session_state.invalid_employees)

    if st.session_state.generated and st.session_state.carry_count:
        new_hires = st.session_state.get("carry_new_hires", [])
        cc1, cc2 = st.columns(2)
        cc1.metric("Cells carried forward", st.session_state.carry_count)
        cc2.metric("New hires (blank, yellow)", len(new_hires))
        if new_hires:
            with st.expander(
                f"New hires — blank row, highlighted yellow ({len(new_hires)})"
            ):
                st.caption(
                    "עובדים אלה לא היו בחודש הקודם, כך שאין מה למלא -- "
                    'ת"ז ושם מסומנים בצהוב בגיליון.'
                )
                st.dataframe(
                    [{"תעודת זהות": tz, "שם": name} for tz, name in new_hires],
                    use_container_width=True,
                )

    with st.expander(f"Component list ({len(components)})"):
        st.dataframe(
            [
                {
                    "actual קוד": (1 if rc == 2 else rc - 1),
                    "name": name,
                    'קוד מה"כ': km,
                    "raw extracted": rc,
                }
                for rc, name, km in components
            ],
            use_container_width=True,
        )

    with st.expander(f"Employee list ({len(employees)})"):
        inactive_set = set(st.session_state.get("included_inactive", []))
        st.dataframe(
            [
                {
                    "תעודת זהות": tz,
                    "name": name,
                    "סטטוס": "לא פעיל" if (tz, name) in inactive_set else "פעיל",
                }
                for tz, name in employees
            ],
            use_container_width=True,
        )

    if st.session_state.generated:
        with st.expander(f'@@QD@@ block — "{template_name}"', expanded=False):
            st.code(st.session_state.template_text, language=None)

        with st.expander("Column mapping", expanded=False):
            first_comp_name = (
                st.session_state.components[0][1]
                if st.session_state.components
                else "משכורת"
            )
            rows = [
                {
                    "actual code": 1,
                    "name": first_comp_name,
                    "col כמות": "E",
                    "col מחיר": "F",
                    "col ברוטו/נטו": "C",
                }
            ]
            for actual, name, ck, cm, cbn in st.session_state.col_map:
                rows.append(
                    {
                        "actual code": actual,
                        "name": name,
                        "col כמות": engine.col_letter(ck),
                        "col מחיר": engine.col_letter(cm),
                        "col ברוטו/נטו": engine.col_letter(cbn),
                    }
                )
            st.dataframe(rows, use_container_width=True)

        dl1, dl2 = st.columns(2)
        with dl1:
            with open(st.session_state.excel_path, "rb") as f:
                st.download_button(
                    "⬇ Download Excel",
                    f,
                    file_name=os.path.basename(st.session_state.excel_path),
                )
        with dl2:
            st.download_button(
                "⬇ Download template .txt",
                st.session_state.template_text,
                file_name=f"template_{company}.txt",
            )

        st.caption("Open the Excel and check it before inserting the template below.")
else:
    st.info("Run extract & generate to continue.")

st.divider()

# ---------------------------------------------------------------------------
# Step 3 -- insert into Q8SRGL26.000 (explicit, separate, backed up)
# ---------------------------------------------------------------------------
st.header("Step 3 — Insert into Q8SRGL26.000")

if not st.session_state.generated:
    st.info("Complete Step 2 first.")
else:
    st.warning(
        f"This will modify: `{srgl_path}`\n\n"
        f"A backup is created automatically first, named:\n"
        f"`{SRGL_FILENAME}.bak-{str(company).zfill(3)}-YYYYMMDD_HHMMSS`"
    )
    confirm = st.checkbox("I've reviewed the generated Excel and template block above.")
    if st.button("Insert template", disabled=not confirm, type="primary"):
        try:
            result = engine.regenerate_and_insert_template(
                srgl_path, template_name, st.session_state.template_text, company
            )
        except PermissionError:
            msg = (
                f"**Permission denied** — cannot write to `{srgl_path}`.\n\n"
                "You need write access to the data folder for insertion to work."
            )
            if os.name == "nt":
                msg += (
                    " If it's a shared drive, ask whoever manages it to grant "
                    "you write permission."
                )
            else:
                msg += (
                    "\n\nOn this dev machine the share usually needs elevation — "
                    "restart with:\n```\nsudo .venv/bin/streamlit run app.py\n```"
                )
            st.error(msg)
        except Exception as e:
            st.error(f"Insertion failed — file NOT modified: {e}")
        else:
            st.session_state.inserted = True
            st.session_state.backup_path = result["insert_backup"]

    if st.session_state.inserted:
        st.success(f"Inserted '{template_name}' into {srgl_path}")
        st.code(f"Backup saved at:\n{st.session_state.backup_path}")
        st.caption("To roll back manually: copy the backup file over Q8SRGL26.000.")
