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
    "invalid_employees": 0,
    "generated": False,
    "template_text": None,
    "col_map": None,
    "stats_cols": None,
    "excel_path": None,
    "inserted": False,
    "backup_path": None,
}
for k, v in _defaults.items():
    st.session_state.setdefault(k, v)


def _reset_downstream():
    st.session_state.extracted = False
    st.session_state.generated = False
    st.session_state.inserted = False


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

existing_templates = []
if os.path.exists(srgl_path):
    existing_templates = engine.list_existing_templates(srgl_path)
else:
    st.warning(
        f"{SRGL_FILENAME} not found at {srgl_path} -- insertion (Step 4) will not be possible until it exists."
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

if st.button("Extract & generate", type="primary"):
    mifl_path = os.path.join(data_dir, f"Q8MIFL26.{company}")
    ovdm_path = os.path.join(data_dir, f"Q8OVDM26.{company}")
    try:
        components, skipped = engine.extract_components(mifl_path)
        employees, invalid = engine.extract_employees(ovdm_path)
    except Exception as e:
        st.error(f"Extraction failed: {e}")
    else:
        st.session_state.components = components
        st.session_state.skipped_components = skipped
        st.session_state.employees = employees
        st.session_state.invalid_employees = invalid
        st.session_state.extracted = True
        st.session_state.generated = False
        st.session_state.inserted = False

        os.makedirs(out_dir, exist_ok=True)
        template_text, col_map, stats_cols = engine.build_template(
            template_name, components
        )
        excel_path = os.path.join(out_dir, f"template_{company}.xlsx")
        try:
            engine.build_excel(
                company,
                components,
                col_map,
                stats_cols,
                employees,
                excel_path,
                int(year),
                int(month),
            )
        except Exception as e:
            st.error(f"Excel generation failed: {e}")
        else:
            st.session_state.template_text = template_text
            st.session_state.col_map = col_map
            st.session_state.stats_cols = stats_cols
            st.session_state.excel_path = excel_path
            st.session_state.generated = True

if st.session_state.extracted:
    components = st.session_state.components
    employees = st.session_state.employees

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Components found", len(components))
    m2.metric("Components filtered out", st.session_state.skipped_components)
    m3.metric("Employees found", len(employees))
    m4.metric("Invalid תז blocks", st.session_state.invalid_employees)

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
        st.dataframe(
            [{"תעודת זהות": tz, "name": name} for tz, name in employees],
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
                }
            ]
            for actual, name, ck, cm in st.session_state.col_map:
                rows.append(
                    {
                        "actual code": actual,
                        "name": name,
                        "col כמות": engine.col_letter(ck),
                        "col מחיר": engine.col_letter(cm),
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
            st.error(
                "**Permission denied** — the app cannot write to the data folder.\n\n"
                "Restart the app with:\n"
                "```\n"
                "sudo ~/workspace/micpal-app/.venv/bin/streamlit run app.py\n"
                "```"
            )
        except Exception as e:
            st.error(f"Insertion failed — file NOT modified: {e}")
        else:
            st.session_state.inserted = True
            st.session_state.backup_path = result["insert_backup"]

    if st.session_state.inserted:
        st.success(f"Inserted '{template_name}' into {srgl_path}")
        st.code(f"Backup saved at:\n{st.session_state.backup_path}")
        st.caption("To roll back manually: copy the backup file over Q8SRGL26.000.")
