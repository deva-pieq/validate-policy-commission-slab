"""
Commission Validators UI

One command:
    streamlit run ui/app.py

Wraps validate_policy_commission_db.py and validate_apl_vs_db_after_import.py.
Uses .env for DB/API secrets (never edited in the UI).
"""

from __future__ import annotations

import html
import os
from datetime import datetime
from pathlib import Path
from typing import List, Optional

import streamlit as st
from dotenv import dotenv_values

from runner import (
    APL_ARTIFACT_PATTERNS,
    DB_ARTIFACT_PATTERNS,
    REPO_ROOT,
    build_apl_argv,
    build_db_post_argv,
    build_db_validate_argv,
    list_artifacts,
    parse_gate_hints,
    run_command,
    save_uploaded_file,
    stdin_decline_post,
)

st.set_page_config(
    page_title="Commission Validators",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)


def _inject_layout_css() -> None:
    """Keep the main pane scrollable; Streamlit defaults often clip tall pages."""
    st.markdown(
        """
<style>
  /* Let the page / main pane scroll instead of clipping content */
  html, body, #root, .stApp {
    height: 100% !important;
    overflow-y: auto !important;
  }
  [data-testid="stAppViewContainer"] {
    overflow: auto !important;
    height: 100% !important;
  }
  [data-testid="stAppViewContainer"] > .main,
  section.main,
  [data-testid="stMain"],
  [data-testid="stMain"] > div {
    overflow-y: auto !important;
    overflow-x: hidden !important;
    height: auto !important;
    max-height: none !important;
  }
  [data-testid="stMain"] {
    height: 100% !important;
  }
  .block-container {
    padding-top: 1.25rem !important;
    padding-bottom: 4rem !important;
    max-width: 1100px !important;
    overflow: visible !important;
  }
  /* Scrollable run logs — keep the rest of the UI reachable without browser zoom */
  pre.ui-log-scroll {
    max-height: min(40vh, 420px) !important;
    overflow: auto !important;
    white-space: pre-wrap !important;
    word-break: break-word !important;
    font-size: 0.8rem !important;
    line-height: 1.35 !important;
    background: #0e1117 !important;
    color: #fafafa !important;
    padding: 0.75rem 1rem !important;
    border-radius: 0.5rem !important;
    margin: 0 !important;
  }
</style>
        """,
        unsafe_allow_html=True,
    )


def _env_defaults() -> dict:
    env_path = REPO_ROOT / ".env"
    if env_path.is_file():
        return {k: (v or "") for k, v in dotenv_values(env_path).items()}
    return {}


def _init_state():
    defaults = {
        "db_log": "",
        "db_result": None,
        "db_output_dir": None,
        "db_file_path": None,
        "db_cancel": False,
        "db_running": False,
        "post_log": "",
        "post_result": None,
        "apl_log": "",
        "apl_result": None,
        "apl_output_dir": None,
        "apl_file_path": None,
        "apl_cancel": False,
        "apl_running": False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def _download_buttons(artifacts: List[Path], key_prefix: str):
    if not artifacts:
        st.info("No report files found yet.")
        return
    cols = st.columns(2)
    for i, path in enumerate(artifacts):
        try:
            data = path.read_bytes()
        except OSError as e:
            cols[i % 2].warning(f"Could not read {path.name}: {e}")
            continue
        cols[i % 2].download_button(
            label=f"⬇️ {path.name}",
            data=data,
            file_name=path.name,
            mime="application/octet-stream",
            key=f"{key_prefix}_{path.name}_{i}",
        )


def _save_optional_upload(uploaded, label: str) -> Optional[str]:
    if uploaded is None:
        return None
    path = save_uploaded_file(uploaded, run_id=f"opt_{label}_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
    return str(path)


def _sidebar_common(env: dict) -> dict:
    st.sidebar.header("Shared settings")
    st.sidebar.caption("Secrets come from `.env` (not shown here).")
    agency_id = st.sidebar.number_input(
        "Agency ID",
        min_value=1,
        value=int(env.get("AGENCY_ID") or 5),
        step=1,
    )
    env_file = st.sidebar.text_input("Env file", value=".env")
    st.sidebar.markdown("---")
    st.sidebar.markdown(
        "**Safety**\n\n"
        "- DELETE SQL is download-only — never executed by this UI.\n"
        "- POST recreate only runs after you confirm in Tab 1.\n"
        "- Use Tab 2 after updates are complete."
    )
    return {"agency_id": int(agency_id), "env_file": env_file}


def _render_log(placeholder, text: str):
    body = html.escape(text[-80000:] if text else "(waiting for output…)")
    placeholder.markdown(f'<pre class="ui-log-scroll">{body}</pre>', unsafe_allow_html=True)


def tab_db_validation(shared: dict, env: dict):
    st.subheader("1 · Validate against DB")
    st.markdown(
        "Upload a commission sheet, run Stage 1 + Stage 2 validation, review reports, "
        "then optionally POST recreate (after manual DELETE if needed)."
    )

    uploaded = st.file_uploader(
        "Commission upload (CSV / Excel)",
        type=["csv", "xlsx", "xls"],
        key="db_upload",
    )

    with st.expander("Optional mapping files", expanded=False):
        c1, c2 = st.columns(2)
        with c1:
            rate_upload = st.file_uploader("Rate sheet", type=["xlsx", "xls", "csv"], key="rate_up")
            rate_path = st.text_input("Rate sheet path", value=env.get("RATE_SHEET_PATH") or "", key="rate_path")
            norm_upload = st.file_uploader("Norm file", type=["xlsx", "xls", "csv"], key="norm_up")
            norm_path = st.text_input("Norm file path", value=env.get("NORM_FILE_PATH") or "", key="norm_path")
        with c2:
            roster_upload = st.file_uploader("Agent roster", type=["xlsx", "xls", "csv"], key="roster_up")
            roster_path = st.text_input(
                "Agent roster path", value=env.get("AGENT_ROSTER_PATH") or "", key="roster_path"
            )
            mapping_upload = st.file_uploader("Policy mapping", type=["xlsx", "xls", "csv"], key="map_up")
            mapping_path = st.text_input(
                "Policy mapping path", value=env.get("POLICY_MAPPING_PATH") or "", key="map_path"
            )

    oc1, oc2, oc3 = st.columns(3)
    with oc1:
        dry_run = st.checkbox("Dry run (no live DB)", value=False, key="db_dry")
    with oc2:
        skip_normalize = st.checkbox("Skip normalize", value=False, key="db_skip_norm")
    with oc3:
        tolerance = st.number_input("Tolerance", min_value=0.0, value=0.02, step=0.01, key="db_tol")

    log_box = st.empty()
    status_box = st.empty()

    run_col, cancel_col = st.columns([3, 1])
    run_clicked = run_col.button("▶ Run validation", type="primary", use_container_width=True, key="db_run")
    cancel_clicked = cancel_col.button("⏹ Cancel", use_container_width=True, key="db_cancel_btn")

    if cancel_clicked:
        st.session_state.db_cancel = True

    if run_clicked:
        if uploaded is None:
            st.error("Upload a commission CSV or Excel file.")
        else:
            saved = save_uploaded_file(uploaded)
            file_path = str(saved)
            st.session_state.db_file_path = file_path
            st.session_state.db_cancel = False
            st.session_state.db_running = True
            st.session_state.db_log = ""
            st.session_state.db_result = None
            st.session_state.post_result = None
            st.session_state.post_log = ""

            rate = _save_optional_upload(rate_upload, "rate") or (rate_path.strip() or None)
            norm = _save_optional_upload(norm_upload, "norm") or (norm_path.strip() or None)
            roster = _save_optional_upload(roster_upload, "roster") or (roster_path.strip() or None)
            mapping = _save_optional_upload(mapping_upload, "mapping") or (mapping_path.strip() or None)

            argv = build_db_validate_argv(
                file_path,
                agency_id=shared["agency_id"],
                tolerance=float(tolerance),
                dry_run=dry_run,
                skip_normalize=skip_normalize,
                rate_sheet=rate,
                norm_file=norm,
                agent_roster=roster,
                policy_mapping=mapping,
                env_file=shared["env_file"],
            )

            status_box.info(
                "Running validation (writes delete SQL; answers n to skip POST)…"
            )
            buf: List[str] = []

            def on_line(line: str):
                buf.append(line)
                st.session_state.db_log = "\n".join(buf)
                _render_log(log_box, st.session_state.db_log)

            result = run_command(
                argv,
                on_line=on_line,
                cancel_check=lambda: bool(st.session_state.get("db_cancel")),
                stdin_text=stdin_decline_post(),
            )
            st.session_state.db_result = result
            st.session_state.db_output_dir = result.output_dir
            st.session_state.db_running = False
            if result.exit_code == 0:
                status_box.success(f"Validation finished (exit 0). Output: `{result.output_dir}`")
            elif result.exit_code == -1:
                status_box.warning("Run cancelled.")
            else:
                status_box.error(f"Validation exited with code {result.exit_code}.")

    if st.session_state.db_log:
        with st.expander("Live / last console log", expanded=True):
            st.code(st.session_state.db_log[-80000:], language="text")

    result = st.session_state.db_result
    if result and result.output_dir:
        st.markdown("### Reports")
        st.caption(f"Output folder: `{result.output_dir}`")
        artifacts = list_artifacts(result.output_dir, DB_ARTIFACT_PATTERNS)
        _download_buttons(artifacts, "db_art")

        hints = parse_gate_hints(result.log_text or "")
        # If validate-only skipped the gate printout, still show POST panel
        st.markdown("---")
        st.markdown("### Update gate (optional POST recreate)")
        st.warning(
            "This UI never runs DELETE. Download `delete_commission_agency5.sql`, "
            "have senior eng run it, confirm remaining=0, then choose a POST option."
        )
        if hints.get("missing") is not None or hints.get("mismatch_delete") is not None:
            m1, m2, m3 = st.columns(3)
            m1.metric("Missing (no delete)", hints.get("missing") if hints.get("missing") is not None else "—")
            m2.metric(
                "Mismatch to delete",
                hints.get("mismatch_delete") if hints.get("mismatch_delete") is not None else "—",
            )
            m3.metric(
                "Left in place (invalid)",
                hints.get("invalid_left_in_place") if hints.get("invalid_left_in_place") is not None else "—",
            )
        else:
            st.caption("Gate counts not found in log — check delete SQL / unmatched report below.")

        post_mode = st.radio(
            "POST action",
            options=["skip", "missing_only", "with_mismatches"],
            format_func=lambda x: {
                "skip": "Skip POST",
                "missing_only": "POST missing-in-DB only (safe without DELETE)",
                "with_mismatches": "POST missing + mismatches (requires DELETE remaining=0)",
            }[x],
            key="post_mode",
        )

        confirm_delete = False
        if post_mode == "with_mismatches":
            confirm_delete = st.checkbox(
                "I confirm mismatch rows were deleted and remaining = 0",
                key="confirm_delete_zero",
            )

        post_log_box = st.empty()
        post_status = st.empty()
        if st.button("Confirm POST recreate", type="primary", disabled=post_mode == "skip", key="post_go"):
            if post_mode == "with_mismatches" and not confirm_delete:
                st.error("Check the DELETE remaining=0 confirmation first.")
            elif not st.session_state.db_file_path:
                st.error("No input file from the validation run.")
            else:
                st.session_state.db_cancel = False
                rate = (st.session_state.get("rate_path") or "").strip() or None
                norm = (st.session_state.get("norm_path") or "").strip() or None
                roster = (st.session_state.get("roster_path") or "").strip() or None
                mapping = (st.session_state.get("map_path") or "").strip() or None
                argv = build_db_post_argv(
                    st.session_state.db_file_path,
                    mode=post_mode,
                    agency_id=shared["agency_id"],
                    tolerance=float(st.session_state.get("db_tol") or 0.02),
                    dry_run=bool(st.session_state.get("db_dry")),
                    skip_normalize=bool(st.session_state.get("db_skip_norm")),
                    rate_sheet=rate,
                    norm_file=norm,
                    agent_roster=roster,
                    policy_mapping=mapping,
                    env_file=shared["env_file"],
                )
                post_status.info(f"Running POST recreate ({post_mode})…")
                buf: List[str] = []

                def on_line(line: str):
                    buf.append(line)
                    st.session_state.post_log = "\n".join(buf)
                    _render_log(post_log_box, st.session_state.post_log)

                post_result = run_command(
                    argv,
                    on_line=on_line,
                    cancel_check=lambda: bool(st.session_state.get("db_cancel")),
                )
                st.session_state.post_result = post_result
                if post_result.output_dir:
                    st.session_state.db_output_dir = post_result.output_dir
                if post_result.exit_code == 0:
                    post_status.success(f"POST run finished. Output: `{post_result.output_dir}`")
                elif post_result.exit_code == -1:
                    post_status.warning("POST run cancelled.")
                else:
                    post_status.error(f"POST run exited with code {post_result.exit_code}.")

        if st.session_state.post_log:
            with st.expander("POST console log", expanded=True):
                st.code(st.session_state.post_log[-80000:], language="text")
            if st.session_state.post_result and st.session_state.post_result.output_dir:
                st.markdown("#### Post-run reports")
                _download_buttons(
                    list_artifacts(st.session_state.post_result.output_dir, DB_ARTIFACT_PATTERNS),
                    "post_art",
                )


def tab_apl_validation(shared: dict, env: dict):
    st.subheader("2 · Post-import APL vs DB")
    st.markdown(
        "Use **after** DB updates / POSTs are complete. "
        "Read-only check that DB commission values match the original APL file."
    )

    uploaded = st.file_uploader(
        "APL / commission file (CSV / Excel)",
        type=["csv", "xlsx", "xls"],
        key="apl_upload",
    )
    exec_report = st.text_input("Optional execution report path", value="", key="apl_exec")

    c1, c2 = st.columns(2)
    with c1:
        dry_run = st.checkbox("Dry run", value=False, key="apl_dry")
    with c2:
        verbose = st.checkbox("Verbose", value=False, key="apl_verbose")

    log_box = st.empty()
    status_box = st.empty()
    run_col, cancel_col = st.columns([3, 1])
    run_clicked = run_col.button("▶ Run post-import validation", type="primary", use_container_width=True, key="apl_run")
    cancel_clicked = cancel_col.button("⏹ Cancel", use_container_width=True, key="apl_cancel_btn")

    if cancel_clicked:
        st.session_state.apl_cancel = True

    if run_clicked:
        if uploaded is None:
            st.error("Upload an APL / commission CSV or Excel file.")
        else:
            saved = save_uploaded_file(uploaded, run_id=f"apl_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
            file_path = str(saved)
            st.session_state.apl_file_path = file_path
            st.session_state.apl_cancel = False
            st.session_state.apl_running = True
            st.session_state.apl_log = ""
            st.session_state.apl_result = None

            argv = build_apl_argv(
                file_path,
                agency_id=shared["agency_id"],
                dry_run=dry_run,
                verbose=verbose,
                execution_report=(exec_report.strip() or None),
                env_file=shared["env_file"],
            )
            status_box.info("Running post-import APL validation…")
            buf: List[str] = []

            def on_line(line: str):
                buf.append(line)
                st.session_state.apl_log = "\n".join(buf)
                _render_log(log_box, st.session_state.apl_log)

            result = run_command(
                argv,
                on_line=on_line,
                cancel_check=lambda: bool(st.session_state.get("apl_cancel")),
            )
            st.session_state.apl_result = result
            st.session_state.apl_output_dir = result.output_dir
            st.session_state.apl_running = False
            if result.exit_code == 0:
                status_box.success(f"APL validation finished. Output: `{result.output_dir}`")
            elif result.exit_code == -1:
                status_box.warning("Run cancelled.")
            else:
                status_box.error(f"APL validation exited with code {result.exit_code}.")

    if st.session_state.apl_log:
        with st.expander("Live / last console log", expanded=True):
            st.code(st.session_state.apl_log[-80000:], language="text")

    result = st.session_state.apl_result
    if result and result.output_dir:
        st.markdown("### Reports")
        st.caption(f"Output folder: `{result.output_dir}`")
        artifacts = list_artifacts(result.output_dir, APL_ARTIFACT_PATTERNS)
        if not artifacts:
            # APL may write under run_* folders — show any xlsx/csv in that dir
            artifacts = list_artifacts(result.output_dir, ["*.xlsx", "*.csv"])
        _download_buttons(artifacts, "apl_art")


def main():
    _init_state()
    _inject_layout_css()
    env = _env_defaults()
    st.title("Commission Validators")
    st.caption(
        "Local UI for `validate_policy_commission_db.py` and `validate_apl_vs_db_after_import.py`. "
        f"Repo: `{REPO_ROOT}`"
    )

    shared = _sidebar_common(env)
    tab1, tab2 = st.tabs(["DB validation + POST gate", "Post-import APL validation"])
    with tab1:
        tab_db_validation(shared, env)
    with tab2:
        tab_apl_validation(shared, env)


if __name__ == "__main__":
    main()
