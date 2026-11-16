"""Subprocess helpers for running commission validators from the Streamlit UI."""

from __future__ import annotations

import glob
import os
import signal
import subprocess
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable, List, Optional

REPO_ROOT = Path(__file__).resolve().parent.parent
UPLOADS_DIR = REPO_ROOT / "ui_uploads"
DB_OUTPUT_ROOT = REPO_ROOT / "pipeline_output" / "commission_validation_db"

LogCallback = Callable[[str], None]


@dataclass
class RunResult:
    exit_code: int
    log_text: str
    output_dir: Optional[str] = None
    versioned_dir: Optional[str] = None
    started_at: float = 0.0
    finished_at: float = 0.0
    argv: List[str] = field(default_factory=list)


def ensure_uploads_dir() -> Path:
    UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
    return UPLOADS_DIR


def save_uploaded_file(uploaded_file, run_id: Optional[str] = None) -> Path:
    """Persist a Streamlit UploadedFile under ui_uploads/<run_id>/."""
    ensure_uploads_dir()
    rid = run_id or datetime.now().strftime("%Y%m%d_%H%M%S")
    dest_dir = UPLOADS_DIR / rid
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / uploaded_file.name
    dest.write_bytes(uploaded_file.getbuffer())
    return dest


def _list_version_dirs(root: Path) -> List[Path]:
    if not root.is_dir():
        return []
    dirs = []
    for p in root.iterdir():
        if p.is_dir() and (p.name.startswith("v") or p.name.startswith("run_")):
            dirs.append(p)
    return dirs


def find_newest_output_dir(after_ts: float, root: Optional[Path] = None) -> tuple[Optional[str], Optional[str]]:
    """
    Find the newest versioned run folder created at/after after_ts.
    Returns (output_dir, versioned_dir).
    """
    base = root or DB_OUTPUT_ROOT
    candidates = []
    for vdir in _list_version_dirs(base):
        try:
            mtime = vdir.stat().st_mtime
        except OSError:
            continue
        if mtime + 1.0 >= after_ts:  # small clock skew tolerance
            out = vdir / "output"
            candidates.append((mtime, str(out) if out.is_dir() else str(vdir), str(vdir)))
    if not candidates:
        # Fallback: absolute newest under root
        for vdir in _list_version_dirs(base):
            try:
                mtime = vdir.stat().st_mtime
            except OSError:
                continue
            out = vdir / "output"
            candidates.append((mtime, str(out) if out.is_dir() else str(vdir), str(vdir)))
    if not candidates:
        return None, None
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1], candidates[0][2]


def list_artifacts(output_dir: Optional[str], patterns: Optional[List[str]] = None) -> List[Path]:
    if not output_dir or not os.path.isdir(output_dir):
        return []
    pats = patterns or ["*"]
    found: List[Path] = []
    seen = set()
    for pat in pats:
        for match in glob.glob(os.path.join(output_dir, pat)):
            p = Path(match)
            if p.is_file() and str(p) not in seen:
                seen.add(str(p))
                found.append(p)
    found.sort(key=lambda p: p.name.lower())
    return found


DB_ARTIFACT_PATTERNS = [
    "validation_summary_*.md",
    "file_validated_clean_*.xlsx",
    "file_validated_errors_*.xlsx",
    "db_matched_commission_*.xlsx",
    "db_unmatched_commission_*.xlsx",
    "db_commission_snapshot_*.xlsx",
    "delete_commission_agency5.sql",
    "rejected_missing_user_uuid_payloads_*.json",
    "db_verification_auto_updated_*.xlsx",
    "db_commission_snapshot_post_update_*.xlsx",
    "file_validated_ambiguous_matches_*.xlsx",
    "file_validated_duplicates_*.xlsx",
    "file_validated_zero_comm_ignored_*.xlsx",
    "normalized_yearband_override_*.xlsx",
    "update_writing_agent_from_file.sql",
    "writing_agent_corrections_*.json",
    "post_api_failures_*.log",
]

APL_ARTIFACT_PATTERNS = [
    "apl_db_post_update_validation_*.xlsx",
    "apl_annotated_validation_*.xlsx",
    "apl_annotated_validation_*.csv",
]


def _db_base_argv(
    file_path: str,
    *,
    agency_id: int = 5,
    tolerance: float = 0.02,
    dry_run: bool = False,
    skip_normalize: bool = False,
    rate_sheet: Optional[str] = None,
    norm_file: Optional[str] = None,
    agent_roster: Optional[str] = None,
    policy_mapping: Optional[str] = None,
    env_file: str = ".env",
    output_dir: str = "pipeline_output/commission_validation_db",
) -> List[str]:
    argv = [
        sys.executable,
        str(REPO_ROOT / "validate_policy_commission_db.py"),
        "--file",
        file_path,
        "-y",
        "--agency-id",
        str(agency_id),
        "--tolerance",
        str(tolerance),
        "--env-file",
        env_file,
        "--output-dir",
        output_dir,
    ]
    if dry_run:
        argv.append("--dry-run")
    if skip_normalize:
        argv.append("--skip-normalize")
    if rate_sheet:
        argv.extend(["--rate-sheet", rate_sheet])
    if norm_file:
        argv.extend(["--norm-file", norm_file])
    if agent_roster:
        argv.extend(["--agent-roster", agent_roster])
    if policy_mapping:
        argv.extend(["--policy-mapping", policy_mapping])
    return argv


def build_db_validate_argv(
    file_path: str,
    *,
    agency_id: int = 5,
    tolerance: float = 0.02,
    dry_run: bool = False,
    skip_normalize: bool = False,
    rate_sheet: Optional[str] = None,
    norm_file: Optional[str] = None,
    agent_roster: Optional[str] = None,
    policy_mapping: Optional[str] = None,
    env_file: str = ".env",
    output_dir: str = "pipeline_output/commission_validation_db",
) -> List[str]:
    """
    Validate-only argv.

    Does NOT pass --skip-delete-gate so delete SQL + gate summary are still written.
    Pair with stdin_decline_post() so prompts answer n and never POST.
    """
    return _db_base_argv(
        file_path,
        agency_id=agency_id,
        tolerance=tolerance,
        dry_run=dry_run,
        skip_normalize=skip_normalize,
        rate_sheet=rate_sheet,
        norm_file=norm_file,
        agent_roster=agent_roster,
        policy_mapping=policy_mapping,
        env_file=env_file,
        output_dir=output_dir,
    )


def stdin_decline_post() -> str:
    """Enough 'n' answers to exit the delete gate without POSTing."""
    return "n\nn\nn\n"


def build_db_post_argv(
    file_path: str,
    *,
    mode: str,
    agency_id: int = 5,
    tolerance: float = 0.02,
    dry_run: bool = False,
    skip_normalize: bool = False,
    rate_sheet: Optional[str] = None,
    norm_file: Optional[str] = None,
    agent_roster: Optional[str] = None,
    policy_mapping: Optional[str] = None,
    env_file: str = ".env",
    output_dir: str = "pipeline_output/commission_validation_db",
) -> List[str]:
    """
    POST recreate re-run (non-interactive flags).
    mode: 'missing_only' | 'with_mismatches'
    """
    argv = _db_base_argv(
        file_path,
        agency_id=agency_id,
        tolerance=tolerance,
        dry_run=dry_run,
        skip_normalize=skip_normalize,
        rate_sheet=rate_sheet,
        norm_file=norm_file,
        agent_roster=agent_roster,
        policy_mapping=policy_mapping,
        env_file=env_file,
        output_dir=output_dir,
    )
    argv.append("--skip-delete-gate")
    if mode == "missing_only":
        argv.append("--post-missing-only")
    elif mode == "with_mismatches":
        argv.append("--force-post-mismatches")
    else:
        raise ValueError(f"Unknown POST mode: {mode}")
    return argv


def build_apl_argv(
    file_path: str,
    *,
    agency_id: int = 5,
    dry_run: bool = False,
    verbose: bool = False,
    execution_report: Optional[str] = None,
    run_id: Optional[str] = None,
    env_file: str = ".env",
    output: Optional[str] = None,
) -> List[str]:
    argv = [
        sys.executable,
        str(REPO_ROOT / "validate_apl_vs_db_after_import.py"),
        "-i",
        file_path,
        "-a",
        str(agency_id),
        "-e",
        env_file,
    ]
    if dry_run:
        argv.append("--dry-run")
    if verbose:
        argv.append("-v")
    if execution_report:
        argv.extend(["-x", execution_report])
    if run_id:
        argv.extend(["-r", run_id])
    if output:
        argv.extend(["-o", output])
    return argv


class ProcessHandle:
    """Track a running subprocess so the UI can cancel it."""

    def __init__(self, proc: subprocess.Popen):
        self.proc = proc

    @property
    def pid(self) -> Optional[int]:
        return self.proc.pid if self.proc else None

    def is_running(self) -> bool:
        return self.proc is not None and self.proc.poll() is None

    def cancel(self) -> None:
        if not self.proc or self.proc.poll() is not None:
            return
        try:
            if sys.platform == "win32":
                self.proc.terminate()
            else:
                os.killpg(os.getpgid(self.proc.pid), signal.SIGTERM)
        except (ProcessLookupError, OSError):
            try:
                self.proc.terminate()
            except Exception:
                pass
        try:
            self.proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            try:
                if sys.platform != "win32":
                    os.killpg(os.getpgid(self.proc.pid), signal.SIGKILL)
                else:
                    self.proc.kill()
            except Exception:
                pass


def run_command(
    argv: List[str],
    *,
    on_line: Optional[LogCallback] = None,
    cancel_check: Optional[Callable[[], bool]] = None,
    handle_out: Optional[dict] = None,
    cwd: Optional[str] = None,
    stdin_text: Optional[str] = None,
) -> RunResult:
    """
    Run argv, stream lines to on_line, support cooperative cancel via cancel_check.
    If handle_out dict is provided, stores ProcessHandle under key 'handle'.
    Optional stdin_text is written once then stdin closed (for declining prompts).
    """
    started = time.time()
    env = os.environ.copy()
    env["PYTHONUNBUFFERED"] = "1"
    workdir = cwd or str(REPO_ROOT)

    popen_kwargs = {
        "args": argv,
        "cwd": workdir,
        "env": env,
        "stdout": subprocess.PIPE,
        "stderr": subprocess.STDOUT,
        "stdin": subprocess.PIPE if stdin_text is not None else subprocess.DEVNULL,
        "text": True,
        "bufsize": 1,
    }
    if sys.platform != "win32":
        popen_kwargs["start_new_session"] = True  # own process group for cancel

    proc = subprocess.Popen(**popen_kwargs)
    handle = ProcessHandle(proc)
    if handle_out is not None:
        handle_out["handle"] = handle

    if stdin_text is not None and proc.stdin is not None:
        try:
            proc.stdin.write(stdin_text)
            proc.stdin.close()
        except BrokenPipeError:
            pass

    lines: List[str] = []
    cancelled = False
    exit_code = -1
    try:
        assert proc.stdout is not None
        for line in proc.stdout:
            lines.append(line)
            if on_line:
                on_line(line.rstrip("\n"))
            if cancel_check and cancel_check():
                cancelled = True
                handle.cancel()
                break
        exit_code = proc.wait()
    except Exception:
        handle.cancel()
        raise

    finished = time.time()
    if cancelled:
        exit_code = -1

    out_dir, ver_dir = find_newest_output_dir(started)
    return RunResult(
        exit_code=exit_code,
        log_text="".join(lines),
        output_dir=out_dir,
        versioned_dir=ver_dir,
        started_at=started,
        finished_at=finished,
        argv=list(argv),
    )


def parse_gate_hints(log_text: str) -> dict:
    """Best-effort extract of missing/mismatch counts from validator log."""
    hints = {
        "missing": None,
        "mismatch_delete": None,
        "invalid_left_in_place": None,
        "rules_available": None,
    }
    for line in log_text.splitlines():
        s = line.strip()
        if "missing in DB" in s and ":" in s:
            try:
                hints["missing"] = int(s.rsplit(":", 1)[-1].strip())
            except ValueError:
                pass
        elif "mismatch rows to delete" in s and ":" in s:
            try:
                hints["mismatch_delete"] = int(s.rsplit(":", 1)[-1].strip())
            except ValueError:
                pass
        elif "left in place" in s and ":" in s:
            try:
                hints["invalid_left_in_place"] = int(s.rsplit(":", 1)[-1].strip())
            except ValueError:
                pass
        elif "Rules available to create" in s and ":" in s:
            try:
                hints["rules_available"] = int(s.rsplit(":", 1)[-1].strip())
            except ValueError:
                pass
        # Also catch unmatched summary lines
        elif "Unmatched Rows:" in s:
            pass
    return hints
