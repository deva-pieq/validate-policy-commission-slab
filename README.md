# Policy Commission Validation & Import

Tools for validating carrier commission upload sheets against PieQ’s database, generating safe delete SQL for mismatched rules, and (re)creating policy commission setups via the PieQ API.

---

## What this project does

Given a commission upload file (CSV/Excel) — or a folder of them — the main validator:

1. **Normalizes** the sheet (product names, agent IDs, year-band From/To ranges, override 3-row expansion, split percentages).
2. **Validates** carrier commission totals against an optional master rate sheet.
3. **Compares** clean rows to live `policy_commission` data in PostgreSQL (read-only).
4. **Writes reports** under a versioned output folder.
5. Optionally **gates** a delete → POST recreate flow for missing/mismatched rules (never PUT).

Two companion scripts cover standalone delete-SQL generation and a separate API import runner.

| Script | Role |
|--------|------|
| `validate_policy_commission_db.py` | Main 2-stage validate + DB verify + optional POST recreate |
| `generate_delete_commission_sql.py` | Build raw `DELETE` SQL only (does **not** execute it) |
| `import_policy_commission.py` | Upload sheet → API → create commission rules (standalone importer) |

---

## Requirements

- Python 3.10+ recommended
- PostgreSQL access (read-only) for DB verification
- PieQ API + Keycloak SSO credentials for POST/import flows

```bash
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate

pip install -r requirements.txt
```

Dependencies: `pandas`, `numpy`, `openpyxl`, `xlrd`, `psycopg2-binary`, `python-dotenv`, `tabulate`.

---

## Configuration (`.env`)

Copy your secrets into a local `.env` (gitignored). CLI flags override `.env` where both exist.

### Database (validator)

| Variable | Purpose |
|----------|---------|
| `DATABASE_URL` | Optional full URL (`postgresql://...`) |
| `DB_USER` / `DB_PASSWORD` | Read-only DB credentials |
| `DB_HOST` / `DB_PORT` / `DB_NAME` | Used if URL is not set (defaults: `localhost`, `5432`, `pieq_db`) |
| `AGENCY_ID` | Agency filter (default `5`) |

### Input paths

| Variable | Purpose |
|----------|---------|
| `UPLOAD_FILE_PATH` | Commission upload file **or** directory of CSV/XLSX |
| `RATE_SHEET_PATH` | Master ProductWise commission rate sheet (optional) |
| `NORM_FILE_PATH` | Product name normalization map (optional) |
| `AGENT_ROSTER_PATH` | Agent name → Agent ID roster (optional) |
| `POLICY_MAPPING_PATH` | Recovered-policies mapping Excel (optional) |

### API / SSO (POST recreate & importer)

| Variable | Purpose |
|----------|---------|
| `API_BASE_URL` | PieQ API base (e.g. `https://prod.api.pieq.ai`) |
| `AUTH_URL` / `SSO_TOKEN_URL` | Keycloak auth host |
| `KEYCLOAK_CLIENT_ID` / `client_id` | SSO client (default `pieq-app`) |
| `USERNAME` / `PASSWORD` | SSO credentials |
| `AUTH_TOKEN` / `BEARER_TOKEN` | Optional pre-issued bearer token |
| `X_CLIENT_ID` / `X_SERVICE_NAME` | Optional API identity headers |
| `BATCH_SIZE` | Parallel workers for importer (default `5`) |
| `REQUEST_TIMEOUT` | Per-request timeout seconds (default `30`) |
| `AUTO_CONFIRM` | `1` = auto-confirm importer payloads |
| `RESUME_IMPORT_UUID` | Resume an existing import without re-upload |

Typical reference sheets live under `Truth+Sheets/`; upload sheets under `Files/`.

---

## Pipeline overview

```
Upload sheet(s)
      │
      ▼
┌─────────────────────────────────────┐
│ Pre-Stage: year-band normalize +    │
│            override 3-row expand    │  (--skip-normalize to skip)
└─────────────────────────────────────┘
      │
      ▼
┌─────────────────────────────────────┐
│ Stage 1 — File validation           │
│  • product / agent normalization    │
│  • split roll-up & % shares         │
│  • rate-sheet check                 │
│  → file_validated_clean_*.xlsx      │
│  → file_validated_errors_*.xlsx     │
└─────────────────────────────────────┘
      │
      ▼
┌─────────────────────────────────────┐
│ Stage 1.5 — DB commission snapshot  │
│  → db_commission_snapshot_*.xlsx    │
└─────────────────────────────────────┘
      │
      ▼
┌─────────────────────────────────────┐
│ Stage 2 — DB verification           │
│  products → policies → commission   │
│  → db_matched_commission_*.xlsx     │
│  → db_unmatched_commission_*.xlsx   │
│  → validation_summary_*.md          │
└─────────────────────────────────────┘
      │
      ▼
┌─────────────────────────────────────┐
│ Delete gate (if mismatches/missing) │
│  1. Write delete_commission_*.sql   │
│  2. Senior eng runs DELETE manually │
│  3. Confirm remaining = 0           │
│  4. AUTO-POST recreate (no PUT)     │
└─────────────────────────────────────┘
```

Outputs are written to:

```
pipeline_output/commission_validation_db/vYYYYMMDD_HHMMSS/
  input/     # archived copies of upload files
  output/    # reports, SQL, summaries
```

---

## Commands

### 1. Validate against DB (primary workflow)

```bash
# Paths from .env
python validate_policy_commission_db.py

# Explicit file or folder
python validate_policy_commission_db.py --file "Files/Cigna_....csv"
python validate_policy_commission_db.py --file Files

# Positional path (same as --file)
python validate_policy_commission_db.py "Files/some_sheet.xlsx"

# Dry-run (no live DB)
python validate_policy_commission_db.py --file Files --dry-run

# After DELETE is already done — skip gate and POST recreate
python validate_policy_commission_db.py --file Files -y --force-post-mismatches --skip-delete-gate
```

#### Flags — `validate_policy_commission_db.py`

| Flag | Default | Description |
|------|---------|-------------|
| `file_path` (positional) | — | Same as `--file` / `UPLOAD_FILE_PATH` |
| `--file` | from `.env` | Input CSV/XLSX file or directory |
| `--rate-sheet` | `RATE_SHEET_PATH` | Master rate sheet |
| `--norm-file` | `NORM_FILE_PATH` | Product normalization table |
| `--agent-roster` | `AGENT_ROSTER_PATH` | Master agent roster |
| `--policy-mapping` | `POLICY_MAPPING_PATH` | Recovered policy number mapping |
| `--env-file` | `.env` | Env file path |
| `--output-dir` | `pipeline_output/commission_validation_db` | Parent output directory |
| `--agency-id` | `5` / `AGENCY_ID` | Agency filter for DB queries |
| `--dry-run` | off | Simulate DB without connecting |
| `--tolerance` | `0.02` | Commission value match tolerance (±%) |
| `-y` / `--yes` | off | Helper flag; delete gate still prompts `y/n`, then POST auto-runs |
| `--skip-normalize` | off | Skip year-band From–To normalize and override 3-row expansion |
| `--skip-delete-gate` | off | Skip delete SQL + confirm (pair with `--force-post-mismatches`) |
| `--force-post-mismatches` | off | Skip confirm and AUTO-POST all missing/mismatch rules |

**Delete → POST flow**

1. Validator writes `delete_commission_agency5.sql` (also under `docs/`).
2. Share SQL with a senior engineer; this repo **never** runs `DELETE`.
3. After delete and remaining count is `0`, answer `y` — or re-run with `--force-post-mismatches --skip-delete-gate`.
4. Script POSTs new commission rules only (no PUT).

---

### 2. Generate delete SQL only

Useful when you want the SQL without running full validation/POST.

```bash
python generate_delete_commission_sql.py --file Files
python generate_delete_commission_sql.py --file "Files/foo.csv"
python generate_delete_commission_sql.py   # uses UPLOAD_FILE_PATH

# Custom output + no DB count preview
python generate_delete_commission_sql.py --file Files --out docs/delete_commission_agency5.sql --skip-db-preview
```

#### Flags — `generate_delete_commission_sql.py`

| Flag | Default | Description |
|------|---------|-------------|
| `file_path` (positional) | — | Same as `--file` / `UPLOAD_FILE_PATH` |
| `--file` | from `.env` | Sheet file or directory |
| `--env-file` | `.env` | Env file path |
| `--agency-id` | `5` | Agency ID in generated SQL |
| `--out` | `docs/delete_commission_agency5.sql` | Output `.sql` path |
| `--skip-db-preview` | off | Skip read-only COUNT of rows that would match |

SQL is wrapped in `BEGIN;` … `ROLLBACK;` so it is safe to review before changing the final transaction.

---

### 3. Standalone API import

Uploads a commission CSV to PieQ, groups pending records into payloads, confirms interactively (or with `-y`), and POSTs create rules in parallel.

```bash
python import_policy_commission.py --file /path/to/Commission_Upload.csv
python import_policy_commission.py --file upload.csv --yes
python import_policy_commission.py --resume <IMPORT_UUID> --yes
python import_policy_commission.py --dry-run --file upload.csv
python import_policy_commission.py --api-base-url https://prod.api.pieq.ai -v
```

#### Flags — `import_policy_commission.py`

| Flag | Default | Description |
|------|---------|-------------|
| `-f` / `--file` | `UPLOAD_FILE_PATH` | Commission upload CSV |
| `-r` / `--resume` | `RESUME_IMPORT_UUID` | Resume existing import UUID (skip upload) |
| `-e` / `--env-file` | `.env` | Env file path |
| `-u` / `--api-base-url` | `API_BASE_URL` | API base URL override |
| `-t` / `--token` | `AUTH_TOKEN` | Bearer token override |
| `-U` / `--username` | `USERNAME` | SSO username |
| `-P` / `--password` | `PASSWORD` | SSO password |
| `--sso-url` | `AUTH_URL` | Keycloak token URL |
| `--sso-client-id` | `pieq-app` | Keycloak client ID |
| `-b` / `--batch-size` | `5` | Parallel POST workers |
| `--timeout` | `30` | Request timeout (seconds) |
| `-y` / `--yes` | `AUTO_CONFIRM=1` | Auto-confirm all payloads |
| `--dry-run` | off | Parse/generate payloads only; no API calls |
| `-v` / `--verbose` | off | Debug logging |

---

## Output artifacts

| File pattern | Stage | Meaning |
|--------------|-------|---------|
| `normalized_yearband_override_*.xlsx` | Pre | After year-band / override expansion |
| `file_validated_clean_*.xlsx` | 1 | Rows that passed file validation |
| `file_validated_errors_*.xlsx` | 1 | Rows that failed file validation |
| `db_commission_snapshot_*.xlsx` | 1.5 | Current DB commission rows for valid policies |
| `db_matched_commission_*.xlsx` | 2 | File vs DB matches |
| `db_unmatched_commission_*.xlsx` | 2 | Missing / mismatched rules |
| `validation_summary_*.md` | 2 | Counts + mismatch detail |
| `delete_commission_agency5.sql` | Gate | Raw DELETE for review/execution |
| `db_commission_snapshot_post_update_*.xlsx` | Post | Snapshot after successful POST recreate |
| `rejected_missing_user_uuid_payloads_*.json` | Post | Payloads skipped for missing user UUID |

---

## Suggested day-to-day flow

1. Point `UPLOAD_FILE_PATH` (or `--file`) at the carrier sheet / `Files/` folder.
2. Ensure `RATE_SHEET_PATH`, `NORM_FILE_PATH`, and `AGENT_ROSTER_PATH` are set if you use them.
3. Run `python validate_policy_commission_db.py`.
4. Review `validation_summary_*.md` and unmatched Excel.
5. If recreate is needed: share `delete_commission_agency5.sql`, wait for DELETE + remaining `0`, then confirm or re-run with `--force-post-mismatches --skip-delete-gate`.
6. Confirm post-update snapshot looks correct.

For a first-time bulk import of a clean sheet (no DB compare), use `import_policy_commission.py` instead.

---

## Project layout

```
.
├── validate_policy_commission_db.py   # Main validator + DB verify + POST gate
├── generate_delete_commission_sql.py  # Standalone DELETE SQL generator
├── import_policy_commission.py        # API upload / create importer
├── requirements.txt
├── .env                               # Local secrets (not committed)
├── Files/                             # Commission upload sheets
├── Truth+Sheets/                      # Rate sheet, norm map, agent roster
├── docs/                              # Shared delete SQL artifacts
├── logs/                              # Run logs
└── pipeline_output/                   # Versioned run outputs
```

---

## Notes & caveats

- Prefer **read-only** DB credentials for validation; DELETE is always manual via generated SQL.
- Commission recreate uses **POST only** (create after delete), never PUT/patch of existing rule UUIDs.
- Matching allows a small numeric tolerance (`--tolerance`, default `0.02`) and flexible agent code/name equivalence.
- Temp scripts (`temp_*.py`) are local helpers and not part of the main pipeline.
- Never commit `.env` or live credentials.
