---
name: audit-repo
description: Perform a structured production-readiness, security, and SOC2 audit of a code repository and produce a styled Excel report. Use when the user asks for a "repo audit", "production-readiness review", "audit report", or wants to generate an xlsx-style findings document like the existing audit_report templates.
---

# Repo Audit Skill

This skill performs an end-to-end audit of a code repository against eight dimensions plus SOC2 Trust Service Criteria, and produces a multi-sheet Excel workbook in the same style as the reference `audit_report.xlsx`.

## When to invoke

- User says: "audit this repo", "do an audit report", "generate audit_report.xlsx for X", "production-readiness review of X"
- User asks to apply the same audit format to a new repository

## Workflow

### Step 1 — Crawl the repo

Use the `Explore` agent (or direct tools) to map the repository. Read every file that matters:

- All `*.py` files
- `Dockerfile`, `docker-compose.yml`
- `requirements.txt`, `pyproject.toml`, `setup.py`, `setup.cfg`
- `.gitignore`, `.dockerignore`, `.env*`
- CI/CD: `.github/workflows/*.yml`, `azure-pipelines.yaml`, `pipelines/*.yaml`, `Jenkinsfile`
- IaC: `*.bicep`, `*.tf`, `containerapp.yaml`
- Config: `*.toml`, `*.cfg`, `*.ini`, `*.yaml`, `*.json` at root or in `config/`
- Docs: `README.md`, `CLAUDE.md`, `docs/`
- Tests: `test*/`, `tests/`, `*_test.py`
- Pre-commit: `.pre-commit-config.yaml`
- Look for committed binaries (model files, PDFs) that should be in LFS or blob storage

Do not skim. Read complete file contents for every file < 500 lines. For larger files, read enough to identify all issues in scope.

### Step 2 — Identify findings

For every file/configuration, evaluate against these dimensions. Each finding gets a unique ID, a severity, a priority, and a category.

#### Dimensions (use these exact category names)

1. **Security** — auth, secrets handling, input validation, CORS, credentials on disk, path traversal, prompt injection, decompression bombs, deprecated crypto packages
2. **Production** — error handling, retries, rate limiting, async vs sync (event-loop blocking), input size limits, /healthz vs /readyz, graceful shutdown, hardcoded config, dependency pinning, API versioning
3. **Container** — base image digest pinning, USER directive, ENV-baked secrets, multi-stage build, .dockerignore, HEALTHCHECK, resource limits, layer caching
4. **CI/CD** — pipeline existence, build → test → lint → scan → push → deploy stages, image scanning (Trivy/Grype), image signing (Cosign/Notary), env promotion, manual approval gates, ACR auth (managed identity vs password)
5. **Code Quality** — linter, type checker, formatter, pre-commit, SAST (Bandit), test coverage, pyproject.toml, dead code, duplicate dependencies, broad except blocks, print() vs logging, naming bugs (typos in module names)
6. **Observability** — structured JSON logging, correlation IDs, OpenTelemetry tracing, /metrics, alerting rules, logger name correctness (`__name__` vs `'__name__'`), Application Insights initialization
7. **Azure** — IaC (Bicep/Terraform/containerapp.yaml), managed identity, Key Vault references, autoscaling rules, ingress config, resource limits, App Insights wiring
8. **Additional** — anything that does not fit the above seven (broken code paths, dead modules, structural bugs, repo hygiene, debug endpoints exposed in prod, committed large binaries)
9. **SOC2** — Trust Service Criteria mapping (separate sheet, see below)

#### SOC2 Trust Service Criteria to cover

For SOC2, generate findings against each TSC where applicable. Use ID prefix `SOC2-<criterion>` (e.g. `SOC2-CC6.1`):

- **CC6** — Logical & Physical Access (auth, access logs, RBAC, credential protection, MFA)
- **CC7** — System Operations (incident response, security monitoring)
- **CC8** — Change Management (PR review, change log)
- **CC9** — Risk Mitigation (vendor risk, third-party DPAs)
- **A1** — Availability (SLA/SLO, DR/BCP, geo-redundancy)
- **PI1** — Processing Integrity (audit trail, output validation)
- **C1** — Confidentiality (data classification, DPAs with AI providers, retention/deletion policy)
- **P** — Privacy (PIA, PII handling, log scrubbing)

It is fine — and expected — for SOC2 findings to overlap with Security findings. The SOC2 sheet frames them through the compliance lens.

#### Severity levels (use exactly these strings)

- `CRITICAL` — direct security breach risk, data exfiltration, unauthenticated public access to sensitive operations, secrets baked into shippable artifacts
- `HIGH` — production-readiness blocker, likely outage trigger, significant compliance gap
- `MEDIUM` — meaningful reliability/security improvement, audit finding
- `LOW` — quality of life, minor hygiene

#### Priority levels (use exactly these strings)

- `P0` — fix immediately
- `P1` — fix in current sprint
- `P2` — fix in next sprint
- `P3` — backlog

### Step 3 — Write findings.json

Save findings to `findings.json` at the audited repo root. Schema:

```json
{
  "metadata": {
    "repo_name": "<repo>",
    "audit_date": "YYYY-MM-DD",
    "auditor": "Claude <model>"
  },
  "findings": [
    {
      "id": "S1",
      "category": "Security",
      "severity": "CRITICAL",
      "priority": "P0",
      "title": "Short, actionable title (one line)",
      "files": "path/to/file.py:LINE or comma-separated list",
      "description": "What the issue is and why it matters. 2–4 sentences.",
      "fix": "Concrete remediation steps. Reference specific code changes."
    }
  ]
}
```

ID conventions (recommended, not enforced):

- `0.1, 0.2, ...` — Additional / beyond-scope findings
- `S1, S2, ...` — Security
- `P1, P2, ...` — Production
- `D1, D2, ...` — Container/Docker
- `C1, C2, ...` — CI/CD
- `Q1, Q2, ...` — Code Quality
- `O1, O2, ...` — Observability
- `A1, A2, ...` — Azure
- `SOC2-CC6.1, SOC2-A1.2, ...` — SOC2 (must use this prefix; the renderer extracts the TSC from this prefix)

### Step 4 — Render the xlsx

Run the bundled renderer:

```bash
python3 <skill-dir>/scripts/generate_audit_report.py findings.json --output audit_report_<repo>.xlsx
```

The script reads `findings.json` and emits a 12-sheet styled workbook:

1. Executive Summary (severity counts, category counts, priority legend)
2. All Findings (full table, sortable)
3. Security
4. Production
5. Container
6. CI-CD
7. Code Quality
8. Observability
9. Azure
10. Additional
11. SOC2 (with extra TSC column)
12. Priority Roadmap (sorted P0 → P3)

The script handles styling. Do not write a custom xlsx renderer per repo.

### Step 5 — Summarise to the user

After the script completes, report:

- Total findings count
- Severity breakdown (CRITICAL / HIGH / MEDIUM / LOW counts)
- The 3–5 most urgent items (P0 + CRITICAL Security/SOC2 items)
- Output file path

## Quality bar

- **Be specific.** "No tests" is weak. "tests/ directory contains demo.py and test_translate_client.py which require a live server, no pytest discovery" is strong.
- **Include file:line references** in the `files` field whenever possible. A reviewer must be able to navigate to the issue immediately.
- **Connect findings.** If S1 (key on disk) is the same root cause as C1 (DPA gap), say so — link them in the description.
- **Don't pad.** A 30-finding precise audit beats a 100-finding vague one.
- **Look for "beyond scope" issues.** Dead code, leftover debug endpoints, structural bugs (e.g. imports inside `if __name__ == "__main__":` referenced by module-level code), committed binaries, misspelled module names. These are some of the most valuable findings.

## Dependencies

The renderer needs `openpyxl`. Install once:

```bash
pip install openpyxl
```

## Files in this skill

- `SKILL.md` — this file
- `scripts/generate_audit_report.py` — the xlsx renderer (reads findings.json, writes styled xlsx)
- `templates/findings_template.json` — empty findings file with all enum values documented as comments

## Distribution

To run on another repo:

1. Copy this `skills/audit/` directory into the target repo (or symlink from a central location).
2. From the target repo root, ask Claude: "audit this repo using the audit skill" or invoke the skill directly.
3. Claude crawls, writes `findings.json`, runs the renderer, and gives a summary.

The renderer script is **stable** — do not modify it per repo. All repo-specific content lives in `findings.json`.
