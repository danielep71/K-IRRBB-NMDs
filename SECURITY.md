<div align="center">

# 🔒 IRRBB Toolkit Security Policy

### Interest Rate Risk in the Banking Book — behavioral models in Excel/VBA

[![Reporting](https://img.shields.io/badge/Reporting-Private-d97706?style=for-the-badge)](#reporting-a-vulnerability)
[![Support](https://img.shields.io/badge/Support-Pre--release-6e7781?style=for-the-badge)](#supported-versions)
[![Scope](https://img.shields.io/badge/Scope-Source_%7C_Automation-0969da?style=for-the-badge)](#security-scope)
[![Disclosure](https://img.shields.io/badge/Disclosure-Coordinated-6f42c1?style=for-the-badge)](#coordinated-disclosure)

<br>

**Protect users · Minimize exposure · Preserve evidence · Coordinate disclosure**

</div>

---

This document is authoritative for **vulnerability scope, private reporting,
security triage, coordinated disclosure and safe harbor**. Contribution workflow
is owned by [`CONTRIBUTING.md`](CONTRIBUTING.md); the release sequence by
[`RELEASING.md`](RELEASING.md); repository governance by
[`docs/GOVERNANCE.md`](docs/GOVERNANCE.md).

> [!IMPORTANT]
> A security policy does not make macros, workbooks, add-ins or release artifacts
> inherently trustworthy. Establish provenance and apply organizational security
> controls before enabling executable content.

## 🧭 Security model

The project assumes Microsoft Excel, the operating system and the VBA runtime
are trusted; the user is authorized to run the project; macros are enabled
through an approved trust mechanism; and source comes from this repository.

These are trust boundaries, not guarantees. VBA projects running in the same
Excel process are not isolated security sandboxes.

<a id="supported-versions"></a>

## 📦 Supported versions

The IRRBB Toolkit is **pre-release**. No version has been released, so no
version carries production security support.

| Source state | Security support |
| --- | --- |
| `release/v0.2.0` (active development) | ⚠️ Best effort |
| `main` | ⚠️ Best effort |
| Modified copies or unofficial mirrors | ❌ Unsupported unless reproduced in official source |

This table will be revised when the first version is released. Reports must
identify a full commit SHA; relative descriptions such as "latest" are
insufficient.

<a id="reporting-a-vulnerability"></a>

## 📣 Reporting a vulnerability

Do **not** disclose a suspected vulnerability in an issue, pull request, commit
message, sample workbook, screenshot or release note.

Report through GitHub private vulnerability reporting: **Security → Report a
vulnerability** on the repository. If that is unavailable to you, report by
email to **danielep71@gmail.com** with the subject
**Private security report — IRRBB Toolkit**.

Include only the information needed to assess the issue:

| Evidence | Requested detail |
| --- | --- |
| Identity | Full commit SHA, component or procedure, affected file or workflow |
| Environment | Excel/Office build, bitness, Windows version and deployment model |
| Impact | Confidentiality, integrity, availability, execution or supply-chain consequence |
| Reproduction | Minimal steps using synthetic data |
| Exploitability | Preconditions, privileges, user interaction and persistence |
| Mitigation | Tested workaround or containment, if known |
| Evidence | Sanitized diagnostics, hashes or proof of concept |

Never send real deposit balances, account or customer records, calibration
datasets, ALM extracts, client or employer workbooks, or personal data. Remove
credentials, internal paths, links, connections, document metadata, cached
values and unrelated content.

If a secret has been exposed, revoke or rotate it immediately before improving
the report.

## ⏱️ Response process

The IRRBB Toolkit is maintained by one person; targets are best-effort, not a
contractual SLA.

| Stage | Target |
| --- | --- |
| Acknowledge | Within 5 business days |
| Initial scope and severity assessment | Within 10 business days after sufficient evidence |
| Active-investigation update | At least every 14 days |
| Remediation and disclosure | Proportionate to severity, exploitability and validation needs |

The normal path is reproduce → scope affected commits → contain risk → fix →
add regression evidence → validate in Excel → publish a correction when
appropriate.

## 🎯 Security issue or ordinary defect?

When uncertain, report privately. Security-relevant reports include credible
risk of:

- unintended code execution or trust-boundary crossing;
- unauthorized reading, modification, deletion or disclosure of data;
- persistent or exploitable loss of availability;
- credential, token, runner or automation compromise;
- a provenance or validation bypass that can present unvalidated model output
  as trusted; or
- a correctness defect deliberately exploitable to defeat an integrity boundary.

An incorrect coefficient, decay profile or backtest statistic, a compatibility
problem, a bounded performance regression or a documentation error is normally
an ordinary bug, reported through an issue, unless it creates concrete security
impact.

<a id="security-scope"></a>

## 🛠️ Security scope

### In scope

- source in this repository, including exported VBA once it is added;
- repository-owned validation tooling in [`tools/`](tools/README.md);
- GitHub Actions workflows, their permissions and pinned actions; and
- security or integrity behavior introduced by project code.

### Current risk surfaces

- **Runtime.** No VBA source exists yet, so there is no runtime surface. Future
  code is expected to read and write only its own workbook; any use of files,
  network, native code (`Declare`), `Shell` or `CreateObject` needs an issue and
  explicit review.
- **Automation.** Every workflow checkout that can run code under review sets
  `persist-credentials: false`, so that code never receives Git credentials.
  The static-check and label-drift jobs run with a read-only token. The
  label-reconciliation job holds `issues: write` and runs only on pushes to
  `main` and manual dispatch, never on pull requests.
  The one exception to the credential rule is the daily traffic export. It
  runs only on the default branch, on a schedule or manual dispatch, keeps Git
  credentials to push its data to the `traffic-history` branch, and holds
  `contents: write` and `issues: write` for that branch and for its alert
  issues. Its `TRAFFIC_TOKEN`, a fine-grained token with `Administration:
  read` on this repository only, lives in the `analytics` environment, which
  is restricted to `main`; rotate it every 90 days and immediately after
  suspected exposure. The environment name is not a boundary that binds the
  token to that one workflow, so keep the environment out of every build,
  test and release job.
  All actions are pinned to full commit SHAs. See
  [`tools/README.md`](tools/README.md), [`docs/LABELS.md`](docs/LABELS.md) and
  [`docs/TRAFFIC.md`](docs/TRAFFIC.md).
- **Artifacts.** No workbook, add-in or other binary is distributed. Office
  packages are ignored by Git unless an exact path is re-included.

### Out of scope

- vulnerabilities in Microsoft Excel, Office, Windows, GitHub, Python, Node.js or
  VBA themselves;
- organization-controlled endpoint, macro, access or deployment policy;
- malicious VBA already trusted in the same Excel process;
- unrelated workbooks, add-ins, dependencies or infrastructure, including
  downstream ALM platforms that consume exported parameters;
- modified copies, mirrors or historical snapshots;
- compromised user credentials not exposed by this project; and
- ordinary defects without concrete security impact.

Upstream vulnerabilities belong with the responsible vendor or platform.

<a id="data-and-secrets"></a>

## 🔐 Data and secret handling

Never commit, upload, log or attach:

- passwords, API keys, personal access tokens, signing keys, certificates or
  connection strings;
- real deposit balances, account histories, customer rates, segment data,
  calibration datasets or any client, employer or personal data;
- proprietary source, internal bank models, vendor code, workbooks or licensed
  data;
- internal URLs, machine-specific paths, environment dumps or unredacted
  screenshots; or
- exploit material beyond what is necessary to establish the issue.

Test fixtures and examples are synthetic. The repository is public; no
visibility setting, private fork or private branch waives client
confidentiality, GDPR or contractual restrictions. Excel files
can contain sensitive material outside visible cells, including document
properties, names, hidden sheets, VBA, cached values, queries, links and
connections.

These rules are enforced by review, not by automation. Repository secrets must
use least privilege, remain unavailable to untrusted pull-request code and be
rotated after suspected exposure.

## 📦 Supply-chain boundary

Trusted source is limited to this repository. The release sequence is defined in
[`RELEASING.md`](RELEASING.md) and is not duplicated here.

Security-sensitive workflow changes require least-privilege permissions,
full-SHA action pins and explicit review. Do not run untrusted code on a
persistent credentialed Excel/Windows machine. Treat logs, screenshots,
workbooks, test artifacts and environment metadata as potentially sensitive.

## ✅ Safe-use guidance

Users should:

- preserve organization-approved macro security and deployment controls;
- obtain source only from this repository and know which commit it is;
- test with synthetic data in a controlled environment before using real data;
  and
- understand that IRRBB Toolkit output is a modeling aid. It is not
  regulatory approval, independent model validation or an authentication or
  authorization control.

<a id="coordinated-disclosure"></a>

## 📣 Coordinated disclosure

Avoid wider disclosure while exploitability is being assessed, a fix is being
prepared, users have not had reasonable time to update, or an exposed secret
remains valid.

The maintainer and reporter agree a plan based on severity, active
exploitation, remediation complexity, workarounds and validation time. The
maintainer may request a sanitized reproduction, more environment detail,
confirmation against a candidate fix or a reasonable embargo.

<a id="safe-harbor"></a>

## 🛡️ Good-faith research and safe harbor

Good-faith research is welcome when it:

- stays within project-owned source, artifacts and documented integrations;
- avoids privacy violations, destructive actions, persistence, social
  engineering and unnecessary access;
- stops after establishing the minimum required evidence;
- reports privately and promptly; and
- allows reasonable investigation and remediation time.

The project will not initiate or recommend legal action solely for research
conducted in good faith and consistently with this policy. This does not
authorize testing third-party systems or bind Microsoft, GitHub, an employer,
client or other third party.

No paid bug bounty is offered.

## 📚 Related documents

- [`CONTRIBUTING.md`](CONTRIBUTING.md) — contribution and review workflow
- [`CODE_OF_CONDUCT.md`](CODE_OF_CONDUCT.md) — participant behavior
- [`INSTALLATION.md`](INSTALLATION.md) — developer setup and safe import
- [`RELEASING.md`](RELEASING.md) — release sequence
- [`docs/GOVERNANCE.md`](docs/GOVERNANCE.md) — branches, issues and review process

Conduct complaints and vulnerability reports are different: use the Code of
Conduct for participant behavior and this policy for software and security risk.

---

<div align="center">

### Security principle

**Trust deliberately · Run minimally · Protect secrets · Preserve evidence · Disclose responsibly**

<br>

Maintained by **Daniele Penza**

</div>
