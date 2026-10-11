# v0.1.0 foundation closeout

This checklist defines milestone closeout (#12). It is not authorization to
publish a release, tag a version, make the repository public, or certify models.

> ✅ **Closed on 2026-10-09.** The owner closed milestone v0.1.0; each item
> below records the evidence it rests on. Excel and model checks were not run
> in this milestone and remain NOT RUN.

- [x] Architecture decision and source/provenance review recorded (#2, #11).
  Architecture in [`REPOSITORY_STRUCTURE.md`](REPOSITORY_STRUCTURE.md);
  full-history scans recorded in #11 and #42.
- [x] Host, data and model baseline acceptance records are linked (#3, #4, #9),
  with audit amendments and implementation blockers explicitly tracked (#26,
  #32, #33, #34, #35, #36, #37). Each amendment issue sits in the milestone
  that needs it (table below).
- [x] Sanitized template, import/export and regression harness obligations are
  satisfied with exact-commit Excel evidence (#5, #6, #8), or the owner explicitly
  records a scope/milestone change for each deferred obligation. Foundation scope
  alone does not waive these existing acceptance criteria. The owner moved all
  three to v0.2.0 on 2026-10-09 (#12, roadmap below); their Excel evidence
  remains mandatory there.
- [x] Static gates pass on the final candidate, with regressions for known defects
  and no unresolved blocking review findings (#7, #22, #25). All 10 gates pass
  on `main` at the #55 merge (a231fee); every review finding on the closing PRs
  was fixed before merge.
- [x] All open and closed issues have `danielep71`, a milestone and one priority;
  default milestone configuration is updated before retiring milestone 1 (#10).
  `ISSUE_MILESTONE_NUMBER` was set to 2 (v0.2.0) before the milestone closed.
- [x] Rulesets and repository visibility are re-read after settings changes; no
  direct protected-branch commits or history rewriting is used to hide past gaps.
  Rulesets read back on 2026-10-09 (#42); the repository became public after the
  distribution review. Earlier direct commits are disclosed in the
  retrospective record below, not hidden.
- [x] Every remaining issue is completed with evidence or explicitly deferred
  by the owner. Excel/model checks not executed are recorded as NOT RUN.
  Milestone v0.1.0 has no open issues; every closing PR records Excel as
  NOT RUN.
- [x] Owner explicitly accepts closeout and separately authorizes any subsequent
  release-to-main integration, tag, release or distribution under RELEASING.md.
  Closeout accepted by closing the milestone; integrations into `main` were made
  on explicit owner instruction (#44, #51, #55); no tag or release exists, and
  public visibility was decided separately in #42.

## Approved roadmap and deferred obligations

Detailed explanations of each phase, its purpose, dependencies and completion
criteria are in the [milestone guides](Milestones/README.md). These guides
explain the approved scope; they do not certify delivery or replace the
authoritative data/model contracts and issue evidence.

On 2026-10-09 the owner approved this roadmap and the corresponding milestone
moves. The scope decision is recorded in #12 and in the affected issue bodies.
The previous requirement for an explicit deferral of #5/#6/#8 is satisfied by
that decision; their actual Excel deliverables remain mandatory in v0.2.0.

| Milestone | Deliverable / tracked prerequisites |
| --- | --- |
| v0.1.0 | Repository foundation and accepted baseline contracts |
| v0.2.0 | Workbook #5, source import/export #6, Excel harness #8, data availability/censoring #32 and validated data import |
| v0.3.0 | Decay model, independent references and Excel evidence #36 |
| v0.4.0 | Rates model after stability/statistical amendments #33 |
| v0.5.0 | Stable model after cohort/estimation amendments #34 |
| v0.6.0 | Integrated calibration/backtesting after Stable-Decay composition #37 |
| v0.7.0 | Scenarios, verified regulatory applicability/overlays #35; amendment tracker #26 |
| v0.8.0 | User workflow, parameter persistence and controlled ALM export |
| v0.9.0 | Feature freeze, complete candidate qualification and distribution review |
| v1.0.0 | Owner-accepted stable NMD application and separately authorized release |

#26 remains open until its original obligations are evidenced. Its v0.7.0
allocation reflects the last amendment delivery, not permission to implement
Rates, Stable or data transformations before their earlier prerequisites.
Independent benchmarks and real-Excel tests accompany each model milestone;
v0.9.0 consolidates qualification rather than starting validation.

The initial v1.0.0 scope is the NMD behavioral-model application. A full
banking-book EVE/NII engine needs separate scope approval. No dates or release
publication are implied by this roadmap. Public visibility is a separate owner
decision, recorded in the distribution review #42; it does not authorize a
tag, release or workbook distribution.

The 2026-10-08 audit found no production VBA or workbook. Portable static tests
cannot satisfy the unexecuted Excel, numerical, provenance or distribution gates.

## Retrospective review record

The accepted documentation entered the release branch without associated PRs:
host `6c6725f10556d07b34252ac3e8ecdca9cf57fcdd`, data
`d17020ed304baf876a65570c3c9a6e1e3a72a5b2`, and model baseline
`dff85bfbc470f2d477fd453ab1200f5521b5434f`. Their acceptance evidence is recorded
in #3, #4 and #9. The audit corrections PR for #25 reviews the resulting snapshot
and links #26 for remaining specification gaps. This is retrospective review,
not a claim that those commits originally passed a PR review. History is retained.
