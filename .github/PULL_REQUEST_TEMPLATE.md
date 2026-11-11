# K-IRRBB — Pull Request

> Target the active release branch, named in [RELEASING.md](../RELEASING.md#current-status), except for an integration into `main` that I have approved. Use synthetic data only. No claim of Excel execution without an actual recorded run.

## Scope

What changes and why?

## Issues

<!-- A task PR uses Resolves # for the issues it completes. An integration PR
(release branch into main), or any PR that only carries earlier work, uses
Refs # for those issues and PRs, so that no issue closes at the wrong time.
Milestones belong to issues only. See docs/GOVERNANCE.md#pull-requests. -->

- Resolves #

## Contract and risks

- Model/feature impact and non-goals:
- Data provenance, model version, units, handling of missing values:
- Backward compatibility and fallback / stale-result handling:

## Evidence

- [ ] `python tools/check.py --ci` completed on this commit (record outcome)
- [ ] Added or updated synthetic regression test(s), or explained why not applicable
- [ ] Independent numerical expected values and tolerances, or N/A
- [ ] Real Excel compile/run on exact SHA, or explicitly NOT RUN with reason
- [ ] Docs and changelog updated

## Final-head review

<!-- After the last push, compare the reviewed commit with the current head.
Any push after this record renews the review and the CI run. Dependabot pull
requests do not use this template; record their final-head review in a
comment on the pull request. -->

- Final PR head SHA:
- Reviewed commit SHA and review type:
- Review result and link:
- Static repository checks result and run link:

- [ ] The review covers the current PR head; every finding is fixed or answered and its thread resolved.
- [ ] Static repository checks pass on the current head.
- [ ] Any push after this record is followed by a renewed review and CI run.

## Reviewer notes

Unresolved risks, platform limits, follow-up work:
