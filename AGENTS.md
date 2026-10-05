# Agent guidance

Read `.github/copilot-instructions.md` and `docs/testing.md` before changing code
or test-reference plumbing. Keep edits in this owned repository; shared reference
content and release coordination belong to the shared-repository owner.

## Priority coordination

Use `chat` with `target_agent_name: "@alias"` and explicit `mode: "steer"` for:

- user scope changes and stop/hold instructions;
- release or pin corrections;
- safety blockers or decisions needed to unblock another agent.

Include the CURRENT shared tag/commit, the action and owner, and the prior notice
being superseded. Read the pin from `spec/reference-distribution.json`; do not copy
an older value from a delayed message. Keep existing pins until an approved
replacement is verified. Use `mode: "queue"` only for routine progress that does
not change another agent's next action.

A receiver acknowledges the latest state once. Do not replay historical notices
or reply to stale progress unless correcting a live decision. Publication holds
remain in force until the responsible coordinator explicitly releases them;
authorisation does not grant repository write permission.

## Project-owned caches, temporary files and evidence

The canonical host disposable root is `/workspace/tmp/go-ooxml/`. The repository
ships `scripts/project-tmp.sh` for CI/other hosts: a validated absolute
`PROJECT_TMP_ROOT` ending in `go-ooxml` wins (an invalid explicit override
fails); otherwise choose writable `/workspace/tmp/go-ooxml`, then
`RUNNER_TEMP/go-ooxml`, original `TMPDIR/go-ooxml`, then platform temp base
plus `go-ooxml`. Resolve once before replacing `TMPDIR`; no home-cache
fallback. `Makefile` exports `TMPDIR`/`TMP`/`TEMP` to `runs/make/<run-id>/`,
`GOTMPDIR` to `build/go/`,
`GOCACHE` to `cache/go/build/`, `GOMODCACHE` to `cache/go/mod/`, `GOPATH` to
`cache/go/path/`, plus project-owned XDG, NuGet and dotnet-home caches. The
profiling runner creates isolated `runs/tests/<run-id>/<module>/pkg-<n>/`
roots for each package. For direct commands, source the repository resolver, use the same variables
and create the corresponding directories first; never use bare `/tmp`, ad-hoc
top-level `/workspace/tmp` output, or home caches. Do not change production atomic-save
semantics: its short-lived staging file must stay alongside the caller-chosen
output to preserve same-filesystem rename. Test-owned output must remain under
an isolated project run root, not a live checkout or shared fixtures.

Disposable output belongs only under `cache/`, `build/` or `runs/`; retained
CPU/heap profiles, matching binaries, logs, Go test JSON, receipts and generated
OOXML quality evidence stay in `artifacts/` or their existing retained analysis
location. Never move or remove active jobs or historical graphics evidence.
`make clean` intentionally deletes nothing: inspect jobs and remove only idle,
confirmed disposable directories beneath this project's root. Release/CI runners outside this host must use the repository resolver or an
explicit validated root to provision the same hierarchy; the platform temp
fallback still appends `/go-ooxml` and never uses a home cache.

## Verification

Run related package/feature batches, not individual tests, with `GOMAXPROCS=2`;
`make test`, `make acceptance`, `make test-batch`, and graphics targets invoke
`scripts/test-profile.sh` per package. The runner retains per-run CPU and heap
profiles, binaries, command/toolchain/revision, logs, and cumulative CPU,
`alloc_space` and `alloc_objects` tables under `artifacts/profiles/`. Every
Go test/benchmark/race/fuzz run, including failed runs and nested modules,
requires this capture; missing or empty CPU samples are reported and need a
representative workload before accepting performance. After each run review
application hotspots separately from test/runtime overhead and compare like
workloads. Unprofiled child processes and LibreOffice oracles are separate
and must not be counted as profiled Go activity. Record exact scope/results,
inspect the diff before each commit, and keep runtime assertions, staged
catalogue review and canonical execution credit separate. Never rebase; use
merge when histories need integration.
