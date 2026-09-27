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

## Verification

Run related package/feature batches, not individual tests, with `GOMAXPROCS=2` and
`-p 2`. Reuse caches. Record exact scope/results and inspect the diff before each
commit. Keep runtime assertions, staged catalogue review and canonical execution
credit separate. Never rebase; use merge when histories need integration.
