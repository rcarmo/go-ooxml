@catalogue @not-executed
Feature: Planned OPC graph mutation and lifecycle custody
  Candidate outcomes for central reconciliation; graph edits do not imply format-level clone or import support.

  @candidate-go-graph-001
  Scenario: Relationship targets resolve only within supported package URI syntax
    Given package-relative or root-relative targets with encoded spaces or a fragment
    When internal relationship targets are resolved
    Then the target identifies the expected package part
    And escaping, malformed, scheme-bearing or unproved URI forms refuse
    And invalid relationship-registry owner paths are rejected

  @candidate-go-graph-002
  Scenario: Graph plans retain private snapshots and cannot cross owners
    Given a planned addition or retarget on an immutable retained-package snapshot
    When caller-owned addition bytes change after planning
    Then applying the plan still uses the planned bytes
    And applying the plan to another package or reapplying a consumed nonempty plan refuses

  @candidate-go-graph-003
  Scenario: Intervening changes invalidate plans even if payload bytes return to their old value
    Given a graph plan held while a payload changes and is restored
    When the earlier plan is applied
    Then it returns a stale-target refusal
    And no-op plans remain repeatable without changing original archive bytes

  @candidate-go-graph-004
  Scenario: Invalid mixed additions do not stage partial state
    Given a graph mutation containing a valid addition and a duplicate or unsafe part identity
    When planning validates the mutation
    Then it refuses before either addition is committed
    And signed-package mutation also refuses
    And a valid empty payload can be added when its identity and content type are proved

  @candidate-go-graph-005
  Scenario: Planned additions and retargets preserve untouched archive members
    Given an existing shared target and a new private target
    When a selected relationship is retargeted through a graph plan
    Then other relationships and the shared original remain unchanged
    And untouched member metadata and compressed bytes are retained
    And repeated serialization yields identical output

  @candidate-go-graph-006
  Scenario: Receipts describe the accumulated graph delta
    Given an added part that is subsequently replaced
    When the package is saved and reopened
    Then its receipt uses schema2 with an add operation and final payload fingerprint
    And chained additions and retargets accumulate the measured changed-part count
    And a retarget-only delta uses the replacement-only receipt schema
    And restoring the original target can restore exact no-op bytes

  @candidate-go-graph-007
  Scenario: Transient additions disappear completely when later deleted
    Given a newly added part that has not existed in the original archive
    When it is replaced and then deleted through successive graph plans
    Then the delivered archive and empty change receipt match the original
    And the transient part can no longer be read

  @candidate-go-graph-008
  Scenario: Refused previews preserve earlier edits and held valid plans
    Given prior committed edits and a valid unconsumed deletion plan
    When a different mixed preview includes a stale replacement
    Then preview returns an error and no private plan
    And prior bytes and the held valid plan remain usable

  @candidate-go-graph-009
  Scenario: Original deleted identities stay reserved and stale older plans
    Given an original detached payload deleted through a graph plan
    When an earlier plan, replacement or same-name addition is attempted
    Then the earlier plan is stale and the deleted payload cannot be replaced
    And neither exact nor case-colliding original names can be reused

  @candidate-go-graph-010
  Scenario: Combined retarget removal and leaf deletion leave a coherent graph
    Given a shared target with multiple inbound edges
    When one edge is retargeted, another removed and the now-detached leaf deleted atomically
    Then only the intended new edge survives and the deleted part is absent after reopen
    And deleting a relationship-owning nonleaf or replacing a registry alongside graph edits refuses
