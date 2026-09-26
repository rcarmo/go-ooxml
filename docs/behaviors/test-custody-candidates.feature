@catalogue @not-executed
Feature: Native test-input identity and reference custody
  Candidate outcomes for central reconciliation; fixture and helper checks do not create document-format execution credit.

  @candidate-go-test-custody-001
  Scenario: Content identity resolves a manifest-selected owned input
    Given a schema2 manifest fixture with a content ID, SHA-256, byte size and format/scenario-group metadata
    When native fixture lookup resolves the content ID
    Then it returns the canonical fixture path only when stored bytes match the identity and size
    And the physical filename and group depth do not need to match a consumer origin layout

  @candidate-go-test-custody-002
  Scenario: Invalid manifest identities and paths refuse
    Given a fixture record with a missing ID, duplicated ID, mismatched identity, unsafe path, wrong size or missing group metadata
    When native fixture lookup resolves the requested input
    Then lookup returns an error
    And altered file bytes are not accepted as the original fixture

  @candidate-go-test-custody-003
  Scenario: Output paths cannot redirect writes into shared references
    Given the shared reference root and a proposed output path
    When the output-custody guard resolves existing ancestors
    Then root paths, descendants and symlink descendants are refused
    And an outside path or a sibling with a similar prefix is allowed

  @candidate-go-test-custody-004
  Scenario: Released references require the exact annotated identity and seals
    Given a clean checkout matching the recorded reference commit
    When release verification checks the annotation, peeled commit and manifest seals
    Then the exact annotated release is accepted
    And wrong commits, lightweight or mismatched tags, retargeted annotations, missing Git metadata and altered seals refuse

  @candidate-go-test-custody-005
  Scenario: Dirty facts and workflows cannot hide behind valid fixture hashes
    Given unchanged asset seals but modified, staged or deleted tracked facts or workflows, or an untracked fact
    When the full reference checkout is verified
    Then verification refuses
    And assume-unchanged does not hide altered tracked bytes
    And verification itself leaves checkout bytes unchanged

  @candidate-go-test-custody-006
  Scenario: Candidate verification does not claim a tagged release
    Given an explicit candidate pin with an empty annotation and a candidate-labelled tag
    When verification checks its exact commit, seals and tracked checkout
    Then a clean matching candidate is accepted and a dirty candidate is refused
    And the root-module pin test requires an explicit candidate root when the candidate override is selected

  @candidate-go-test-custody-007
  Scenario: Temporary-path helpers do not imply that an output already exists
    Given a native test helper and a requested temporary filename
    When the helper resolves that filename
    Then the returned path is nonempty and the file does not yet exist

  @candidate-go-test-custody-008
  Scenario: Successful assertion helper calls return without reporting failure
    Given equal integers, true and false conditions and text containing the expected substring
    When the corresponding native assertion helpers receive those values
    Then the calls complete without a test failure
    And this characterisation does not cover failing assertion diagnostics

  @candidate-go-test-custody-009
  Scenario: Helper labels describe format options and document types
    Given supported bold, italic, size and colour combinations or one of the three Office document types
    When helper labels are requested
    Then formatting labels use the expected plain or combined tokens
    And Word, Excel and PowerPoint have their corresponding names and file extensions

  @candidate-go-test-custody-010
  Scenario: Shared test-case collections are populated
    Given the common native string, numeric, formatting, cell-reference and range collections
    When their entries are inspected
    Then each collection has at least one entry

  @candidate-go-test-custody-011
  Scenario: Resource helpers return an open object from successful factories
    Given a successful resource factory or path-based opener
    When the matching native resource helper invokes it
    Then it returns a non-nil resource that has not yet been closed
    And this assertion does not establish later cleanup timing or error-path behaviour
