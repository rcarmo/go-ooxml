@catalogue @not-executed
Feature: Bounded ZIP intake and retained archive delivery
  Candidate outcomes for central reconciliation; no catalogue steps execute.

  @candidate-go-archive-001
  Scenario: Configured intake budgets produce typed resource refusals
    Given a package exceeding its source-byte, entry-count, per-part or total-inflated-byte budget
    When the bounded archive reader opens the source
    Then it returns a resource-limit refusal
    And negative caller budgets are rejected as invalid arguments

  @candidate-go-archive-002
  Scenario: Payload CRC corruption refuses without source changes
    Given a stored archive payload altered without updating its CRC
    When the package reader verifies the archive
    Then it returns an invalid-package refusal
    And directory entries are never exposed as ordinary parts in a valid archive
    And an empty binary part remains a valid ordinary part

  @candidate-go-archive-003
  Scenario: Local and central ZIP metadata must identify the same entry
    Given mismatched local and central names, compression methods or flags, or an offset redirected to another local header
    When the archive reader validates physical entry metadata
    Then it refuses the archive
    And a prepended non-archive prefix is refused even if central offsets were adjusted

  @candidate-go-archive-004
  Scenario: Ambiguous descriptor signatures still require payload integrity
    Given an unsigned data descriptor whose CRC equals the descriptor signature value
    When descriptor geometry is inspected
    Then the geometry can be recognised without assuming the signature is present
    But opening a fixture with a deliberately incorrect payload CRC still refuses

  @candidate-go-archive-005
  Scenario: Valid ZIP64 variants retain byte-identical no-op delivery
    Given a valid ZIP64 archive using store or deflate and supported signed or unsigned descriptor forms
    When it is opened as a retained package and saved without edits
    Then both stream and file delivery preserve exact source bytes
    And an ordinary stored empty member remains readable

  @candidate-go-archive-006
  Scenario: Malformed ZIP64 fields cannot borrow or invent bounds
    Given missing or truncated ZIP64 extras, inconsistent sizes or offsets, malformed descriptors, duplicate size metadata or multidisk values
    When the reader validates the archive
    Then it refuses the archive
    And a required ZIP64 field cannot consume bytes from the next unrelated extra record
    And gaps or invalid lengths around ZIP64 terminal records also refuse

  @candidate-go-archive-007
  Scenario: Individually readable members cannot overlap physical ranges
    Given distinct named members whose headers and payload CRCs are readable by the standard ZIP reader but whose physical ranges overlap
    When the retained package reader validates member extents
    Then it refuses the overlap
    And the caller's source buffer remains byte-identical

  @candidate-go-archive-008
  Scenario: Edited ZIP64 delivery preserves unrelated payloads and reopens
    Given a retained ZIP64 package with a selected XML payload
    When a fingerprinted replacement is saved and reopened
    Then the selected payload contains the replacement and untouched payloads are unchanged
    And a separately planned new part remains readable after graph delivery

  @candidate-go-archive-009
  Scenario: Writer entry-count boundary emits a readable ZIP64 archive
    Given a package whose parts and registry produce 65535 ZIP entries
    When the package is written and reopened at that allowed entry limit
    Then all entries are readable and retained no-op delivery preserves bytes
    And reopening with a lower entry limit returns a resource-limit refusal

  @candidate-go-archive-010
  Scenario: Every enrolled corpus archive preserves exact no-op bytes
    Given the enrolled owned document corpus
    When each archive is opened and written through retained-source delivery without edits
    Then every output is byte-identical to its input
    And an empty or missing corpus fails verification
