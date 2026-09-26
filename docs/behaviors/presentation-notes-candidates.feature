@catalogue @not-executed
Feature: Native presentation notes preservation and template predicates
  Two native declarations describe pinned-fixture edits and synthetic paragraph variants.
  Byte comparisons and model readback remain separate from schema or Office rendering checks.

  @candidate-go-presentation-notes-001
  Scenario: Existing notes change only the selected text splice and survive disk reopen
    Given the hash-pinned notes fixture and its known Gothic-elements speaker text
    When the slide notes are replaced with Updated speaker notes and saved
    Then the notes XML equals one byte replacement of the prior text
    And the receipt contains exactly the notes part as its only change
    And each other original graph part has identical delivered payload bytes
    And reopened FindNotes returns the new text without a rendering assertion

  @candidate-go-presentation-notes-002
  Scenario: Refusal and identical text preserve source before clearing and refilling notes
    Given a held notes target and a separate session opened from the same fixture
    When a foreign-target write and NUL invalid-UTF8 and tab inputs are attempted
    Then those calls error and an identical-text call succeeds with exact original archive bytes afterward
    And clearing through the original target succeeds while another held target refuses as stale
    And a fresh target reads empty text and accepts refilled text
    And no exact refusal kind or disk readback of the refill is asserted

  @candidate-go-presentation-notes-003
  Scenario: A direct low-level notes edit invalidates a held target
    Given a target found before the retained notes part changes Gothic to Victorian
    When ReplaceNotes is called with the original target
    Then it returns an error
    And no exact error kind or post-refusal byte equality is asserted in this subtest

  @candidate-go-presentation-notes-004
  Scenario: A self-closing notes text node can be filled
    Given a fixture copy whose first paragraph is replaced with a bold run and self-closing text node
    When its notes are replaced with filled
    Then a fresh FindNotes returns "filled"
    And this subtest does not independently assert retained bold formatting or serialized archive bytes

  @candidate-go-presentation-notes-005
  Scenario: Edge whitespace authors the XML namespace preservation attribute
    Given a fixture copy with one old text leaf marked xml:space default
    When notes are replaced with a leading and trailing space around leading
    Then parsed notes XML contains the exact spaced text and an XML-namespace space attribute valued preserve
    And the assertion does not measure rendered whitespace

  @candidate-go-presentation-notes-006
  Scenario: Multirun no-op is exact and later run formatting is not borrowed
    Given a notes paragraph with an unformatted first run and a bold second run
    When identical notes text is submitted
    Then the notes-part bytes remain identical
    And a later replacement with boundary newlines spaced A-and-B and a Unicode snow character reads back exactly as four paragraphs
    And the entire notes part contains no literal bold-one attribute
    And no independent alignment or unrelated-placeholder assertion is added by this subtest

  @candidate-go-presentation-notes-007
  Scenario: Two-line replacement repeats four selected template fragments
    Given a notes paragraph with bullet default run size colour font and end-paragraph language properties
    When notes are replaced by a newline b
    Then each exact colour112233 escaped-fontF-and-F languageen-US and bullet-character fragment occurs twice in the notes XML
    And only these four fragment counts are checked without structural placement default-size or rendering assertions
