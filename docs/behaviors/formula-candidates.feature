@catalogue @not-executed
Feature: Static formula references and insertion remapping
  Candidate outcomes for central reconciliation. These scenarios are not runner bindings.

  @candidate-go-formula-001
  Scenario: Static analysis retains exact reference locations
    Given a supported expression containing cell references, ranges, quoted sheet names and string literals
    When the native analyser identifies static dependencies
    Then each reference has a nonempty source byte span and bounded row and column coordinates
    And absolute row and column flags and decoded quoted sheet names are retained
    And reference-looking string contents do not become dependencies

  @candidate-go-formula-002
  Scenario: Unsupported expressions return no partial dependency list
    Given an expression containing an unresolved name, unsupported function, external or three-dimensional reference, structured reference, union, intersection, array, spill, error token or malformed syntax
    When the native analyser identifies static dependencies
    Then analysis returns an error and no references

  @candidate-go-formula-003
  Scenario: String punctuation cannot act as formula punctuation
    Given a string token whose contents are a formula operator or delimiter
    When the token occurs in an otherwise invalid expression position
    Then analysis refuses without references
    But when the string occurs as an ordinary function argument or concatenation operand
    Then its punctuation remains string data and the supported expression is accepted

  @candidate-go-formula-004
  Scenario: Insertion changes only affected references
    Given an expression with local references, an unrelated sheet reference and reference-looking string text
    When rows or columns are inserted into the named worksheet through the static remap primitive
    Then affected coordinates shift or the referenced range expands
    And worksheet identity matches without regard to case
    And unrelated reference tokens and string text remain byte-identical
    And the resulting expression is accepted by the static analyser

  @candidate-go-formula-005
  Scenario: Invalid or overflowing remaps return no replacement expression
    Given an insertion with an invalid axis, blank worksheet, nonpositive coordinate or count, or an out-of-grid result
    When the insertion remapper processes the expression
    Then it returns an error and an empty replacement expression
    And unsupported dynamic references also refuse without a partial replacement

  @candidate-go-formula-006
  Scenario: Direct ranges distinguish cells and whole axes
    Given a direct cell, rectangular range, whole-row range or whole-column range with optional absolute flags and quoted sheet name
    When the direct-range parser reads the source
    Then it returns the decoded worksheet and bounded endpoints
    And the missing whole-axis coordinate is zero for the caller to bound
    And general expression analysis does not gain whole-axis support

  @candidate-go-formula-007
  Scenario: A direct-range parser refuses expressions and ambiguous endpoints
    Given a bare name or number, mixed endpoint kinds, external or three-dimensional reference, multi-area or arithmetic expression, spill syntax or out-of-grid endpoint
    When the direct-range parser reads the source
    Then it returns an error

  @candidate-go-formula-008
  Scenario: Analysed references have safe source spans for arbitrary bounded input
    Given a formula byte string within the native fuzz input bound
    When the analyser processes the input
    Then refusal returns no references
    And every accepted reference lies within the input and worksheet grid
    And accepted reference spans are nonempty
