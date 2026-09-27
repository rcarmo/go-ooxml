@catalogue @not-executed
Feature: Native coordinate colour unit and XML utilities
  Candidate outcomes for central reconciliation; these legacy utilities are separate from conservative formula and retained XML APIs.

  @candidate-go-utilities-001
  Scenario: Column numbering round-trips through spreadsheet letters
    Given positive column numbers including single and multiple letter boundaries
    When numbers are converted to letters and letters back to numbers
    Then the original positive number is retained
    And lowercase letters resolve to the same column
    And nonpositive numbers produce an empty label while an empty label produces zero

  @candidate-go-utilities-002
  Scenario: Cell reference parsing and rendering preserve addressing fields
    Given a cell reference with optional worksheet and absolute row or column flags
    When the utility parses and renders that reference
    Then worksheet, coordinates and absolute flags are retained
    And empty, coordinate-incomplete or zero-row examples return errors

  @candidate-go-utilities-003
  Scenario: Range utility requires two endpoints and includes its boundaries
    Given a two-endpoint range with optional absolute flags and worksheet prefix
    When the range utility parses it and checks member coordinates
    Then the corner and interior examples are included and outside coordinates excluded
    And a bare cell or empty input is refused by this range utility

  @candidate-go-utilities-004
  Scenario: Hex colours retain channel values and alpha policy
    Given six-digit RGB or eight-digit ARGB text with optional hash prefix or lowercase letters
    When colour parsing and serialisation run
    Then channels match the encoded bytes and output uses uppercase hexadecimal
    And opaque ToHex output omits alpha while translucent output includes it
    And ToARGB always includes alpha
    And malformed or wrong-length examples return errors

  @candidate-go-utilities-005
  Scenario: Drawing units convert at defined scale factors
    Given an inch, point, centimetre, 96-DPI pixel, twip or half-point measurement
    When the corresponding conversion utility runs
    Then it uses 914400 EMUs per inch, 12700 per point, 360000 per centimetre, 9525 per pixel and 635 per twip
    And 24 half-points represents 12 points
    And the covered zero and multi-unit examples retain their expected conversions

  @candidate-go-utilities-006
  Scenario: Validation errors retain field and optional offending value
    Given a field, validation message and either a string, number or absent value
    When a validation error is formatted
    Then its field and message remain accessible
    And the message includes the offending value only when supplied
    And the standard sentinel error set contains non-nil errors with nonempty messages

  @candidate-go-utilities-007
  Scenario: XML authoring helpers supply declaration and requested indentation
    Given a simple or empty XML-serialisable structure
    When the XML helper serialises it
    Then the output begins with an XML declaration
    And the declaration constant names version 1.0, UTF-8 and standalone yes followed by LF
    And the indented variant includes the requested indentation

  @candidate-go-utilities-008
  Scenario: XML decoding handles a UTF-8 BOM and refuses malformed input
    Given well-formed XML with or without a UTF-8 BOM
    When the decoding helper reads its value element
    Then the decoded text matches the input value
    And an unclosed XML example returns an error

  @candidate-go-utilities-009
  Scenario: XML escaping preserves text through explicit entities
    Given plain text, markup delimiters, ampersands, quotes, LF or empty text
    When the text escape helper encodes it
    Then delimiters and ampersands are escaped, quotes use numeric references and LF uses its character reference
    And plain or empty text retains its content

  @candidate-go-utilities-010
  Scenario: Pointer and dereference helpers retain supplied scalar values
    Given boolean, integer, int64, string or float64 values
    When pointer helpers wrap them
    Then the returned pointers are non-nil and contain the supplied values
    And supported dereference helpers return existing values or the caller default for nil pointers

  @candidate-go-utilities-011
  Scenario: Successful utility fuzz inputs retain round-trip fields
    Given seed or fuzz inputs for bounded column numbers, cell references or two-endpoint ranges
    When parsing succeeds and the result is rendered and parsed again
    Then column identity, all cell-address fields and both range endpoints are retained
    And accepted cell and range coordinates are positive
    And this invariant does not claim an exploratory fuzz campaign has run
