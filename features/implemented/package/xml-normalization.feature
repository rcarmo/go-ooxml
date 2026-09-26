@implemented @go @package
Feature: XML line endings and attribute values retain original edit offsets
  Decoded values follow XML normalization without reserializing untouched source bytes.

  @NORM-001
  Scenario: Literal line endings normalize but no-op and edits use original offsets
    Given XML text containing mixed CRLF and CR line endings
    When I inspect its decoded character data
    Then literal line endings are LF while numeric CR remains CR
    And a decoded-text no-op preserves the original bytes
    And a changed text leaf preserves surrounding original line endings

  @NORM-002
  Scenario: Literal attribute whitespace differs from numeric character references
    Given an XML attribute with literal whitespace and numeric whitespace references
    When I inspect its decoded attribute value
    Then literal attribute whitespace is spaces and numeric references retain their characters
    And a decoded-attribute no-op preserves its original lexical bytes

  @NORM-003
  Scenario: Changed attribute whitespace round-trips without losing numeric characters
    Given an XML attribute with literal whitespace and numeric whitespace references
    When I replace the attribute with tab CR and LF characters
    Then the changed attribute decodes to the requested characters
    And unrelated attributes and start-tag spacing remain byte-identical
