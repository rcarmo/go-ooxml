@catalogue @not-executed
Feature: Lossless namespace-aware XML inspection and bounded edits
  Candidate outcomes for central reconciliation; no steps execute from this catalogue.

  @candidate-go-xml-001
  Scenario: Namespace-aware parsing retains original lexical bytes
    Given XML with scoped namespace declarations, comments, CDATA, entities and self-closing elements
    When the native parser produces an immutable element snapshot
    Then element and attribute names use their expanded namespace identities
    And decoded text is available without changing the retained source bytes
    And an empty edit or identical-value edit returns the exact original bytes

  @candidate-go-xml-002
  Scenario: Invalid namespace identities and document boundaries refuse
    Given XML with an undeclared prefix, duplicate expanded attribute, invalid reserved namespace, mismatched qualified tag, multiple roots or document type declaration
    When the native parser reads the document
    Then parsing returns an error
    And character references or CDATA outside the root also refuse even if their decoded content is whitespace

  @candidate-go-xml-003
  Scenario: XML whitespace normalisation retains byte offsets
    Given literal tabs or line endings in attributes, numeric whitespace references and CR or CRLF in text or CDATA
    When the parser decodes attribute and text values
    Then literal attribute whitespace becomes spaces and numeric whitespace references retain their characters
    And text and CDATA line endings become LF
    And replacing a text leaf changes its original byte span without changing neighbouring line endings

  @candidate-go-xml-004
  Scenario: Text-leaf replacement rejects ambiguous targets atomically
    Given an immutable document and text edits targeting mixed content, comment-containing content, a duplicate target, invalid XML characters or a foreign handle
    When the edits are planned together
    Then the edit returns an error without changing the snapshot
    And nonempty text cannot be assigned through a self-closing leaf splice

  @candidate-go-xml-005
  Scenario: Attribute edits preserve untouched spelling and namespace scope
    Given an existing attribute written with spacing and a chosen quote character
    When its value changes or an unqualified attribute is added
    Then existing quote style and unrelated lexical bytes remain unchanged
    And XML characters in the new value are escaped
    And namespace declarations, undeclared qualified names and duplicate conflicting edits refuse

  @candidate-go-xml-006
  Scenario: Structured children preserve expanded names across namespace environments
    Given a parent with a default namespace, existing prefixes or no namespace
    When structured child elements and attributes with requested expanded names are inserted
    Then each inserted element and attribute has its requested namespace and local name after reparsing
    And text content is retained and conflicting prefixes do not change existing names
    And an unqualified child resets an inherited default namespace when needed

  @candidate-go-xml-007
  Scenario: Invalid structured insertion batches return no partial document
    Given an insertion with an invalid name or character, foreign parent, duplicate parent or overlapping parent targets
    When the insertion batch is applied
    Then it returns an error and no partial XML bytes
    And the original immutable snapshot remains reusable

  @candidate-go-xml-008
  Scenario: Element removal preserves non-target lexical regions
    Given disjoint non-root elements separated by comments and whitespace
    When those elements are removed
    Then only their original element spans disappear
    And root, duplicate, overlapping, invalid and foreign targets refuse
    And an empty removal preserves exact source bytes

  @candidate-go-xml-009
  Scenario: Structured subtree replacement uses the surviving parent scope
    Given a non-root element whose own namespace declarations are being removed
    When structured replacement elements are rendered in its parent
    Then requested expanded names remain correct without relying on removed declarations
    And multiple replacement nodes or deletion affect only the selected span
    And root, duplicate, overlapping, foreign or invalid-content replacements return no partial XML
