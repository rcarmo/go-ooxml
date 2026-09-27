@implemented @go @package
Feature: Lossless namespace-aware XML leaf edits
  XML inspection identifies expanded names while edits retain unrelated source bytes.

  @XML-001
  Scenario: Editing a text leaf preserves unknown markup and namespace context
    Given XML with prefixed text and an opaque extension
    When I replace the selected text leaf with "A&B <new>"
    Then only the leaf character-data bytes change
    And the replacement decodes to "A&B <new>"

  @XML-002
  Scenario Outline: Namespace-invalid XML is rejected
    Given XML with "<defect>"
    When I parse the XML for lossless editing
    Then parsing refuses without changing the source

    Examples:
      | defect |
      | undeclared element prefix |
      | undeclared attribute prefix |
      | duplicate expanded attribute |
      | invalid xml binding |
      | mismatched end prefix |

  @XML-003
  Scenario: A foreign target refuses the entire edit batch
    Given XML with prefixed text and an opaque extension
    When an edit batch contains a target from another parsed document
    Then the batch refuses and a valid target remains reusable
