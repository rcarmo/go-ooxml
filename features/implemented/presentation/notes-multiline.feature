@implemented @go @presentation
Feature: Multiline notes preserve the first paragraph and run formatting template
  Empty lines become empty paragraphs in the existing notes body only.

  @NOTELINES-001
  Scenario Outline: Replace ordinary notes paragraphs with templated lines
    Given multiline notes condition "<condition>"
    When I replace the existing notes with four lines including an empty line
    Then reopened notes contain exactly the four requested lines
    And each new paragraph keeps the first paragraph formatting
    And each nonempty line keeps the first run formatting without formatting from later runs
    And all non-body notes XML and other package payloads stay unchanged

    Examples:
      | condition |
      | multiple runs and paragraphs |
      | empty existing paragraph |
      | prefixed template namespaces |

  @NOTELINES-002
  Scenario: Clearing multiline notes creates one empty paragraph then permits refill
    Given multiline notes condition "multiple runs and paragraphs"
    When I clear and refill the existing notes body
    Then the cleared body is one empty paragraph and refill has the requested text

  @NOTELINES-003
  Scenario Outline: Unproved formatting and late invalid text refuse before paragraph removal
    Given multiline notes condition "<condition>"
    When I attempt an unsupported multiline notes edit
    Then the notes snapshot and package remain unchanged

    Examples:
      | condition |
      | hyperlink in later run |
      | unknown paragraph property |
      | invalid character after newline |
