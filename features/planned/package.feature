@planned @go @package
Feature: Preservation-safe Office package edits
  Unrelated content survives an edit, including content the editor cannot interpret.

  @PKG-001
  Scenario Outline: Untouched package members retain their payload bytes
    Given an existing <format> package with an opaque custom XML part
    And a supported text or cell target in a different part
    When I apply one bounded edit and save to a new path
    Then every unrelated member has its original payload hash
    And the custom XML part remains reachable through its original relationship
    And the saved package reopens with the requested change

    Examples:
      | format |
      | DOCX   |
      | PPTX   |
      | XLSX   |

  @PKG-002
  Scenario: Unknown content inside an edited XML part survives
    Given an existing document with an unknown extension beside editable content
    When I change only that content and save to a new path
    Then the unknown extension and its namespace context remain intact
    And the package diff stays within the permitted changed-part budget

  @PKG-003
  Scenario: An unsafe archive is refused before editing
    Given an Office archive with duplicate member names
    When I open it for preservation-safe editing
    Then I receive an invalid-package refusal
    And the source archive remains byte-identical
    And no output file is created

  @PKG-004
  Scenario: A failed delivery leaves an existing destination untouched
    Given an edited package ready for delivery
    And a destination containing an earlier document
    When archive finalisation fails before replacement
    Then delivery reports the write failure
    And the destination remains byte-identical to the earlier document
    And no temporary output remains
