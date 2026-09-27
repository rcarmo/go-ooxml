@implemented @go @presentation
Feature: Guarded edits to an existing notes body
  Reading and replacing notes never creates a missing notes graph or changes other placeholders.

  @NOTES-001
  Scenario: Replace a plain existing notes body and retain every other payload
    Given a presentation with notes condition "ordinary body"
    When I find the existing notes body
    Then the notes text is "Original notes"
    When I replace that notes body with "Updated notes"
    Then only the notes body text bytes change after delivery
    And the old notes target refuses reuse

  @NOTES-002
  Scenario: An empty existing notes text leaf can be filled without other changes
    Given a presentation with notes condition "empty text leaf"
    When I find the existing notes body
    Then the notes text is ""
    When I replace that notes body with "Updated notes"
    Then only the notes body text bytes change after delivery

  @NOTES-005
  Scenario: A self-closing text leaf can be filled through structural authoring
    Given a presentation with notes condition "self-closing text leaf"
    When I find the existing notes body
    Then the notes text is ""
    When I replace that notes body with "Updated notes"
    Then only the notes body text bytes change after delivery

  @NOTES-004
  Scenario: Grouping-only notes shape lock does not prevent text correction
    Given a presentation with notes condition "grouping-only lock"
    When I find the existing notes body
    Then the notes text is "Original notes"
    When I replace that notes body with "Updated notes"
    Then only the notes body text bytes change after delivery

  @NOTES-003
  Scenario Outline: Unproved notes selection or replacement refuses atomically
    Given a presentation with notes condition "<condition>"
    When I attempt a guarded notes replacement
    Then notes replacement refuses and the archive is unchanged

    Examples:
      | condition |
      | absent notes |
      | duplicate body placeholder |
      | shared notes part |
      | field in body |
      | locked body |
      | malformed grouping lock |
      | bad replacement character |
      | tab replacement |
