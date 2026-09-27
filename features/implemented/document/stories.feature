@implemented @go @document
Feature: Read-only Word story projections
  Inspection includes related stories and nested paragraphs without double-counting.

  @STORY-001
  Scenario Outline: Revision view selects the appropriate visible text
    Given a Word document with body revisions and a related header
    When I inspect its "<view>" story view
    Then the body paragraph reads "<text>"
    And the related header paragraph reads "Header text"
    And story inspection leaves the source archive unchanged

    Examples:
      | view     | text          |
      | current  | base new      |
      | original | base old      |
      | all      | base oldnew   |

  @STORY-002
  Scenario: Nested text box paragraphs do not duplicate their text in the parent
    Given a Word document with a nested text box paragraph
    When I inspect its "current" story view
    Then the outer paragraph reads "Outer"
    And one text-box paragraph reads "Inner"
