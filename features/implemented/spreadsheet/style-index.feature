@implemented @go @spreadsheet
Feature: Numeric edits require resolvable cell styles
  An omitted cell style attribute still selects style zero when a style table exists.

  @STYLE-001
  Scenario Outline: Invalid style references refuse numeric edits unchanged
    Given a numeric workbook with "<condition>"
    When I attempt a numeric edit under the style policy
    Then style validation refuses with the original archive unchanged

    Examples:
      | condition |
      | empty cellXfs and omitted s |
      | empty cellXfs and explicit zero |
      | out of range style |
      | mismatched cellXfs count |
      | malformed style index |
      | unrelated cell with invalid style |

  @STYLE-002
  Scenario: Implicit style zero resolves through a nonempty table
    Given a numeric workbook with "one style and omitted s"
    When I set its style-validated numeric value to 125
    Then the value changes without adding a style attribute or modifying styles
