@implemented @go @spreadsheet
Feature: Whole-axis validation sources use stored worksheet extents
  Worksheet dimension hints do not authorise inferred cells or unbounded allocation.

  @VOCABAXIS-001
  Scenario Outline: One-dimensional whole-axis ranges resolve in source order
    Given a validation workbook with "<condition>"
    When I inspect its allowed values
    Then the expected validation vocabulary is returned without mutation

    Examples:
      | condition |
      | whole column |
      | whole row reversed columns |
      | quoted cross-sheet whole column |
      | whole column empty sheet |
      | whole column ignores dimension hint |

  @VOCABAXIS-002
  Scenario Outline: Unproved or excessive whole-axis sources refuse
    Given a validation workbook with "<condition>"
    When I inspect its allowed values
    Then validation inspection refuses and returns no partial values

    Examples:
      | condition |
      | whole multiple columns |
      | bare column name |
      | mixed range endpoint types |
      | excessive whole-column result |
