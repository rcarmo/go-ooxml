@implemented @go @spreadsheet
Feature: Dependency analysis validates worksheet structure
  A static expression is not enough when its cell ownership or input value structure is ambiguous.

  @DEPEND-001
  Scenario Outline: Malformed non-formula cells refuse the complete dependency plan
    Given a dependency workbook with "<condition>"
    When I attempt an explicit invalidating input edit
    Then invalidation returns a typed refusal without any package mutation

    Examples:
      | condition |
      | duplicate unrelated value |
      | unclassified cell metadata |
      | duplicate sheetData containers |
