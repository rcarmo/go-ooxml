@implemented @go @spreadsheet
Feature: Bounded formula-free numeric cell correction
  Existing numeric cells can change only when dependent structures are absent.

  @CELL-001
  Scenario: Changing a formula-free numeric cell preserves styles and opaque parts
    Given a formula-free workbook with numeric cell "B2"
    When I set the guarded numeric cell to 125
    Then delivery changes only the cell value bytes
    And the consumed cell target refuses reuse

  @CELL-002
  Scenario Outline: Unproved workbook dependencies refuse numeric writes
    Given a workbook containing "<condition>"
    When I attempt the guarded numeric correction
    Then the numeric correction returns a typed refusal
    And saving the numeric session preserves the original archive

    Examples:
      | condition |
      | formula on another sheet |
      | protected worksheet |
      | defined name |
      | duplicate cell address |
      | chart dependency |
      | data validation |
      | unknown worksheet element |
      | unknown cell attribute namespace |
