@implemented @go @spreadsheet
Feature: Structural reference remapping uses parsed formula tokens
  Static reference edits preserve quoted literals and unrelated sheet references.

  @REMAP-001
  Scenario Outline: Row and column insertion follows structural coordinates
    Given a formula "<formula>" located on "<owner>"
    When I insert <count> "<axis>" before coordinate <at> on "<sheet>"
    Then the remapped formula is "<expected>"

    Examples:
      | formula                  | owner | count | axis   | at | sheet | expected                 |
      | $A$1+A2+Other!A2          | Main  | 2     | row    | 2  | Main  | $A$1+A4+Other!A2          |
      | SUM(Main!A1:B3)           | Other | 1     | row    | 2  | Main  | SUM(Main!A1:B4)           |
      | Main!$B2+C$4              | Main  | 2     | column | 2  | Main  | Main!$D2+E$4             |
      | 'Input Data'!A2+1         | Main  | 3     | row    | 2  | Input Data | 'Input Data'!A5+1     |

  @REMAP-002
  Scenario: A remap beyond the worksheet grid refuses without a partial formula
    Given a formula "A1048576+B1" located on "Main"
    When I attempt a row insertion that shifts it beyond the grid
    Then remapping refuses and returns no partial expression
