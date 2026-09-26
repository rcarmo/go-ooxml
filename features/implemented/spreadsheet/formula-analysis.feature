@implemented @go @spreadsheet
Feature: Static formula dependency analysis
  Reference analysis distinguishes literal strings from actual A1 references and refuses unknown dependencies.

  @FORMULA-001
  Scenario Outline: Static expressions expose exact reference ranges
    Given the formula expression "<formula>"
    When I analyse its static references
    Then exactly <count> reference ranges are reported

    Examples:
      | formula                     | count |
      | Input!A1*2                  | 1     |
      | SUM($A$1:B3)+C$4            | 2     |
      | 'Input Data'!$B2+Sheet2!C3   | 2     |
      | 1+2*3                       | 0     |

  @FORMULA-002
  Scenario Outline: Unproved dynamic or extended syntax refuses analysis
    Given the formula expression "<formula>"
    When I analyse its static references
    Then dependency analysis refuses instead of returning a partial answer

    Examples:
      | formula                  |
      | INDIRECT(A1)             |
      | OFFSET(A1,1,1)           |
      | Sheet1:Sheet3!A1         |
      | Table1[Column]           |
      | [Book.xlsx]Sheet1!A1     |
      | A1#                      |
      | A1 B2                    |
      | A1+                      |
