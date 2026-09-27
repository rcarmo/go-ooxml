@implemented @go @spreadsheet
Feature: Formula punctuation is distinct from quoted string content
  Operator-looking text inside string literals never becomes grammar punctuation.

  @TOKEN-001
  Scenario: A quoted closing parenthesis is a literal function argument
    Given a static formula with a quoted closing-parenthesis argument
    When I analyse its static references
    Then literal punctuation is accepted without hiding cell references

  @TOKEN-002
  Scenario: Adjacent quoted operator text cannot stand in for a binary operator
    Given a formula with two references separated only by a quoted plus sign
    When I analyse its static references
    Then dependency analysis refuses instead of returning a partial answer
