@implemented @go @package
Feature: Caller-defined package intake budgets
  Budget violations are refused before payload allocation or decompression.

  @LIMIT-001
  Scenario Outline: Intake enforces a configured resource ceiling
    Given an archive exceeding its "<budget>" budget
    When I open it with the configured resource limits
    Then I receive a typed resource-limit refusal

    Examples:
      | budget |
      | source |
      | entries |
      | part |
      | total |

  @LIMIT-002
  Scenario: Finite sufficient budgets permit normal intake
    Given an archive within all configured budgets
    When I open it with the configured resource limits
    Then its XML payload is available without changes
