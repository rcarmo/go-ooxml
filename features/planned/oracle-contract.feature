@planned @go @spreadsheet
Feature: Calculation adapter failure contracts
  Adapter failures are testable without an installed calculation engine.

  @ORACLE-002
  Scenario: An unavailable configured oracle returns a typed refusal
    Given a valid supported workbook
    And an oracle configuration pointing to a deliberately nonexistent executable
    When I request recalculation
    Then I receive an oracle-unavailable refusal
    And the source workbook remains byte-identical
    And no output file is delivered
