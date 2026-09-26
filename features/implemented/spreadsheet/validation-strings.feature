@implemented @go @spreadsheet
Feature: Validation string sources preserve original shared-string identities
  Shared-string indices refer to source entries in order, including duplicates.

  @VOCABSTR-001
  Scenario Outline: Stored strings resolve without deduplicating shared entries
    Given a validation workbook with "<condition>"
    When I inspect its allowed values
    Then the expected validation vocabulary is returned without mutation

    Examples:
      | condition |
      | shared strings with duplicate entries |
      | rich shared string |
      | rich inline string |

  @VOCABSTR-002
  Scenario Outline: Unproved string sources refuse instead of returning empty text
    Given a validation workbook with "<condition>"
    When I inspect its allowed values
    Then validation inspection refuses and returns no partial values

    Examples:
      | condition |
      | out of range shared index |
      | missing shared string relationship |
      | duplicate shared string relationships |
      | mixed shared string structure |
      | missing rich run text |
