@implemented @go @spreadsheet
Feature: Validation vocabularies reject unknown types and non-decimal scalars
  Literal lists share the finite-range result limit. Inspection preserves the package.

  @VOCABLEX-001
  Scenario Outline: Declared defaults and bounded decimal values are readable
    Given a validation workbook with "<condition>"
    When I inspect its allowed values
    Then the expected validation vocabulary is returned without mutation

    Examples:
      | condition |
      | omitted validation type |
      | explicit none validation type |
      | decimal scalar spellings |
      | literal at vocabulary limit |

  @VOCABLEX-002
  Scenario Outline: Unknown types and unproved scalar spellings refuse
    Given a validation workbook with "<condition>"
    When I inspect its allowed values
    Then validation inspection refuses and returns no partial values

    Examples:
      | condition |
      | unknown validation type |
      | empty validation type |
      | uppercase validation type |
      | hexadecimal stored number |
      | underscored stored number |
      | literal exceeds vocabulary limit |
