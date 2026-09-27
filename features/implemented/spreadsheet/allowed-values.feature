@implemented @go @spreadsheet
Feature: Deterministic list validation inspection
  Values are read without creating cells, calculating formulas or mutating source XML.

  @VOCAB-001
  Scenario Outline: Literal and finite static list vocabularies retain order and blanks
    Given a validation workbook with "<condition>"
    When I inspect its allowed values
    Then the expected validation vocabulary is returned without mutation

    Examples:
      | condition |
      | literal whitespace quotes and empty item |
      | no matching validation |
      | non-list validation |
      | reversed local range with blank |
      | quoted cross-sheet range |
      | overlapping sqref areas in one list |
      | boolean numeric and empty string |

  @VOCAB-002
  Scenario Outline: Unproved validation inputs refuse without a partial vocabulary
    Given a validation workbook with "<condition>"
    When I inspect its allowed values
    Then validation inspection refuses and returns no partial values

    Examples:
      | condition |
      | two matching lists |
      | invalid literal quote |
      | dynamic source |
      | absent source sheet |
      | two-dimensional source |
      | source formula with cached value |
      | source error |
      | merged source interior |
      | merged target interior |
      | extended validation |
      | duplicate source cell |
      | arithmetic after range |
