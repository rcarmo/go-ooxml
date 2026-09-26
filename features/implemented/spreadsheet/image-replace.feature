@implemented @go @spreadsheet
Feature: Replace one loaded worksheet picture without rewriting drawing XML
  Image bytes are added to a fresh part; old shared media and other consumers remain.

  @IMAGE-001
  Scenario: Retarget a selected picture while another picture shares its old media
    Given a worksheet with two pictures sharing old image bytes
    When I replace picture one with a valid PNG
    Then drawing geometry and the other picture remain byte-identical
    And the selected image uses fresh media while original media is retained
    And the consumed image target refuses reuse

  @IMAGE-002
  Scenario Outline: Ambiguous or protected image replacement refuses unchanged
    Given an image workbook with "<condition>"
    When I attempt to replace its selected picture
    Then image replacement refuses without package mutation

    Examples:
      | condition |
      | shared relationship ID |
      | shared drawing part |
      | protected worksheet |
      | malformed image data |
      | external picture |
