@implemented @go @document
Feature: Word span replacement with exact affix alignment
  A replacement changes one proved interval and leaves surrounding run markup intact.

  @REPLACE-001
  Scenario Outline: Replacement keeps unchanged fragments in their source runs
    Given editable Word runs "<first>" and "<second>"
    When I replace the exact span "<needle>" with "<replacement>"
    Then the resulting run texts are "<after_first>" and "<after_second>"
    And all run properties and surrounding XML are byte-preserved

    Examples:
      | first     | second    | needle | replacement | after_first | after_second |
      | pre Al    | pha post  | Alpha  | Omega       | pre Omeg    | a post       |
      | Al        | pha       | Alpha  | Beta        | Bet         | a            |
      | fee😀     | caféEnd   | 😀café | tea         | feetea      | End          |

  @REPLACE-002
  Scenario: Repeated affix ambiguity refuses without consuming the target
    Given editable Word runs "Term" and "Term"
    When I attempt to replace "TermTerm" with "Term"
    Then ambiguous alignment refuses with the original archive intact
    And the same span remains reusable for a different replacement

  @REPLACE-003
  Scenario: A visible span across a hyperlink boundary cannot authorise an edit
    Given text split across a plain run and a hyperlink wrapper
    When I attempt to replace "Alpha" with "Omega"
    Then wrapper-crossing replacement refuses with the original archive intact
