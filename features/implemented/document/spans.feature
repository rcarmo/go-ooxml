@implemented @go @document
Feature: Exact Word spans across run fragmentation
  Search exposes all exact matches without selecting an arbitrary occurrence.

  @SPAN-001
  Scenario Outline: Exact matching crosses formatting runs and Unicode boundaries
    Given a Word paragraph with run text "<first>" and "<second>"
    When I find exact Word spans for "<needle>"
    Then exactly one span reports the original text "<needle>"

    Examples:
      | first       | second       | needle        |
      | Payment     | terms        | Paymentterms  |
      | pre Al      | pha post     | Alpha         |
      | fee 😀      | café         | 😀café        |

  @SPAN-002
  Scenario: Repeated matches remain explicit and single-target lookup refuses
    Given a Word paragraph with run text "red red" and " red"
    When I find exact Word spans for "red"
    Then all three spans remain available in source order
    And single-target lookup reports ambiguous-target

  @SPAN-003
  Scenario: Text separated by paragraph boundaries is not silently concatenated
    Given Word paragraphs containing "first" and "second"
    When I find exact Word spans for "firstsecond"
    Then no spans are returned
