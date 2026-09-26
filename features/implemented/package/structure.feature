@implemented @go @package
Feature: Local and central ZIP structures agree
  Member identity cannot change between directory lookup and payload reading.

  @ZIP-001
  Scenario Outline: Disagreeing local headers refuse before reading XML
    Given a synthetic archive with a mismatched local "<field>"
    When I attempt to open the synthetic archive
    Then archive intake fails without changing the source bytes

    Examples:
      | field |
      | name |
      | method |
      | flags |
