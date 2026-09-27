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

  @ZIP-002
  Scenario: An adjusted archive prefix refuses standalone package intake
    Given a synthetic archive with an adjusted prepended prefix
    When I attempt to open the synthetic archive
    Then archive intake fails without changing the source bytes

  @ZIP-003
  Scenario: Distinct CRC-valid member payloads cannot physically overlap
    Given a ZIP whose distinct CRC-valid members overlap physically
    When I validate the ZIP64 package for editing
    Then ZIP64 intake refuses with a typed error and unchanged bytes
