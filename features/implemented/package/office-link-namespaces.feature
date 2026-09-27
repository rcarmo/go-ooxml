@implemented @go @package
Feature: Office relationship identity follows expanded names and exact type URIs
  Prefix spellings and default namespaces cannot redirect a sheet or slide relationship.

  @OFFLINK-001
  Scenario Outline: Correct expanded relationship attributes open for inspection
    Given an Office link fixture for "<format>" with "<condition>"
    When I inspect the Office link identity
    Then exactly the intended sheet or slide is found without mutation

    Examples:
      | format | condition |
      | xlsx | alias prefix |
      | pptx | alias prefix |
      | xlsx | local correct binding |
      | pptx | local correct binding |
      | xlsx | foreign r alongside correct link |
      | pptx | foreign r alongside correct link |

  @OFFLINK-002
  Scenario Outline: Wrong expanded attributes or type URIs refuse intake
    Given an Office link fixture for "<format>" with "<condition>"
    When I inspect the Office link identity
    Then the Office link is refused and input bytes are unchanged

    Examples:
      | format | condition |
      | xlsx | locally shadowed wrong URI |
      | pptx | locally shadowed wrong URI |
      | xlsx | plain id with default namespace |
      | pptx | plain id with default namespace |
      | xlsx | foreign relationship type suffix |
      | pptx | foreign relationship type suffix |
