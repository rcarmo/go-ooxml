@implemented @go @package
Feature: Bounded ZIP64 declarations agree with physical archive ranges
  Tiny fixtures exercise ZIP64 without allocating multi-gigabyte payloads.

  @ZSIX-001
  Scenario Outline: Verified local ZIP64 payloads survive no-op delivery
    Given a tiny ZIP64 package using "<layout>"
    When I open and round-trip its retained payloads
    Then all ZIP64 source bytes and the empty member are preserved

    Examples:
      | layout |
      | stored local sizes |
      | deflated local sizes |
      | signed descriptor |
      | unsigned descriptor |

  @ZSIX-002
  Scenario Outline: Conflicting or undeclared ZIP64 bytes refuse intake
    Given a tiny ZIP64 package with "<defect>"
    When I validate the ZIP64 package for editing
    Then ZIP64 intake refuses with a typed error and unchanged bytes

    Examples:
      | defect |
      | gap before ZIP64 end |
      | gap before locator |
      | ZIP64 length overflow |
      | ZIP64 length crosses locator |
      | short local extra |
      | local size mismatch |
      | short central extra |
      | duplicate central extra |
      | descriptor mismatch |
      | classic count mismatch |
