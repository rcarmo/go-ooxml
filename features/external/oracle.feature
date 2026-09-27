@external @office @spreadsheet
Feature: Optional spreadsheet calculation checks
  An external calculation engine produces scoped results without replacing the original workbook structure.

  @ORACLE-001
  Scenario: Recalculation delivers only eligible caches in the preserved package
    Given a supported workbook with formulas and an opaque extension part
    And a configured LibreOffice version available in an isolated worker
    When I request recalculation to a different output path
    Then the source file remains byte-identical
    And eligible calculated caches are spliced into a preserved candidate
    And the opaque extension part retains its original bytes
    And LibreOffice's rewritten archive is not delivered
    And the result records the engine version and calculation exclusions
