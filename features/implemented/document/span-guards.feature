@implemented @go @document
Feature: Hidden run content cannot be crossed by Word text edits
  Text matching does not grant authority to cross non-text content or differing namespace semantics.

  @GUARD-001
  Scenario: A drawing-only run inside a changed interval refuses replacement
    Given Word text split around a drawing-only run
    When I attempt the matched cross-run correction
    Then the hidden-content correction refuses without changing the archive

  @GUARD-002
  Scenario: Lexically equal properties with different namespace bindings refuse insertion
    Given boundary runs with equal property bytes but different prefix bindings
    When I attempt the matched boundary insertion
    Then the hidden-content correction refuses without changing the archive
