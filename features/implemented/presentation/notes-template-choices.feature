@implemented @go @presentation
Feature: Notes templates must have unambiguous formatting choices
  Cloning supported properties requires a coherent property tree, not merely known names.

  @NOTEPROP-001
  Scenario Outline: Conflicting template choices refuse the whole edit
    Given multiline notes condition "<condition>"
    When I attempt an unsupported multiline notes edit
    Then the notes snapshot and package remain unchanged

    Examples:
      | condition |
      | conflicting run fill |
      | conflicting colour choice |
      | conflicting paragraph spacing |
      | conflicting bullet type |
      | conflicting bullet size |
      | empty solid fill |
      | missing colour value |
      | reversed property order |
