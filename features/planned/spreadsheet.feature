@planned @go @spreadsheet
Feature: Reference-aware spreadsheet edits
  Edits either preserve dependent structures or refuse before changing the workbook.

  @XLSX-001
  Scenario: Row insertion updates every supported dependent reference
    Given a workbook with supported formulas, names, tables and charts referencing a row
    When I insert a row before that row
    Then each supported reference follows the returned address remap
    And affected formula and chart caches are invalidated
    And unrelated cached results remain unchanged
    And the edit receipt lists the direct changes and derived effects

  @XLSX-002
  Scenario: An unresolved dependent reference refuses a structural edit
    Given a workbook with a dependent dynamic reference that cannot be safely remapped
    And a cell handle acquired before the edit
    When I request a row insertion affecting that dependency
    Then I receive an unsupported-structure refusal
    And the workbook and the acquired cell handle remain unchanged
    And no output is delivered

  @XLSX-003
  Scenario: A purely cosmetic style change leaves formula caches intact
    Given a workbook with cached formula results independent of the edited style
    When I change only a cell border and save
    Then the formula cache payloads remain unchanged
    And the receipt contains no formula-cache invalidation effect

  @XLSX-004
  Scenario: Table append does not infer a calculated column from a neighbouring cell
    Given a supported table with a formula in its last ordinary data row
    And no declared calculated-column formula for that column
    When I append a row with an explicit value for that column
    Then the new cell contains the supplied value
    And no formula is inferred from the neighbouring row

  @XLSX-005
  Scenario: Pivot refresh permission reports the remaining stale cache
    Given a workbook edit affecting the local source of a known pivot
    And explicit refresh-on-open permission for that pivot cache
    When I save the supported edit
    Then the pivot cache is flagged for refresh on open
    And the receipt states that the cached pivot result remains stale until refresh
    And the operation does not claim that pivot results were recalculated
