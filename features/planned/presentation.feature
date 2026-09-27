@planned @go @presentation
Feature: Relationship-safe presentation edits
  Slide edits respect object identity and ownership of shared parts.

  @PPTX-001
  Scenario: A cloned chart has an independent embedded workbook
    Given a slide with a supported chart and its embedded workbook
    When I clone the slide
    And replace the cloned chart data
    Then the original chart and workbook payloads remain byte-identical
    And the cloned chart references a separate updated workbook
    And all internal relationships resolve in the saved presentation

  @PPTX-002
  Scenario: Reordering cannot redirect an anchored text edit
    Given a unique inspected text block with a structural identity and full fingerprint
    When another shape is inserted before its owning shape
    And I replace text through the inspected anchor
    Then only the originally inspected block changes

  @PPTX-003
  Scenario: An ambiguous destination layout refuses the whole import
    Given a destination with two equally eligible layouts
    And a source slide with a supported dependency graph
    When I import the slide without an explicit layout choice
    Then I receive an ambiguous-target refusal identifying both layouts
    And the destination and source presentations remain unchanged

  @PPTX-004
  Scenario: Replacing a shared image changes only its selected occurrence
    Given two picture shapes using the same media part
    When I replace the image in one selected shape
    Then the other shape still references the original image bytes
    And the selected shape retains its position, size and crop
    And the saved package contains the replacement media and valid relationships
