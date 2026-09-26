@catalogue @not-executed
Feature: Native mutable presentation fixture readback predicates
  One declaration dispatches eleven literal fixture rows through dynamic tc.name subtests.
  Each row opens its source, mutates, saves to a temporary file, closes and reopens before verification.
  These assertions do not compare all original payloads or independently render the deck.

  @candidate-go-presentation-fixtures-001
  Scenario: Minimal fixture retains selected text box formatting geometry and notes
    Given minimal.pptx receives a text box repositioned to 120000 and 140000 with size 4100000 by 900000 and formatted Fixture Run text
    When the native mutable presentation is saved and reopened
    Then slide one Notes equals "Fixture notes" and some shape text contains "Fixture Minimal"
    And Shape at string selector zero has those exact position and size values and normal autofit
    And a paragraph containing Fixture Run is centered and an exact Fixture Run run is bold italic underlined with size 18 font Calibri and colour FF0000
    And the source shape selector and substring helpers do not prove all unrelated parts were retained

  @candidate-go-presentation-fixtures-002
  Scenario: Title fixture retains text with conditional placeholder verification
    Given title_slide.pptx receives Fixture Title through its title placeholder or a new text box fallback
    When saved and reopened
    Then some slide-one shape text contains "Fixture Title"
    And a returned nonnil title placeholder must report IsPlaceholder true
    And absence of a reopened title placeholder does not fail that conditional check

  @candidate-go-presentation-fixtures-003
  Scenario: Bullet fixture checks text and placeholder presence rather than bullet markup
    Given bullet_points.pptx receives a text box paragraph assigned character bullets and Fixture Bullet text
    When saved and reopened
    Then some slide-one shape text contains "Fixture Bullet" and Placeholders is nonempty
    And no bullet kind character font or list XML is read back

  @candidate-go-presentation-fixtures-004
  Scenario: Shape fixture requires some shape and conditional nonzero geometry
    Given shapes.pptx receives a rectangle assigned colours and nonzero position and size
    When saved and reopened
    Then slide-one Shapes is nonempty
    And if Shape at string selector zero succeeds its position and size components are nonzero
    And selector failure does not fail the check and exact assigned geometry and colours are not compared

  @candidate-go-presentation-fixtures-005
  Scenario: Table fixture requires some table and conditional nonzero row height
    Given tables.pptx receives a two-by-two table with Fixture Table text and row-zero height 500000
    When saved and reopened
    Then slide-one Tables is nonempty
    And if its first table contains rows the first row height must be nonzero
    And neither the new table identity cell text dimensions nor exact assigned height is compared

  @candidate-go-presentation-fixtures-006
  Scenario: Notes fixture reads back exact assigned notes
    Given notes.pptx has slide-one notes assigned "Fixture Notes"
    When saved and reopened
    Then slide-one Notes equals "Fixture Notes"
    And the mutation ignores the SetNotes return error and verifies no independent placeholder or payload preservation

  @candidate-go-presentation-fixtures-007
  Scenario: Comment fixture requires some comment after adding one
    Given comments.pptx accepts adding Fixture Comment by Tester at position 100 and 100
    When saved and reopened
    Then slide-one Comments is nonempty
    And new comment identity text author position and existing comment preservation are not compared

  @candidate-go-presentation-fixtures-008
  Scenario: Hidden-slide fixture requires at least one hidden slide
    Given hidden_slides.pptx receives a layout-zero slide set hidden
    When saved and reopened
    Then at least one slide reports Hidden true
    And the assertion does not identify the added slide or distinguish an already hidden source slide

  @candidate-go-presentation-fixtures-009
  Scenario: Multiple-master fixture retains nonempty master and layout collections
    Given multiple_masters.pptx receives a layout-zero slide
    When saved and reopened
    Then Masters and Layouts are both nonempty
    And counts identities and the newly added slide are not compared

  @candidate-go-presentation-fixtures-010
  Scenario: Layout fixture retains some layout and slide-one layout presence
    Given layouts.pptx receives a layout-zero slide
    When saved and reopened
    Then Layouts is nonempty and Slide one succeeds with a nonnil Layout
    And no exact layout identity or added-slide relationship is checked

  @candidate-go-presentation-fixtures-011
  Scenario: Image fixture checks picture presence and added text only
    Given images.pptx accepts AddPicture and ReplacePictureImage at selector zero using image1.png and receives Fixture Image Placeholder text
    When saved and reopened
    Then some slide-one shape text contains "Fixture Image Placeholder" and Pictures is nonempty
    And image payload geometry or the replacement target identity is not compared
