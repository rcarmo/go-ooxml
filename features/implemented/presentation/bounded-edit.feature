@implemented @go @presentation
Feature: Bounded presentation text correction
  A shape-ID target changes only its plain text while preserving source formatting.

  @SLIDE-001
  Scenario: Shape-ID targeting preserves geometry and opaque extension bytes
    Given a presentation with a plain text shape and opaque slide markup
    When I replace "old text" in shape 7 on "ppt/slides/slide1.xml"
    Then only the selected slide text bytes differ after delivery
    And the consumed presentation target refuses reuse

  @SLIDE-002
  Scenario Outline: Unsupported presentation text context refuses unchanged
    Given a presentation text shape with "<condition>"
    When I request a guarded presentation text correction
    Then the presentation operation returns a typed refusal
    And its delivered archive is byte-identical to the source

    Examples:
      | condition |
      | duplicate shape IDs |
      | repeated leaf text |
      | field-bearing paragraph |
      | foreign slide part |
