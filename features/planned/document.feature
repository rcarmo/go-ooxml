@planned @go @document
Feature: Guarded Word review and editing
  Review operations keep the unaffected document and its formatting intact.

  @DOCX-001
  Scenario: Replacement spans several runs with one formatting region
    Given a Word paragraph whose unique text "Payment terms" is split across two runs
    And both runs have identical direct formatting
    And the matched text has one wrapper owner and no internal positional marker
    When I replace that match with "Settlement terms"
    Then the paragraph contains "Settlement terms"
    And the replacement retains the source direct formatting
    And text and formatting outside the match remain unchanged
    And the consumed match cannot authorise another edit

  @DOCX-002
  Scenario: A historical target cannot authorise a current edit
    Given a Word document with a pending tracked change
    And a live block selected from the original view
    When I request a paragraph mutation through that block
    Then I receive an unsupported-structure refusal
    And the document and all previously held targets remain unchanged

  @DOCX-003
  Scenario: A redline reproduces both documents through review resolution
    Given two Word documents differing only in supported unambiguous text edits
    When I generate a tracked-change comparison
    Then accepting every generated revision on a private copy yields the revised content
    And rejecting every generated revision on another private copy yields the original content
    And unrelated package members retain their payload bytes

  @DOCX-004
  Scenario: Comments-only protection permits comments but refuses body edits
    Given a Word document protected for comments-only editing
    When I add a comment to a valid current-view span
    Then the new comment is anchored to that exact span
    When I attempt to replace the protected body text
    Then I receive a protected-operation refusal
    And the document remains as it was immediately before the refused edit
