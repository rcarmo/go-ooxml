@implemented @go @package
Feature: Structured subtree replacements preserve surrounding source bytes
  Replacement nodes resolve names using the surviving parent's bindings.

  @SUBTREE-001
  Scenario: Replace a shadowing subtree with two siblings under parent bindings
    Given an XML subtree that shadows its parent's prefix
    When I replace it with structured sibling nodes
    Then inserted names use the parent scope and surrounding bytes stay exact

  @SUBTREE-002
  Scenario: A later invalid subtree replacement leaves the snapshot unchanged
    Given an XML subtree that shadows its parent's prefix
    When one of the structured sibling replacements contains invalid XML text
    Then subtree replacement refuses without publishing partial output
