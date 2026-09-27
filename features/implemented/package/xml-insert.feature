@implemented @go @package
Feature: Structured XML insertion preserves namespace meaning
  New elements use expanded names and cannot silently inherit the wrong default namespace.

  @NODE-001
  Scenario: A new child reuses the existing prefix for its namespace
    Given a prefixed XML parent with opaque surrounding markup
    When I append a same-namespace child with a namespaced attribute
    Then the new child and attribute have the requested expanded names
    And every original byte outside the insertion remains identical

  @NODE-002
  Scenario: An unqualified child resets an inherited default namespace
    Given an XML parent in a default namespace
    When I append an explicitly no-namespace child
    Then the child has no namespace and the parent keeps its namespace

  @NODE-003
  Scenario: A self-closing parent expands without rewriting its attributes
    Given a self-closing prefixed parent with unusual attribute spacing
    When I append a same-namespace child with a namespaced attribute
    Then the new child and attribute have the requested expanded names
    And the parent attributes retain their original lexical bytes

  @NODE-004
  Scenario: An invalid insertion refuses the complete batch
    Given a prefixed XML parent with opaque surrounding markup
    When one insertion attempts to supply a namespace declaration
    Then the insertion batch refuses with a reusable unchanged snapshot
