@implemented @go @package
Feature: Lossless XML attribute edits
  Attribute edits preserve unrelated start-tag bytes and cannot alter namespace bindings.

  @ATTR-001
  Scenario: Add a namespaced attribute using an existing binding
    Given XML with a text element and unrelated attributes
    When I add XML space preservation and replace its text
    Then the original start-tag attributes retain their bytes
    And only the requested attribute and text are changed

  @ATTR-002
  Scenario: A namespace-changing batch refuses without changing the snapshot
    Given XML with a text element and unrelated attributes
    When an attribute edit attempts to change a namespace declaration
    Then the attribute batch refuses and the original snapshot remains reusable
