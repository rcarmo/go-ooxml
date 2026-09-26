@implemented @go @package
Feature: Namespace names and document boundaries obey XML rules
  Namespace expansion must not turn invalid qualified names into accepted nodes.

  @QNAME-001
  Scenario Outline: Invalid qualified names refuse before editing
    Given XML with invalid qualified-name case "<case>"
    When I parse it as an editable XML snapshot
    Then invalid XML structure is refused without changing the input

    Examples:
      | case |
      | numeric local name |
      | numeric attribute local name |
      | numeric namespace prefix |
      | non XML whitespace outside root |

  @QNAME-002
  Scenario: Fresh insertion bindings do not leak to existing or later siblings
    Given a parent with an occupied generated prefix and default namespace
    When I append children with independent new namespace requirements
    Then all new and existing sibling names resolve to their intended namespaces
