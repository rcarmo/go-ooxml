@implemented @go @package
Feature: Document boundary whitespace must be literal XML whitespace
  Decoding a character reference or CDATA into whitespace cannot make it legal outside the root.

  @PROLOG-001
  Scenario Outline: Lexically illegal boundary text refuses unchanged
    Given XML with lexical boundary case "<case>"
    When I parse it as an editable XML snapshot
    Then invalid XML structure is refused without changing the input

    Examples:
      | case |
      | decimal space prolog |
      | hexadecimal tab epilog |
      | whitespace CDATA prolog |
      | empty CDATA epilog |
