@implemented @go @package
Feature: Explicit graph deletion is atomic with payload edits
  A detached leaf can be removed without deleting any unselected shared content.

  @DELETE-001
  Scenario: Delete a leaf and its selected relationships in one transaction
    Given a package with two relationships to a removable leaf
    When I plan removal of both edges and the leaf with an owner edit
    Then the deletion preview preserves all package bytes
    When I apply and deliver the deletion plan
    Then the leaf and its content-type override and edges are absent
    And the owner edit and unrelated registry bytes are preserved

  @DELETE-002
  Scenario Outline: Unsafe deletion refuses without committing any payload
    Given a package with two relationships to a removable leaf
    When I attempt a deletion plan with "<condition>"
    Then deletion planning refuses with unchanged package bytes

    Examples:
      | condition |
      | remaining inbound edge |
      | stale leaf fingerprint |
      | stale owner replacement |
      | duplicate edge removal |
      | delete and replace conflict |
      | missing leaf |
