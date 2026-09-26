@implemented @go @package
Feature: Retained-source delivery receipts
  Receipts describe observed payload changes and failed delivery preserves prior output.

  @RECEIPT-001
  Scenario: A saved receipt agrees with independent package hashes
    Given a staged main-part edit with an opaque custom part
    When I save a main-part replacement with a delivery receipt
    Then the receipt identifies exactly the changed member and both hashes
    And the delivered archive reopens with unchanged opaque bytes

  @RECEIPT-002
  Scenario: A directory destination refuses without consuming staged changes
    Given a staged main-part edit with an opaque custom part
    When I deliver staged changes to a directory destination
    Then delivery fails without changing the directory
    And the staged change remains available for another save
