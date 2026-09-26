@implemented @go @package
Feature: Archive intake and failure-safe package delivery
  Ambiguous archive names and failed writes must not produce a successful edit.

  @IO-001
  Scenario Outline: Unsafe archive names refuse before package parsing
    Given a synthetic archive with "<kind>" member names
    When I attempt to open the synthetic archive
    Then archive intake fails without changing the source bytes

    Examples:
      | kind      |
      | duplicate |
      | collision |
      | traversal |
      | absolute  |
      | backslash |

  @IO-002
  Scenario: Stream finalisation failure is returned to the caller
    Given a small new Office package
    When I write it to a stream that fails during finalisation
    Then the write reports the injected failure

  @IO-003
  Scenario: A serialisation failure preserves an existing destination
    Given an existing destination and an unserialisable package
    When I attempt to save over the destination
    Then saving fails and the destination and package state remain unchanged
    And no temporary files remain in the destination directory
