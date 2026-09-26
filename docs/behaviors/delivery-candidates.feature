@catalogue @not-executed
Feature: Retained payload custody and failure-safe delivery
  Candidate outcomes for central reconciliation; no feature steps execute here.

  @candidate-go-delivery-001
  Scenario: Retained packages own independent source and part buffers
    Given a package opened from caller-owned archive bytes
    When the caller mutates that input or a returned part buffer
    Then the retained archive and subsequent part readback remain unchanged
    And part fingerprints identify the retained payload
    And no-op stream delivery preserves the original archive bytes

  @candidate-go-delivery-002
  Scenario: Replacement batches refuse atomically on stale or unproved entries
    Given a batch containing a valid replacement and another stale fingerprint
    When the retained package applies the batch
    Then no replacement is committed
    And malformed XML, unknown namespace prefixes, duplicate expanded attributes, registry replacements and duplicate targets also refuse
    And the selected original payload remains readable after refusal

  @candidate-go-delivery-003
  Scenario: Short stream writes return an error
    Given a destination writer that reports fewer bytes than supplied without its own error
    When a retained archive is written
    Then delivery returns the short-write error

  @candidate-go-delivery-004
  Scenario: Failed atomic delivery preserves the existing destination
    Given an existing destination and an injected serializer or reopen-verifier failure
    When atomic file delivery is attempted
    Then the original destination bytes remain unchanged
    And no temporary file is left in the destination directory
    And a symlink destination is refused with the same preservation guarantees

  @candidate-go-delivery-005
  Scenario: A closed package cannot emit a replacement archive
    Given a closed mutable package
    When file save or stream writing is requested
    Then the closed-document error is returned
    And an existing destination remains unchanged and the stream receives no bytes

  @candidate-go-delivery-006
  Scenario: Repeated mutable serialization is deterministic
    Given parts and relationship registries added in a nonlexical order
    When the unchanged mutable package is serialized repeatedly
    Then every produced archive is byte-identical
    And members are ordered consistently within the registry and ordinary-part groups

  @candidate-go-delivery-007
  Scenario: Successful save updates state only after destination replacement
    Given an existing regular destination with restricted permissions
    When a modified package is saved successfully
    Then its permissions are retained and the package path becomes the destination
    And package modified state clears
    But when a later save targets an existing directory
    Then it fails without changing the package path, directory or previously saved archive
