@catalogue @not-executed
Feature: Mutable OPC package parts and registries
  Candidate outcomes for central reconciliation; mutable authoring is separate from retained-source guarantees.

  @candidate-go-package-model-001
  Scenario: A new package exposes an initial registry and modified state
    Given no existing package input
    When a native mutable package is created
    Then the package and content-type registry exist
    And the new package is marked modified

  @candidate-go-package-model-002
  Scenario: Added XML and binary parts retain content and content type
    Given XML or binary bytes and a part URI with or without a leading slash
    When the part is added and retrieved by URI
    Then the returned content and content type equal the supplied values
    And listing the package returns the number of added parts

  @candidate-go-package-model-003
  Scenario: Part deletion and closed-package operations report errors
    Given an existing mutable package part
    When that part is deleted
    Then it no longer exists
    And deleting an absent part returns an error
    And after closing the package, getting, adding or deleting a part returns an error

  @candidate-go-package-model-004
  Scenario: Saved mutable packages reopen with parts and registries
    Given a document part and an office-document relationship
    When the package is saved to a file and reopened or opened from its bytes
    Then the document part exists with its content type and payload
    And its package relationship is retained
    And stream writing emits nonempty ZIP-signature bytes

  @candidate-go-package-model-005
  Scenario: Mutable part streams and content updates expose stored bytes
    Given a part with known bytes
    When its stream is read or its size is queried
    Then the returned bytes and length match the stored content
    When its content is replaced
    Then retrieval returns the new bytes and the part is marked modified

  @candidate-go-package-model-006
  Scenario: Core properties retain metadata and establish one relationship
    Given title, author, subject, descriptive fields, revision and timestamps
    When core properties are assigned and read from the package
    Then each assigned field and timestamp is retained
    And exactly one core-properties relationship targets the core-properties part
    And a new package without assigned properties returns empty defaults

  @candidate-go-package-model-007
  Scenario: Content-type resolution uses overrides before extension defaults
    Given specific part overrides and registered extension defaults
    When a URI with or without a leading slash is resolved
    Then the explicit override is returned when present
    And otherwise the extension default is returned or an unknown type is empty

  @candidate-go-package-model-008
  Scenario: Registry updates do not duplicate existing keys
    Given a content-type override or extension default already exists
    When the same key is added with a new value
    Then its value changes without adding another entry
    And removing an override returns true once and false when already absent
    And ensuring an already matching default creates no unnecessary override

  @candidate-go-package-model-009
  Scenario: Relationship helper paths resolve relative to the owning part
    Given a root-level or nested source part and a relative or package-absolute target
    When relationship part paths and target paths are resolved
    Then the relationship file is in the owner's rels directory
    And the target resolves to the expected package-relative part path

  @candidate-go-package-model-010
  Scenario: Relationship IDs and selections are scoped to their owner
    Given empty package-level and part-level relationship sets
    When relationships are added
    Then each owner starts ID allocation at rId1 and subsequent additions advance its IDs
    And adding an explicit existing ID updates its type and target without duplication
    And type selection returns all matching edges or the first match, with no match returning nil
    And removing an existing ID returns true once and false after removal
