@catalogue @not-executed
Feature: Native mutable presentation API test predicates
  These declarations exercise legacy mutable authoring and selected readback outcomes.
  NewResource and OpenResource require successful construction or opening and ignore Close errors at cleanup.

  @candidate-go-presentation-api-001
  Scenario: Default constructor exposes four-by-three size and no slides
    Given a presentation created through NewResource and New
    When SlideSize and SlideCount are queried
    Then dimensions equal SlideWidth4x3 and SlideHeight4x3 and count is zero

  @candidate-go-presentation-api-002
  Scenario: Widescreen constructor exposes its named dimension constants
    Given a presentation created through NewWidescreen
    When SlideSize is queried
    Then dimensions equal SlideWidth16x9 and SlideHeight16x9 without persistence or rendering checks

  @candidate-go-presentation-api-003
  Scenario: Custom constructor retains one explicit dimension pair
    Given NewWithSize receives 7200000 and 5400000
    When SlideSize is queried
    Then both requested integers are returned without invalid-size coverage

  @candidate-go-presentation-api-004
  Scenario: Core property setters read back title and creator in memory
    Given a new presentation and Title Presentation Title and Creator Presentation Author
    When SetCoreProperties and CoreProperties both succeed
    Then Title and Creator equal those inputs without a disk round trip or other field assertion

  @candidate-go-presentation-api-005
  Scenario: Comments fixture exposes at least one master and layout with paths
    Given comments.pptx opens successfully
    When Masters and Layouts are inspected
    Then both collections are nonempty and their first Path strings are nonempty
    And target resolution counts identities and rendering are not compared

  @candidate-go-presentation-api-006
  Scenario: Three selected index errors match ErrInvalidIndex
    Given an empty presentation followed by two added slides
    When Slide zero and DeleteSlide one are called before addition and duplicate-order one-one afterward
    Then each call returns ErrInvalidIndex
    And no post-refusal state or byte comparison is asserted

  @candidate-go-presentation-api-007
  Scenario: Appending slides reports consecutive one-based indices
    Given an empty presentation
    When two slides are appended with layout zero
    Then the first returned slide is nonnil and counts progress through one and two
    And returned indices are one and two without relationship or payload checks

  @candidate-go-presentation-api-008
  Scenario: Inserting first reindexes the three in-memory slides
    Given two slides already exist
    When InsertSlide inserts at index one
    Then count is three and the new slide index is one
    And each SlidesRaw entry reports its zero-based position plus one

  @candidate-go-presentation-api-009
  Scenario: Deleting index one reduces count and an out-of-range delete errors
    Given three slides in a presentation
    When DeleteSlide one succeeds
    Then count is two and DeleteSlide ten returns an error
    And despite the source comment saying middle the exercised deletion argument is one with no identity assertion

  @candidate-go-presentation-api-010
  Scenario: Duplicating a slide checks only count and returned index
    Given one slide containing Original Text
    When DuplicateSlide one is called
    Then count is two and the returned slide index is two
    And copied text independence graph closure and original immutability are not checked

  @candidate-go-presentation-api-011
  Scenario: Reversing three slides retains the selected in-memory IDs
    Given three captured slide IDs in their original order
    When ReorderSlides succeeds for three-two-one
    Then SlidesRaw IDs equal the third second and first captured IDs in that order
    And a later wrong-length one-two request errors without an asserted rollback snapshot
    And no archive payload preservation or saved slide-order assertion is made

  @candidate-go-presentation-api-012
  Scenario: Hidden flag toggles false true false in memory
    Given one new slide
    When its initial Hidden value is read and then SetHidden true and false are called
    Then the three observed values are false true and false

  @candidate-go-presentation-api-013
  Scenario: Deleting one shape decrements the shape count
    Given a slide with two newly added text boxes
    When DeleteShape at string selector zero succeeds
    Then Shapes has one fewer entry without checking which identity or payload disappeared

  @candidate-go-presentation-api-014
  Scenario: Notes append retains both text substrings in memory
    Given a new slide initially reports no notes
    When Speaker notes here is assigned and Additional notes appended
    Then HasNotes becomes true and initial Notes equals Speaker notes here
    And afterward Notes contains both strings without an exact separator or persisted-readback assertion

  @candidate-go-presentation-api-015
  Scenario: Saved notes master ID resolves to a relationship of the expected type
    Given a new slide accepts speaker notes and the presentation saves and reopens
    When the first private NotesMasterId entry supplies its relationship ID
    Then the presentation relationship lookup finds that ID with type RelTypeNotesMaster
    And no notes-master target payload content or broader relationship closure is checked

  @candidate-go-presentation-api-016
  Scenario: One newly added comment retains text and author after reopen
    Given a new slide accepts Needs review by Test Author at position 100 and 200
    When the presentation saves and reopens and slide one is selected
    Then it has exactly one comment with that exact text and author
    And coordinates IDs dates and multi-slide ownership are not compared

  @candidate-go-presentation-api-017
  Scenario: SaveAs success is accompanied by a file-existence predicate
    Given a new presentation with Test Presentation text
    When SaveAs succeeds and the presentation is closed without checking the close error
    Then os.Stat does not return an IsNotExist error
    And other stat errors or file contents are not independently rejected by the existence predicate

  @candidate-go-presentation-api-018
  Scenario: Two-slide round trip checks only slide count
    Given one text slide and one hidden slide
    When saving and reopening both succeed
    Then SlideCount is two without text hidden-state or geometry readback

  @candidate-go-presentation-api-019
  Scenario: Image fixture round trip retains a theme relationship entry
    Given images.pptx opens and saves to a temporary archive
    When the archive reopens through OpenResource
    Then presentation relationships contain some RelTypeTheme entry
    And target identity theme payload and other advanced parts are not compared

  @candidate-go-presentation-api-020
  Scenario: Text-frame autofit enum transitions through none normal and shape
    Given a fresh text-box frame
    When the default is read and normal then shape autofit modes are set
    Then the three getter values equal none normal and shape respectively without rendering checks

  @candidate-go-presentation-api-021
  Scenario: Shape fill operations provide only a no-panic smoke path
    Given a new rectangle shape
    When blue fill red line width 12700 and no-fill operations are invoked in sequence
    Then the test body reaches its end without panic
    And no getter XML payload receipt save or visual assertion is made
