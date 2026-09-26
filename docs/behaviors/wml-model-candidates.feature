@catalogue @not-executed
Feature: Native WordprocessingML model predicates
  These staging outcomes retain the exact strength of existing native assertions.
  Model decoding and seed predicates do not establish schema validity or package custody.

  @candidate-go-wml-model-001
  Scenario: Field character seed serialization contains a name marker
    Given an input containing "fldChar" and "xmlns:w" that decodes as FldChar
    When the native model is marshalled with an XML header
    Then serialization succeeds and the output contains "fldChar"
    And the begin field type and malformed-input refusal codes are not asserted

  @candidate-go-wml-model-002
  Scenario: Admitted paragraph inputs serialize paragraph and namespace markers
    Given an input containing "<w:p" and "xmlns:w" that decodes as P
    When the native model is marshalled with an XML header
    Then the output contains "<p" and the WordprocessingML namespace URI
    And five seed forms supply no text revision or paragraph-property equality assertion

  @candidate-go-wml-model-003
  Scenario: Admitted run inputs serialize run and namespace markers
    Given an input containing "<w:r" and "xmlns:w" that decodes as R
    When the native model is marshalled with an XML header
    Then the output contains "<r" and the WordprocessingML namespace URI
    And four seed forms supply no text whitespace tab break or formatting equality assertion

  @candidate-go-wml-model-004
  Scenario: Admitted body inputs serialize body and namespace markers
    Given an input containing "<w:body" and "xmlns:w" that decodes as Body
    When the native model is marshalled with an XML header
    Then the output contains "<body" and the WordprocessingML namespace URI
    And neither seed asserts paragraph or table identity after serialization

  @candidate-go-wml-model-005
  Scenario: Minimal document serialization contains document and body substrings
    Given a Document with an empty nonnil Body
    When encoding/xml marshals the model
    Then serialization succeeds and contains "document" and "body"
    And no namespace or structural readback assertion is made

  @candidate-go-wml-model-006
  Scenario: A document decodes one paragraph and one run with expected text
    Given a document XML body with one paragraph containing one run and Hello World text
    When encoding/xml unmarshals it as Document
    Then Body has one P whose Content has one R
    And concatenating direct T nodes from that run yields "Hello World"

  @candidate-go-wml-model-007
  Scenario: Three paragraph body input decodes three content slots
    Given a body XML containing First Second and Third paragraphs
    When encoding/xml unmarshals it as Body
    Then the body Content length is three
    And the individual types and paragraph text are not compared

  @candidate-go-wml-model-008
  Scenario: Body decoding retains paragraph table paragraph type order
    Given a body XML with a paragraph followed by a one-cell table and another paragraph
    When encoding/xml unmarshals it as Body
    Then exactly three content slots have types P Tbl and P in that order
    And their text and table geometry are not compared

  @candidate-go-wml-model-009
  Scenario: Body decoding exposes section properties
    Given a body XML with a paragraph and section properties containing page size
    When encoding/xml unmarshals it as Body
    Then SectPr is nonnil
    And page width height and paragraph content are not compared

  @candidate-go-wml-model-010
  Scenario: Paragraph decoding retains explicit style and alignment values
    Given paragraph XML with style Heading1 and alignment center
    When encoding/xml unmarshals it as P
    Then PPr is nonnil and PStyle.Val is "Heading1" and Jc.Val is "center"
    And this asserts direct model values without resolving inherited formatting

  @candidate-go-wml-model-011
  Scenario: Insertion decoding retains its position type and author
    Given paragraph XML with an ordinary run and an inserted run
    When encoding/xml unmarshals it as P
    Then Content has two slots and the second is Ins with Author "Author"
    And insertion ID date and inserted text are not compared

  @candidate-go-wml-model-012
  Scenario: Deletion decoding retains its position type and author
    Given paragraph XML with deleted content followed by an ordinary run
    When encoding/xml unmarshals it as P
    Then Content has two slots and the first is Del with Author "Editor"
    And deletion ID date and deleted text are not compared

  @candidate-go-wml-model-013
  Scenario: Run decoding retains selected direct properties and a text node
    Given run XML with bold italic single underline size 28 red colour and Formatted Text
    When encoding/xml unmarshals it as R
    Then RPr and its B and I members are nonnil
    And U.Val is "single" and Sz.Val is 28 and Color.Val is "FF0000"
    And at least one direct T node equals "Formatted Text"

  @candidate-go-wml-model-014
  Scenario: Run decoding of text break and tab produces five slots
    Given run XML with three text nodes one break and one tab
    When encoding/xml unmarshals it as R
    Then Content has five slots
    And the slot types order and values are not independently compared

  @candidate-go-wml-model-015
  Scenario: Table decoding checks selected row cell and paragraph structure
    Given table XML with style grid widths two rows and two cells per row
    When encoding/xml unmarshals it as Tbl
    Then there are two rows and two cells in the first row
    And the first cell contains one P with one content slot
    And text grid widths style second-row cells and the last slot type are not compared

  @candidate-go-wml-model-016
  Scenario: Table round trip checks selected row and cell counts
    Given a native table model with TableGrid style two grid columns and four text cells
    When encoding/xml marshals and unmarshals it
    Then there are two rows and two cells in the first row
    And text style widths and the second-row cell count are not compared

  @candidate-go-wml-model-017
  Scenario: Cell decoding retains merge and shading values
    Given cell XML with width grid span 2 vertical merge restart and yellow shading
    When encoding/xml unmarshals it as Tc
    Then TcPr is nonnil and GridSpan.Val is 2
    And VMerge.Val is "restart" and Shd.Fill is "FFFF00"
    And width shading mode and cell text are not compared

  @candidate-go-wml-model-018
  Scenario: OnOff distinguishes absence presence and explicit booleans
    Given nil empty explicitly true and explicitly false OnOff pointers
    When Enabled is called for each of the four named rows
    Then the results are false true true and false respectively
    And no XML boolean spelling is decoded in these rows

  @candidate-go-wml-model-019
  Scenario: Enabled OnOff helper creates an enabled nonnil value
    Given the NewOnOffEnabled helper
    When it creates a native value
    Then the pointer is nonnil and Enabled returns true
    And no XML serialization assertion is made

  @candidate-go-wml-model-020
  Scenario: Boolean OnOff helper preserves true and false inputs
    Given true and false arguments to NewOnOff
    When Enabled is queried on each returned value
    Then the true argument yields true and the false argument yields false
    And no serialized boolean attribute is compared

  @candidate-go-wml-model-021
  Scenario: Run property serialization satisfies single-character predicates
    Given run properties with bold italic underline strike sizes colour highlight and fonts
    When encoding/xml marshals RPr
    Then the output contains each literal substring "b" and "i" and "u"
    And these character checks do not identify corresponding XML elements or property values

  @candidate-go-wml-model-022
  Scenario: Paragraph property serialization satisfies name-marker predicates
    Given paragraph properties with Heading1 center spacing values and KeepNext
    When encoding/xml marshals PPr
    Then the output contains "pStyle" and "jc" and "spacing"
    And style alignment spacing values and KeepNext are not read back

  @candidate-go-wml-model-023
  Scenario: Mixed document round trip retains selected structure and style
    Given a native document with two paragraphs a one-row table and page size properties
    When encoding/xml marshals and unmarshals it
    Then Body is nonnil with three content slots
    And the first slot is P with style "Heading1" and the third is Tbl with one row
    And text bold formatting second-slot type table cells and page size are not compared

  @candidate-go-wml-model-024
  Scenario: Text helper preserves input and sets space flag for five literal rows
    Given simple leading-space trailing-space double-space and no-space text inputs
    When NewT creates a native text value for each row
    Then Text equals its input and Space equals "preserve" only for the three space-sensitive rows
    And tabs line breaks non-ASCII whitespace and serialized xml:space are not asserted
