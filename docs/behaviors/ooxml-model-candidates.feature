@catalogue @not-executed
Feature: Native OOXML model serialization assertions
  These staging outcomes describe selected Go model tests, not package editing or schema completeness.
  Fuzz properties below describe predicates and seed coverage; no exploratory fuzz execution is credited.

  @candidate-go-ooxml-model-001
  Scenario: Default chart model retains its basic object structure
    Given the default native ChartSpace model
    When it is marshalled with an XML header and unmarshalled
    Then Chart and PlotArea are nonnil and PlotArea Content is nonempty
    And Title and Legend are nonnil without asserting their text or chart data

  @candidate-go-ooxml-model-002
  Scenario: Admitted chart fuzzer inputs serialize a chartSpace marker
    Given an input containing "<c:chartSpace" and "xmlns:c" that unmarshals as ChartSpace
    When the native model is marshalled with an XML header
    Then serialization succeeds and contains "chartSpace"
    And other inputs and decode errors return without a refusal assertion

  @candidate-go-ooxml-model-003
  Scenario: Core property seed round trips compare three exact strings
    Given title and creator strings at most 128 bytes and a subject at most 256 bytes
    When CoreProperties is marshalled with an XML header and unmarshalled
    Then Title Creator and Subject equal their input strings
    And longer inputs return without asserting a result

  @candidate-go-ooxml-model-004
  Scenario: Shared string seed round trips only check the first lookup
    Given two strings at most 256 bytes each added to a SharedStrings model
    When the model is marshalled with an XML header and unmarshalled
    Then GetString at index zero equals or contains the first input string
    And no assertion checks the second lookup or declared count or uniqueness

  @candidate-go-ooxml-model-005
  Scenario: Five default diagram models serialize and decode
    Given the default dataModel layoutDef styleDef colorsDef and drawing models
    When each is marshalled with an XML header and unmarshalled into its corresponding type
    Then both operations succeed for each of the five named rows
    And no equality assertion checks the decoded properties or diagram contents

  @candidate-go-ooxml-model-006
  Scenario: Admitted diagram fuzzer inputs retain the selected name marker
    Given an input containing "xmlns:dgm" whose first matching selector is dataModel layoutDef styleDef colorsDef or drawing
    When it successfully decodes into that selected native type and is marshalled
    Then serialization succeeds and contains the selected name
    And unmatched inputs and decode errors return without a refusal assertion

  @candidate-go-ooxml-model-007
  Scenario: Admitted drawing text bodies serialize an unprefixed root marker and namespace
    Given an input containing "<a:txBody" and "xmlns:a" that unmarshals as TxBody
    When the native model is marshalled with an XML header
    Then the output contains "<txBody" and the DrawingML namespace URI
    And text and formatting equality are not asserted

  @candidate-go-ooxml-model-008
  Scenario: GraphicData retains a chart relationship identifier
    Given GraphicData with the chart namespace URI and chart relationship "rId5"
    When the model is marshalled with an XML header and unmarshalled
    Then the decoded Chart is nonnil and its relationship ID is "rId5"
    And no package relationship target or chart payload is resolved

  @candidate-go-ooxml-model-009
  Scenario: Admitted slide fuzzer inputs serialize a slide marker
    Given an input containing "<p:sld" and "xmlns:p" that unmarshals as Sld
    When the native model is marshalled with an XML header
    Then the output contains "<sld" or "<p:sld"
    And slide text geometry and shape identity equality are not asserted

  @candidate-go-ooxml-model-010
  Scenario: Admitted worksheet fuzzer inputs serialize a worksheet marker
    Given an input containing "<worksheet" and the spreadsheet namespace URI that unmarshals as Worksheet
    When the native model is marshalled with an XML header
    Then the output contains "worksheet"
    And the value and drawing relationship in the two seeds are not independently compared

  @candidate-go-ooxml-model-011
  Scenario: Admitted table fuzzer inputs serialize a table marker
    Given an input containing "<table" and the spreadsheet namespace URI that unmarshals as Table
    When the native model is marshalled with an XML header
    Then the output contains "table"
    And name geometry column names and count equality are not asserted

  @candidate-go-ooxml-model-012
  Scenario: Worksheet model retains a drawing relationship identifier
    Given a Worksheet with drawing relationship "rId7"
    When the model is marshalled with an XML header and unmarshalled
    Then the decoded Drawing is nonnil and its relationship ID is "rId7"
    And no drawing target anchor or picture is resolved

  @candidate-go-ooxml-model-013
  Scenario: Admitted theme fuzzer inputs serialize a theme marker
    Given an input containing "<a:theme" and "xmlns:a" that unmarshals as Theme
    When the native model is marshalled with an XML header
    Then the output contains "theme"
    And colour font and format equality are not asserted

  @candidate-go-ooxml-model-014
  Scenario: Default theme retains its name and colour scheme container
    Given the default native Theme model
    When it is marshalled with an XML header and unmarshalled
    Then its Name is "Office" and ThemeElements and ClrScheme are nonnil
    And no assertion compares actual colours or font scheme or formatting
