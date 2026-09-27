@implemented @go @package
Feature: Existing Office fixture intake
  Repository fixtures can be inspected without a fixed workspace mount.

  @BASE-001
  Scenario Outline: Open the principal part from a repository fixture
    Given the repository fixture "<fixture>"
    When I open its Office package
    Then the package contains a nonempty part "<part>"

    Examples:
      | fixture                         | part                |
      | word/minimal.docx               | word/document.xml   |
      | pptx/minimal.pptx               | ppt/presentation.xml |
      | excel/minimal.xlsx              | xl/workbook.xml     |
      | pptx/comments.pptx              | ppt/presentation.xml |
      | excel/formatting.xlsx           | xl/workbook.xml     |
      | excel/formulas.xlsx             | xl/workbook.xml     |
      | excel/conditional_format.xlsx   | xl/workbook.xml     |
