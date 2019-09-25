# HtmlGenerator
Using EPPlus to read Excel named range and generate a HTML table.

## HTML Templates for Form Elements
Template is used to store the HTML content for the each form elements.

Template Filename: _ELEMENT_ID_.txt
Template Location: \Template\

Available templates idenfiers:
- Drop down list (**_DDL_.txt**)
- Radio buttons (**_RB_.txt**)
- Textbox (**_TB_.txt**)


## Excel Format
Uses the following rules for the template definition. Each worksheet (tab) is for ONE result file. 

The region for the form is defined in an Excel named range with the following format TABLE_<WORKSHEET_NAME>

Example: TABLE_FORM_A

HTML form elements elements is using _ELEMENT_ID_#FormID format where _ELEMENT_ID_ refers to the HTML template identifier.


