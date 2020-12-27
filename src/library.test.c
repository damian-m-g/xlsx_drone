#include "../ext/unity.h"
#include "library.h"

/* even empty, this functions need to be here */
void setUp(void) {}
void tearDown(void) {}


void test_xlsx_open(void) {
  xlsx_workbook_t wb;

  TEST_ASSERT_EQUAL_INT(1, xlsx_open("test\\helpers\\empty_sample.xlsx", &wb));  // ATTENTION: Must be closed
  // an empty sample has no shared strings xml (nor a file w/o strings)
  TEST_ASSERT_NULL(wb.shared_strings_xml);
  xlsx_close(&wb);

  xlsx_open("test\\helpers\\sample.xlsx", &wb); // ATTENTION: Must be closed
  TEST_ASSERT_NOT_NULL(wb.shared_strings_xml);

  TEST_ASSERT_EQUAL_INT(55, wb.n_styles);

  // if the style id < 50, uses one of the predefined type and format code. Assertions sorts by style appearance on XML
  struct xlsx_style_t * style = wb.styles[0];
  TEST_ASSERT_EQUAL_INT(0, style->style_id);
  TEST_ASSERT_EQUAL_INT(xlsx_predefined_style_types[style->style_id], style->related_type);
  TEST_ASSERT_EQUAL_STRING(xlsx_predefined_styles_format_code[style->style_id], style->format_code);

  // appears in cell B2 & B6 as numbers
  style = wb.styles[1];
  TEST_ASSERT_EQUAL_INT(165, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_type);
  TEST_ASSERT_EQUAL_STRING("0.000", style->format_code);

  style = wb.styles[2];
  TEST_ASSERT_EQUAL_INT(4, style->style_id);
  TEST_ASSERT_EQUAL_INT(xlsx_predefined_style_types[style->style_id], style->related_type);
  TEST_ASSERT_EQUAL_STRING(xlsx_predefined_styles_format_code[style->style_id], style->format_code);

  style = wb.styles[3];
  TEST_ASSERT_EQUAL_INT(2, style->style_id);
  TEST_ASSERT_EQUAL_INT(xlsx_predefined_style_types[style->style_id], style->related_type);
  TEST_ASSERT_EQUAL_STRING(xlsx_predefined_styles_format_code[style->style_id], style->format_code);

  // appears in cell B5 as number
  style = wb.styles[4];
  TEST_ASSERT_EQUAL_INT(166, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_type);
  TEST_ASSERT_EQUAL_STRING("0.000;[Red]0.000", style->format_code);

  // appears in cell C2 & C3 as currency
  style = wb.styles[5];
  TEST_ASSERT_EQUAL_INT(167, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_type);
  TEST_ASSERT_EQUAL_STRING("\"$\"#,##0.00", style->format_code);

  // appears in cell C4 as currency
  style = wb.styles[6];
  TEST_ASSERT_EQUAL_INT(168, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_type);
  // ATTENTION: the ASCII char isn't shown in the debug but does the match (assert passes)
  TEST_ASSERT_EQUAL_STRING("#,##0.00\\ [$֏-42B]", style->format_code);

  // appears in cell D2 as currency (accounting)
  style = wb.styles[7];
  TEST_ASSERT_EQUAL_INT(169, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_type);
  TEST_ASSERT_EQUAL_STRING("_(\"$\"* #,##0.000_);_(\"$\"* \\(#,##0.000\\);_(\"$\"* \"-\"\?\?\?_);_(@_)",
                           style->format_code);

  // appears in cell D3 as currency (accounting)
  style = wb.styles[8];
  TEST_ASSERT_EQUAL_INT(170, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_type);
  TEST_ASSERT_EQUAL_STRING("_-* #,##0.000\\ [$֏-42B]_-;\\-* #,##0.000\\ [$֏-42B]_-;_-* \"-\"\?\?\?\\ [$֏-42B]_-;_-@_-",
                           style->format_code);

  style = wb.styles[9];
  TEST_ASSERT_EQUAL_INT(14, style->style_id);
  TEST_ASSERT_EQUAL_INT(xlsx_predefined_style_types[style->style_id], style->related_type);
  TEST_ASSERT_EQUAL_STRING(xlsx_predefined_styles_format_code[style->style_id], style->format_code);

  // appears in cell E3 as date
  style = wb.styles[10];
  TEST_ASSERT_EQUAL_INT(171, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_type);
  TEST_ASSERT_EQUAL_STRING("[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy", style->format_code);

  // appears in cell E4 as date
  style = wb.styles[11];
  TEST_ASSERT_EQUAL_INT(172, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_type);
  TEST_ASSERT_EQUAL_STRING("yyyy\\-mm\\-dd;@", style->format_code);

  // appears in cell E5 as date
  style = wb.styles[12];
  TEST_ASSERT_EQUAL_INT(173, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_type);
  TEST_ASSERT_EQUAL_STRING("m/d;@", style->format_code);

  // appears in cell E6 & E21 & E22 & E23 as date
  style = wb.styles[13];
  TEST_ASSERT_EQUAL_INT(174, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_type);
  TEST_ASSERT_EQUAL_STRING("m/d/yy;@", style->format_code);

  // appears in cell E7 as date
  style = wb.styles[14];
  TEST_ASSERT_EQUAL_INT(175, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_type);
  TEST_ASSERT_EQUAL_STRING("mm/dd/yy;@", style->format_code);

  // appears in cell E8 as date
  style = wb.styles[15];
  TEST_ASSERT_EQUAL_INT(176, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_type);
  TEST_ASSERT_EQUAL_STRING("[$-409]d\\-mmm;@", style->format_code);

  // appears in cell E9 as date
  style = wb.styles[16];
  TEST_ASSERT_EQUAL_INT(177, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_type);
  TEST_ASSERT_EQUAL_STRING("[$-409]d\\-mmm\\-yy;@", style->format_code);

  // appears in cell E10 as date
  style = wb.styles[17];
  TEST_ASSERT_EQUAL_INT(178, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_type);
  TEST_ASSERT_EQUAL_STRING("[$-409]mmm\\-yy;@", style->format_code);

  // appears in cell E11 as date
  style = wb.styles[18];
  TEST_ASSERT_EQUAL_INT(179, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_type);
  TEST_ASSERT_EQUAL_STRING("[$-409]mmmm\\-yy;@", style->format_code);

  // appears in cell E12 as date
  style = wb.styles[19];
  TEST_ASSERT_EQUAL_INT(180, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_type);
  TEST_ASSERT_EQUAL_STRING("[$-409]mmmm\\ d\\,\\ yyyy;@", style->format_code);

  // appears in cell F9 & F12 & E13 & F13 & F14 as date-time
  style = wb.styles[20];
  TEST_ASSERT_EQUAL_INT(181, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE_TIME, style->related_type);
  TEST_ASSERT_EQUAL_STRING("[$-409]m/d/yy\\ h:mm\\ AM/PM;@", style->format_code);

  // appears in cell F10 & E14 as date-time
  style = wb.styles[21];
  TEST_ASSERT_EQUAL_INT(182, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE_TIME, style->related_type);
  TEST_ASSERT_EQUAL_STRING("m/d/yy\\ h:mm;@", style->format_code);

  // appears in cell E15 as date
  style = wb.styles[22];
  TEST_ASSERT_EQUAL_INT(183, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_type);
  TEST_ASSERT_EQUAL_STRING("[$-409]mmmmm;@", style->format_code);

  // appears in cell E16 as date
  style = wb.styles[23];
  TEST_ASSERT_EQUAL_INT(184, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_type);
  TEST_ASSERT_EQUAL_STRING("[$-409]mmmmm\\-yy;@", style->format_code);

  // appears in cell E17 as date
  style = wb.styles[24];
  TEST_ASSERT_EQUAL_INT(185, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_type);
  TEST_ASSERT_EQUAL_STRING("m/d/yyyy;@", style->format_code);

  // appears in cell E18 as date
  style = wb.styles[25];
  TEST_ASSERT_EQUAL_INT(186, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_type);
  TEST_ASSERT_EQUAL_STRING("[$-409]d\\-mmm\\-yyyy;@", style->format_code);

  // appears in cell E19 as date
  style = wb.styles[26];
  TEST_ASSERT_EQUAL_INT(187, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_type);
  TEST_ASSERT_EQUAL_STRING("[$-2010000]yyyy/mm/dd;@", style->format_code);

  // appears in cell F2 as time
  style = wb.styles[27];
  TEST_ASSERT_EQUAL_INT(188, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_TIME, style->related_type);
  TEST_ASSERT_EQUAL_STRING("[$-F400]h:mm:ss\\ AM/PM", style->format_code);

  // appears in cell F3 as time
  style = wb.styles[28];
  TEST_ASSERT_EQUAL_INT(189, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_TIME, style->related_type);
  TEST_ASSERT_EQUAL_STRING("h:mm;@", style->format_code);

  // appears in cell F4 as time
  style = wb.styles[29];
  TEST_ASSERT_EQUAL_INT(190, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_TIME, style->related_type);
  TEST_ASSERT_EQUAL_STRING("[$-409]h:mm\\ AM/PM;@", style->format_code);

  // appears in cell F5 as time
  style = wb.styles[30];
  TEST_ASSERT_EQUAL_INT(191, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_TIME, style->related_type);
  TEST_ASSERT_EQUAL_STRING("h:mm:ss;@", style->format_code);

  // appears in cell F6 as time
  style = wb.styles[31];
  TEST_ASSERT_EQUAL_INT(192, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_TIME, style->related_type);
  TEST_ASSERT_EQUAL_STRING("[$-409]h:mm:ss\\ AM/PM;@", style->format_code);

  // appears in cell F7 as time
  style = wb.styles[32];
  TEST_ASSERT_EQUAL_INT(193, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_TIME, style->related_type);
  TEST_ASSERT_EQUAL_STRING("mm:ss.0;@", style->format_code);

  // appears in cell F8 as time
  style = wb.styles[33];
  TEST_ASSERT_EQUAL_INT(194, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_TIME, style->related_type);
  TEST_ASSERT_EQUAL_STRING("[h]:mm:ss;@", style->format_code);

  style = wb.styles[34];
  TEST_ASSERT_EQUAL_INT(10, style->style_id);
  TEST_ASSERT_EQUAL_INT(xlsx_predefined_style_types[style->style_id], style->related_type);
  TEST_ASSERT_EQUAL_STRING(xlsx_predefined_styles_format_code[style->style_id], style->format_code);

  // appears in cell G3 as percentage
  style = wb.styles[35];
  TEST_ASSERT_EQUAL_INT(195, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_type);
  TEST_ASSERT_EQUAL_STRING("0.000%", style->format_code);

  style = wb.styles[36];
  TEST_ASSERT_EQUAL_INT(12, style->style_id);
  TEST_ASSERT_EQUAL_INT(xlsx_predefined_style_types[style->style_id], style->related_type);
  TEST_ASSERT_EQUAL_STRING(xlsx_predefined_styles_format_code[style->style_id], style->format_code);

  style = wb.styles[37];
  TEST_ASSERT_EQUAL_INT(13, style->style_id);
  TEST_ASSERT_EQUAL_INT(xlsx_predefined_style_types[style->style_id], style->related_type);
  TEST_ASSERT_EQUAL_STRING(xlsx_predefined_styles_format_code[style->style_id], style->format_code);

  // appears in cell H4 as a fraction
  style = wb.styles[38];
  TEST_ASSERT_EQUAL_INT(196, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_type);
  TEST_ASSERT_EQUAL_STRING("#\\ \?\?\?/\?\?\?", style->format_code);

  // appears in cell H5 as a fraction
  style = wb.styles[39];
  TEST_ASSERT_EQUAL_INT(197, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_type);
  TEST_ASSERT_EQUAL_STRING("#\\ ?/2", style->format_code);

  // appears in cell H6 as a fraction
  style = wb.styles[40];
  TEST_ASSERT_EQUAL_INT(198, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_type);
  TEST_ASSERT_EQUAL_STRING("#\\ ?/4", style->format_code);

  // appears in cell H7 as a fraction
  style = wb.styles[41];
  TEST_ASSERT_EQUAL_INT(199, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_type);
  TEST_ASSERT_EQUAL_STRING("#\\ ?/8", style->format_code);

  // appears in cell H8 as a fraction
  style = wb.styles[42];
  TEST_ASSERT_EQUAL_INT(200, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_type);
  TEST_ASSERT_EQUAL_STRING("#\\ \?\?/16", style->format_code);

  // appears in cell H9 as a fraction
  style = wb.styles[43];
  TEST_ASSERT_EQUAL_INT(201, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_type);
  TEST_ASSERT_EQUAL_STRING("#\\ ?/10", style->format_code);

  // appears in cell H10 as a fraction
  style = wb.styles[44];
  TEST_ASSERT_EQUAL_INT(202, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_type);
  TEST_ASSERT_EQUAL_STRING("#\\ \?\?/100", style->format_code);

  style = wb.styles[45];
  TEST_ASSERT_EQUAL_INT(11, style->style_id);
  TEST_ASSERT_EQUAL_INT(xlsx_predefined_style_types[style->style_id], style->related_type);
  TEST_ASSERT_EQUAL_STRING(xlsx_predefined_styles_format_code[style->style_id], style->format_code);

  // appears in cell I3 as a number (scientific notation)
  style = wb.styles[46];
  TEST_ASSERT_EQUAL_INT(203, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_type);
  TEST_ASSERT_EQUAL_STRING("0.0E+00", style->format_code);

  // 49, default text style
  style = wb.styles[47];
  TEST_ASSERT_EQUAL_INT(49, style->style_id);
  TEST_ASSERT_EQUAL_INT(xlsx_predefined_style_types[style->style_id], style->related_type);
  TEST_ASSERT_EQUAL_STRING(xlsx_predefined_styles_format_code[style->style_id], style->format_code);

  // appears in cell K2 as a number (0 padded)
  style = wb.styles[48];
  TEST_ASSERT_EQUAL_INT(204, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_type);
  TEST_ASSERT_EQUAL_STRING("00000", style->format_code);

  // appears in cell K3 as a number
  style = wb.styles[49];
  TEST_ASSERT_EQUAL_INT(205, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_type);
  TEST_ASSERT_EQUAL_STRING("00000\\-0000", style->format_code);

  // appears in cell K4 as a number
  style = wb.styles[50];
  TEST_ASSERT_EQUAL_INT(206, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_type);
  TEST_ASSERT_EQUAL_STRING("[<=9999999]###\\-####;\\(###\\)\\ ###\\-####", style->format_code);

  // appears in cell K5 as a number
  style = wb.styles[51];
  TEST_ASSERT_EQUAL_INT(207, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_type);
  TEST_ASSERT_EQUAL_STRING("000\\-00\\-0000", style->format_code);

  // appears in cell L2 as a sort of currency
  style = wb.styles[52];
  TEST_ASSERT_EQUAL_INT(164, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_type);
  TEST_ASSERT_EQUAL_STRING("_(\"$\"* #,##0_);_(\"$\"* \\(#,##0\\);_(\"$\"* \"-\"_);_(@_)", style->format_code);

  style = wb.styles[53];
  TEST_ASSERT_EQUAL_INT(16, style->style_id);
  TEST_ASSERT_EQUAL_INT(xlsx_predefined_style_types[style->style_id], style->related_type);
  TEST_ASSERT_EQUAL_STRING(xlsx_predefined_styles_format_code[style->style_id], style->format_code);

  style = wb.styles[54];
  TEST_ASSERT_EQUAL_INT(1, style->style_id);
  TEST_ASSERT_EQUAL_INT(xlsx_predefined_style_types[style->style_id], style->related_type);
  TEST_ASSERT_EQUAL_STRING(xlsx_predefined_styles_format_code[style->style_id], style->format_code);

  // sheet related things
  TEST_ASSERT_EQUAL_INT(3, wb.n_sheets);

  TEST_ASSERT_EQUAL_STRING("Sheet1", wb.sheets[0]->name);
  TEST_ASSERT_EQUAL_STRING("second sheet", wb.sheets[1]->name);
  TEST_ASSERT_EQUAL_STRING("third one", wb.sheets[2]->name);

  xlsx_close(&wb);
}


void test_xlsx_load_sheet(void) {
  xlsx_workbook_t wb;
  xlsx_open("test\\helpers\\sample.xlsx", &wb); // ATTENTION: Must be closed

  xlsx_sheet_t *sheet_1, *sheet_2, *sheet_3;
  sheet_1 = xlsx_load_sheet(&wb, 1, NULL);
  TEST_ASSERT_NOT_NULL(sheet_1);
  TEST_ASSERT_NOT_NULL(sheet_1->sheet_xml);
  TEST_ASSERT_NOT_NULL(sheet_1->sheetdata);
  TEST_ASSERT_EQUAL_INT(23, sheet_1->last_row);

  sheet_2 = xlsx_load_sheet(&wb, 0, "second sheet");
  TEST_ASSERT_NOT_NULL(sheet_2);
  TEST_ASSERT_NOT_NULL(sheet_2->sheet_xml);
  TEST_ASSERT_NOT_NULL(sheet_2->sheetdata);
  TEST_ASSERT_EQUAL_INT(5, sheet_2->last_row);

  sheet_3 = xlsx_load_sheet(&wb, 0, "third one");
  TEST_ASSERT_NOT_NULL(sheet_3);
  TEST_ASSERT_NOT_NULL(sheet_3->sheet_xml);
  TEST_ASSERT_NOT_NULL(sheet_3->sheetdata);
  TEST_ASSERT_EQUAL_INT(0, sheet_3->last_row);

  xlsx_close(&wb);
}


void test_xlsx_unload_sheet(void) {
  xlsx_workbook_t wb;
  xlsx_open("test\\helpers\\sample.xlsx", &wb); // ATTENTION: Must be closed

  xlsx_sheet_t *sheet_1 = xlsx_load_sheet(&wb, 0, "Sheet1");
  xlsx_unload_sheet(sheet_1);
  TEST_ASSERT_NULL(sheet_1->name);
  TEST_ASSERT_NULL(sheet_1->sheet_xml);

  xlsx_close(&wb);
}


void test_xlsx_read_cell(void) {
  // setup
  xlsx_workbook_t wb;
  xlsx_open("test\\helpers\\sample.xlsx", &wb); // ATTENTION: Must be closed
  xlsx_sheet_t *sheet_1 = xlsx_load_sheet(&wb, 1, NULL);
  xlsx_cell_t cell_data_holder;

  /* WIP:
  * - All 3 members of cell_data_holder must be tested per cell read
  * - sheet_1->last_row_looked.row_n & sheet_1->last_row_looked.sheetdata_child_i must be tested, is the only way to
  * test the searching algoritm through public functions.
  */
  xlsx_read_cell(sheet_1, 2, "A", &cell_data_holder);


  // teardown
  xlsx_close(&wb);
}


void test_xlsx_close(void) {
  xlsx_workbook_t wb;
  xlsx_open("test\\helpers\\sample.xlsx", &wb); // ATTENTION: Must be closed
  TEST_ASSERT_EQUAL_INT(1, xlsx_close(&wb));
  TEST_ASSERT_NULL(wb.shared_strings_xml);
  TEST_ASSERT_NULL(wb.styles);
  TEST_ASSERT_NULL(wb.sheets);
  TEST_ASSERT_NULL(wb.deployment_path);
}


int main(void) {
  UNITY_BEGIN();
  // put every test to be run here
  RUN_TEST(test_xlsx_open);
  RUN_TEST(test_xlsx_load_sheet);
  RUN_TEST(test_xlsx_unload_sheet);
  RUN_TEST(test_xlsx_read_cell);
  RUN_TEST(test_xlsx_close);
  return UNITY_END();
}
