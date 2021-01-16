#include "../ext/unity.h"
#include "../src/xlsx_drone.h"

/* even empty, this functions need to be here */
void setUp(void) {}
void tearDown(void) {}


void test_xlsx_open(void) {
  // ATTENTION: pass 1 to show error output, 0 if not needed
  xlsx_set_print_err_messages(0);

  xlsx_workbook_t wb;

  // non-existent path
  TEST_ASSERT_EQUAL_INT(0, xlsx_open("test\\helpers\\non_existent.xlsx", &wb));
  TEST_ASSERT_EQUAL_INT(-3, xlsx_get_xlsx_errno());

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
  TEST_ASSERT_EQUAL_INT(xlsx_predefined_style_types[style->style_id], style->related_category);
  TEST_ASSERT_EQUAL_STRING(xlsx_predefined_styles_format_code[style->style_id], style->format_code);

  // appears in cell B2 & B6 as numbers
  style = wb.styles[1];
  TEST_ASSERT_EQUAL_INT(165, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_category);
  TEST_ASSERT_EQUAL_STRING("0.000", style->format_code);

  style = wb.styles[2];
  TEST_ASSERT_EQUAL_INT(4, style->style_id);
  TEST_ASSERT_EQUAL_INT(xlsx_predefined_style_types[style->style_id], style->related_category);
  TEST_ASSERT_EQUAL_STRING(xlsx_predefined_styles_format_code[style->style_id], style->format_code);

  style = wb.styles[3];
  TEST_ASSERT_EQUAL_INT(2, style->style_id);
  TEST_ASSERT_EQUAL_INT(xlsx_predefined_style_types[style->style_id], style->related_category);
  TEST_ASSERT_EQUAL_STRING(xlsx_predefined_styles_format_code[style->style_id], style->format_code);

  // appears in cell B5 as number
  style = wb.styles[4];
  TEST_ASSERT_EQUAL_INT(166, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_category);
  TEST_ASSERT_EQUAL_STRING("0.000;[Red]0.000", style->format_code);

  // appears in cell C2 & C3 as currency
  style = wb.styles[5];
  TEST_ASSERT_EQUAL_INT(167, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_category);
  TEST_ASSERT_EQUAL_STRING("\"$\"#,##0.00", style->format_code);

  // appears in cell C4 as currency
  style = wb.styles[6];
  TEST_ASSERT_EQUAL_INT(168, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_category);
  // ATTENTION: the ASCII char isn't shown in the debug but does the match (assert passes)
  TEST_ASSERT_EQUAL_STRING("#,##0.00\\ [$֏-42B]", style->format_code);

  // appears in cell D2 as currency (accounting)
  style = wb.styles[7];
  TEST_ASSERT_EQUAL_INT(169, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_category);
  TEST_ASSERT_EQUAL_STRING("_(\"$\"* #,##0.000_);_(\"$\"* \\(#,##0.000\\);_(\"$\"* \"-\"\?\?\?_);_(@_)",
                           style->format_code);

  // appears in cell D3 as currency (accounting)
  style = wb.styles[8];
  TEST_ASSERT_EQUAL_INT(170, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_category);
  TEST_ASSERT_EQUAL_STRING("_-* #,##0.000\\ [$֏-42B]_-;\\-* #,##0.000\\ [$֏-42B]_-;_-* \"-\"\?\?\?\\ [$֏-42B]_-;_-@_-",
                           style->format_code);

  style = wb.styles[9];
  TEST_ASSERT_EQUAL_INT(14, style->style_id);
  TEST_ASSERT_EQUAL_INT(xlsx_predefined_style_types[style->style_id], style->related_category);
  TEST_ASSERT_EQUAL_STRING(xlsx_predefined_styles_format_code[style->style_id], style->format_code);

  // appears in cell E3 as date
  style = wb.styles[10];
  TEST_ASSERT_EQUAL_INT(171, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_category);
  TEST_ASSERT_EQUAL_STRING("[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy", style->format_code);

  // appears in cell E4 as date
  style = wb.styles[11];
  TEST_ASSERT_EQUAL_INT(172, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_category);
  TEST_ASSERT_EQUAL_STRING("yyyy\\-mm\\-dd;@", style->format_code);

  // appears in cell E5 as date
  style = wb.styles[12];
  TEST_ASSERT_EQUAL_INT(173, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_category);
  TEST_ASSERT_EQUAL_STRING("m/d;@", style->format_code);

  // appears in cell E6 & E21 & E22 & E23 as date
  style = wb.styles[13];
  TEST_ASSERT_EQUAL_INT(174, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_category);
  TEST_ASSERT_EQUAL_STRING("m/d/yy;@", style->format_code);

  // appears in cell E7 as date
  style = wb.styles[14];
  TEST_ASSERT_EQUAL_INT(175, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_category);
  TEST_ASSERT_EQUAL_STRING("mm/dd/yy;@", style->format_code);

  // appears in cell E8 as date
  style = wb.styles[15];
  TEST_ASSERT_EQUAL_INT(176, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_category);
  TEST_ASSERT_EQUAL_STRING("[$-409]d\\-mmm;@", style->format_code);

  // appears in cell E9 as date
  style = wb.styles[16];
  TEST_ASSERT_EQUAL_INT(177, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_category);
  TEST_ASSERT_EQUAL_STRING("[$-409]d\\-mmm\\-yy;@", style->format_code);

  // appears in cell E10 as date
  style = wb.styles[17];
  TEST_ASSERT_EQUAL_INT(178, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_category);
  TEST_ASSERT_EQUAL_STRING("[$-409]mmm\\-yy;@", style->format_code);

  // appears in cell E11 as date
  style = wb.styles[18];
  TEST_ASSERT_EQUAL_INT(179, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_category);
  TEST_ASSERT_EQUAL_STRING("[$-409]mmmm\\-yy;@", style->format_code);

  // appears in cell E12 as date
  style = wb.styles[19];
  TEST_ASSERT_EQUAL_INT(180, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_category);
  TEST_ASSERT_EQUAL_STRING("[$-409]mmmm\\ d\\,\\ yyyy;@", style->format_code);

  // appears in cell F9 & F12 & E13 & F13 & F14 as date-time
  style = wb.styles[20];
  TEST_ASSERT_EQUAL_INT(181, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE_TIME, style->related_category);
  TEST_ASSERT_EQUAL_STRING("[$-409]m/d/yy\\ h:mm\\ AM/PM;@", style->format_code);

  // appears in cell F10 & E14 as date-time
  style = wb.styles[21];
  TEST_ASSERT_EQUAL_INT(182, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE_TIME, style->related_category);
  TEST_ASSERT_EQUAL_STRING("m/d/yy\\ h:mm;@", style->format_code);

  // appears in cell E15 as date
  style = wb.styles[22];
  TEST_ASSERT_EQUAL_INT(183, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_category);
  TEST_ASSERT_EQUAL_STRING("[$-409]mmmmm;@", style->format_code);

  // appears in cell E16 as date
  style = wb.styles[23];
  TEST_ASSERT_EQUAL_INT(184, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_category);
  TEST_ASSERT_EQUAL_STRING("[$-409]mmmmm\\-yy;@", style->format_code);

  // appears in cell E17 as date
  style = wb.styles[24];
  TEST_ASSERT_EQUAL_INT(185, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_category);
  TEST_ASSERT_EQUAL_STRING("m/d/yyyy;@", style->format_code);

  // appears in cell E18 as date
  style = wb.styles[25];
  TEST_ASSERT_EQUAL_INT(186, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_category);
  TEST_ASSERT_EQUAL_STRING("[$-409]d\\-mmm\\-yyyy;@", style->format_code);

  // appears in cell E19 as date
  style = wb.styles[26];
  TEST_ASSERT_EQUAL_INT(187, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, style->related_category);
  TEST_ASSERT_EQUAL_STRING("[$-2010000]yyyy/mm/dd;@", style->format_code);

  // appears in cell F2 as time
  style = wb.styles[27];
  TEST_ASSERT_EQUAL_INT(188, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_TIME, style->related_category);
  TEST_ASSERT_EQUAL_STRING("[$-F400]h:mm:ss\\ AM/PM", style->format_code);

  // appears in cell F3 as time
  style = wb.styles[28];
  TEST_ASSERT_EQUAL_INT(189, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_TIME, style->related_category);
  TEST_ASSERT_EQUAL_STRING("h:mm;@", style->format_code);

  // appears in cell F4 as time
  style = wb.styles[29];
  TEST_ASSERT_EQUAL_INT(190, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_TIME, style->related_category);
  TEST_ASSERT_EQUAL_STRING("[$-409]h:mm\\ AM/PM;@", style->format_code);

  // appears in cell F5 as time
  style = wb.styles[30];
  TEST_ASSERT_EQUAL_INT(191, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_TIME, style->related_category);
  TEST_ASSERT_EQUAL_STRING("h:mm:ss;@", style->format_code);

  // appears in cell F6 as time
  style = wb.styles[31];
  TEST_ASSERT_EQUAL_INT(192, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_TIME, style->related_category);
  TEST_ASSERT_EQUAL_STRING("[$-409]h:mm:ss\\ AM/PM;@", style->format_code);

  // appears in cell F7 as time
  style = wb.styles[32];
  TEST_ASSERT_EQUAL_INT(193, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_TIME, style->related_category);
  TEST_ASSERT_EQUAL_STRING("mm:ss.0;@", style->format_code);

  // appears in cell F8 as time
  style = wb.styles[33];
  TEST_ASSERT_EQUAL_INT(194, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_TIME, style->related_category);
  TEST_ASSERT_EQUAL_STRING("[h]:mm:ss;@", style->format_code);

  style = wb.styles[34];
  TEST_ASSERT_EQUAL_INT(10, style->style_id);
  TEST_ASSERT_EQUAL_INT(xlsx_predefined_style_types[style->style_id], style->related_category);
  TEST_ASSERT_EQUAL_STRING(xlsx_predefined_styles_format_code[style->style_id], style->format_code);

  // appears in cell G3 as percentage
  style = wb.styles[35];
  TEST_ASSERT_EQUAL_INT(195, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_category);
  TEST_ASSERT_EQUAL_STRING("0.000%", style->format_code);

  style = wb.styles[36];
  TEST_ASSERT_EQUAL_INT(12, style->style_id);
  TEST_ASSERT_EQUAL_INT(xlsx_predefined_style_types[style->style_id], style->related_category);
  TEST_ASSERT_EQUAL_STRING(xlsx_predefined_styles_format_code[style->style_id], style->format_code);

  style = wb.styles[37];
  TEST_ASSERT_EQUAL_INT(13, style->style_id);
  TEST_ASSERT_EQUAL_INT(xlsx_predefined_style_types[style->style_id], style->related_category);
  TEST_ASSERT_EQUAL_STRING(xlsx_predefined_styles_format_code[style->style_id], style->format_code);

  // appears in cell H4 as a fraction
  style = wb.styles[38];
  TEST_ASSERT_EQUAL_INT(196, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_category);
  TEST_ASSERT_EQUAL_STRING("#\\ \?\?\?/\?\?\?", style->format_code);

  // appears in cell H5 as a fraction
  style = wb.styles[39];
  TEST_ASSERT_EQUAL_INT(197, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_category);
  TEST_ASSERT_EQUAL_STRING("#\\ ?/2", style->format_code);

  // appears in cell H6 as a fraction
  style = wb.styles[40];
  TEST_ASSERT_EQUAL_INT(198, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_category);
  TEST_ASSERT_EQUAL_STRING("#\\ ?/4", style->format_code);

  // appears in cell H7 as a fraction
  style = wb.styles[41];
  TEST_ASSERT_EQUAL_INT(199, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_category);
  TEST_ASSERT_EQUAL_STRING("#\\ ?/8", style->format_code);

  // appears in cell H8 as a fraction
  style = wb.styles[42];
  TEST_ASSERT_EQUAL_INT(200, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_category);
  TEST_ASSERT_EQUAL_STRING("#\\ \?\?/16", style->format_code);

  // appears in cell H9 as a fraction
  style = wb.styles[43];
  TEST_ASSERT_EQUAL_INT(201, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_category);
  TEST_ASSERT_EQUAL_STRING("#\\ ?/10", style->format_code);

  // appears in cell H10 as a fraction
  style = wb.styles[44];
  TEST_ASSERT_EQUAL_INT(202, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_category);
  TEST_ASSERT_EQUAL_STRING("#\\ \?\?/100", style->format_code);

  style = wb.styles[45];
  TEST_ASSERT_EQUAL_INT(11, style->style_id);
  TEST_ASSERT_EQUAL_INT(xlsx_predefined_style_types[style->style_id], style->related_category);
  TEST_ASSERT_EQUAL_STRING(xlsx_predefined_styles_format_code[style->style_id], style->format_code);

  // appears in cell I3 as a number (scientific notation)
  style = wb.styles[46];
  TEST_ASSERT_EQUAL_INT(203, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_category);
  TEST_ASSERT_EQUAL_STRING("0.0E+00", style->format_code);

  // 49, default text style
  style = wb.styles[47];
  TEST_ASSERT_EQUAL_INT(49, style->style_id);
  TEST_ASSERT_EQUAL_INT(xlsx_predefined_style_types[style->style_id], style->related_category);
  TEST_ASSERT_EQUAL_STRING(xlsx_predefined_styles_format_code[style->style_id], style->format_code);

  // appears in cell K2 as a number (0 padded)
  style = wb.styles[48];
  TEST_ASSERT_EQUAL_INT(204, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_category);
  TEST_ASSERT_EQUAL_STRING("00000", style->format_code);

  // appears in cell K3 as a number
  style = wb.styles[49];
  TEST_ASSERT_EQUAL_INT(205, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_category);
  TEST_ASSERT_EQUAL_STRING("00000\\-0000", style->format_code);

  // appears in cell K4 as a number
  style = wb.styles[50];
  TEST_ASSERT_EQUAL_INT(206, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_category);
  TEST_ASSERT_EQUAL_STRING("[<=9999999]###\\-####;\\(###\\)\\ ###\\-####", style->format_code);

  // appears in cell K5 as a number
  style = wb.styles[51];
  TEST_ASSERT_EQUAL_INT(207, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_category);
  TEST_ASSERT_EQUAL_STRING("000\\-00\\-0000", style->format_code);

  // appears in cell L2 as a sort of currency
  style = wb.styles[52];
  TEST_ASSERT_EQUAL_INT(164, style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, style->related_category);
  TEST_ASSERT_EQUAL_STRING("_(\"$\"* #,##0_);_(\"$\"* \\(#,##0\\);_(\"$\"* \"-\"_);_(@_)", style->format_code);

  style = wb.styles[53];
  TEST_ASSERT_EQUAL_INT(16, style->style_id);
  TEST_ASSERT_EQUAL_INT(xlsx_predefined_style_types[style->style_id], style->related_category);
  TEST_ASSERT_EQUAL_STRING(xlsx_predefined_styles_format_code[style->style_id], style->format_code);

  style = wb.styles[54];
  TEST_ASSERT_EQUAL_INT(1, style->style_id);
  TEST_ASSERT_EQUAL_INT(xlsx_predefined_style_types[style->style_id], style->related_category);
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
  TEST_ASSERT_EQUAL_INT(0, xlsx_get_xlsx_errno());
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

  // non-existent sheet
  TEST_ASSERT_NULL(xlsx_load_sheet(&wb, 0, "non existent"));
  TEST_ASSERT_EQUAL_INT(XLSX_LOAD_SHEET_ERRNO_NON_EXISTENT, xlsx_get_xlsx_errno());
  TEST_ASSERT_NULL(xlsx_load_sheet(&wb, 6, NULL));
  TEST_ASSERT_EQUAL_INT(XLSX_LOAD_SHEET_ERRNO_INDEX_OUT_OF_BOUNDS, xlsx_get_xlsx_errno());

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


void test_xlsx_get_last_column(void) {
  xlsx_workbook_t wb;
  xlsx_open("test\\helpers\\empty_sample.xlsx", &wb); // ATTENTION: Must be closed

  // before loading a sheet, we expect an error
  TEST_ASSERT_NULL(xlsx_get_last_column(wb.sheets[0]));
  TEST_ASSERT_EQUAL_INT(-31, xlsx_get_xlsx_errno());

  // if the sheet is empty, should return NULL
  xlsx_sheet_t *sheet;
  sheet = xlsx_load_sheet(&wb, 1, NULL);
  TEST_ASSERT_NULL(xlsx_get_last_column(sheet));
  TEST_ASSERT_EQUAL_INT(0, xlsx_get_xlsx_errno());

  // this sheet is also empty, but has styles defined in several cells
  sheet = xlsx_load_sheet(&wb, 2, NULL);
  TEST_ASSERT_NULL(xlsx_get_last_column(sheet));
  TEST_ASSERT_EQUAL_INT(0, xlsx_get_xlsx_errno());

  xlsx_close(&wb);

  xlsx_open("test\\helpers\\sample.xlsx", &wb); // ATTENTION: Must be closed
  // if the sheet is empty, should return NULL
  sheet = xlsx_load_sheet(&wb, 1, NULL);
  TEST_ASSERT_EQUAL_STRING("L", xlsx_get_last_column(sheet));
  // if asked again, should return same result
  TEST_ASSERT_EQUAL_STRING("L", xlsx_get_last_column(sheet));
  // after asking once, you can directly retrieve the value
  TEST_ASSERT_EQUAL_STRING("L", sheet->last_column);

  // this sheet has its las value on column "B" but has styles defined at the right, should return "B"
  sheet = xlsx_load_sheet(&wb, 2, NULL);
  TEST_ASSERT_EQUAL_STRING("B", xlsx_get_last_column(sheet));

  xlsx_close(&wb);
}


void test_xlsx_read_cell(void) {
  // setup
  xlsx_workbook_t wb;
  xlsx_open("test\\helpers\\sample.xlsx", &wb); // ATTENTION: Must be closed
  xlsx_sheet_t *sheet_1 = xlsx_load_sheet(&wb, 1, NULL);
  xlsx_cell_t cell_data_holder;

  // A2, General, "Foo"
  xlsx_read_cell(sheet_1, 2, "A", &cell_data_holder);
  TEST_ASSERT_NULL(cell_data_holder.style);
  TEST_ASSERT_EQUAL_INT(XLSX_POINTER_TO_CHAR, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_STRING("Foo", cell_data_holder.value.pointer_to_char_value);
  TEST_ASSERT_EQUAL_INT(2, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(1, sheet_1->last_row_looked.sheetdata_child_i);

  // A3, General, 235
  xlsx_read_cell(sheet_1, 3, "A", &cell_data_holder);
  TEST_ASSERT_NULL(cell_data_holder.style);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(235, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(3, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(2, sheet_1->last_row_looked.sheetdata_child_i);

  // A4, General, 17.89
  xlsx_read_cell(sheet_1, 4, "A", &cell_data_holder);
  TEST_ASSERT_NULL(cell_data_holder.style);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(17.89, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(4, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(3, sheet_1->last_row_looked.sheetdata_child_i);

  // B2, Number, 1000.000
  xlsx_read_cell(sheet_1, 2, "B", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(165, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(1000, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(2, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(1, sheet_1->last_row_looked.sheetdata_child_i);

  // B3, Number, 1,000.00
  xlsx_read_cell(sheet_1, 3, "B", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(4, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(1000, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(3, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(2, sheet_1->last_row_looked.sheetdata_child_i);

  // B4, Number, -1000.00
  xlsx_read_cell(sheet_1, 4, "B", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(2, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(-1000, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(4, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(3, sheet_1->last_row_looked.sheetdata_child_i);

  // B5, Number, 1000.000 (painted in red, hence negative)
  xlsx_read_cell(sheet_1, 5, "B", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(166, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(-1000, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(5, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(4, sheet_1->last_row_looked.sheetdata_child_i);

  // B6, Number, 1200.561 (painted in red, hence negative)
  xlsx_read_cell(sheet_1, 6, "B", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(165, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(1200.561, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(6, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(5, sheet_1->last_row_looked.sheetdata_child_i);

  // B7, Number, 123456789
  xlsx_read_cell(sheet_1, 7, "B", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(1, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(123456789, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(7, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(6, sheet_1->last_row_looked.sheetdata_child_i);

  // B8, Number, 1234567890
  xlsx_read_cell(sheet_1, 8, "B", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(1, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_LONG_LONG, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT64(1234567890, cell_data_holder.value.long_long_value);
  TEST_ASSERT_EQUAL_INT(8, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(7, sheet_1->last_row_looked.sheetdata_child_i);

  // B9, Number, 2345678901
  xlsx_read_cell(sheet_1, 9, "B", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(1, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_LONG_LONG, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT64(2345678901, cell_data_holder.value.long_long_value);
  TEST_ASSERT_EQUAL_INT(9, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(8, sheet_1->last_row_looked.sheetdata_child_i);

  // B10, Number, 5678901234
  xlsx_read_cell(sheet_1, 10, "B", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(1, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_LONG_LONG, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT64(5678901234, cell_data_holder.value.long_long_value);
  TEST_ASSERT_EQUAL_INT(10, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(9, sheet_1->last_row_looked.sheetdata_child_i);

  // B11, Number, 123456789012345
  xlsx_read_cell(sheet_1, 11, "B", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(1, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_LONG_LONG, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT64(123456789012345, cell_data_holder.value.long_long_value);
  TEST_ASSERT_EQUAL_INT(11, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(10, sheet_1->last_row_looked.sheetdata_child_i);

  // B12, Number, 1234567890123450
  xlsx_read_cell(sheet_1, 12, "B", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(1, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_LONG_LONG, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT64(1234567890123450, cell_data_holder.value.long_long_value);
  TEST_ASSERT_EQUAL_INT(12, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(11, sheet_1->last_row_looked.sheetdata_child_i);

  // B13, Number, 1234567890123450000 (see documentation to understand why this number is set as a double)
  xlsx_read_cell(sheet_1, 13, "B", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(1, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(1.2345678901234501E+18, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(13, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(12, sheet_1->last_row_looked.sheetdata_child_i);

  // B14, Number, 12345678901234500000 (see documentation to understand why this number is set as a double)
  xlsx_read_cell(sheet_1, 14, "B", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(1, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(1.23456789012345E+19, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(14, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(13, sheet_1->last_row_looked.sheetdata_child_i);

  // B15, Number, 12345678901234500 (see documentation to understand why this number is set as a double)
  xlsx_read_cell(sheet_1, 15, "B", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(1, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(1.23456789012345E+16, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(15, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(14, sheet_1->last_row_looked.sheetdata_child_i);

  // B16, Number, 123456789012345000 (see documentation to understand why this number is set as a double)
  xlsx_read_cell(sheet_1, 16, "B", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(1, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(1.23456789012345E+17, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(16, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(15, sheet_1->last_row_looked.sheetdata_child_i);

  // C2, Currency, $1,000.00
  xlsx_read_cell(sheet_1, 2, "C", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(167, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(1000, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(2, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(1, sheet_1->last_row_looked.sheetdata_child_i);

  // C3, Currency, -$14,562.74
  xlsx_read_cell(sheet_1, 3, "C", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(167, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(-14562.74, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(3, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(2, sheet_1->last_row_looked.sheetdata_child_i);

  // C4, Currency, 584.00 (with non-ASCII character as currency)
  xlsx_read_cell(sheet_1, 4, "C", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(168, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(584, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(4, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(3, sheet_1->last_row_looked.sheetdata_child_i);

  // D2, Accounting, 147.000
  xlsx_read_cell(sheet_1, 2, "D", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(169, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(147, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(2, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(1, sheet_1->last_row_looked.sheetdata_child_i);

  // D3, Accounting, 1,200.874 (with non-ASCII character as currency)
  xlsx_read_cell(sheet_1, 3, "D", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(170, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(1200.874, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(3, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(2, sheet_1->last_row_looked.sheetdata_child_i);

  // E2, Date, 24/12/2018
  xlsx_read_cell(sheet_1, 2, "E", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(14, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(43458, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(2, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(1, sheet_1->last_row_looked.sheetdata_child_i);

  // E3, Date, "lunes, 24 de diciembre de 2018" (tested with spanish lang set)
  xlsx_read_cell(sheet_1, 3, "E", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(171, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(43458, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(3, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(2, sheet_1->last_row_looked.sheetdata_child_i);

  // E4, Date, 2018-12-24
  xlsx_read_cell(sheet_1, 4, "E", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(172, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(43458, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(4, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(3, sheet_1->last_row_looked.sheetdata_child_i);

  // E5, Date, 12/24
  xlsx_read_cell(sheet_1, 5, "E", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(173, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(43458, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(5, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(4, sheet_1->last_row_looked.sheetdata_child_i);

  // E6, Date, 12/24/18
  xlsx_read_cell(sheet_1, 6, "E", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(174, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(43458, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(6, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(5, sheet_1->last_row_looked.sheetdata_child_i);

  // E7, Date, 12/24/18
  xlsx_read_cell(sheet_1, 7, "E", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(175, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(43458, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(7, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(6, sheet_1->last_row_looked.sheetdata_child_i);

  // E8, Date, 24-Dec
  xlsx_read_cell(sheet_1, 8, "E", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(176, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(43458, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(8, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(7, sheet_1->last_row_looked.sheetdata_child_i);

  // E9, Date, 24-Dec-18
  xlsx_read_cell(sheet_1, 9, "E", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(177, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(43458, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(9, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(8, sheet_1->last_row_looked.sheetdata_child_i);

  // E10, Date, Dec-18
  xlsx_read_cell(sheet_1, 10, "E", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(178, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(43458, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(10, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(9, sheet_1->last_row_looked.sheetdata_child_i);

  // E11, Date, December-18
  xlsx_read_cell(sheet_1, 11, "E", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(179, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(43458, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(11, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(10, sheet_1->last_row_looked.sheetdata_child_i);

  // E12, Date, "December 24, 2018"
  xlsx_read_cell(sheet_1, 12, "E", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(180, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(43458, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(12, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(11, sheet_1->last_row_looked.sheetdata_child_i);

  // E13, Date, 12/24/18 12:00 AM
  xlsx_read_cell(sheet_1, 13, "E", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(181, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE_TIME, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(43458, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(13, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(12, sheet_1->last_row_looked.sheetdata_child_i);

  // E14, Date, 12/24/18 0:00
  xlsx_read_cell(sheet_1, 14, "E", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(182, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE_TIME, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(43458, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(14, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(13, sheet_1->last_row_looked.sheetdata_child_i);

  // E15, Date, D
  xlsx_read_cell(sheet_1, 15, "E", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(183, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(43458, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(15, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(14, sheet_1->last_row_looked.sheetdata_child_i);

  // E16, Date, D-18
  xlsx_read_cell(sheet_1, 16, "E", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(184, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(43458, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(16, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(15, sheet_1->last_row_looked.sheetdata_child_i);

  // E17, Date, 12/24/2018
  xlsx_read_cell(sheet_1, 17, "E", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(185, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(43458, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(17, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(16, sheet_1->last_row_looked.sheetdata_child_i);

  // E18, Date, 24-Dec-2018
  xlsx_read_cell(sheet_1, 18, "E", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(186, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(43458, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(18, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(17, sheet_1->last_row_looked.sheetdata_child_i);

  // E19, Date, (Arabic (Algeria) style)
  xlsx_read_cell(sheet_1, 19, "E", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(187, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(43458, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(19, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(18, sheet_1->last_row_looked.sheetdata_child_i);

  // E20, Date, 43458
  xlsx_read_cell(sheet_1, 20, "E", &cell_data_holder);
  TEST_ASSERT_NULL(cell_data_holder.style);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(43458, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(20, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(19, sheet_1->last_row_looked.sheetdata_child_i);

  // E21, Date, 12/24/1154 (special case, is considered date, but is saved as string since "out of range" date)
  xlsx_read_cell(sheet_1, 21, "E", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(174, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_POINTER_TO_CHAR, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_STRING("12/24/1154", cell_data_holder.value.pointer_to_char_value);
  TEST_ASSERT_EQUAL_INT(21, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(20, sheet_1->last_row_looked.sheetdata_child_i);

  // E22, Date, 12/24/99
  xlsx_read_cell(sheet_1, 22, "E", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(174, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(2958458, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(22, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(21, sheet_1->last_row_looked.sheetdata_child_i);

  /* E23, Date, text
  * Here we are presented with a great contradiction. Cell format is set do "date", but user wrote "text" in the cell.
  * The data was saved as a string, but the style associated has a "date" format. So what should we do here regarding
  * the related_category? Date or text? Big contradiction since in cell E21 we have this problem but as we saw, there the
  * user inputted an old date, out of range for the standard that relates a number to a date, but the user actually
  * tried to input a date. In my opinion, we should state this realted_type as Date, having in consideration the case of
  * cell E21 which could be more usual than inputting text there.
  */
  xlsx_read_cell(sheet_1, 23, "E", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(174, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_POINTER_TO_CHAR, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_STRING("text", cell_data_holder.value.pointer_to_char_value);
  TEST_ASSERT_EQUAL_INT(23, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(22, sheet_1->last_row_looked.sheetdata_child_i);

  // F2, Time, 2:30:54 a. m.
  xlsx_read_cell(sheet_1, 2, "F", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(188, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_TIME, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(0.10479166666666667, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(2, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(1, sheet_1->last_row_looked.sheetdata_child_i);

  // F3, Time, 2:30
  xlsx_read_cell(sheet_1, 3, "F", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(189, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_TIME, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(0.10479166666666667, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(3, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(2, sheet_1->last_row_looked.sheetdata_child_i);

  // F4, Time, 2:30 AM
  xlsx_read_cell(sheet_1, 4, "F", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(190, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_TIME, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(0.10479166666666667, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(4, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(3, sheet_1->last_row_looked.sheetdata_child_i);

  // F5, Time, 2:30:54
  xlsx_read_cell(sheet_1, 5, "F", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(191, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_TIME, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(0.10479166666666667, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(5, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(4, sheet_1->last_row_looked.sheetdata_child_i);

  // F6, Time, 2:30:54 AM
  xlsx_read_cell(sheet_1, 6, "F", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(192, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_TIME, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(0.10479166666666667, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(6, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(5, sheet_1->last_row_looked.sheetdata_child_i);

  // F7, Time, 30:54.0
  xlsx_read_cell(sheet_1, 7, "F", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(193, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_TIME, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(0.10479166666666667, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(7, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(6, sheet_1->last_row_looked.sheetdata_child_i);

  // F8, Time, 2:30:54
  xlsx_read_cell(sheet_1, 8, "F", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(194, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_TIME, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(0.10479166666666667, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(8, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(7, sheet_1->last_row_looked.sheetdata_child_i);

  // F9, Time, 1/3/56 2:30 AM
  xlsx_read_cell(sheet_1, 9, "F", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(181, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE_TIME, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(20457.104791666668, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(9, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(8, sheet_1->last_row_looked.sheetdata_child_i);

  // F10, Time, 1/0/00 2:30 (is day "0" of 1900, computes as 0 for date)
  xlsx_read_cell(sheet_1, 10, "F", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(182, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE_TIME, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(0.10479166666666667, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(10, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(9, sheet_1->last_row_looked.sheetdata_child_i);

  // F11, Time, 0,104791667
  xlsx_read_cell(sheet_1, 11, "F", &cell_data_holder);
  TEST_ASSERT_NULL(cell_data_holder.style);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(0.104791666666667, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(11, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(10, sheet_1->last_row_looked.sheetdata_child_i);

  // F12, Time, 1/1/1889  2:30:54 AM (OOR for date representation as int/float (1900~9999?), it's saved as text)
  xlsx_read_cell(sheet_1, 12, "F", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(181, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE_TIME, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_POINTER_TO_CHAR, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_STRING("1/1/1889  2:30:54 AM", cell_data_holder.value.pointer_to_char_value);
  TEST_ASSERT_EQUAL_INT(12, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(11, sheet_1->last_row_looked.sheetdata_child_i);

  // F13, Time, 1/1/89  2:30 (comparing with the previous value, this is 1989, not 1889)
  xlsx_read_cell(sheet_1, 13, "F", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(181, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE_TIME, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(32509.104791666668, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(13, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(12, sheet_1->last_row_looked.sheetdata_child_i);

  // F14, Time, 1/1/10000  2:30:54 AM (OOR for date representation as int/float (1900~9999?), it's saved as text)
  xlsx_read_cell(sheet_1, 14, "F", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(181, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE_TIME, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_POINTER_TO_CHAR, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_STRING("1/1/10000  2:30:54 AM", cell_data_holder.value.pointer_to_char_value);
  TEST_ASSERT_EQUAL_INT(14, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(13, sheet_1->last_row_looked.sheetdata_child_i);

  // G2, Percentage, 50.00%
  xlsx_read_cell(sheet_1, 2, "G", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(10, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(0.5, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(2, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(1, sheet_1->last_row_looked.sheetdata_child_i);

  // G3, Percentage, 45.000%
  xlsx_read_cell(sheet_1, 3, "G", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(195, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(0.45, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(3, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(2, sheet_1->last_row_looked.sheetdata_child_i);

  // G4, Percentage, 160.00%
  xlsx_read_cell(sheet_1, 4, "G", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(10, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(1.6, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(4, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(3, sheet_1->last_row_looked.sheetdata_child_i);

  // H2, Fraction (always 1.5), 1 1/2
  xlsx_read_cell(sheet_1, 2, "H", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(12, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(1.5, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(2, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(1, sheet_1->last_row_looked.sheetdata_child_i);

  // H3, Fraction (always 1.5), 1  1/2
  xlsx_read_cell(sheet_1, 3, "H", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(13, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(1.5, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(3, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(2, sheet_1->last_row_looked.sheetdata_child_i);

  // H4, Fraction (always 1.5), 1   1/2
  xlsx_read_cell(sheet_1, 4, "H", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(196, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(1.5, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(4, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(3, sheet_1->last_row_looked.sheetdata_child_i);

  // H5, Fraction (always 1.5), 1 1/2
  xlsx_read_cell(sheet_1, 5, "H", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(197, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(1.5, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(5, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(4, sheet_1->last_row_looked.sheetdata_child_i);

  // H6, Fraction (always 1.5), 1 2/4
  xlsx_read_cell(sheet_1, 6, "H", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(198, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(1.5, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(6, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(5, sheet_1->last_row_looked.sheetdata_child_i);

  // H7, Fraction (always 1.5), 1 4/8
  xlsx_read_cell(sheet_1, 7, "H", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(199, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(1.5, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(7, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(6, sheet_1->last_row_looked.sheetdata_child_i);

  // H8, Fraction (always 1.5), 1  8/16
  xlsx_read_cell(sheet_1, 8, "H", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(200, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(1.5, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(8, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(7, sheet_1->last_row_looked.sheetdata_child_i);

  // H9, Fraction (always 1.5), 15/10
  xlsx_read_cell(sheet_1, 9, "H", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(201, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(1.5, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(9, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(8, sheet_1->last_row_looked.sheetdata_child_i);

  // H10, Fraction (always 1.5), 150/100
  xlsx_read_cell(sheet_1, 10, "H", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(202, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(1.5, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(10, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(9, sheet_1->last_row_looked.sheetdata_child_i);


  // H11, Fraction (always 1.5) (actually, just General), 1,5
  xlsx_read_cell(sheet_1, 11, "H", &cell_data_holder);
  TEST_ASSERT_NULL(cell_data_holder.style);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(1.5, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(11, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(10, sheet_1->last_row_looked.sheetdata_child_i);

  // H13, Fraction (always 1.5), #VALUE! (this shows an error, an error has no style associated)
  xlsx_read_cell(sheet_1, 13, "H", &cell_data_holder);
  TEST_ASSERT_NULL(cell_data_holder.style);
  TEST_ASSERT_EQUAL_INT(XLSX_POINTER_TO_CHAR, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_STRING("#VALUE!", cell_data_holder.value.pointer_to_char_value);
  TEST_ASSERT_EQUAL_INT(13, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(12, sheet_1->last_row_looked.sheetdata_child_i);

  // I2, Scientific (always 0.001), 1.00E-03
  xlsx_read_cell(sheet_1, 2, "I", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(11, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(1E-3, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(2, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(1, sheet_1->last_row_looked.sheetdata_child_i);

  // I3, Scientific (always 0.001), 1.0E-03
  xlsx_read_cell(sheet_1, 3, "I", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(203, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(1E-3, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(3, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(2, sheet_1->last_row_looked.sheetdata_child_i);

  // I4, Scientific (always 0.001) (actually General), 0.001
  xlsx_read_cell(sheet_1, 4, "I", &cell_data_holder);
  TEST_ASSERT_NULL(cell_data_holder.style);
  TEST_ASSERT_EQUAL_INT(XLSX_DOUBLE, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_DOUBLE(1E-3, cell_data_holder.value.double_value);
  TEST_ASSERT_EQUAL_INT(4, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(3, sheet_1->last_row_looked.sheetdata_child_i);

  // J2, Text, 1875 (Typed as '1875)
  xlsx_read_cell(sheet_1, 2, "J", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(49, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_TEXT, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_POINTER_TO_CHAR, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_STRING("1875", cell_data_holder.value.pointer_to_char_value);
  TEST_ASSERT_EQUAL_INT(2, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(1, sheet_1->last_row_looked.sheetdata_child_i);

  // J3, Text, Just text
  xlsx_read_cell(sheet_1, 3, "J", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(49, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_TEXT, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_POINTER_TO_CHAR, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_STRING("Just text", cell_data_holder.value.pointer_to_char_value);
  TEST_ASSERT_EQUAL_INT(3, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(2, sheet_1->last_row_looked.sheetdata_child_i);

  // J4, Text (Unicode), 𐐀34
  xlsx_read_cell(sheet_1, 4, "J", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(49, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_TEXT, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_POINTER_TO_CHAR, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_STRING(u8"𐐀34", cell_data_holder.value.pointer_to_char_value);
  TEST_ASSERT_EQUAL_INT(4, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(3, sheet_1->last_row_looked.sheetdata_child_i);

  // J5, Text (Unicode), foo你bar好qaz
  xlsx_read_cell(sheet_1, 5, "J", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(49, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_TEXT, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_POINTER_TO_CHAR, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_STRING(u8"foo你bar好qaz", cell_data_holder.value.pointer_to_char_value);
  TEST_ASSERT_EQUAL_INT(5, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(4, sheet_1->last_row_looked.sheetdata_child_i);

  // K2, Special, 02000 (Typed 2000)
  xlsx_read_cell(sheet_1, 2, "K", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(204, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(2000, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(2, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(1, sheet_1->last_row_looked.sheetdata_child_i);

  // K3, Special, 00000-2000 (Typed 2000)
  xlsx_read_cell(sheet_1, 3, "K", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(205, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(2000, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(3, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(2, sheet_1->last_row_looked.sheetdata_child_i);

  // K4, Special, (54341) 563-5644 (Typed 543415635644)
  xlsx_read_cell(sheet_1, 4, "K", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(206, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_LONG_LONG, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT64(543415635644, cell_data_holder.value.long_long_value);
  TEST_ASSERT_EQUAL_INT(4, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(3, sheet_1->last_row_looked.sheetdata_child_i);

  // K5, Special, 034-58-0585 (Typed 34580585)
  xlsx_read_cell(sheet_1, 5, "K", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(207, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(34580585, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(5, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(4, sheet_1->last_row_looked.sheetdata_child_i);

  // L2, Custom, $ 12 (Typed 12)
  xlsx_read_cell(sheet_1, 2, "L", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(164, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_NUMBER, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(12, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(2, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(1, sheet_1->last_row_looked.sheetdata_child_i);

  // L3, Custom, 16-feb (Typed 16/2/2012)
  xlsx_read_cell(sheet_1, 3, "L", &cell_data_holder);
  TEST_ASSERT_EQUAL_INT(16, cell_data_holder.style->style_id);
  TEST_ASSERT_EQUAL_INT(XLSX_DATE, cell_data_holder.style->related_category);
  TEST_ASSERT_EQUAL_INT(XLSX_INT, cell_data_holder.value_type);
  TEST_ASSERT_EQUAL_INT(40955, cell_data_holder.value.int_value);
  TEST_ASSERT_EQUAL_INT(3, sheet_1->last_row_looked.row_n);
  TEST_ASSERT_EQUAL_INT(2, sheet_1->last_row_looked.sheetdata_child_i);

  // Empty cell
  xlsx_read_cell(sheet_1, 13, "D", &cell_data_holder);
  TEST_ASSERT_NULL(cell_data_holder.style);
  TEST_ASSERT_EQUAL_INT(XLSX_NULL, cell_data_holder.value_type);

  // Empty cell
  xlsx_read_cell(sheet_1, 50, "A", &cell_data_holder);
  TEST_ASSERT_NULL(cell_data_holder.style);
  TEST_ASSERT_EQUAL_INT(XLSX_NULL, cell_data_holder.value_type);

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
  RUN_TEST(test_xlsx_get_last_column);
  RUN_TEST(test_xlsx_read_cell);
  RUN_TEST(test_xlsx_close);
  return UNITY_END();
}
