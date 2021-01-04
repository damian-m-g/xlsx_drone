#ifndef PORCUPINE_LIBRARY_H
#define PORCUPINE_LIBRARY_H

// includes
// core libraries
#include <stdlib.h>
#include <string.h>
#include <dirent.h>
#include <unistd.h>
#include <sys/stat.h>

// external libraries
#include "../ext/zip.h" // https://github.com/kuba--/zip | using version 0.1.21 (2020/12)
#include "../ext/sxmlc.h" // http://sxmlc.sourceforge.net/ | using version 4.5.1 (2020/08)
#include "../ext/sxmlsearch.h"


// constants
#ifndef true
  #define true 1
#endif
#ifndef false
  #define false 0
#endif
#ifdef __MINGW32__
  #include <afxres.h>
  #define XLSX_SET_ERRNO(x) SetLastError(x)
#else
  #define XLSX_SET_ERRNO(x) errno=(x)
#endif
#define ENVIRONMENT_VARIABLE_TEMP "TEMP" // target is Windows OS > Win 3.x for now
#define REL_PATH_TO_STYLES "\\xl\\styles.xml"
#define REL_PATH_TO_SHARED_STRINGS "\\xl\\sharedStrings.xml"
#define REL_PATH_TO_WORKBOOK "\\xl\\workbook.xml"
#define REL_PATH_TO_WORKSHEETS "\\xl\\worksheets\\"
#define STYLES_CELLXFS_TAG "cellXfs"
#define STYLES_NUMFMT_TAG "numFmt"
#define AMOUNT_OF_PREDEFINED_STYLE_TYPES 50
#define STYLES_NUMFMTID_ATTR_NAME "numFmtId"
#define STYLES_FORMATCODE_ATTR_NAME "formatCode"
#define WORKBOOK_SHEETS_TAG "sheets"
#define WORKBOOK_NAME_ATTR_NAME "name"
#define SHEET_SHEETDATA_TAG "sheetData"
#define SHEET_ROW_ATTR_NAME "r"
#define SHEET_TYPE_ATTR_NAME "t"
#define SHEET_STYLE_ATTR_NAME "s"
#define SHEET_ROW_TAG "row"
#define SHEET_COL_TAG "c"
#define SHEET_VALUE_TAG "v"


// structures
// represents a workbook
typedef struct xlsx_workbook_t {
  char *deployment_path;
  XMLDoc *shared_strings_xml;
  int n_styles;
  struct xlsx_style_t **styles;
  int n_sheets;
  struct xlsx_sheet_t **sheets;
} xlsx_workbook_t;

typedef enum xlsx_formatter {
  XLSX_FORMATTER_UNKNOWN, // others
  XLSX_FORMATTER_AMBIGUOUS_M, // unescaped 'm'
  XLSX_FORMATTER_TIME, // unescaped 'h' or 's'
  XLSX_FORMATTER_DATE // unescaped 'y' or 'd'
} xlsx_formatter;

typedef enum xlsx_cell_kind {
  XLSX_NUMBER, // int, long long, or double
  XLSX_TEXT,
  XLSX_DATE, // int
  XLSX_TIME, // double
  XLSX_DATE_TIME, // double
  XLSX_UNKNOWN
} xlsx_cell_kind;

// represents a single style
typedef struct xlsx_style_t {
  int style_id;
  enum xlsx_cell_kind related_type;
  char *format_code;
} xlsx_style_t;

// for speeding (caching) purpose, useful for xlsx_read_cell()
typedef struct xlsx_reference_to_row_t {
  int row_n;
  int sheetdata_child_i;
} xlsx_reference_to_row_t;

// represents a sheet of a workbook
typedef struct xlsx_sheet_t {
  struct xlsx_workbook_t *xlsx;
  char *name;
  XMLDoc *sheet_xml;
  XMLNode *sheetdata; // speeding purpose
  int last_row;
  struct xlsx_reference_to_row_t last_row_looked; // speeding purpose
} xlsx_sheet_t;

typedef enum xlsx_value_type {
  XLSX_POINTER_TO_CHAR,
  XLSX_INT,
  XLSX_LONG_LONG,
  XLSX_DOUBLE,
  XLSX_NULL
} xlsx_value_type;

typedef union xlsx_cell_value {
  char *pointer_to_char_value;
  int int_value;
  long long long_long_value;
  double double_value;
} xlsx_cell_value;

// represents a cell value, returned when you read the content of a cell
typedef struct xlsx_cell_t {
  struct xlsx_style_t *style;
  enum xlsx_value_type value_type;
  union xlsx_cell_value value;
} xlsx_cell_t;


// global variables
static int xlsx_errno;
static xlsx_cell_kind xlsx_predefined_style_types[AMOUNT_OF_PREDEFINED_STYLE_TYPES] = {
  XLSX_UNKNOWN, // 0
  XLSX_NUMBER, // 1
  XLSX_NUMBER, // 2
  XLSX_NUMBER, // 3
  XLSX_NUMBER, // 4
  XLSX_NUMBER, // 5
  XLSX_NUMBER, // 6
  XLSX_NUMBER, // 7
  XLSX_NUMBER, // 8
  XLSX_NUMBER, // 9
  XLSX_NUMBER, // 10
  XLSX_NUMBER, // 11
  XLSX_NUMBER, // 12
  XLSX_NUMBER, // 13
  XLSX_DATE, // 14
  XLSX_DATE, // 15
  XLSX_DATE, // 16
  XLSX_DATE, // 17
  XLSX_DATE_TIME, // 18
  XLSX_DATE_TIME, // 19
  XLSX_TIME, // 20
  XLSX_TIME, // 21
  XLSX_DATE_TIME, // 22
  XLSX_UNKNOWN, // 23
  XLSX_UNKNOWN, // 24
  XLSX_UNKNOWN, // 25
  XLSX_UNKNOWN, // 26
  XLSX_UNKNOWN, // 27
  XLSX_UNKNOWN, // 28
  XLSX_UNKNOWN, // 29
  XLSX_UNKNOWN, // 30
  XLSX_UNKNOWN, // 31
  XLSX_UNKNOWN, // 32
  XLSX_UNKNOWN, // 33
  XLSX_UNKNOWN, // 34
  XLSX_UNKNOWN, // 35
  XLSX_UNKNOWN, // 36
  XLSX_NUMBER, // 37
  XLSX_NUMBER, // 38
  XLSX_NUMBER, // 39
  XLSX_NUMBER, // 40
  XLSX_UNKNOWN, // 41
  XLSX_UNKNOWN, // 42
  XLSX_UNKNOWN, // 43
  XLSX_UNKNOWN, // 44
  XLSX_TIME, // 45
  XLSX_TIME, // 46
  XLSX_TIME, // 47
  XLSX_NUMBER, // 48
  XLSX_TEXT // 49
};
static char *xlsx_predefined_styles_format_code[AMOUNT_OF_PREDEFINED_STYLE_TYPES] = {
    NULL, // 0
    "0", // 1
    "0.00", // 2
    "#,##0", // 3
    "#,##0.00", // 4
    "$#,##0_);($#,##0)", // 5
    "$#,##0_);[Red]($#,##0)", // 6
    "$#,##0.00_);($#,##0.00)", // 7
    "$#,##0.00_);[Red]($#,##0.00)", // 8
    "0%", // 9
    "0.00%", // 10
    "0.00E+00", // 11
    "# ?/?", // 12
    "# ?\?/??", // 13
    "d/m/yyyy", // 14
    "d-mmm-yy", // 15
    "d-mmm", // 16
    "mmm-yy", // 17
    "h:mm AM/PM", // 18
    "h:mm:ss AM/PM", // 19
    "h:mm", // 20
    "h:mm:ss", // 21
    "m/d/yyyy h:mm", // 22
    NULL, // 23
    NULL, // 24
    NULL, // 25
    NULL, // 26
    NULL, // 27
    NULL, // 28
    NULL, // 29
    NULL, // 30
    NULL, // 31
    NULL, // 32
    NULL, // 33
    NULL, // 34
    NULL, // 35
    NULL, // 36
    "#,##0_);(#,##0)", // 37
    "#,##0_);[Red](#,##0)", // 38
    "#,##0.00_);(#,##0.00)", // 39
    "#,##0.00_);[Red](#,##0.00)", // 40
    NULL, // 41
    NULL, // 42
    NULL, // 43
    NULL, // 44
    "mm:ss", // 45
    "[h]:mm:ss", // 46
    "mm:ss.0", // 47
    "##0.0E+0", // 48
    "@" // 49
};
static int xlsx_print_err_messages = true;
static char *e_position; // speeding purpose


// functions
// public
void xlsx_set_print_err_messages(int flag);
int xlsx_open(const char *src, xlsx_workbook_t *xlsx);

enum xlsx_open_errno {
  XLSX_OPEN_ERRNO_MALFORMED_PARAMS = -1,
  XLSX_OPEN_ERRNO_OUT_OF_MEMORY = -2,
  XLSX_OPEN_ERRNO_CANT_DEPLOY_FILE = -3,
  XLSX_OPEN_ERRNO_XML_PARSING_ERROR = -4
};

xlsx_sheet_t * xlsx_load_sheet(const xlsx_workbook_t *deployed_xlsx, int sheet_number, const char *sheet_name);

enum xlsx_load_sheet_errno {
  XLSX_LOAD_SHEET_ERRNO_MALFORMED_PARAMS = -1,
  XLSX_LOAD_SHEET_ERRNO_OUT_OF_MEMORY = -2,
  XLSX_LOAD_SHEET_ERRNO_INDEX_OUT_OF_BOUNDS = -3,
  XLSX_LOAD_SHEET_ERRNO_XML_PARSING_ERROR = -4,
  XLSX_LOAD_SHEET_ERRNO_NON_EXISTENT = -5
};

void xlsx_unload_sheet(xlsx_sheet_t *sheet);

void xlsx_read_cell(xlsx_sheet_t *sheet, unsigned row, const char *column, xlsx_cell_t *cell_data_holder);

enum xlsx_read_cell_errno {
  XLSX_READ_CELL_ERRNO_MALFORMED_PARAMS = -1,
  XLSX_READ_CELL_ERRNO_OUT_OF_MEMORY = -2,
  XLSX_READ_CELL_ERRNO_SHEET_NOT_LOADED = -3
};

int xlsx_close(xlsx_workbook_t *deployed_xlsx);

// private
static void init_xlsx_workbook_t_struct(xlsx_workbook_t *xlsx);
static void init_xlsx_sheet_t_struct(xlsx_sheet_t *sheet, xlsx_workbook_t *deployed_xlsx);
static xlsx_cell_kind get_related_type(const char *format_code, int format_code_length);
static xlsx_formatter get_formatter(const char *format_code, int current_analyzed_index);
static int parse_sheet(int sheet_number, xlsx_sheet_t * sheet);
static XMLNode * find_row_node(xlsx_sheet_t *sheet, unsigned row, int start_from_child);
static XMLNode * find_cell_node(XMLNode *row, const char *cell);
static void interpret_cell_node(XMLNode *cell, xlsx_sheet_t *sheet, xlsx_cell_t * cell_data_holder);
static int delete_folder(const char *folder_path);
static void set_cell_data_values_for_number(const char *cell_text, xlsx_cell_t *cell_data_holder);

#endif
