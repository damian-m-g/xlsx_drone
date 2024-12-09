/*
* xlsx_drone - Copyright (c) 2021, Damian M. Gonzalez.
* Released under MIT license, please refer to LICENSE file for details.
* VERSION: 0.4.0
*/
#ifndef XLSX_DRONE_LIBRARY_H
#define XLSX_DRONE_LIBRARY_H

// includes
// core libraries
#include <stdlib.h>
#include <string.h>

// include needed for delete_folder(); for Visual Studio, windows.h must be included before zip.h
#if defined(_MSC_VER)
  #include <windows.h>
  #include <malloc.h>
#else
  #include <dirent.h>
  #include <unistd.h>
#endif

#include <sys/stat.h>
#include <math.h>

// code from https://github.com/michael-hartmann/parsefloat
#include <ctype.h>

#ifdef __cplusplus
extern "C" {
#endif

// external libraries
#include "../ext/zip.h" // https://github.com/kuba--/zip | using version 0.3.2 (2024/12)
#include "../ext/sxmlc.h" // http://sxmlc.sourceforge.net/ | using version 4.5.4 (2024/12)
#include "../ext/sxmlsearch.h"


// constants
// basic
#ifndef true
  #define true 1
#endif
#ifndef false
  #define false 0
#endif

// error output managment
#ifdef __MINGW32__
  #include <afxres.h>
  #define XLSX_SET_ERRNO(x) SetLastError(x)
#else
  #include <errno.h>
  #define XLSX_SET_ERRNO(x) errno=(x)
#endif

// OS awareness
#if defined(_WIN32) || defined(_WIN64)
  #define WINDOWS 1 // Windows OS > Win 3.x
#else
  #define WINDOWS 0 // Other OS
#endif

// temp location managment
#if defined(_WIN32) || defined(_WIN64)
  #define ENVIRONMENT_VARIABLE_TEMP "TEMP" // Windows OS > Win 3.x
#elif defined(__unix__) || defined(__linux__) || defined(__APPLE__) || defined(__MACH__)
  #define ENVIRONMENT_VARIABLE_TEMP "TMPDIR" // Ubuntu/Linux/OSX
#else
  #define ENVIRONMENT_VARIABLE_TEMP "TMPDIR" // Other OS probably uses this symbol
#endif

// other
#if defined(_WIN32) || defined(_WIN64)
  #define REL_PATH_TO_STYLES "\\xl\\styles.xml"
  #define REL_PATH_TO_SHARED_STRINGS "\\xl\\sharedStrings.xml"
  #define REL_PATH_TO_WORKBOOK "\\xl\\workbook.xml"
  #define REL_PATH_TO_WORKSHEETS "\\xl\\worksheets\\"
#else
  #define REL_PATH_TO_STYLES "/xl/styles.xml"
  #define REL_PATH_TO_SHARED_STRINGS "/xl/sharedStrings.xml"
  #define REL_PATH_TO_WORKBOOK "/xl/workbook.xml"
  #define REL_PATH_TO_WORKSHEETS "/xl/worksheets/"
#endif
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

typedef enum xlsx_cell_category {
  XLSX_NUMBER, // int, long long, or double
  XLSX_TEXT, // string
  XLSX_DATE, // int
  XLSX_TIME, // double
  XLSX_DATE_TIME, // double
  XLSX_UNKNOWN
} xlsx_cell_category;

// represents a single style
typedef struct xlsx_style_t {
  int style_id;
  enum xlsx_cell_category related_category;
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
  char *last_column;
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
static xlsx_cell_category xlsx_predefined_style_types[AMOUNT_OF_PREDEFINED_STYLE_TYPES] = {
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
static const char *xlsx_predefined_styles_format_code[AMOUNT_OF_PREDEFINED_STYLE_TYPES] = {
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

/*
* summary:
*   Flags if the error messages must or must not be printed. Perhaps the user of the library wants to manage errors
*   silently, without making them get written to stderr. By DEFAULT it is set to true.
* params:
*   flag: pass 0 if you want to cancel this feature, pass another value if you want to enable it.
* notes:
*   Out of memory errors never get printed.
*/
void xlsx_set_print_err_messages(int flag);

/*
* summary:
*   Useful to understand what went wrong in functions that failed. Compare this returned value against error codes
*   related to the failing function.
*/
int xlsx_get_xlsx_errno(void);

/*
* params:
*   - src: source XLSX.
*   - xlsx: handler. It will be written with data gathered after deploying the XLSX.
* returns:
*   - 1: everything went OK.
*   - 0: the process FAILED. Compare xlsx_errno against enum xlsx_open_errno to know why.
*/
int xlsx_open(const char *src, xlsx_workbook_t *xlsx);

enum xlsx_open_errno {
  XLSX_OPEN_ERRNO_MALFORMED_PARAMS = -1,
  XLSX_OPEN_ERRNO_OUT_OF_MEMORY = -2,
  XLSX_OPEN_ERRNO_CANT_DEPLOY_FILE = -3,
  XLSX_OPEN_ERRNO_XML_PARSING_ERROR = -4
};

/*
* summary:
*   Loads certain sheet, by its index (starts at 1) or by its name. Doing it by index is a bit faster.
* params:
*   - deployed_xlsx: deployed_xlsx parameter already passed to xlsx_open() with result 1 (OK).
*   - sheet_number: sheet index, the first sheet is the 1, and so on. Pass 0 if you pass a valid *sheet_name*.
*   - sheet_name: sheet name string. Pass NULL if you're passing a valid *sheet_number*.
* returns:
*   - NULL: FAIL. Check xlsx_errno and compare it against enum xlsx_load_sheet_errno to know what happened.
*   - xlsx_sheet_t *: SUCCESS.
*/
xlsx_sheet_t * xlsx_load_sheet(const xlsx_workbook_t *deployed_xlsx, int sheet_number, const char *sheet_name);

enum xlsx_load_sheet_errno {
  XLSX_LOAD_SHEET_ERRNO_MALFORMED_PARAMS = -11,
  XLSX_LOAD_SHEET_ERRNO_OUT_OF_MEMORY = -12,
  XLSX_LOAD_SHEET_ERRNO_INDEX_OUT_OF_BOUNDS = -13,
  XLSX_LOAD_SHEET_ERRNO_XML_PARSING_ERROR = -14,
  XLSX_LOAD_SHEET_ERRNO_NON_EXISTENT = -15
};

/*
* summary:
*   Manual way of freeing the memory allocated to treat this *sheet*. You may invoke this function once you're done
*   reading from it (you won't be able to load it again). This is not mandatory, is available in cases in which RAM
*   availability really concerns you. Useful when the *sheet* is very crowded with data, a good practice to call this
*   func if you finished reading it.
* params:
*   - sheet: the sheet to unload.
*/
void xlsx_unload_sheet(xlsx_sheet_t *sheet);

/*
* summary:
*   As to actually get the last column with a non-empty value requires some effort (run-time), it is not gathered in
*   xlsx_load_sheet(). This is because maybe you already know what columns are you interested in, and it could be
*   superflous to get the last column used.
*   So after the first time you call this function, sheet->last_column gets value, and then, you can directly ask for
*   sheet->last_column, or you can keep calling xlsx_get_last_column(), your choice.
* params:
*   sheet: A loaded sheet.
* returns:
*   A string with the last column value, i.e.: "AF", or "B", etc. Or, will return NULL if:
*     * an error happened. Check xlsx_errno against xlsx_get_last_column_errno values, that will be 0 if were no error.
*     * the sheet is empty.
*/
char* xlsx_get_last_column(xlsx_sheet_t *sheet);

enum xlsx_get_last_column_errno {
  XLSX_GET_LAST_COLUMN_ERRNO_SHEET_NOT_LOADED = -31,
  XLSX_GET_LAST_COLUMN_ERRNO_OUT_OF_MEMORY = -32
};

/*
* summary:
*   Uses *cell_data_holder* as carrier of the content read. This function zero-initialize all its fields, so you don't
*   have to do it. This means that you can pass the same structure over and over again, in fact, this is the
*   recommended way to go, because, as you can see, a cell_data_holder reserves a lot of memory. This was thought to
*   save run-time.
* params:
*   - sheet: the sheet where to look, it had to be loaded.
*   - row: cell row.
*   - column: cell column.
*   - cell_data_holder: read data will be written here, read it after the function returns.
* notes:
*   This function prioritizes speed over other concerns.
*   *cell_data_holder* will have an xlsx_value_type equal to XLSX_NULL if the cell has not content at all.
*/
int xlsx_read_cell(xlsx_sheet_t *sheet, unsigned row, const char *column, xlsx_cell_t *cell_data_holder);

enum xlsx_read_cell_errno {
  XLSX_READ_CELL_ERRNO_MALFORMED_PARAMS = -21,
  XLSX_READ_CELL_ERRNO_OUT_OF_MEMORY = -22,
  XLSX_READ_CELL_ERRNO_SHEET_NOT_LOADED = -23
};

/*
* summary:
*   Closes (cleans) the deployed XLSX, freeing all the memory dinamically allocated by this library related to the
*   object passed as argument. It is your responsability to pass an actually deployed xlsx, otherwise the behaviour
*   is undefined.
* params:
*   - deployed_xlsx: deployed_xlsx parameter already passed to xlsx_open() with result 1 (OK).
* notes:
*   - *deployed_xlsx* won't be freed by this function, it's responsability of the library user.
* returns:
*   - 1: everything went OK.
*   - 0: something happened and the process FAILED. Check errno against errno.h constant values.
*/
int xlsx_close(xlsx_workbook_t *deployed_xlsx);

// private
static void init_xlsx_workbook_t_struct(xlsx_workbook_t *xlsx);
static void init_xlsx_sheet_t_struct(xlsx_sheet_t *sheet, xlsx_workbook_t *deployed_xlsx);
static xlsx_cell_category get_related_category(const char *format_code, int format_code_length);
static xlsx_formatter get_formatter(const char *format_code, int current_analyzed_index);
static int parse_sheet(int sheet_number, xlsx_sheet_t * sheet);
static XMLNode * find_row_node(xlsx_sheet_t *sheet, unsigned row, int start_from_child);
static XMLNode * find_cell_node(XMLNode *row, const char *cell);
static void interpret_cell_node(XMLNode *cell, xlsx_sheet_t *sheet, xlsx_cell_t * cell_data_holder);
static int delete_folder(const char *folder_path);
static void set_cell_data_values_for_number(const char *cell_text, xlsx_cell_t *cell_data_holder);
static void withdraw_alphabetic_chars(const char *s_input, char s_output[5]);


#ifdef __cplusplus
}
#endif

#endif
