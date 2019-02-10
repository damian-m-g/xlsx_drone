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
#include "zip.h" // https://github.com/kuba--/zip
#include "sxmlc.h" // http://sxmlc.sourceforge.net/
#include "sxmlsearch.h"


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
static char **xlsx_predefined_styles_format_code = NULL;
static xlsx_cell_kind xlsx_predefined_style_types[AMOUNT_OF_PREDEFINED_STYLE_TYPES];
static int xlsx_print_err_messages = true;


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
static int initialize_predefined_style_data(void);
static xlsx_cell_kind get_related_type(const char *format_code, int format_code_length);
static xlsx_formatter get_formatter(const char *format_code, int current_analyzed_index);
static int parse_sheet(int sheet_number, xlsx_sheet_t * sheet);
static XMLNode * find_row_node(xlsx_sheet_t *sheet, unsigned row, int start_from_child);
static XMLNode * find_cell_node(XMLNode *row, const char *cell);
static void interpret_cell_node(XMLNode *cell, xlsx_sheet_t *sheet, xlsx_cell_t * cell_data_holder);
static int delete_folder(const char *folder_path);
static void set_cell_data_values_for_number(const char *cell_text, xlsx_cell_t *cell_data_holder);

#endif
