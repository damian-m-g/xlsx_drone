// includes
// this library header
#include "xlsx_drone.h"


// functions
/*
* summary:
*   Flags if the error messages must or must not be printed. Perhaps the user of the library wants to manage errors
*   silently, without making them get written to stderr. By DEFAULT it is set to true.
* params:
*   flag: pass 0 if you want to cancel this feature, pass another value if you want to enable it.
* notes:
*   Out of memory errors never get printed.
*/
void xlsx_set_print_err_messages(int flag) {
  xlsx_print_err_messages = flag;
}

/*
* summary:
*   Useful to understand what went wrong in functions that failed. Compare this returned value against error codes
*   related to the failing function.
*/
int xlsx_get_xlsx_errno(void) {
  return xlsx_errno;
}

/*
* params:
*   - src: source XLSX.
*   - xlsx: handler. It will be written with data gathered after deploying the XLSX.
* returns:
*   - 1: everything went OK.
*   - 0: the process FAILED. Compare xlsx_errno against enum xlsx_open_errno to know why.
*/
int xlsx_open(const char *src, xlsx_workbook_t *xlsx)
{
  xlsx_errno = 0;

  if(!src || !xlsx) {
    if(xlsx_print_err_messages)
      fprintf(stderr, "XLSX_C ERROR: Malformed parameters.\n");
    xlsx_errno = XLSX_OPEN_ERRNO_MALFORMED_PARAMS;
    return 0; // FAIL
  }

  init_xlsx_workbook_t_struct(xlsx);

  // build the temporary path where the excel will be deployed (fallback to /tmp if env var not defined)
  const char *temp_path = getenv(ENVIRONMENT_VARIABLE_TEMP) ? getenv(ENVIRONMENT_VARIABLE_TEMP) : "/tmp";
#if !(WINDOWS)
  if(temp_path[0] == '\0')
      temp_path = "/tmp";
#endif
  // non-Windows users use mkdtemp() procedure
  const char *temp_folder = WINDOWS ? tmpnam(NULL) : "/XXXXXX";
  int deployed_xlsx_path_len = (int)strlen(temp_path) + (int)strlen(temp_folder);
  char *deployed_xlsx_path = malloc(sizeof(char) * (deployed_xlsx_path_len + 1));
  if(!deployed_xlsx_path) {
    xlsx_errno = XLSX_OPEN_ERRNO_OUT_OF_MEMORY;
    fprintf(stderr, "XLSX_C ERROR: memory error.\n");
    return 0; // FAIL
  }
  strcpy(deployed_xlsx_path, temp_path);
  strcat(deployed_xlsx_path, temp_folder);
  // ensure the char array to be a string
  deployed_xlsx_path[deployed_xlsx_path_len] = '\0';
  // non-Windows users are suggested to use mkdtemp()
#if !defined(_MSC_VER)
  if(!WINDOWS) {
    if(!mkdtemp(deployed_xlsx_path)) {
      xlsx_errno = XLSX_OPEN_ERRNO_CANT_DEPLOY_FILE;
      fprintf(stderr, "XLSX_C ERROR: XLSX_OPEN_ERRNO_CANT_DEPLOY_FILE to %s\n", deployed_xlsx_path);
      return 0; // FAIL
    }
  }
#endif

  // deploy there
  if(zip_extract(src, deployed_xlsx_path, NULL, NULL) != 0) {
    free(deployed_xlsx_path);
    if(xlsx_print_err_messages)
      fprintf(stderr, "XLSX_C ERROR: \"%s\" couldn't be deployed. Check that the file isn't already opened.\n", src);
    xlsx_errno = XLSX_OPEN_ERRNO_CANT_DEPLOY_FILE;
    return 0; // FAIL
  }

  // fill *xlsx* with data
  xlsx->deployment_path = deployed_xlsx_path;

  // load sharedStrings.xml
  if(!(xlsx->shared_strings_xml = malloc(sizeof(XMLDoc)))) {
    xlsx_close(xlsx);
    xlsx_errno = XLSX_OPEN_ERRNO_OUT_OF_MEMORY;
    return 0; // FAIL
  }
  XMLDoc_init(xlsx->shared_strings_xml);
#if defined(_MSC_VER)
  char *path_to_shared_strings_xml = (char*)_alloca(strlen(deployed_xlsx_path) + strlen(REL_PATH_TO_SHARED_STRINGS) + 1);
#else
  char path_to_shared_strings_xml[strlen(deployed_xlsx_path) + strlen(REL_PATH_TO_SHARED_STRINGS) + 1];
#endif
  strcpy(path_to_shared_strings_xml, deployed_xlsx_path);
  strcat(path_to_shared_strings_xml, REL_PATH_TO_SHARED_STRINGS);
  // next function returns false if something went wrong in the parsing OR if the file doesn't exist, which may happen
  // when the XLSX has no strings
  if(!XMLDoc_parse_file_DOM(path_to_shared_strings_xml, xlsx->shared_strings_xml)) {
    // if prev function fails, it calls XMLDoc_free(), so no need to call it again
    free(xlsx->shared_strings_xml);
    xlsx->shared_strings_xml = NULL;
  }

  // load and parse a bit of styles.xml
  XMLDoc styles_xml;
  XMLDoc_init(&styles_xml);
#if defined(_MSC_VER)
  char *path_to_styles_xml = (char*)_alloca(strlen(deployed_xlsx_path) + strlen(REL_PATH_TO_STYLES) + 1);
#else
  char path_to_styles_xml[strlen(deployed_xlsx_path) + strlen(REL_PATH_TO_STYLES) + 1];
#endif
  strcpy(path_to_styles_xml, deployed_xlsx_path);
  strcat(path_to_styles_xml, REL_PATH_TO_STYLES);
  if(!(XMLDoc_parse_file_DOM(path_to_styles_xml, &styles_xml))) {
    xlsx_close(xlsx);
    if(xlsx_print_err_messages)
      fprintf(stderr, "XLSX_C ERROR: \"%s\" can't be parsed or doesn't exist.\n", REL_PATH_TO_STYLES);
    xlsx_errno = XLSX_OPEN_ERRNO_XML_PARSING_ERROR;
    XMLDoc_free(&styles_xml);
    return 0; // FAIL
  }

  // start the search for cellXfs
  XMLSearch search_engine;
  XMLSearch_init(&search_engine);
  XMLSearch_search_set_tag(&search_engine, STYLES_CELLXFS_TAG);
  // from the root tag
  XMLNode *cell_xfs_node = XMLSearch_next(styles_xml.nodes[styles_xml.i_root], &search_engine);
  if(!cell_xfs_node) {
    xlsx_close(xlsx);
    fprintf(stderr, "XLSX_C ERROR: \"%s\" node can't be found on \"%s\".\n", STYLES_CELLXFS_TAG, REL_PATH_TO_STYLES);
    xlsx_errno = XLSX_OPEN_ERRNO_XML_PARSING_ERROR;
    XMLSearch_free(&search_engine, false);
    XMLDoc_free(&styles_xml);
    return 0; // FAIL
  }

  // allocate memory for the different styles
  xlsx->n_styles = cell_xfs_node->n_children;
  if(!(xlsx->styles = calloc(cell_xfs_node->n_children, sizeof(xlsx_style_t *)))) {
    xlsx_close(xlsx);
    xlsx_errno = XLSX_OPEN_ERRNO_OUT_OF_MEMORY;
    XMLSearch_free(&search_engine, false);
    XMLDoc_free(&styles_xml);
    return 0; // FAIL
  }

  // initialize variables that will be used on the loop
  int xf_index, attr_index, format_code_length, xf_node_numfmtid_value_as_int;
  XMLNode *xf_node, *num_fmt_node = NULL;
  char *xf_node_numfmtid_value = NULL;
  // loop over *cell_xfs_node* children
  for(xf_index = 0; xf_index < cell_xfs_node->n_children; ++xf_index) {
    // allocate memory for this style
    if(!(xlsx->styles[xf_index] = malloc(sizeof(xlsx_style_t)))) {
      xlsx_close(xlsx);
      xlsx_errno = XLSX_OPEN_ERRNO_OUT_OF_MEMORY;
      XMLSearch_free(&search_engine, false);
      XMLDoc_free(&styles_xml);
      return 0; // FAIL
    }
    // zero initialize all its fields that need memory allocation
    xlsx->styles[xf_index]->format_code = NULL;
    // work with the node
    xf_node = cell_xfs_node->children[xf_index];
    // the format code has to be captured
    for(attr_index = 0; attr_index < xf_node->n_attributes; ++attr_index) {
      if(strcmp(xf_node->attributes[attr_index].name, STYLES_NUMFMTID_ATTR_NAME) == 0) {
        xf_node_numfmtid_value = xf_node->attributes[attr_index].value;
        break;
      }
    }
    if(attr_index == xf_node->n_attributes) {
      xlsx_close(xlsx);
      if(xlsx_print_err_messages)
        fprintf(stderr, "XLSX_C ERROR: \"%s\" attr can't be found on \"%s\" children over \"%s\".\n",
                STYLES_NUMFMTID_ATTR_NAME, STYLES_CELLXFS_TAG, REL_PATH_TO_STYLES);
      xlsx_errno = XLSX_OPEN_ERRNO_XML_PARSING_ERROR;
      XMLSearch_free(&search_engine, false);
      XMLDoc_free(&styles_xml);
      return 0; // FAIL
    }
    // once *xf_node_numfmtid_value* was found, see if it points to the predefined ones,
    // otherwise, look the matching numFmt element
    xf_node_numfmtid_value_as_int = (int)strtol(xf_node_numfmtid_value, NULL, 10);
    xlsx->styles[xf_index]->style_id = xf_node_numfmtid_value_as_int;
    if(xf_node_numfmtid_value_as_int < AMOUNT_OF_PREDEFINED_STYLE_TYPES) {
      // predefined style
      xlsx->styles[xf_index]->related_category = xlsx_predefined_style_types[xf_node_numfmtid_value_as_int];
      // note that next value could be NULL
      xlsx->styles[xf_index]->format_code = (char*)xlsx_predefined_styles_format_code[xf_node_numfmtid_value_as_int];
    } else {
      // custom style
      XMLSearch_free(&search_engine, false);
      XMLSearch_init(&search_engine);
      XMLSearch_search_set_tag(&search_engine, STYLES_NUMFMT_TAG);
      XMLSearch_search_add_attribute(&search_engine, STYLES_NUMFMTID_ATTR_NAME, xf_node_numfmtid_value, true);
      if(!(num_fmt_node = XMLSearch_next(styles_xml.nodes[styles_xml.i_root], &search_engine))) {
        xlsx_close(xlsx);
        if(xlsx_print_err_messages)
          fprintf(stderr, "XLSX_C ERROR: There's no \"%s\" with \"%s\" equal to \"%s\" in \"%s\".\n",
                  STYLES_NUMFMT_TAG, STYLES_NUMFMTID_ATTR_NAME, xf_node_numfmtid_value, REL_PATH_TO_STYLES);
        xlsx_errno = XLSX_OPEN_ERRNO_XML_PARSING_ERROR;
        XMLSearch_free(&search_engine, false);
        XMLDoc_free(&styles_xml);
        return 0; // FAIL
      }
      for(attr_index = (num_fmt_node->n_attributes - 1); attr_index >= 0; --attr_index) {
        if(strcmp(num_fmt_node->attributes[attr_index].name, STYLES_FORMATCODE_ATTR_NAME) == 0) {
          format_code_length = (int)strlen(num_fmt_node->attributes[attr_index].value);
          if(!(xlsx->styles[xf_index]->format_code = malloc(sizeof(char) * (format_code_length + 1)))) {
            xlsx_close(xlsx);
            xlsx_errno = XLSX_OPEN_ERRNO_OUT_OF_MEMORY;
            XMLSearch_free(&search_engine, false);
            XMLDoc_free(&styles_xml);
            return 0; // FAIL
          }
          strcpy(xlsx->styles[xf_index]->format_code, num_fmt_node->attributes[attr_index].value);
          // find out what kind of category this style is, inspecting its formatCode
          xlsx->styles[xf_index]->related_category = \
            get_related_category(xlsx->styles[xf_index]->format_code, format_code_length);
          break;
        }
      }
      if(attr_index == -1) {
        xlsx_close(xlsx);
        if(xlsx_print_err_messages)
          fprintf(stderr, "XLSX_C ERROR: \"%s\" attr can't be found on \"%s\" elements over \"%s\".\n",
                  STYLES_FORMATCODE_ATTR_NAME, STYLES_NUMFMT_TAG, REL_PATH_TO_STYLES);
        xlsx_errno = XLSX_OPEN_ERRNO_XML_PARSING_ERROR;
        XMLSearch_free(&search_engine, false);
        XMLDoc_free(&styles_xml);
        return 0; // FAIL
      }
      XMLSearch_free(&search_engine, false);
    }
  }

  // look for all sheets on the workbook and partially initialize the sheets members
  XMLDoc workbook_xml;
  XMLDoc_init(&workbook_xml);
#if defined(_MSC_VER)
  char *path_to_workbook_xml = (char*)_alloca(strlen(deployed_xlsx_path) + strlen(REL_PATH_TO_WORKBOOK) + 1);
#else
  char path_to_workbook_xml[strlen(deployed_xlsx_path) + strlen(REL_PATH_TO_WORKBOOK) + 1];
#endif
  strcpy(path_to_workbook_xml, deployed_xlsx_path);
  strcat(path_to_workbook_xml, REL_PATH_TO_WORKBOOK);
  if(!(XMLDoc_parse_file_DOM(path_to_workbook_xml, &workbook_xml))) {
    xlsx_close(xlsx);
    if(xlsx_print_err_messages)
      fprintf(stderr, "XLSX_C ERROR: \"%s\" can't be parsed or doesn't exist.\n", REL_PATH_TO_WORKBOOK);
    xlsx_errno = XLSX_OPEN_ERRNO_XML_PARSING_ERROR;
    XMLDoc_free(&workbook_xml);
    XMLDoc_free(&styles_xml);
    return 0; // FAIL
  }
  // look for sheet elements
  XMLSearch_free(&search_engine, false);
  XMLSearch_init(&search_engine);
  XMLSearch_search_set_tag(&search_engine, WORKBOOK_SHEETS_TAG);
  // from the root tag
  XMLNode *sheets_node = XMLSearch_next(workbook_xml.nodes[workbook_xml.i_root], &search_engine);
  if(!sheets_node) {
    xlsx_close(xlsx);
    if(xlsx_print_err_messages)
      fprintf(stderr, "XLSX_C ERROR: There's no \"%s\" element inside \"%s\".\n",
              WORKBOOK_SHEETS_TAG, REL_PATH_TO_WORKBOOK);
    xlsx_errno = XLSX_OPEN_ERRNO_XML_PARSING_ERROR;
    XMLSearch_free(&search_engine, false);
    XMLDoc_free(&workbook_xml);
    XMLDoc_free(&styles_xml);
    return 0; // FAIL
  }

  xlsx->n_sheets = sheets_node->n_children;
  if(!(xlsx->sheets = malloc(sheets_node->n_children * sizeof(xlsx_sheet_t *)))) {
    xlsx_close(xlsx);
    xlsx_errno = XLSX_OPEN_ERRNO_OUT_OF_MEMORY;
    XMLSearch_free(&search_engine, false);
    XMLDoc_free(&workbook_xml);
    XMLDoc_free(&styles_xml);
    return 0; // FAIL
  }

  int sheet_index;
  for(sheet_index = 0; sheet_index < sheets_node->n_children; ++sheet_index) {
    if(!(xlsx->sheets[sheet_index] = malloc(sizeof(xlsx_sheet_t)))) {
      xlsx_close(xlsx);
      xlsx_errno = XLSX_OPEN_ERRNO_OUT_OF_MEMORY;
      XMLSearch_free(&search_engine, false);
      XMLDoc_free(&workbook_xml);
      XMLDoc_free(&styles_xml);
      return 0; // FAIL
    }
    // initialize all members of this *xlsx_sheet_t* struct
    init_xlsx_sheet_t_struct(xlsx->sheets[sheet_index], xlsx);
    // get its name
    for(attr_index = 0; attr_index < sheets_node->children[sheet_index]->n_attributes; ++attr_index) {
      if(strcmp(sheets_node->children[sheet_index]->attributes[attr_index].name, WORKBOOK_NAME_ATTR_NAME) == 0) {
        if(!(xlsx->sheets[sheet_index]->name = \
          malloc(strlen(sheets_node->children[sheet_index]->attributes[attr_index].value) + 1))) {
          xlsx_close(xlsx);
          xlsx_errno = XLSX_OPEN_ERRNO_OUT_OF_MEMORY;
          XMLSearch_free(&search_engine, false);
          XMLDoc_free(&workbook_xml);
          return 0; // FAIL
        }
        strcpy(xlsx->sheets[sheet_index]->name, sheets_node->children[sheet_index]->attributes[attr_index].value);
        break;
      }
    }
    if(attr_index == sheets_node->children[sheet_index]->n_attributes) {
      xlsx_close(xlsx);
      if(xlsx_print_err_messages)
        fprintf(stderr, "XLSX_C ERROR: \"%s\" attr can't be found on \"%s\" children over \"%s\".\n",
                WORKBOOK_NAME_ATTR_NAME, WORKBOOK_SHEETS_TAG, REL_PATH_TO_WORKBOOK);
      xlsx_errno = XLSX_OPEN_ERRNO_XML_PARSING_ERROR;
      XMLSearch_free(&search_engine, false);
      XMLDoc_free(&workbook_xml);
      XMLDoc_free(&styles_xml);
      return 0; // FAIL
    }
  }

  XMLSearch_free(&search_engine, false);
  XMLDoc_free(&workbook_xml);
  XMLDoc_free(&styles_xml);
  return 1;
}


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
xlsx_sheet_t * xlsx_load_sheet(const xlsx_workbook_t *deployed_xlsx, int sheet_number, const char *sheet_name)
{
  xlsx_errno = 0;

  if((sheet_number <= 0) && (!sheet_name)) {
    if(xlsx_print_err_messages)
      fprintf(stderr, "XLSX_C ERROR: Malformed parameters.\n");
    xlsx_errno = XLSX_LOAD_SHEET_ERRNO_MALFORMED_PARAMS;
    return NULL; // FAIL
  }

  if(sheet_number > 0)

  {
    // the user is trying to seek a sheet by index
    if(sheet_number <= deployed_xlsx->n_sheets) {
      xlsx_sheet_t *sheet = deployed_xlsx->sheets[sheet_number - 1];

      if(!sheet->sheet_xml) {
        if(!(parse_sheet(sheet_number, sheet)))
          return NULL; // FAIL
      }

      return(sheet); // SUCCESS
    } else {
      // index out of bounds
      if(xlsx_print_err_messages)
        fprintf(stderr, "XLSX_C ERROR: Index out of bounds.\n");
      xlsx_errno = XLSX_LOAD_SHEET_ERRNO_INDEX_OUT_OF_BOUNDS;
      return NULL; // FAIL
    }
  }

  else if(sheet_name)

  {
    // the user is trying to seek a sheet by its name
    int sheet_index;
    for(sheet_index = 0; sheet_index < deployed_xlsx->n_sheets; ++sheet_index) {
      if(strcmp(deployed_xlsx->sheets[sheet_index]->name, sheet_name) == 0) {
        // sheet found
        xlsx_sheet_t *sheet = deployed_xlsx->sheets[sheet_index];
        if(!sheet->sheet_xml) {
          if(!(parse_sheet((sheet_index + 1), sheet)))
            return NULL; // FAIL
        }
        return sheet; // SUCCESS
      }
    }
    // if you reach this line, then there's no sheet with such name
    if(xlsx_print_err_messages)
      fprintf(stderr, "XLSX_C ERROR: There's no sheet named \"%s\".\n", sheet_name);
    xlsx_errno = XLSX_LOAD_SHEET_ERRNO_NON_EXISTENT;
    return NULL; // FAIL
  }

  else

  {
    if(xlsx_print_err_messages)
      fprintf(stderr, "XLSX_C ERROR: Malformed parameters.\n");
    xlsx_errno = XLSX_LOAD_SHEET_ERRNO_MALFORMED_PARAMS;
    return NULL; // FAIL
  }
}


/*
* summary:
*   Manual way of freeing the memory allocated to treat this *sheet*. You may invoke this function once you're done
*   reading from it (you won't be able to load it again). This is not mandatory, is available in cases in which RAM
*   availability really concerns you. Useful when the *sheet* is very crowded with data, a good practice to call this
*   func if you finished reading it.
* params:
*   - sheet: the sheet to unload.
*/
void xlsx_unload_sheet(xlsx_sheet_t *sheet) {
  if(sheet->name) {
    free(sheet->name);
    sheet->name = NULL;
    if(sheet->last_column) {
      free(sheet->last_column);
      sheet->last_column = NULL;
    }
    // if the name isn't allocated, neither the sheet_xml field
    if(sheet->sheet_xml) {
      XMLDoc_free(sheet->sheet_xml);
      free(sheet->sheet_xml);
      sheet->sheet_xml = NULL;
    }
  }
}


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
char* xlsx_get_last_column(xlsx_sheet_t *sheet) {

  // reset xlsx_errno value
  xlsx_errno = 0;

  // the sheet must be loaded
  if(!sheet->sheet_xml) {
    if(xlsx_print_err_messages)
      fprintf(stderr, "XLSX_C ERROR: The sheet isn't loaded.\n");
    xlsx_errno = XLSX_GET_LAST_COLUMN_ERRNO_SHEET_NOT_LOADED;
    return NULL; // FAIL
  }

  // it could have been obtained before (or be an empty sheet, in which case its value is NULL)
  if(sheet->last_column || (sheet->sheetdata->n_children == 0)) {
    return sheet->last_column;
  }

  // the sheet at least have some value, so will have an actual string for last_column
  int row_node_index, cell_node_index;
  XMLNode *row, *cell;
  char last_column[5] = {0}, column[5] = {0};

  // iterate over all rows, starting from the first one
  for(row_node_index = 0; row_node_index < sheet->sheetdata->n_children; ++row_node_index) {
    row = sheet->sheetdata->children[row_node_index];

    // iterate over all cells of thy row, starting from the last one
    for(cell_node_index = row->n_children - 1; cell_node_index >= 0; --cell_node_index) {
      cell = row->children[cell_node_index];

      // does it have a <v> child with value within? If has a child, it has value
      if(cell->n_children > 0) {
        // first attribute should be "r"
        withdraw_alphabetic_chars(cell->attributes[0].value, column);
        // compare this column value with the "last_column"
        if(strlen(column) == strlen(last_column)) {
          // only use strcmp when both strings have the same length, otherwise doesn't work for the purpose
          if(strcmp(column, last_column) > 0) {
            strcpy(last_column, column);
          }
        } else if(strlen(column) > strlen(last_column)) {
          // *column* has a bigger column value
          strcpy(last_column, column);
        }
        // don't go any further, above you'll find lesser column values
        break;
      }
    }
  }

  // in the case the sheet is empty but has style inserted, last_column won't have a valid string
  if(last_column[0] != '\0') {
    sheet->last_column = (char*)malloc(sizeof(char) * 5);
    if(!sheet->last_column) {
      xlsx_errno = XLSX_GET_LAST_COLUMN_ERRNO_OUT_OF_MEMORY;
      return NULL; // FAIL
    }
    strcpy(sheet->last_column, last_column);
    return sheet->last_column;
  } else {
    return NULL;
  }
}


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
* returns:
*   - 1: everything went OK.
*   - 0: the process FAILED. Compare xlsx_errno against enum xlsx_read_cell_errno to know why.
* notes:
*   This function prioritizes speed over other concerns.
*   *cell_data_holder* will have an xlsx_value_type equal to XLSX_NULL if the cell has not content at all.
*/
int xlsx_read_cell(xlsx_sheet_t *sheet, unsigned row, const char *column, xlsx_cell_t *cell_data_holder) {

  // the sheet must be loaded
  if(!sheet->sheet_xml) {
    if(xlsx_print_err_messages)
      fprintf(stderr, "XLSX_C ERROR: The sheet isn't loaded.\n");
    xlsx_errno = XLSX_READ_CELL_ERRNO_SHEET_NOT_LOADED;
    return 0;
  }

  // if non-sense parameter, return
  if(strlen(column) > 4) {
    if(xlsx_print_err_messages)
      fprintf(stderr, "XLSX_C ERROR: *column* must be a null terminated char array, with no more than 4 chars.\n");
    xlsx_errno = XLSX_READ_CELL_ERRNO_MALFORMED_PARAMS;
    return 0;
  }

  // reset *cell_data_holder*
  cell_data_holder->style = NULL;
  cell_data_holder->value_type = XLSX_NULL;

  // figure out the cell as string
  char row_as_s[8]; // max n° of rows: 1_048_576; max n° of cols: 16_384 (worth with 4 letters A-Z)
  snprintf(row_as_s, 8, "%d", row);
  char cell_as_s[12];
  strcpy(cell_as_s, column);
  strcat(cell_as_s, row_as_s);

  // check if the row looked is the last row looked
  if(row == sheet->last_row_looked.row_n) {

    // it is, so don't look for the row, just for the cell
    XMLNode *cell_node = \
      find_cell_node(sheet->sheetdata->children[sheet->last_row_looked.sheetdata_child_i], cell_as_s);
    if(cell_node) {
      interpret_cell_node(cell_node, sheet, cell_data_holder);
    }

  } else if((int)row > sheet->last_row_looked.row_n) {

    // probably the next row will contain what you're looking fore
    XMLNode * row_node = find_row_node(sheet, row, (sheet->last_row_looked.sheetdata_child_i + 1));
    if(row_node) {
      XMLNode * cell_node = find_cell_node(row_node, cell_as_s);
      if(cell_node) {
        interpret_cell_node(cell_node, sheet, cell_data_holder);
      }
    }

  } else {

    // look for the row, from the beginning
    XMLNode * row_node = find_row_node(sheet, row, 0);
    if(row_node) {
      XMLNode * cell_node = find_cell_node(row_node, cell_as_s);
      if(cell_node) {
        interpret_cell_node(cell_node, sheet, cell_data_holder);
      }
    }
  }

  return 1;
}


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
int xlsx_close(xlsx_workbook_t *deployed_xlsx)
{
  if(deployed_xlsx->shared_strings_xml) {
    XMLDoc_free(deployed_xlsx->shared_strings_xml);
    free(deployed_xlsx->shared_strings_xml);
    deployed_xlsx->shared_strings_xml = NULL;
  }

  int index;
  if(deployed_xlsx->styles) {
    for(index = 0; index < deployed_xlsx->n_styles; ++index) {
      if(deployed_xlsx->styles[index]) {
        if((deployed_xlsx->styles[index]->format_code) &&
           (deployed_xlsx->styles[index]->style_id >= AMOUNT_OF_PREDEFINED_STYLE_TYPES)) {
          // if the style is a custom one, free the allocated memory
          free(deployed_xlsx->styles[index]->format_code);
        }
        free(deployed_xlsx->styles[index]);
      } else {
        // if at least one of this items has no assigned memory, the rest of the items neither have
        break;
      }
    }
    free(deployed_xlsx->styles);
    deployed_xlsx->styles = NULL;
  }

  if(deployed_xlsx->sheets) {
    for(index = 0; index < deployed_xlsx->n_sheets; ++index) {
      xlsx_unload_sheet(deployed_xlsx->sheets[index]);
      free(deployed_xlsx->sheets[index]);
    }
    free(deployed_xlsx->sheets);
    deployed_xlsx->sheets = NULL;
  }

  int xlsx_delete_folder_res = 1;
  if(deployed_xlsx->deployment_path) {
    // delete deployed XLSX
    xlsx_delete_folder_res = delete_folder(deployed_xlsx->deployment_path);

    // free what's left
    free(deployed_xlsx->deployment_path);
    deployed_xlsx->deployment_path = NULL;
  }

  return xlsx_delete_folder_res;
}


// private functions


static void init_xlsx_workbook_t_struct(xlsx_workbook_t *xlsx) {
  xlsx->deployment_path = NULL;
  xlsx->shared_strings_xml = NULL;
  xlsx->n_styles = 0;
  xlsx->styles = NULL;
  xlsx->n_sheets = 0;
  xlsx->sheets = NULL;
}


static void init_xlsx_sheet_t_struct(xlsx_sheet_t *sheet, xlsx_workbook_t *deployed_xlsx) {
  sheet->xlsx = deployed_xlsx;
  sheet->name = NULL;
  sheet->sheet_xml = NULL; // won't be loaded until xlsx_load_sheet()
  sheet->sheetdata = NULL; // won't be loaded until xlsx_load_sheet()
  sheet->last_row = -1; // won't be known until xlsx_load_sheet()
  sheet->last_column = NULL; // won't be known until xlsx_get_last_column() is called
  sheet->last_row_looked.row_n = -1; // no row was seeked yet
  sheet->last_row_looked.sheetdata_child_i = -1; // no row was seeked yet
}


/*
* summary:
*   Parses a *format_code*, in search for clues, that takes the program to understand which is the type the
*   *format_code* is talking about. Don't pass *format_code* that formats raw text, this function expects to work
*   with numbers.
* params:
*   - format_code: format_code withdrawn from a numFmt node in the styles XML.
*   - format_code_length: for speeding purpose.
* returns:
*   One of:
*     - XLSX_NUMBER
*     - XLSX_DATE
*     - XLSX_TIME
*     - XLSX_DATE_TIME
*/
static xlsx_cell_category get_related_category(const char *format_code, int format_code_length) {
  // *m_found* means an 'm' char found, it's ambiguous between date and time
  int current_analyzed_index, is_date = 0, is_time = 0, m_found = 0;
  for(current_analyzed_index = 0; current_analyzed_index < format_code_length; ++current_analyzed_index) {

    // note that this could also return XLSX_FORMATTER_UNKNOWN
    switch(get_formatter(format_code, current_analyzed_index)) {
      case XLSX_FORMATTER_AMBIGUOUS_M:
        m_found = true;
        break;
      case XLSX_FORMATTER_TIME:
        is_time = true;
        break;
      case XLSX_FORMATTER_DATE:
        is_date = true;
        break;
      default: {}
    }

    if(is_date && is_time)
      // there's no need to keep researching
      return XLSX_DATE_TIME;
  }

  // inspect the result of the research
  if(is_time) {
    return XLSX_TIME;
  } else if(is_date || m_found) {
    return XLSX_DATE;
  } else {
    return XLSX_NUMBER;
  }
}


/*
* summary:
*   Analyze an explicit char of the *format_code* and tells if it's part of some specific formatting.
* params:
*   - format_code: format_code withdrawn from a numFmt node in the styles XML.
*   - current_analyzed_index: the index pointing to a specific char of *format_code*.
*/
static xlsx_formatter get_formatter(const char *format_code, int current_analyzed_index) {
  switch(format_code[current_analyzed_index]) {
    case 'm': case 'h': case 's': case 'y': case 'd': {
      // "[Red]" case
      if(format_code[current_analyzed_index] == 'd' && current_analyzed_index >= 3 &&
         format_code[current_analyzed_index - 3] == '[' && format_code[current_analyzed_index - 2] == 'R' &&
         format_code[current_analyzed_index - 1] == 'e' && format_code[current_analyzed_index + 1] == ']') {
        return XLSX_FORMATTER_UNKNOWN;
      }

      // so far it isn't escaped
      int char_is_escaped = false;
      // check if it's escaped
      if(current_analyzed_index > 0) {

        // being preceded by '\'
        if(format_code[current_analyzed_index - 1] == '\\') {
          // check if the '\' is actually being escaped
          if(current_analyzed_index > 1) {

            // so far it is escaped
            char_is_escaped = true;
            int i = current_analyzed_index - 2;

            // look backwards until find a non '\' char or the start of the string
            while(true) {
              if(format_code[i] == '\\') {
                if(char_is_escaped) {
                  char_is_escaped = false;
                } else {
                  char_is_escaped = true;
                }
              } else {
                break;
              }

              if(--i < 0) {
                break;
              }
            }
          } else {
            char_is_escaped = true;
          }
        } else {

          // check being surrounded by '"'
          int quotes_open = false;
          int i = 0;
          int continue_loop = true;

          while(continue_loop) {
            switch(format_code[i]) {
              case '\0':
                // end of string found, get out
                continue_loop = false;
                break;
              case '"':
                if(!quotes_open) {
                  quotes_open = true;
                } else {
                  // see if our currently analyzed index is in-between
                  if(current_analyzed_index < i) {
                    // char definitely escaped, get out
                    char_is_escaped = true;
                    continue_loop = false;
                    // quotes_open logically should be set to false, but doesn't matter anymore
                    break;
                  }
                  quotes_open = false;
                }
            }

            if((++i == current_analyzed_index) && (!quotes_open)) {
              // char definitely not escaped
              continue_loop = false;
            }
          }
        }

      }

      if(char_is_escaped) {
        return XLSX_FORMATTER_UNKNOWN;
      } else {
        switch(format_code[current_analyzed_index]) {
          case 'm':
            return XLSX_FORMATTER_AMBIGUOUS_M;
          case 'h': case 's':
            return XLSX_FORMATTER_TIME;
          default:
            // 'y' or 'd'
            return XLSX_FORMATTER_DATE;
        }
      }

    } default:
      return XLSX_FORMATTER_UNKNOWN;
  }
}


/*
* summary:
*   Seeks certain sheet previously deployed and load it into RAM. Makes sheet->sheet_xml point to the loaded data.
* params:
*   - sheet_number: sheet index among all existent sheets.
*   - sheet: the sheet data container.
* returns:
*   - 0: FAIL. Check xlsx_errno and compare it against enum xlsx_load_sheet_errno to know what happened.
*   - 1: SUCCESS.
*/
static int parse_sheet(int sheet_number, xlsx_sheet_t * sheet) {
  XMLDoc *sheet_xml = malloc(sizeof(XMLDoc));
  if(!sheet_xml) {
    xlsx_errno = XLSX_LOAD_SHEET_ERRNO_OUT_OF_MEMORY;
    return 0; // FAIL
  }

  XMLDoc_init(sheet_xml);
  char sheet_number_as_string[12]; // an int can't occupy more than 11 chars
  sprintf(sheet_number_as_string, "%d", sheet_number);
  char path_to_sheet[260];
  strcpy(path_to_sheet, sheet->xlsx->deployment_path);
  strcat(path_to_sheet, REL_PATH_TO_WORKSHEETS);
  strcat(path_to_sheet, "sheet");
  strcat(path_to_sheet, sheet_number_as_string);
  strcat(path_to_sheet, ".xml");

  if(!(XMLDoc_parse_file_DOM(path_to_sheet, sheet_xml))) {
    XMLDoc_free(sheet_xml);
    free(sheet_xml);
    if(xlsx_print_err_messages)
      fprintf(stderr, "XLSX_C ERROR: \"%s\" can't be parsed or doesn't exist.\n", path_to_sheet);
    xlsx_errno = XLSX_LOAD_SHEET_ERRNO_XML_PARSING_ERROR;
    return 0; // FAIL
  }

  // initialize references_to_rows_ll
  XMLSearch search_engine;
  XMLSearch_init(&search_engine);
  XMLSearch_search_set_tag(&search_engine, SHEET_SHEETDATA_TAG);
  XMLNode *sheet_data_node = XMLSearch_next(sheet_xml->nodes[sheet_xml->i_root], &search_engine);
  if(!sheet_data_node) {
    XMLDoc_free(sheet_xml);
    free(sheet_xml);
    if(xlsx_print_err_messages)
      fprintf(stderr, "XLSX_C ERROR: There's no \"%s\" element inside #\"%d\" sheet.\n",
              SHEET_SHEETDATA_TAG, sheet_number);
    xlsx_errno = XLSX_LOAD_SHEET_ERRNO_XML_PARSING_ERROR;
    XMLSearch_free(&search_engine, false);
    return 0; // FAIL
  }

  // find out the last row
  if(sheet_data_node->n_children > 0) {
    int row_node_index, attr_index;
    XMLNode *row;
    for(row_node_index = sheet_data_node->n_children - 1; row_node_index >= 0; --row_node_index) {
      row = sheet_data_node->children[row_node_index];
      // check out if it has content (seek inside a <v> with content)
      XMLSearch_free(&search_engine, false);
      XMLSearch_init(&search_engine);
      XMLSearch_search_set_tag(&search_engine, SHEET_VALUE_TAG);
      XMLSearch_search_set_text(&search_engine, "*?*");
      if(XMLSearch_next(row, &search_engine)) {
        for(attr_index = 0; attr_index < row->n_attributes; ++attr_index) {
          if(strcmp(row->attributes[attr_index].name, SHEET_ROW_ATTR_NAME) == 0) {
            sheet->last_row = (int)strtol(row->attributes[attr_index].value, NULL, 10);
            break;
          }
        }
        // check for problem
        if(attr_index == row->n_attributes) {
          XMLDoc_free(sheet_xml);
          free(sheet_xml);
          if(xlsx_print_err_messages)
            fprintf(stderr, "XLSX_C ERROR: There's no \"%s\" attribute, inside node \"%s\", inside #\"%d\" sheet.\n",
                    SHEET_ROW_ATTR_NAME, SHEET_ROW_TAG, sheet_number);
          xlsx_errno = XLSX_LOAD_SHEET_ERRNO_XML_PARSING_ERROR;
          XMLSearch_free(&search_engine, false);
          return 0; // FAIL
        }
        break;
      }
    }
    if(row_node_index < 0) {
      // the sheet is completely empty
      sheet->last_row = 0;
    }
  } else {
    // the sheet is completely empty
    sheet->last_row = 0;
  }

  sheet->sheet_xml = sheet_xml;
  sheet->sheetdata = sheet_data_node;

  XMLSearch_free(&search_engine, false);

  return 1;
}


/*
* summary:
*   Finds and returns the XMLNode * row node seeked, NULL if doesn't find it. The seeking is made from certain child
*   of the sheetnode (*start_from_child*) of the sheet. While this looks for the sheet, updates
*   sheet->references_to_rows_ll and sheet->last_row_looked.
* params:
*   sheet: where to look in.
*   row: the row to be found.
*   start_from_child: the index of sheet->sheetdata->children[index] from where to start looking.
* returns:
*   XMLNode *: if found.
*   NULL: if not found.
*/
static XMLNode * find_row_node(xlsx_sheet_t *sheet, unsigned row, int start_from_child) {
  // *start_from_child* is re-used to index
  int row_inspected;
  XMLNode *row_node;
  for(; start_from_child < sheet->sheetdata->n_children; ++start_from_child) {
    row_node = sheet->sheetdata->children[start_from_child];
    row_inspected = (int)strtol(row_node->attributes[0].value, NULL, 10);
    if(row_inspected == row) {
      // row found
      sheet->last_row_looked.row_n = (int)row;
      sheet->last_row_looked.sheetdata_child_i = start_from_child;
      return(row_node);
    } else if(row_inspected > (int)row) {
      return NULL;
    }
  }
  // reached this point, the node wasn't found
  return NULL;
}


/*
* summary:
*   Given a *row* node, look in it for certain *cell*. Returns the XMLNode * cell node seeked, NULL if doesn't find
*   it.
* params:
*   row: the row node where to look for the cell node.
*   cell: the cell represented as a string, i.e.: "A5".
* returns:
*   XMLNode *: if found.
*   NULL: if not found.
*/
static XMLNode * find_cell_node(XMLNode *row, const char *cell) {
  int index;
  for(index = 0; index < row->n_children; ++index) {
    if(strcmp(row->children[index]->attributes[0].value, cell) == 0) {
      // cell found
      return row->children[index];
    }
  }
  // reached this point, the node wasn't found
  return NULL;
}


/*
* summary:
*   Given certain *cell* node, interpret its data, and fill the *cell_data_holder*.
* params:
*   cell: XMLNode * returned by a call to find_cell_node().
*   sheet: where this cell is.
*   cell_data_holder: will get filled with interpreted data.
* notes:
*   *cell_data_holder* will have:
*     - style == NULL && value_type == XLSX_POINTER_TO_CHAR when is TEXT
*     - style == NULL && value_type != XLSX_POINTER_TO_CHAR && value_type != XLSX_NULL when is NUMBER
*     - if style != NULL, inspect style.related_category to see what it is
*/
static void interpret_cell_node(XMLNode *cell, xlsx_sheet_t *sheet, xlsx_cell_t * cell_data_holder) {

  // check if has "t"
  if(strcmp(cell->attributes[cell->n_attributes - 1].name, SHEET_TYPE_ATTR_NAME) == 0) {

    // is some kind of text
    cell_data_holder->value_type = XLSX_POINTER_TO_CHAR;
    // check which one
    if(strcmp(cell->attributes[cell->n_attributes - 1].value, "s") == 0) {
      // it's a shared string
      int shared_strings_index = (int)strtol(cell->children[cell->n_children - 1]->text, NULL, 10);
      cell_data_holder->value.pointer_to_char_value = \
        sheet->xlsx->shared_strings_xml->nodes[1]->children[shared_strings_index]->children[0]->text;
      // it could have some associated style (i.e.: see sample.xlsx cell E21)
      if(strcmp(cell->attributes[cell->n_attributes - 2].name, SHEET_STYLE_ATTR_NAME) == 0) {
        int style_index = (int)strtol(cell->attributes[cell->n_attributes - 2].value, NULL, 10);
        cell_data_holder->style = sheet->xlsx->styles[style_index];
      }
    } else {
      // it's an inlineStr or an error. An error doesn't have style associated.
      cell_data_holder->value.pointer_to_char_value = cell->children[cell->n_children - 1]->text;
    }

  } else if(cell->n_children == 0) {

    // it has not value (i.e.: a cell that has a style set, but no value will still appear in sheet.xml)
    return;

  } else {

    // it's not a string, check if has value
    const char *cell_text = cell->children[cell->n_children - 1]->text;
    if(cell_text) {
      // check if it's a plain number or could be a complex type
      if(strcmp(cell->attributes[cell->n_attributes - 1].name, SHEET_STYLE_ATTR_NAME) == 0) {

        // could be a complex type
        int style_index = (int)strtol(cell->attributes[cell->n_attributes - 1].value, NULL, 10);
        cell_data_holder->style = sheet->xlsx->styles[style_index];
        // save the value where it should be
        set_cell_data_values_for_number(cell_text, cell_data_holder);

      } else {

        // it's just a plain number, see where should be the value saved
        set_cell_data_values_for_number(cell_text, cell_data_holder);

      }
    }
  }
}


/*
* summary:
*   Known that the *cell_text* represents a number, see to which primitive type commits better and reflect that inside
*   *cell_data_holder*. Note that bigger numbers (bigger than the largest number long long can hold, like > to 18
*   digits) will be represented by double. In this last case, the number appears with "E+" as part of the value
*   representation in the XML.
* params:
*   cell_text: text representation of the number.
*   cell_data_holder: xlsx_cell_t structure which holds data regarding a specific cell.
*/
static void set_cell_data_values_for_number(const char *cell_text, xlsx_cell_t *cell_data_holder) {
  if(strchr(cell_text, '.') || strchr(cell_text, 'E')) {
    // it's double (even if it's a big number, since it can't be represented by long long)
    cell_data_holder->value_type = XLSX_DOUBLE;
    cell_data_holder->value.double_value = strtod(cell_text, NULL);
  } else {
    // it's int or long long
    if(strlen(cell_text) > 9) {
      // fits on a long long
      cell_data_holder->value_type = XLSX_LONG_LONG;
      cell_data_holder->value.long_long_value = strtoll(cell_text, NULL, 10);
    } else {
      // fits on an int
      cell_data_holder->value_type = XLSX_INT;
      cell_data_holder->value.int_value = (int)(strtol(cell_text, NULL, 10));
    }
  }
}


/*
* returns:
*   - 1: everything went OK.
*   - 0: something happened and the process FAILED. Check errno against errno.h constant values.
* notes:
*   - remove() and rmdir() deals with both path separators.
*/
#if defined(_MSC_VER)
#include <fileapi.h>


static BOOL IsDots(const char* str) {
	if(strcmp(str, (".")) && strcmp(str, (".."))) return FALSE;
	return TRUE;
}

static int delete_folder(const char *sPath) {
	HANDLE hFind;  // file handle
	WIN32_FIND_DATAA FindFileData;

	TCHAR DirPath[MAX_PATH];
	TCHAR FileName[MAX_PATH];

	strcpy(DirPath, sPath);
	strcat(DirPath, ("\\*"));    // searching all files
	strcpy(FileName, sPath);
	strcat(FileName, ("\\"));

	hFind = FindFirstFileA(DirPath, &FindFileData); // find the first file
	if(hFind == INVALID_HANDLE_VALUE) return FALSE;
	strcpy(DirPath, FileName);

	BOOL bSearch = TRUE;
	while(bSearch) { // until we finds an entry
		if(FindNextFile(hFind, &FindFileData)) {
			if(IsDots(FindFileData.cFileName)) continue;
			strcat(FileName, FindFileData.cFileName);
			if((FindFileData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)) {

				// we have found a directory, recurse
				if(!delete_folder(FileName)) {
					FindClose(hFind);
					return 0; // directory couldn't be deleted
				}
				RemoveDirectory(FileName); // remove the empty directory
        strcpy(FileName, DirPath);
			}
			else {
				if(!DeleteFile(FileName)) {  // delete the file
					FindClose(hFind);
					return 0;
				}
				strcpy(FileName, DirPath);
			}
		}
		else {
			if(GetLastError() == ERROR_NO_MORE_FILES) // no more files there
				bSearch = FALSE;
			else {
				// some error occured, close the handle and return FALSE
				FindClose(hFind);
				return 0;
			}
		}
	}
	FindClose(hFind);  // closing file handle

	return RemoveDirectory(sPath); // remove the empty directory
}
#else
static int delete_folder(const char *folder_path) {

  DIR *dir = opendir(folder_path);

  XLSX_SET_ERRNO(0);
  long int index;
  struct dirent *entry;
  struct stat entry_statistics;
  char *f_basename;
  char f_fullname[260];

  while((entry = readdir(dir))) {
    f_basename = entry->d_name;
    strcpy(f_fullname, folder_path);
#if WINDOWS
    strcat(f_fullname, "\\");
#else
    strcat(f_fullname, "/");
#endif
    strcat(f_fullname, f_basename);

    stat(f_fullname, &entry_statistics); // can set errno

    if((S_ISDIR(entry_statistics.st_mode)) && (strcmp(f_basename, ".") != 0) && (strcmp(f_basename, "..") != 0)) {
      // it is a folder
      delete_folder(f_fullname);
    } else if(S_ISREG(entry_statistics.st_mode)) {
      // it is a file
      remove(f_fullname); // can set errno
    }
  }

  closedir(dir);

  // remove the dir looked
  rmdir(folder_path); // can set errno

  if(errno)
    return 0; // FAIL
  else
    return 1; // OK
}
#endif


/*
* summary:
*   After passing a string, it's parsed from the beginning stopping when a non-alphabetic char is found. A substring
*   till that point is returned. I.e.: You pass "ABC451", "ABC" is returned. It is known that the string passed will
*   always start at least with one alphabetic char, and that the alphabetic chars won't be more than 4 chars.
* params:
*   s_input: String input.
*   s_output: String output. Zero initialized.
*/
static void withdraw_alphabetic_chars(const char *s_input, char s_output[5]) {
  int char_index = 0;
  // as maximum will have 4 chars
  while(1) {
    if(s_input[char_index] >= 65 && s_input[char_index] <= 90) {
      s_output[char_index] = s_input[char_index];
      ++char_index;
    } else {
      // needed so it behaves as a string
      s_output[char_index] = '\0';
      break;
    }
  }
}