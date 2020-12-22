#include "../ext/unity.h"
#include "library.h"

/* even empty, this functions need to be here */
void setUp(void) {}
void tearDown(void) {}


void test_xlsx_open(void) {
  xlsx_workbook_t wb;
  xlsx_open("C:\\code\\c\\porcupine\\helpers\\sample.xlsx", &wb);

  // DEBUGGING
  // printf("Deployment path: %s", wb->deployment_path);
  // printf("Number of sheets: %u", wb->n_sheets);

  xlsx_close(&wb);
}

void test_xlsx_load_sheet(void) {
  TEST_IGNORE_MESSAGE("Not written yet.");
}

void test_xlsx_unload_sheet(void) {
  TEST_IGNORE_MESSAGE("Not written yet.");
}

void test_xlsx_read_cell(void) {
  TEST_IGNORE_MESSAGE("Not written yet.");
}

void test_xlsx_close(void) {
  TEST_IGNORE_MESSAGE("Not written yet.");
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
