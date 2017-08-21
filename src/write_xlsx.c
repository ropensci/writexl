#define R_NO_REMAP
#define STRICT_R_HEADERS
#include <Rinternals.h>
#include <R_ext/Rdynload.h>
#include <R_ext/Visibility.h>
#include <xlsxwriter.h>

#define assert_that(a, b) bail_if(!a, b)

void bail_if(int check, const char * error){
  if(check)
    Rf_error("Error %s", error);
}

void assert_lxw(lxw_error err){
  if(err != LXW_NO_ERROR)
    Rf_errorcall(R_NilValue, "Error in libxlsxwriter: '%s'", lxw_strerror(err));
}

attribute_visible SEXP C_write_data_frame(SEXP df, SEXP file, SEXP headers){
  assert_that(Rf_inherits(df, "data.frame"), "Object is not a data frame");
  assert_that(Rf_isString(file) && Rf_length(file), "Invalid file path");
  assert_that(Rf_isString(headers), "Headers must be character vector");

  //create workbook
  lxw_workbook *workbook = workbook_new(Rf_translateChar(STRING_ELT(file, 0)));
  assert_that(workbook, "failed to create workbook");

  //create sheet
  lxw_worksheet *sheet = workbook_add_worksheet(workbook, NULL);
  assert_that(sheet, "failed to create workbook");

  //for headers
  lxw_format * bold = workbook_add_format(workbook);
  format_set_bold(bold);

  //for dates
  lxw_format * date = workbook_add_format(workbook);
  format_set_num_format(date, "yyyy-mm-d HH:MM AM/PM");

  //create header row
  size_t cursor = 0;
  if(Rf_isString(headers) && Rf_length(headers)){
    for(size_t i = 0; i < Rf_length(headers); i++)
      worksheet_write_string(sheet, cursor, i, CHAR(STRING_ELT(headers, i)), bold);
    cursor++;
  }

  // number of records
  size_t cols = Rf_length(df);
  size_t rows = cols ? Rf_length(VECTOR_ELT(df, 0)) : 0;
  if(rows == 0 || cols == 0)
    goto done;

  // Need to iterate by row first
  for (size_t i = 0; i < rows; i++) {
    for(size_t j = 0; j < cols; j++){
      SEXP col = VECTOR_ELT(df, j);

      // unclassed types
      switch(TYPEOF(col)){
      case STRSXP:
        assert_lxw(worksheet_write_string(sheet, cursor, j, CHAR(STRING_ELT(col, i)), NULL));
        continue;
      case INTSXP:
        assert_lxw(worksheet_write_number(sheet, cursor, j, INTEGER(col)[i], NULL));
        continue;
      case REALSXP:
        if(Rf_inherits(col, "POSIXct")){
          double val = REAL(col)[i];
          assert_lxw(worksheet_write_number(sheet, cursor, j, 25569 + val / (24*60*60) , date));
        } else {
          assert_lxw(worksheet_write_number(sheet, cursor, j, REAL(col)[i], NULL));
        }
        continue;
      case LGLSXP:
        assert_lxw(worksheet_write_boolean(sheet, cursor, j, LOGICAL(col)[i], NULL));
        continue;
      default:
        assert_lxw(worksheet_write_blank(sheet, cursor, j, NULL));
        continue;
      };
    }
    cursor++;
  }

  //this both writes the xlsx file and frees the memory
  done:
  workbook_close(workbook);
  return file;
}

attribute_visible SEXP C_lxw_version(){
  return Rf_mkString(LXW_VERSION);
}

attribute_visible void R_init_writexl(DllInfo* info) {
  R_registerRoutines(info, NULL, NULL, NULL, NULL);
  R_useDynamicSymbols(info, TRUE);
}
