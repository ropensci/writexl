#define R_NO_REMAP
#define STRICT_R_HEADERS
#include <Rinternals.h>
#include <R_ext/Rdynload.h>
#include <R_ext/Visibility.h>
#include <xlsxwriter.h>

typedef enum {
  COL_LOGCIAL,
  COL_REAL,
  COL_INTEGER,
  COL_STRING,
  COL_POSIXCT,
  COL_BLANK,
  COL_UNKNOWN
} R_COL_TYPE;

#define assert_that(a, b) bail_if(!a, b)

#define max(a,b) (a > b) ? a : b

void bail_if(int check, const char * error){
  if(check)
    Rf_error("Error %s", error);
}

void assert_lxw(lxw_error err){
  if(err != LXW_NO_ERROR)
    Rf_errorcall(R_NilValue, "Error in libxlsxwriter: '%s'", lxw_strerror(err));
}

R_COL_TYPE get_type(SEXP col){
  if(Rf_inherits(col, "POSIXct"))
    return COL_POSIXCT;
  switch(TYPEOF(col)){
  case STRSXP:
    return COL_STRING;
  case INTSXP:
    return COL_INTEGER;
  case REALSXP:
    return COL_REAL;
  case LGLSXP:
    return COL_LOGCIAL;
  default:
    return COL_UNKNOWN;
  };
}

attribute_visible SEXP C_write_data_frame(SEXP df, SEXP file, SEXP headers, SEXP date_format){
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
  lxw_format * title = workbook_add_format(workbook);
  format_set_bold(title);
  format_set_align(title, LXW_ALIGN_CENTER);

  //for dates
  lxw_format * date = workbook_add_format(workbook);
  format_set_num_format(date, CHAR(STRING_ELT(date_format, 0)));

  //create header row
  size_t cursor = 0;
  if(Rf_isString(headers) && Rf_length(headers)){
    for(size_t i = 0; i < Rf_length(headers); i++)
      worksheet_write_string(sheet, cursor, i, CHAR(STRING_ELT(headers, i)), title);
    cursor++;
  }

  // number of records
  size_t cols = Rf_length(df);
  size_t rows = 0;

  // determinte how to format each column
  R_COL_TYPE coltypes[cols];
  for(size_t i = 0; i < cols; i++){
    SEXP COL = VECTOR_ELT(df, i);
    coltypes[i] = get_type(COL);
    if(!Rf_isMatrix(COL) && !Rf_inherits(COL, "data.frame"))
      rows = max(rows, Rf_length(COL));
  }

  // Need to iterate by row first for performance
  for (size_t i = 0; i < rows; i++) {
    for(size_t j = 0; j < cols; j++){
      SEXP col = VECTOR_ELT(df, j);
      switch(coltypes[j]){
      case COL_POSIXCT: {
        double val = REAL(col)[i];
        if(R_FINITE(val))
          assert_lxw(worksheet_write_number(sheet, cursor, j, 25569 + val / (24*60*60) , date));
      }; continue;
      case COL_STRING:{
        SEXP val = STRING_ELT(col, i);
        if(val != NA_STRING && Rf_length(val))
          assert_lxw(worksheet_write_string(sheet, cursor, j, Rf_translateCharUTF8(val), NULL));
        else  // xlsx does string not supported it seems?
          assert_lxw(worksheet_write_string(sheet, cursor, j, " ", NULL));
      }; continue;
      case COL_REAL:{
        double val = REAL(col)[i];
        if(R_FINITE(val)) // TODO: distingish NA, NaN, Inf
          assert_lxw(worksheet_write_number(sheet, cursor, j, val, NULL));
      }; continue;
      case COL_INTEGER:{
        int val = INTEGER(col)[i];
        if(val != NA_INTEGER)
          assert_lxw(worksheet_write_number(sheet, cursor, j, val, NULL));
      }; continue;
      case COL_LOGCIAL:{
        int val = LOGICAL(col)[i];
        if(val != NA_LOGICAL)
          assert_lxw(worksheet_write_boolean(sheet, cursor, j, val, NULL));
        }; continue;
      default:
        continue;
      };
    }
    cursor++;
  }

  //this both writes the xlsx file and frees the memory
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
