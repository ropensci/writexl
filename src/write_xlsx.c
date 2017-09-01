#define R_NO_REMAP
#define STRICT_R_HEADERS
#include <string.h>
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

char TEMPDIR[2048];

attribute_visible SEXP C_set_tempdir(SEXP dir){
  strcpy(TEMPDIR, Rf_translateChar(STRING_ELT(dir, 0)));
  return Rf_mkString(TEMPDIR);
}

attribute_visible SEXP C_write_data_frame_list(SEXP df_list, SEXP file, SEXP col_names){
  assert_that(Rf_isVectorList(df_list), "Object is not a list");
  assert_that(Rf_isString(file) && Rf_length(file), "Invalid file path");
  assert_that(Rf_isLogical(col_names), "col_names must be logical");

  lxw_workbook_options options;
  options.constant_memory = 1; //use less memory
  options.tmpdir = TEMPDIR;

  //create workbook
  lxw_workbook *workbook = workbook_new_opt(Rf_translateChar(STRING_ELT(file, 0)), &options);
  assert_that(workbook, "failed to create workbook");

  //how to format dates
  lxw_format * date = workbook_add_format(workbook);
  format_set_num_format(date, "yyyy-mm-dd HH:mm:ss UTC");

  //how to format headers (bold + center)
  lxw_format * title = workbook_add_format(workbook);
  format_set_bold(title);
  format_set_align(title, LXW_ALIGN_CENTER);

  //iterate over sheets
  SEXP df_names = PROTECT(Rf_getAttrib(df_list, R_NamesSymbol));
  for(size_t s = 0; s < Rf_length(df_list); s++){

    //create sheet
    const char * sheet_name = Rf_length(df_names) > s ? Rf_translateCharUTF8(STRING_ELT(df_names, s)) : NULL;
    lxw_worksheet *sheet = workbook_add_worksheet(workbook, sheet_name);
    assert_that(sheet, "failed to create workbook");

    //get data frame
    size_t cursor = 0;
    SEXP df = VECTOR_ELT(df_list, s);
    assert_that(Rf_inherits(df, "data.frame"), "object is not a data frame");

    //create header row
    if(Rf_asLogical(col_names)){
      SEXP names = PROTECT(Rf_getAttrib(df, R_NamesSymbol));
      for(size_t i = 0; i < Rf_length(names); i++)
        worksheet_write_string(sheet, cursor, i, Rf_translateCharUTF8(STRING_ELT(names, i)), NULL);
      worksheet_set_row(sheet, cursor, 15, title);
      UNPROTECT(1);
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
      if(coltypes[i] == COL_POSIXCT)
        worksheet_set_column(sheet, i, i, 20, date);
    }

    // Need to iterate by row first for performance
    for (size_t i = 0; i < rows; i++) {
      for(size_t j = 0; j < cols; j++){
        SEXP col = VECTOR_ELT(df, j);
        switch(coltypes[j]){
        case COL_POSIXCT: {
          double val = REAL(col)[i];
          if(R_FINITE(val))
            assert_lxw(worksheet_write_number(sheet, cursor, j, 25569 + val / (24*60*60) , NULL));
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
          if(val == R_PosInf)
            assert_lxw(worksheet_write_string(sheet, cursor, j, "Inf", NULL));
          else if(val == R_NegInf)
            assert_lxw(worksheet_write_string(sheet, cursor, j, "-Inf", NULL));
          else if(R_FINITE(val)) // skips NA and NAN
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
  }

  //this both writes the xlsx file and frees the memory
  UNPROTECT(1);
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
