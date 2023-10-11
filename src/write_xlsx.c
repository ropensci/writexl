#define R_NO_REMAP
#define STRICT_R_HEADERS
#include <string.h>
#include <Rinternals.h>
#include <R_ext/Rdynload.h>
#include <R_ext/Visibility.h>
#include <xlsxwriter.h>

typedef enum {
  COL_LOGICAL,
  COL_REAL,
  COL_INTEGER,
  COL_STRING,
  COL_DATE,
  COL_POSIXCT,
  COL_HYPERLINK,
  COL_FORMULA,
  COL_BLANK,
  COL_UNKNOWN
} R_COL_TYPE;

#define assert_that(a, b) bail_if(!a, b)

#define max(a,b) (a > b) ? a : b

static void bail_if(int check, const char * error){
  if(check)
    Rf_errorcall(R_NilValue, "Error in writexl: %s", error);
}

static void assert_lxw(lxw_error err){
  if(err != LXW_NO_ERROR)
    Rf_errorcall(R_NilValue, "Error in libxlsxwriter: '%s'", lxw_strerror(err));
}

static R_COL_TYPE get_type(SEXP col){
  if(Rf_inherits(col, "Date"))
    return COL_DATE;
  if(Rf_inherits(col, "POSIXct"))
    return COL_POSIXCT;
  if(Rf_inherits(col, "xl_hyperlink"))
    return COL_HYPERLINK;
  if(Rf_isString(col) && Rf_inherits(col, "xl_formula"))
    return COL_FORMULA;
  switch(TYPEOF(col)){
  case STRSXP:
    return COL_STRING;
  case INTSXP:
    return COL_INTEGER;
  case REALSXP:
    return COL_REAL;
  case LGLSXP:
    return COL_LOGICAL;
  default:
    return COL_UNKNOWN;
  };
}

//global options
static char TEMPDIR[2048] = {0};

//set to R tempdir when pkg is loaded
SEXP C_set_tempdir(SEXP dir){
  strcpy(TEMPDIR, Rf_translateChar(STRING_ELT(dir, 0)));
  return Rf_mkString(TEMPDIR);
}

SEXP C_write_data_frame_list(SEXP df_list,
                             SEXP file,
                             SEXP col_names,
                             SEXP format_headers,
                             SEXP use_zip64,
                             SEXP freeze_rows,
                             SEXP freeze_cols,
                             SEXP header_bg_color,
                             SEXP autofilter,
                             SEXP col_widths){
  assert_that(Rf_isVectorList(df_list), "Object is not a list");
  assert_that(Rf_isString(file) && Rf_length(file), "Invalid file path");
  assert_that(Rf_isLogical(col_names), "col_names must be logical");
  assert_that(Rf_isLogical(format_headers), "format_headers must be logical");
  assert_that(Rf_isInteger(freeze_rows), "freeze_rows must be integer");
  assert_that(Rf_isInteger(freeze_cols), "freeze_cols must be integer");
  assert_that(Rf_isInteger(header_bg_color), "header_bg_color must be integer");
  assert_that(Rf_isLogical(autofilter), "autofilter must be logical");
  assert_that(Rf_isVectorList(col_widths), "col_widths must be a list of vectors");

  //create workbook
  lxw_workbook_options options = {
    .constant_memory = 1,
    .tmpdir = TEMPDIR,
    .use_zip64 = Rf_asLogical(use_zip64)
  };
  lxw_workbook *workbook = workbook_new_opt(Rf_translateChar(STRING_ELT(file, 0)), &options);
  assert_that(workbook, "failed to create workbook");

  //how to format dates
  lxw_format * date = workbook_add_format(workbook);
  format_set_num_format(date, "yyyy-mm-dd");

  //how to format timetamps
  lxw_format * datetime = workbook_add_format(workbook);
  format_set_num_format(datetime, "yyyy-mm-dd HH:mm:ss UTC");

  //how to format headers (bold + center + background color)
  lxw_format * title = workbook_add_format(workbook);
  format_set_bold(title);
  format_set_align(title, LXW_ALIGN_CENTER);

  //set bg color
  if(Rf_asInteger(header_bg_color) > -1){
    format_set_bg_color(title, Rf_asInteger(header_bg_color));
  }




  //how to format hyperlinks (underline + blue)
  lxw_format * hyperlink = workbook_add_format(workbook);
  format_set_underline(hyperlink, LXW_UNDERLINE_SINGLE);
  format_set_font_color(hyperlink, LXW_COLOR_BLUE);

  //iterate over sheets
  SEXP df_names = PROTECT(Rf_getAttrib(df_list, R_NamesSymbol));
  for(size_t s = 0; s < Rf_length(df_list); s++){

    //create sheet
    const char * sheet_name = Rf_length(df_names) > s && Rf_length(STRING_ELT(df_names, s)) ? \
      Rf_translateCharUTF8(STRING_ELT(df_names, s)) : NULL;
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
        //format the headers
        if(Rf_asLogical(format_headers))
          worksheet_write_string(sheet, cursor, i, Rf_translateCharUTF8(STRING_ELT(names, i)), title);
        else
          worksheet_write_string(sheet, cursor, i, Rf_translateCharUTF8(STRING_ELT(names, i)), NULL);

      UNPROTECT(1);
      cursor++;
    }

    // number of records
    size_t cols = Rf_length(df);
    size_t rows = 0;

    // set column widths and determine format
    SEXP sheet_col_widths = VECTOR_ELT(col_widths, s);
    size_t len_widths = Rf_length(sheet_col_widths);
    double *widths_value;
    int widths_type = TYPEOF(sheet_col_widths);
    if (widths_type == REALSXP) {
      widths_value = REAL(sheet_col_widths);
    } else {
      widths_value = (double *)malloc(len_widths * sizeof(double));
      for(size_t i = 0; i < len_widths; i++){
        widths_value[i] = LXW_DEF_COL_WIDTH;
      }
    }

    R_COL_TYPE coltypes[cols];
    for(size_t i = 0; i < cols; i++){
      // set column width
      double width;
      width = ISNA(widths_value[i]) ? LXW_DEF_COL_WIDTH : widths_value[i];
      worksheet_set_column(sheet, i, i, width, NULL);

      // determine format
      SEXP COL = VECTOR_ELT(df, i);
      coltypes[i] = get_type(COL);
      if(!Rf_isMatrix(COL) && !Rf_inherits(COL, "data.frame"))
        rows = max(rows, Rf_length(COL));
      if(coltypes[i] == COL_DATE)
        assert_lxw(worksheet_set_column(sheet, i, i, width, date));
      if(coltypes[i] == COL_POSIXCT)
        assert_lxw(worksheet_set_column(sheet, i, i, width, datetime));
    }

    // Need to iterate by row first for performance
    for (size_t i = 0; i < rows; i++) {
      for(size_t j = 0; j < cols; j++){
        SEXP col = VECTOR_ELT(df, j);
        switch(coltypes[j]){
        case COL_DATE:{
          double val = Rf_isReal(col) ? REAL(col)[i] : INTEGER(col)[i];
          if(Rf_isReal(col) ? R_FINITE(val) : val != NA_INTEGER)
            assert_lxw(worksheet_write_number(sheet, cursor, j, 25569 + val, NULL));
        }; continue;
        case COL_POSIXCT: {
          double val = REAL(col)[i];
          if(R_FINITE(val)){
            val = 25568.0 + val / (24*60*60);
            if(val >= 60.0)
              val = val + 1.0;
            assert_lxw(worksheet_write_number(sheet, cursor, j, val , NULL));
          }
        }; continue;
        case COL_STRING:{
          SEXP val = STRING_ELT(col, i);
          // NB: xlsx does distinguish between empty string and NA
          if(val != NA_STRING && Rf_length(val))
            assert_lxw(worksheet_write_string(sheet, cursor, j, Rf_translateCharUTF8(val), NULL));
        }; continue;
        case COL_FORMULA:{
          SEXP val = STRING_ELT(col, i);
          if(val != NA_STRING && Rf_length(val))
            assert_lxw(worksheet_write_formula(sheet, cursor, j, Rf_translateCharUTF8(val), NULL));
        }; continue;
        case COL_HYPERLINK:{
          SEXP val = STRING_ELT(col, i);
          if(val != NA_STRING && Rf_length(val))
            assert_lxw(worksheet_write_formula(sheet, cursor, j, Rf_translateCharUTF8(val), hyperlink));
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
        case COL_LOGICAL:{
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
    //freeze rows
    if((Rf_asInteger(freeze_rows) + Rf_asInteger(freeze_cols)) > 0){
      worksheet_freeze_panes(sheet, Rf_asInteger(freeze_rows), Rf_asInteger(freeze_cols));
    }

    if(Rf_asLogical(autofilter)){
         worksheet_autofilter(sheet, 0, 0, rows, cols-1);
    }

  }



  //this both writes the xlsx file and frees the memory
  assert_lxw(workbook_close(workbook));
  UNPROTECT(1);
  return file;
}

SEXP C_lxw_version(void){
  return Rf_mkString(LXW_VERSION);
}

static const R_CallMethodDef CallEntries[] = {
  {"C_lxw_version",           (DL_FUNC) &C_lxw_version,           0},
  {"C_set_tempdir",           (DL_FUNC) &C_set_tempdir,           1},
  {"C_write_data_frame_list", (DL_FUNC) &C_write_data_frame_list, 10},
  {NULL, NULL, 0}
};

attribute_visible void R_init_writexl(DllInfo *dll) {
  R_registerRoutines(dll, NULL, CallEntries, NULL, NULL);
  R_useDynamicSymbols(dll, FALSE);
  R_forceSymbols(dll, TRUE);
}
