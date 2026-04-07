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
  const char *src = Rf_translateChar(STRING_ELT(dir, 0));
  if(strlen(src) >= sizeof(TEMPDIR))
    Rf_errorcall(R_NilValue, "Error in writexl: tempdir path too long (max %zu bytes)", sizeof(TEMPDIR) - 1);
  strncpy(TEMPDIR, src, sizeof(TEMPDIR) - 1);
  return Rf_mkString(TEMPDIR);
}

/* Context struct shared across all cell-writing functions */
typedef struct {
  lxw_workbook  *workbook;      /* reserved for future use in cell-format support */
  lxw_worksheet *sheet;
  lxw_format    *date_fmt;
  lxw_format    *datetime_fmt;
  lxw_format    *hyperlink_fmt;
} cell_write_ctx;

/* --- Individual per-type cell writers ------------------------------------ */

static void write_cell_date(cell_write_ctx *ctx, lxw_row_t row, lxw_col_t col,
                             SEXP col_data, lxw_row_t i){
  double val = Rf_isReal(col_data) ? REAL(col_data)[i] : INTEGER(col_data)[i];
  if(Rf_isReal(col_data) ? R_FINITE(val) : val != NA_INTEGER)
    assert_lxw(worksheet_write_number(ctx->sheet, row, col, 25569 + val, NULL));
}

static void write_cell_posixct(cell_write_ctx *ctx, lxw_row_t row, lxw_col_t col,
                                SEXP col_data, lxw_row_t i){
  double val = REAL(col_data)[i];
  if(R_FINITE(val)){
    val = 25568.0 + val / (24*60*60);
    if(val >= 60.0)
      val = val + 1.0;
    assert_lxw(worksheet_write_number(ctx->sheet, row, col, val, NULL));
  }
}

static void write_cell_string(cell_write_ctx *ctx, lxw_row_t row, lxw_col_t col,
                               SEXP col_data, lxw_row_t i){
  SEXP val = STRING_ELT(col_data, i);
  // NB: xlsx does distinguish between empty string and NA
  if(val != NA_STRING && Rf_length(val))
    assert_lxw(worksheet_write_string(ctx->sheet, row, col, Rf_translateCharUTF8(val), NULL));
}

static void write_cell_formula(cell_write_ctx *ctx, lxw_row_t row, lxw_col_t col,
                                SEXP col_data, lxw_row_t i){
  SEXP val = STRING_ELT(col_data, i);
  if(val != NA_STRING && Rf_length(val))
    assert_lxw(worksheet_write_formula(ctx->sheet, row, col, Rf_translateCharUTF8(val), NULL));
}

static void write_cell_hyperlink(cell_write_ctx *ctx, lxw_row_t row, lxw_col_t col,
                                  SEXP col_data, lxw_row_t i){
  SEXP val = STRING_ELT(col_data, i);
  if(val != NA_STRING && Rf_length(val))
    assert_lxw(worksheet_write_formula(ctx->sheet, row, col, Rf_translateCharUTF8(val),
                                        ctx->hyperlink_fmt));
}

static void write_cell_real(cell_write_ctx *ctx, lxw_row_t row, lxw_col_t col,
                             SEXP col_data, lxw_row_t i){
  double val = REAL(col_data)[i];
  if(val == R_PosInf)
    assert_lxw(worksheet_write_string(ctx->sheet, row, col, "Inf", NULL));
  else if(val == R_NegInf)
    assert_lxw(worksheet_write_string(ctx->sheet, row, col, "-Inf", NULL));
  else if(R_FINITE(val)) // skips NA and NAN
    assert_lxw(worksheet_write_number(ctx->sheet, row, col, val, NULL));
}

static void write_cell_integer(cell_write_ctx *ctx, lxw_row_t row, lxw_col_t col,
                                SEXP col_data, lxw_row_t i){
  int val = INTEGER(col_data)[i];
  if(val != NA_INTEGER)
    assert_lxw(worksheet_write_number(ctx->sheet, row, col, val, NULL));
}

static void write_cell_logical(cell_write_ctx *ctx, lxw_row_t row, lxw_col_t col,
                                SEXP col_data, lxw_row_t i){
  int val = LOGICAL(col_data)[i];
  if(val != NA_LOGICAL)
    assert_lxw(worksheet_write_boolean(ctx->sheet, row, col, val, NULL));
}

/* --- General cell dispatcher --------------------------------------------- */

static void write_cell(cell_write_ctx *ctx, lxw_row_t row, lxw_col_t col,
                        SEXP col_data, R_COL_TYPE type, lxw_row_t i){
  switch(type){
  case COL_DATE:
    write_cell_date(ctx, row, col, col_data, i);
    break;
  case COL_POSIXCT:
    write_cell_posixct(ctx, row, col, col_data, i);
    break;
  case COL_STRING:
    write_cell_string(ctx, row, col, col_data, i);
    break;
  case COL_FORMULA:
    write_cell_formula(ctx, row, col, col_data, i);
    break;
  case COL_HYPERLINK:
    write_cell_hyperlink(ctx, row, col, col_data, i);
    break;
  case COL_REAL:
    write_cell_real(ctx, row, col, col_data, i);
    break;
  case COL_INTEGER:
    write_cell_integer(ctx, row, col, col_data, i);
    break;
  case COL_LOGICAL:
    write_cell_logical(ctx, row, col, col_data, i);
    break;
  default:
    break;
  }
}

/* --- Main entry point ---------------------------------------------------- */

SEXP C_write_data_frame_list(SEXP df_list, SEXP file, SEXP col_names, SEXP format_headers, SEXP use_zip64){
  assert_that(Rf_isVectorList(df_list), "Object is not a list");
  assert_that(Rf_isString(file) && Rf_length(file), "Invalid file path");
  assert_that(Rf_isLogical(col_names), "col_names must be logical");
  assert_that(Rf_isLogical(format_headers), "format_headers must be logical");

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

  //how to format headers (bold + center)
  lxw_format * title = workbook_add_format(workbook);
  format_set_bold(title);
  format_set_align(title, LXW_ALIGN_CENTER);

  //how to format hyperlinks (underline + blue)
  lxw_format * hyperlink = workbook_add_format(workbook);
  format_set_underline(hyperlink, LXW_UNDERLINE_SINGLE);
  format_set_font_color(hyperlink, LXW_COLOR_BLUE);

  //iterate over sheets
  SEXP df_names = PROTECT(Rf_getAttrib(df_list, R_NamesSymbol));
  R_xlen_t nsheets = Rf_length(df_list);
  for(R_xlen_t s = 0; s < nsheets; s++){

    //create sheet
    const char * sheet_name = Rf_length(df_names) > s && Rf_length(STRING_ELT(df_names, s)) ? \
      Rf_translateCharUTF8(STRING_ELT(df_names, s)) : NULL;
    lxw_worksheet *sheet = workbook_add_worksheet(workbook, sheet_name);
    assert_that(sheet, "failed to create workbook");

    //get data frame
    lxw_row_t cursor = 0;
    SEXP df = VECTOR_ELT(df_list, s);
    assert_that(Rf_inherits(df, "data.frame"), "object is not a data frame");
    SEXP names = PROTECT(Rf_getAttrib(df, R_NamesSymbol));

    // validate and get column count
    R_xlen_t ncols_r = Rf_length(df);
    bail_if(ncols_r > LXW_COL_MAX, "data frame has too many columns for xlsx (max 16384)");
    lxw_col_t cols = (lxw_col_t) ncols_r;

    //create header row
    if(Rf_asLogical(col_names)){
      for(lxw_col_t i = 0; i < cols; i++)
        assert_lxw(worksheet_write_string(sheet, cursor, i, Rf_translateCharUTF8(STRING_ELT(names, i)), NULL));
      if(Rf_asLogical(format_headers))
        assert_lxw(worksheet_set_row(sheet, cursor, 15, title));
      cursor++;
    }

    // number of records
    lxw_row_t rows = 0;

    // determinte how to format each column
    R_COL_TYPE coltypes[cols];
    for(lxw_col_t i = 0; i < cols; i++){
      SEXP COL = VECTOR_ELT(df, i);
      coltypes[i] = get_type(COL);
      if(!Rf_isMatrix(COL) && !Rf_inherits(COL, "data.frame")){
        lxw_row_t col_rows = (lxw_row_t) Rf_length(COL);
        if(col_rows > rows) rows = col_rows;
      }
      if(coltypes[i] == COL_DATE)
        assert_lxw(worksheet_set_column(sheet, i, i, 20, date));
      if(coltypes[i] == COL_POSIXCT)
        assert_lxw(worksheet_set_column(sheet, i, i, 20, datetime));
      if(coltypes[i] == COL_UNKNOWN)
        Rf_warning("Column '%s' has unrecognized data type.", CHAR(STRING_ELT(names, i)));
    }
    UNPROTECT(1); //names

    // validate row count against xlsx limit
    lxw_row_t max_rows = (lxw_row_t)(LXW_ROW_MAX - (Rf_asLogical(col_names) ? 1 : 0));
    bail_if(rows > max_rows, "data frame has too many rows for xlsx (max 1048576)");

    // Build context for cell writers
    cell_write_ctx ctx = {workbook, sheet, date, datetime, hyperlink};

    // Need to iterate by row first for performance
    for (lxw_row_t i = 0; i < rows; i++) {
      for(lxw_col_t j = 0; j < cols; j++){
        SEXP col = VECTOR_ELT(df, j);
        if(Rf_length(col) <= (R_xlen_t) i) continue;
        write_cell(&ctx, cursor, j, col, coltypes[j], i);
      }
      cursor++;
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
  {"C_write_data_frame_list", (DL_FUNC) &C_write_data_frame_list, 5},
  {NULL, NULL, 0}
};

attribute_visible void R_init_writexl(DllInfo *dll) {
  R_registerRoutines(dll, NULL, CallEntries, NULL, NULL);
  R_useDynamicSymbols(dll, FALSE);
  R_forceSymbols(dll, TRUE);
}
