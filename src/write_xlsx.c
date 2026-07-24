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
  COL_BLANK,
  COL_CELL_GENERAL,
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
  if(Rf_inherits(col, "xl_cell_general"))
    return COL_CELL_GENERAL;
  if(Rf_inherits(col, "Date"))
    return COL_DATE;
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

/* --- xl_cell_general support --------------------------------------------- */

/* Return the element named 'name' from a named R list, or R_NilValue */
static SEXP list_get(SEXP lst, const char *name){
  SEXP names = Rf_getAttrib(lst, R_NamesSymbol);
  if(names == R_NilValue) return R_NilValue;
  int n = Rf_length(names);
  for(int k = 0; k < n; k++){
    if(strcmp(CHAR(STRING_ELT(names, k)), name) == 0)
      return VECTOR_ELT(lst, k);
  }
  return R_NilValue;
}

/* --- Atomic value dispatcher (all non-xl_cell_general types) ------------- */

/* Called by both write_cell() and write_cell_general() to write a single
   atomic value.  No forward declaration needed: defined before both callers. */
static void write_atomic_value(cell_write_ctx *ctx, lxw_row_t row, lxw_col_t col,
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

/*
 * write_cell_general: write the i-th cell of an xl_cell_general column.
 * col is the full xl_cell_general list-column; VECTOR_ELT(col, i) is one
 * per-cell named list with elements: value, formula, hyperlink.
 *
 * Priority: hyperlink > formula > value.
 * For hyperlinks, a character value provides the optional display string.
 * For formulas, a numeric or character value is stored as a pre-calculated
 * result.  Plain values are dispatched through write_atomic_value().
 */
static void write_cell_general(cell_write_ctx *ctx,
                                lxw_row_t row, lxw_col_t col_idx,
                                SEXP col, lxw_row_t i){
  SEXP cell      = VECTOR_ELT(col, (R_xlen_t) i);
  SEXP value     = list_get(cell, "value");
  SEXP formula   = list_get(cell, "formula");
  SEXP hyperlink = list_get(cell, "hyperlink");

  /* --- hyperlink (character value provides the optional display string) --- */
  const char *display = NULL;
  if(value != R_NilValue && TYPEOF(value) == STRSXP &&
     Rf_length(value) > 0 && STRING_ELT(value, 0) != NA_STRING)
    display = Rf_translateCharUTF8(STRING_ELT(value, 0));

  if(hyperlink != R_NilValue && !Rf_isNull(hyperlink)){
    if(TYPEOF(hyperlink) == STRSXP && STRING_ELT(hyperlink, 0) != NA_STRING){
      assert_lxw(worksheet_write_url_opt(ctx->sheet, row, col_idx,
                   Rf_translateCharUTF8(STRING_ELT(hyperlink, 0)),
                   ctx->hyperlink_fmt, display, NULL));
      return;
    } else if(TYPEOF(hyperlink) == VECSXP){
      SEXP url_s = list_get(hyperlink, "url");
      SEXP tip_s = list_get(hyperlink, "tooltip");
      if(url_s != R_NilValue && TYPEOF(url_s) == STRSXP &&
         STRING_ELT(url_s, 0) != NA_STRING){
        const char *tooltip = (tip_s != R_NilValue && TYPEOF(tip_s) == STRSXP &&
                                STRING_ELT(tip_s, 0) != NA_STRING)
                               ? Rf_translateCharUTF8(STRING_ELT(tip_s, 0)) : NULL;
        assert_lxw(worksheet_write_url_opt(ctx->sheet, row, col_idx,
                     Rf_translateCharUTF8(STRING_ELT(url_s, 0)),
                     ctx->hyperlink_fmt, display, tooltip));
        return;
      }
    }
    /* hyperlink is NA or invalid — fall through to formula / value */
  }

  /* --- formula (with optional pre-calculated value) ----------------------- */
  if(formula != R_NilValue && TYPEOF(formula) == STRSXP &&
     STRING_ELT(formula, 0) != NA_STRING && Rf_length(formula) > 0){
    const char *fstr = Rf_translateCharUTF8(STRING_ELT(formula, 0));
    if(value != R_NilValue && TYPEOF(value) == REALSXP &&
       Rf_length(value) > 0 && R_FINITE(REAL(value)[0])){
      assert_lxw(worksheet_write_formula_num(ctx->sheet, row, col_idx,
                                              fstr, NULL, REAL(value)[0]));
    } else if(value != R_NilValue && TYPEOF(value) == STRSXP &&
              Rf_length(value) > 0 && STRING_ELT(value, 0) != NA_STRING){
      assert_lxw(worksheet_write_formula_str(ctx->sheet, row, col_idx, fstr, NULL,
                   Rf_translateCharUTF8(STRING_ELT(value, 0))));
    } else {
      assert_lxw(worksheet_write_formula(ctx->sheet, row, col_idx, fstr, NULL));
    }
    return;
  }

  /* --- value only: dispatch through the atomic writer --------------------- */
  if(value == R_NilValue || Rf_isNull(value) || Rf_length(value) == 0) return;
  write_atomic_value(ctx, row, col_idx, value, get_type(value), 0);
}

/* --- Top-level cell dispatcher ------------------------------------------- */

static void write_cell(cell_write_ctx *ctx, lxw_row_t row, lxw_col_t col,
                        SEXP col_data, R_COL_TYPE type, lxw_row_t i){
  if(type == COL_CELL_GENERAL)
    write_cell_general(ctx, row, col, col_data, i);
  else
    write_atomic_value(ctx, row, col, col_data, type, i);
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
