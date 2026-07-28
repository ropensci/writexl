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
  TEMPDIR[sizeof(TEMPDIR) - 1] = '\0';
  return Rf_mkString(TEMPDIR);
}

/* Return the element named 'name' from a named R list, or R_NilValue */
static SEXP list_get(SEXP lst, const char *name){
  if(lst == R_NilValue || !Rf_isVectorList(lst)) return R_NilValue;
  SEXP names = Rf_getAttrib(lst, R_NamesSymbol);
  if(names == R_NilValue) return R_NilValue;
  int n = Rf_length(names);
  for(int k = 0; k < n; k++){
    if(strcmp(CHAR(STRING_ELT(names, k)), name) == 0)
      return VECTOR_ELT(lst, k);
  }
  return R_NilValue;
}

/* --- format payload accessors -------------------------------------------- */

static int payload_has(SEXP p, const char *key){
  return list_get(p, key) != R_NilValue;
}
static int payload_int(SEXP p, const char *key){
  SEXP v = list_get(p, key);
  return v == R_NilValue ? 0 : Rf_asInteger(v);
}
static double payload_dbl(SEXP p, const char *key){
  SEXP v = list_get(p, key);
  return v == R_NilValue ? 0.0 : Rf_asReal(v);
}
static const char *payload_str(SEXP p, const char *key){
  SEXP v = list_get(p, key);
  if(v == R_NilValue || TYPEOF(v) != STRSXP || Rf_length(v) < 1) return NULL;
  if(STRING_ELT(v, 0) == NA_STRING) return NULL;
  return Rf_translateCharUTF8(STRING_ELT(v, 0));
}

/*
 * Build a libxlsxwriter format object from a payload (a named R list produced
 * by .xl_format_payload() in R).  Each recognized key maps 1:1 to a
 * format_set_* call.  All enum/color translation is done in R, so this is a
 * faithful applier only.
 */
static lxw_format *build_lxw_format(lxw_workbook *wb, SEXP p){
  lxw_format *f = workbook_add_format(wb);
  const char *s;

  /* font */
  if((s = payload_str(p, "font_name")))   format_set_font_name(f, s);
  if(payload_has(p, "font_size"))         format_set_font_size(f, payload_dbl(p, "font_size"));
  if(payload_has(p, "font_color"))        format_set_font_color(f, (lxw_color_t) payload_int(p, "font_color"));
  if(payload_has(p, "bold"))              format_set_bold(f);
  if(payload_has(p, "italic"))            format_set_italic(f);
  if(payload_has(p, "underline"))         format_set_underline(f, (uint8_t) payload_int(p, "underline"));
  if(payload_has(p, "strikeout"))         format_set_font_strikeout(f);
  if(payload_has(p, "script"))            format_set_font_script(f, (uint8_t) payload_int(p, "script"));
  if(payload_has(p, "font_family"))       format_set_font_family(f, (uint8_t) payload_int(p, "font_family"));
  if(payload_has(p, "font_charset"))      format_set_font_charset(f, (uint8_t) payload_int(p, "font_charset"));
  if(payload_has(p, "font_outline"))      format_set_font_outline(f);
  if(payload_has(p, "font_shadow"))       format_set_font_shadow(f);
  if(payload_has(p, "font_condense"))     format_set_font_condense(f);
  if(payload_has(p, "font_extend"))       format_set_font_extend(f);
  if((s = payload_str(p, "font_scheme"))) format_set_font_scheme(f, s);
  if(payload_has(p, "theme"))             format_set_theme(f, (uint8_t) payload_int(p, "theme"));
  if(payload_has(p, "color_indexed"))     format_set_color_indexed(f, (uint8_t) payload_int(p, "color_indexed"));
  if(payload_has(p, "font_only"))         format_set_font_only(f);

  /* fill */
  if(payload_has(p, "bg_color"))          format_set_bg_color(f, (lxw_color_t) payload_int(p, "bg_color"));
  if(payload_has(p, "fg_color"))          format_set_fg_color(f, (lxw_color_t) payload_int(p, "fg_color"));
  if(payload_has(p, "pattern"))           format_set_pattern(f, (uint8_t) payload_int(p, "pattern"));

  /* border */
  if(payload_has(p, "border_left"))       format_set_left(f, (uint8_t) payload_int(p, "border_left"));
  if(payload_has(p, "border_right"))      format_set_right(f, (uint8_t) payload_int(p, "border_right"));
  if(payload_has(p, "border_top"))        format_set_top(f, (uint8_t) payload_int(p, "border_top"));
  if(payload_has(p, "border_bottom"))     format_set_bottom(f, (uint8_t) payload_int(p, "border_bottom"));
  if(payload_has(p, "left_color"))        format_set_left_color(f, (lxw_color_t) payload_int(p, "left_color"));
  if(payload_has(p, "right_color"))       format_set_right_color(f, (lxw_color_t) payload_int(p, "right_color"));
  if(payload_has(p, "top_color"))         format_set_top_color(f, (lxw_color_t) payload_int(p, "top_color"));
  if(payload_has(p, "bottom_color"))      format_set_bottom_color(f, (lxw_color_t) payload_int(p, "bottom_color"));
  if(payload_has(p, "diag_type"))         format_set_diag_type(f, (uint8_t) payload_int(p, "diag_type"));
  if(payload_has(p, "diag_border"))       format_set_diag_border(f, (uint8_t) payload_int(p, "diag_border"));
  if(payload_has(p, "diag_color"))        format_set_diag_color(f, (lxw_color_t) payload_int(p, "diag_color"));

  /* alignment */
  if(payload_has(p, "align_h"))           format_set_align(f, (uint8_t) payload_int(p, "align_h"));
  if(payload_has(p, "align_v"))           format_set_align(f, (uint8_t) payload_int(p, "align_v"));
  if(payload_has(p, "text_wrap"))         format_set_text_wrap(f);
  if(payload_has(p, "rotation"))          format_set_rotation(f, (int16_t) payload_int(p, "rotation"));
  if(payload_has(p, "indent"))            format_set_indent(f, (uint8_t) payload_int(p, "indent"));
  if(payload_has(p, "shrink"))            format_set_shrink(f);
  if(payload_has(p, "reading_order"))     format_set_reading_order(f, (uint8_t) payload_int(p, "reading_order"));

  /* number format */
  if((s = payload_str(p, "num_format")))  format_set_num_format(f, s);
  if(payload_has(p, "num_format_index"))  format_set_num_format_index(f, (uint8_t) payload_int(p, "num_format_index"));

  /* protection & misc */
  if(payload_has(p, "unlocked"))          format_set_unlocked(f);
  if(payload_has(p, "hidden"))            format_set_hidden(f);
  if(payload_has(p, "quote_prefix"))      format_set_quote_prefix(f);
  if(payload_has(p, "set_hyperlink"))     format_set_hyperlink(f);

  return f;
}

/* Context struct shared across all cell-writing functions */
typedef struct {
  lxw_workbook  *workbook;
  lxw_worksheet *sheet;
  lxw_format   **fmts;   /* 1-based table of user formats; fmts[0] == NULL   */
  int            nfmts;  /* number of entries in fmts (excluding slot 0)     */
  const int     *fmt_prot;         /* fmt_prot[id]: format sets unlock/hidden */
  int            sheet_protected;  /* is this worksheet protected?            */
  int           *warn_unlocked;    /* set when protection formatting is used  */
                                   /* on an unprotected sheet                 */
} cell_write_ctx;

/* Resolve a 1-based format id to a format pointer (0 or out-of-range -> NULL) */
static lxw_format *ctx_format(cell_write_ctx *ctx, int id){
  if(id <= 0 || id > ctx->nfmts) return NULL;
  return ctx->fmts[id];
}

/* Note when a format that unlocks/hides a cell is applied on a sheet that is
   not protected (in which case the cell-locking has no effect). */
static void note_protection(cell_write_ctx *ctx, int id){
  if(!ctx->sheet_protected && ctx->fmt_prot && id > 0 && id <= ctx->nfmts &&
     ctx->fmt_prot[id])
    *ctx->warn_unlocked = 1;
}

/* --- worksheet-level option appliers ------------------------------------- */

static int opt_scalar_int(SEXP opts, const char *key, int dflt){
  SEXP v = list_get(opts, key);
  if(v == R_NilValue || Rf_length(v) < 1) return dflt;
  int x = Rf_asInteger(v);
  return (x == NA_INTEGER) ? dflt : x;
}
static double opt_scalar_dbl(SEXP opts, const char *key, double dflt){
  SEXP v = list_get(opts, key);
  if(v == R_NilValue || Rf_length(v) < 1) return dflt;
  double x = Rf_asReal(v);
  return ISNA(x) ? dflt : x;
}

/* Apply per-column width / format / hidden / outline-level options. */
static void apply_columns(cell_write_ctx *ctx, SEXP opts, lxw_col_t cols){
  if(opts == R_NilValue) return;
  SEXP w  = list_get(opts, "col_width");
  SEXP fi = list_get(opts, "col_format_id");
  SEXP hd = list_get(opts, "col_hidden");
  SEXP lv = list_get(opts, "col_level");
  for(lxw_col_t i = 0; i < cols; i++){
    double width = (w  != R_NilValue && Rf_length(w)  > i) ? REAL(w)[i]     : NA_REAL;
    int    fid   = (fi != R_NilValue && Rf_length(fi) > i) ? INTEGER(fi)[i] : 0;
    int    hid   = (hd != R_NilValue && Rf_length(hd) > i) ? INTEGER(hd)[i] : NA_INTEGER;
    int    lev   = (lv != R_NilValue && Rf_length(lv) > i) ? INTEGER(lv)[i] : NA_INTEGER;
    lxw_format *fmt = ctx_format(ctx, fid);
    if(ISNA(width) && fmt == NULL && hid == NA_INTEGER && lev == NA_INTEGER)
      continue;
    note_protection(ctx, fid);
    lxw_row_col_options o = {0, 0, 0};
    if(hid != NA_INTEGER && hid > 0) o.hidden = 1;
    if(lev != NA_INTEGER && lev > 0) o.level = (uint8_t) lev;
    double wv = ISNA(width) ? LXW_DEF_COL_WIDTH : width;
    assert_lxw(worksheet_set_column_opt(ctx->sheet, i, i, wv, fmt, &o));
  }
}

/* Apply a row option for worksheet row `wrow`, if one is defined. */
static void apply_row(cell_write_ctx *ctx, SEXP opts, lxw_row_t wrow){
  if(opts == R_NilValue) return;
  SEXP rr = list_get(opts, "row_row");
  if(rr == R_NilValue) return;
  R_xlen_t n = Rf_length(rr);
  SEXP hh = list_get(opts, "row_height");
  SEXP fi = list_get(opts, "row_format_id");
  SEXP hd = list_get(opts, "row_hidden");
  SEXP lv = list_get(opts, "row_level");
  for(R_xlen_t k = 0; k < n; k++){
    if((lxw_row_t) INTEGER(rr)[k] != wrow) continue;
    double height = (hh != R_NilValue && !ISNA(REAL(hh)[k])) ? REAL(hh)[k] : LXW_DEF_ROW_HEIGHT;
    int    fid    = (fi != R_NilValue) ? INTEGER(fi)[k] : 0;
    int    hid    = (hd != R_NilValue) ? INTEGER(hd)[k] : NA_INTEGER;
    int    lev    = (lv != R_NilValue) ? INTEGER(lv)[k] : NA_INTEGER;
    lxw_row_col_options o = {0, 0, 0};
    if(hid != NA_INTEGER && hid > 0) o.hidden = 1;
    if(lev != NA_INTEGER && lev > 0) o.level = (uint8_t) lev;
    note_protection(ctx, fid);
    assert_lxw(worksheet_set_row_opt(ctx->sheet, wrow, height, ctx_format(ctx, fid), &o));
    return;
  }
}

/*
 * Apply the page setup: how the sheet prints.  Every value is a scalar
 * resolved on the R side, and every setter here is unconditional in
 * libxlsxwriter, so an absent key simply means "leave Excel's default".
 */
static void apply_page_setup(cell_write_ctx *ctx, SEXP opts){
  SEXP p = list_get(opts, "page");
  if(p == R_NilValue || !Rf_isVectorList(p)) return;
  lxw_worksheet *sheet = ctx->sheet;

  if(payload_has(p, "landscape")){
    if(payload_int(p, "landscape")) worksheet_set_landscape(sheet);
    else                            worksheet_set_portrait(sheet);
  }
  if(payload_has(p, "scale"))
    worksheet_set_print_scale(sheet, (uint16_t) payload_int(p, "scale"));
  if(payload_has(p, "first_page"))
    worksheet_set_start_page(sheet, (uint16_t) payload_int(p, "first_page"));

  /* worksheet_set_start_page() is the one setter that does not raise
     page_setup_changed, so a first_page on its own would emit no <pageSetup>
     element at all and be silently dropped.  worksheet_set_paper() does raise
     it, and paper size 0 ("printer default") writes no attribute, so it is a
     harmless way to say "this sheet has a page setup".  Called for first_page
     even when no paper was requested, and it doubles as the paper setter. */
  if(payload_has(p, "paper") || payload_has(p, "first_page"))
    worksheet_set_paper(sheet, (uint8_t) payload_int(p, "paper"));

  /* Margins are all-or-nothing in libxlsxwriter: a negative value means "keep
     Excel's default for this side", so an unset side passes -1 rather than 0. */
  if(payload_has(p, "margin_left") || payload_has(p, "margin_right") ||
     payload_has(p, "margin_top")  || payload_has(p, "margin_bottom")){
    double l = payload_has(p, "margin_left")   ? payload_dbl(p, "margin_left")   : -1;
    double r = payload_has(p, "margin_right")  ? payload_dbl(p, "margin_right")  : -1;
    double t = payload_has(p, "margin_top")    ? payload_dbl(p, "margin_top")    : -1;
    double b = payload_has(p, "margin_bottom") ? payload_dbl(p, "margin_bottom") : -1;
    worksheet_set_margins(sheet, l, r, t, b);
  }

  if(payload_has(p, "fit_width") || payload_has(p, "fit_height"))
    worksheet_fit_to_pages(sheet,
                           (uint16_t) payload_int(p, "fit_width"),
                           (uint16_t) payload_int(p, "fit_height"));

  /* Header and footer.  The _opt form carries the margin and any images; R has
     already checked that the &G/&[Picture] placeholders and the images agree
     in number, so libxlsxwriter's own count check is a backstop rather than
     the first line of defence. */
  const char *hdr = payload_str(p, "header");
  if(hdr){
    lxw_header_footer_options o;
    memset(&o, 0, sizeof(o));
    o.image_left   = (char *) payload_str(p, "header_image_left");
    o.image_center = (char *) payload_str(p, "header_image_center");
    o.image_right  = (char *) payload_str(p, "header_image_right");
    if(payload_has(p, "header_margin"))
      o.margin = payload_dbl(p, "header_margin");
    if(o.margin > 0.0 || o.image_left || o.image_center || o.image_right)
      assert_lxw(worksheet_set_header_opt(sheet, hdr, &o));
    else
      assert_lxw(worksheet_set_header(sheet, hdr));
  }
  const char *ftr = payload_str(p, "footer");
  if(ftr){
    lxw_header_footer_options o;
    memset(&o, 0, sizeof(o));
    o.image_left   = (char *) payload_str(p, "footer_image_left");
    o.image_center = (char *) payload_str(p, "footer_image_center");
    o.image_right  = (char *) payload_str(p, "footer_image_right");
    if(payload_has(p, "footer_margin"))
      o.margin = payload_dbl(p, "footer_margin");
    if(o.margin > 0.0 || o.image_left || o.image_center || o.image_right)
      assert_lxw(worksheet_set_footer_opt(sheet, ftr, &o));
    else
      assert_lxw(worksheet_set_footer(sheet, ftr));
  }

  if(payload_int(p, "center_horizontally")) worksheet_center_horizontally(sheet);
  if(payload_int(p, "center_vertically"))   worksheet_center_vertically(sheet);
  if(payload_int(p, "page_view"))           worksheet_set_page_view(sheet);
  if(payload_int(p, "across"))              worksheet_print_across(sheet);
  if(payload_int(p, "black_and_white"))     worksheet_print_black_and_white(sheet);
  if(payload_int(p, "row_col_headers"))     worksheet_print_row_col_headers(sheet);

  /* Print area and repeat rows/columns.  These write defined names into
     workbook.xml rather than anything in the worksheet part.  Both quads are
     already 0-based, resolved on the R side. */
  SEXP pa = list_get(p, "print_area");
  if(pa != R_NilValue && Rf_length(pa) >= 4){
    SEXP q = PROTECT(Rf_coerceVector(pa, INTSXP));
    int *a = INTEGER(q);
    assert_lxw(worksheet_print_area(sheet, (lxw_row_t) a[0], (lxw_col_t) a[1],
                                    (lxw_row_t) a[2], (lxw_col_t) a[3]));
    UNPROTECT(1);
  }
  SEXP rr = list_get(p, "repeat_rows");
  if(rr != R_NilValue && Rf_length(rr) >= 2){
    SEXP q = PROTECT(Rf_coerceVector(rr, INTSXP));
    int *a = INTEGER(q);
    assert_lxw(worksheet_repeat_rows(sheet, (lxw_row_t) a[0], (lxw_row_t) a[1]));
    UNPROTECT(1);
  }
  SEXP rc = list_get(p, "repeat_cols");
  if(rc != R_NilValue && Rf_length(rc) >= 2){
    SEXP q = PROTECT(Rf_coerceVector(rc, INTSXP));
    int *a = INTEGER(q);
    assert_lxw(worksheet_repeat_columns(sheet, (lxw_col_t) a[0], (lxw_col_t) a[1]));
    UNPROTECT(1);
  }

  /* Page breaks.  libxlsxwriter takes a zero-terminated array, so the buffer is
     one longer than the break count and R has already guaranteed no value is
     0 (which would end the list early). */
  SEXP hb = list_get(p, "h_breaks");
  if(hb != R_NilValue && Rf_length(hb) > 0){
    SEXP q = PROTECT(Rf_coerceVector(hb, INTSXP));
    R_xlen_t n = Rf_length(q);
    lxw_row_t *b = (lxw_row_t *) R_alloc((size_t) n + 1, sizeof(lxw_row_t));
    for(R_xlen_t k = 0; k < n; k++) b[k] = (lxw_row_t) INTEGER(q)[k];
    b[n] = 0;
    assert_lxw(worksheet_set_h_pagebreaks(sheet, b));
    UNPROTECT(1);
  }
  SEXP vb = list_get(p, "v_breaks");
  if(vb != R_NilValue && Rf_length(vb) > 0){
    SEXP q = PROTECT(Rf_coerceVector(vb, INTSXP));
    R_xlen_t n = Rf_length(q);
    lxw_col_t *b = (lxw_col_t *) R_alloc((size_t) n + 1, sizeof(lxw_col_t));
    for(R_xlen_t k = 0; k < n; k++) b[k] = (lxw_col_t) INTEGER(q)[k];
    b[n] = 0;
    assert_lxw(worksheet_set_v_pagebreaks(sheet, b));
    UNPROTECT(1);
  }
}

/*
 * Apply the sheet's tab state.  The rules that make these mutually exclusive
 * span the whole workbook and are checked in R by .resolve_sheet_visibility();
 * libxlsxwriter enforces none of them, so by the time we get here the
 * combination is known to be sane.
 */
/* How the outline (grouping) controls are drawn.  Grouping itself comes from
   the per-column/row `level`; this only changes the symbols. */
static void apply_outline(cell_write_ctx *ctx, SEXP opts){
  SEXP o = list_get(opts, "outline");
  if(o == R_NilValue || !Rf_isVectorList(o)) return;
  worksheet_outline_settings(ctx->sheet,
                             (uint8_t) payload_int(o, "visible"),
                             (uint8_t) payload_int(o, "symbols_below"),
                             (uint8_t) payload_int(o, "symbols_right"),
                             (uint8_t) payload_int(o, "auto_style"));
}

static void apply_sheet_view(cell_write_ctx *ctx, SEXP opts){
  SEXP v = list_get(opts, "view");
  if(v == R_NilValue || !Rf_isVectorList(v)) return;
  if(payload_int(v, "active"))    worksheet_activate(ctx->sheet);
  if(payload_int(v, "selected"))  worksheet_select(ctx->sheet);
  if(payload_has(v, "visible") && !payload_int(v, "visible"))
    worksheet_hide(ctx->sheet);
  if(payload_int(v, "first_tab")) worksheet_set_first_sheet(ctx->sheet);
  if(payload_int(v, "hide_zero"))     worksheet_hide_zero(ctx->sheet);
  if(payload_int(v, "right_to_left")) worksheet_right_to_left(ctx->sheet);

  /* Selected range and scroll position, both 0-based quads from R. */
  SEXP sel = list_get(v, "selection");
  if(sel != R_NilValue && Rf_length(sel) >= 4){
    SEXP q = PROTECT(Rf_coerceVector(sel, INTSXP));
    int *a = INTEGER(q);
    assert_lxw(worksheet_set_selection(ctx->sheet, (lxw_row_t) a[0],
                                       (lxw_col_t) a[1], (lxw_row_t) a[2],
                                       (lxw_col_t) a[3]));
    UNPROTECT(1);
  }
  /* Split panes.  R has already converted the cell reference into
     libxlsxwriter's row-height / column-width distances. */
  SEXP sp = list_get(v, "split");
  if(sp != R_NilValue && Rf_length(sp) >= 2){
    SEXP q = PROTECT(Rf_coerceVector(sp, REALSXP));
    worksheet_split_panes(ctx->sheet, REAL(q)[0], REAL(q)[1]);
    UNPROTECT(1);
  }

  SEXP tl = list_get(v, "top_left");
  if(tl != R_NilValue && Rf_length(tl) >= 2){
    SEXP q = PROTECT(Rf_coerceVector(tl, INTSXP));
    int *a = INTEGER(q);
    worksheet_set_top_left_cell(ctx->sheet, (lxw_row_t) a[0], (lxw_col_t) a[1]);
    UNPROTECT(1);
  }
}

/*
 * Merge each requested range.  Called after the sheet's rows have been written;
 * see the call site for why.  Ranges are 0-based quads resolved in R, which has
 * also rejected the single-cell case libxlsxwriter refuses.
 */
static void apply_merges(cell_write_ctx *ctx, SEXP opts){
  SEXP ms = list_get(opts, "merges");
  if(ms == R_NilValue || !Rf_isVectorList(ms)) return;
  for(R_xlen_t k = 0; k < Rf_length(ms); k++){
    SEXP m = VECTOR_ELT(ms, k);
    SEXP rng = list_get(m, "range");
    if(rng == R_NilValue || Rf_length(rng) < 4) continue;
    SEXP q = PROTECT(Rf_coerceVector(rng, INTSXP));
    int *a = INTEGER(q);
    const char *txt = payload_str(m, "text");
    int fid = payload_int(m, "format_id");
    note_protection(ctx, fid);
    /* a NULL string would be rejected, so an empty merge writes "" */
    assert_lxw(worksheet_merge_range(ctx->sheet, (lxw_row_t) a[0],
                                     (lxw_col_t) a[1], (lxw_row_t) a[2],
                                     (lxw_col_t) a[3], txt ? txt : "",
                                     ctx_format(ctx, fid)));
    UNPROTECT(1);
  }
}

/*
 * Place the sheet's images.  Applied after the row loop beside the merges and
 * tables: an embedded image writes a cell, and a floating one anchors to a row
 * whose height the row loop may still be setting.
 *
 * R sends either `filename` or `buffer`, which picks between the file and the
 * _buffer variants.  Options are only filled in where R supplied them, so an
 * unset field keeps libxlsxwriter's own default rather than a zero.
 */
static void apply_images(cell_write_ctx *ctx, SEXP opts){
  SEXP ims = list_get(opts, "images");
  if(ims == R_NilValue || !Rf_isVectorList(ims)) return;
  for(R_xlen_t k = 0; k < Rf_length(ims); k++){
    SEXP im = VECTOR_ELT(ims, k);
    lxw_row_t row = (lxw_row_t) payload_int(im, "row");
    lxw_col_t col = (lxw_col_t) payload_int(im, "col");
    int embed = payload_int(im, "embed");

    lxw_image_options o = {0};
    if(payload_has(im, "x_scale"))  o.x_scale  = payload_dbl(im, "x_scale");
    if(payload_has(im, "y_scale"))  o.y_scale  = payload_dbl(im, "y_scale");
    if(payload_has(im, "x_offset")) o.x_offset = payload_int(im, "x_offset");
    if(payload_has(im, "y_offset")) o.y_offset = payload_int(im, "y_offset");
    if(payload_has(im, "object_position"))
      o.object_position = (uint8_t) payload_int(im, "object_position");
    o.description = (char *) payload_str(im, "description");
    o.url         = (char *) payload_str(im, "url");
    o.tip         = (char *) payload_str(im, "tip");
    if(payload_has(im, "decorative")) o.decorative = 1;
    if(payload_has(im, "cell_format_id")){
      int fid = payload_int(im, "cell_format_id");
      note_protection(ctx, fid);
      o.cell_format = ctx_format(ctx, fid);
    }

    SEXP buf = list_get(im, "buffer");
    if(buf != R_NilValue && TYPEOF(buf) == RAWSXP){
      const unsigned char *bytes = (const unsigned char *) RAW(buf);
      size_t n = (size_t) Rf_xlength(buf);
      if(embed)
        assert_lxw(worksheet_embed_image_buffer_opt(ctx->sheet, row, col,
                                                    bytes, n, &o));
      else
        assert_lxw(worksheet_insert_image_buffer_opt(ctx->sheet, row, col,
                                                     bytes, n, &o));
    } else {
      const char *file = payload_str(im, "filename");
      bail_if(file == NULL, "image payload has neither a filename nor a buffer");
      if(embed)
        assert_lxw(worksheet_embed_image_opt(ctx->sheet, row, col, file, &o));
      else
        assert_lxw(worksheet_insert_image_opt(ctx->sheet, row, col, file, &o));
    }
  }
}

/* Apply per-sheet scalar options (freeze panes, gridlines, tab color, ...). */
static void apply_sheet_scalars(cell_write_ctx *ctx, SEXP opts){
  if(opts == R_NilValue) return;
  int fr = opt_scalar_int(opts, "freeze_row", -1);
  int fc = opt_scalar_int(opts, "freeze_col", -1);
  if(fr >= 0 && fc >= 0) worksheet_freeze_panes(ctx->sheet, (lxw_row_t) fr, (lxw_col_t) fc);
  int gl = opt_scalar_int(opts, "gridlines", -1);
  if(gl >= 0) worksheet_gridlines(ctx->sheet, (uint8_t) gl);
  int tc = opt_scalar_int(opts, "tab_color", -1);
  if(tc >= 0) worksheet_set_tab_color(ctx->sheet, (lxw_color_t) tc);
  int zoom = opt_scalar_int(opts, "zoom", 0);
  if(zoom > 0) worksheet_set_zoom(ctx->sheet, (uint16_t) zoom);
  double drh = opt_scalar_dbl(opts, "default_row_height", NA_REAL);
  if(!ISNA(drh)) worksheet_set_default_row(ctx->sheet, drh, 0);

  /* comment defaults (set before any per-cell comment is written) */
  const char *comment_author = payload_str(opts, "comment_author");
  if(comment_author) worksheet_set_comments_author(ctx->sheet, comment_author);
  if(opt_scalar_int(opts, "show_comments", 0)) worksheet_show_comments(ctx->sheet);

  /* worksheet protection */
  if(opt_scalar_int(opts, "protect", 0)){
    const char *pw = payload_str(opts, "protect_password");
    SEXP po = list_get(opts, "protect_options");
    if(po == R_NilValue){
      worksheet_protect(ctx->sheet, pw, NULL);
    } else {
      lxw_protection prot;
      memset(&prot, 0, sizeof(prot));
      prot.no_select_locked_cells   = (uint8_t) opt_scalar_int(po, "no_select_locked_cells", 0);
      prot.no_select_unlocked_cells = (uint8_t) opt_scalar_int(po, "no_select_unlocked_cells", 0);
      prot.format_cells      = (uint8_t) opt_scalar_int(po, "format_cells", 0);
      prot.format_columns    = (uint8_t) opt_scalar_int(po, "format_columns", 0);
      prot.format_rows       = (uint8_t) opt_scalar_int(po, "format_rows", 0);
      prot.insert_columns    = (uint8_t) opt_scalar_int(po, "insert_columns", 0);
      prot.insert_rows       = (uint8_t) opt_scalar_int(po, "insert_rows", 0);
      prot.insert_hyperlinks = (uint8_t) opt_scalar_int(po, "insert_hyperlinks", 0);
      prot.delete_columns    = (uint8_t) opt_scalar_int(po, "delete_columns", 0);
      prot.delete_rows       = (uint8_t) opt_scalar_int(po, "delete_rows", 0);
      prot.sort              = (uint8_t) opt_scalar_int(po, "sort", 0);
      prot.autofilter        = (uint8_t) opt_scalar_int(po, "autofilter", 0);
      prot.pivot_tables      = (uint8_t) opt_scalar_int(po, "pivot_tables", 0);
      prot.scenarios         = (uint8_t) opt_scalar_int(po, "scenarios", 0);
      prot.objects           = (uint8_t) opt_scalar_int(po, "objects", 0);
      worksheet_protect(ctx->sheet, pw, &prot);
    }
  }
}

/* --- Sheet overlays ------------------------------------------------------
 *
 * Range-scoped worksheet features (the autofilter today; merged cells, data
 * validation, conditional formats, images, charts and tables later) are all
 * applied here, in one pass after the per-sheet scalars and *before* the row
 * loop: under constant_memory each row is flushed as it is written, so
 * anything spanning rows has to be declared up front.
 *
 * Each payload is a named R list built by .resolve_sheet_plan() carrying a
 * `kind` string.  A new feature adds a kind here rather than another argument
 * to C_write_data_frame_list().
 */

/* Read a length-4 0-based range (first_row, first_col, last_row, last_col). */
static int overlay_range(SEXP p, const char *key, int *out){
  SEXP v = list_get(p, key);
  if(v == R_NilValue || Rf_length(v) < 4) return 0;
  SEXP iv = PROTECT(Rf_coerceVector(v, INTSXP));
  for(int k = 0; k < 4; k++) out[k] = INTEGER(iv)[k];
  UNPROTECT(1);
  return 1;
}

/* Read a datetime broken out into fields on the R side into an lxw_datetime. */
static void payload_datetime(SEXP p, const char *key, lxw_datetime *dt){
  SEXP v = list_get(p, key);
  if(v == R_NilValue) return;
  dt->year  = payload_int(v, "year");
  dt->month = payload_int(v, "month");
  dt->day   = payload_int(v, "day");
  dt->hour  = payload_int(v, "hour");
  dt->min   = payload_int(v, "min");
  dt->sec   = payload_dbl(v, "sec");
}

/*
 * Build and apply one lxw_data_validation.  Every field arrives resolved from
 * R -- enums as integers, limits already sorted into the number / formula /
 * datetime slot each belongs in -- so this is a faithful applier only.
 *
 * The three "on by default" toggles are only sent when the caller set them, so
 * an absent key must leave libxlsxwriter's default alone rather than writing 0.
 */
static void apply_validation(cell_write_ctx *ctx, SEXP p, int *r){
  lxw_data_validation v;
  memset(&v, 0, sizeof(v));
  v.validate = (uint8_t) payload_int(p, "validate");
  v.criteria = (uint8_t) payload_int(p, "criteria");
  if(payload_has(p, "error_type"))
    v.error_type = (uint8_t) payload_int(p, "error_type");

  /* lxw_validation_boolean is DEFAULT/OFF/ON, and the memset above already
     leaves DEFAULT, which libxlsxwriter reads as "on" for all four of these.
     R sends OFF or ON explicitly when the caller set one. */
  v.ignore_blank = (uint8_t) payload_int(p, "ignore_blank");
  v.show_input   = (uint8_t) payload_int(p, "show_input");
  v.show_error   = (uint8_t) payload_int(p, "show_error");
  v.dropdown     = (uint8_t) payload_int(p, "dropdown");

  if(payload_has(p, "value_number"))   v.value_number   = payload_dbl(p, "value_number");
  if(payload_has(p, "minimum_number")) v.minimum_number = payload_dbl(p, "minimum_number");
  if(payload_has(p, "maximum_number")) v.maximum_number = payload_dbl(p, "maximum_number");

  v.value_formula   = payload_str(p, "value_formula");
  v.minimum_formula = payload_str(p, "minimum_formula");
  v.maximum_formula = payload_str(p, "maximum_formula");

  payload_datetime(p, "value_datetime",   &v.value_datetime);
  payload_datetime(p, "minimum_datetime", &v.minimum_datetime);
  payload_datetime(p, "maximum_datetime", &v.maximum_datetime);

  v.input_title   = payload_str(p, "input_title");
  v.input_message = payload_str(p, "input_message");
  v.error_title   = payload_str(p, "error_title");
  v.error_message = payload_str(p, "error_message");

  /* A dropdown's choices become a NULL-terminated char* array, which
     libxlsxwriter joins into the CSV formula Excel stores. */
  SEXP lst = list_get(p, "value_list");
  if(lst != R_NilValue && TYPEOF(lst) == STRSXP && Rf_length(lst) > 0){
    R_xlen_t n = Rf_length(lst);
    const char **items = (const char **) R_alloc((size_t) n + 1, sizeof(char *));
    for(R_xlen_t k = 0; k < n; k++)
      items[k] = Rf_translateCharUTF8(STRING_ELT(lst, k));
    items[n] = NULL;
    v.value_list = items;
  }

  assert_lxw(worksheet_data_validation_range(ctx->sheet, (lxw_row_t) r[0],
                                             (lxw_col_t) r[1], (lxw_row_t) r[2],
                                             (lxw_col_t) r[3], &v));
}

/*
 * Build and apply one lxw_conditional_format.  As with data validation, every
 * enum arrives resolved from R, including the type/criteria pairing that
 * libxlsxwriter itself does not check.
 */
static void apply_conditional(cell_write_ctx *ctx, SEXP p, int *r){
  lxw_conditional_format c;
  memset(&c, 0, sizeof(c));
  c.type     = (uint8_t) payload_int(p, "type");
  c.criteria = (uint8_t) payload_int(p, "criteria");

  if(payload_has(p, "value"))     c.value     = payload_dbl(p, "value");
  if(payload_has(p, "min_value")) c.min_value = payload_dbl(p, "min_value");
  if(payload_has(p, "max_value")) c.max_value = payload_dbl(p, "max_value");
  c.value_string     = payload_str(p, "value_string");
  c.min_value_string = payload_str(p, "min_value_string");
  c.max_value_string = payload_str(p, "max_value_string");

  int fid = payload_int(p, "format_id");
  note_protection(ctx, fid);
  c.format = ctx_format(ctx, fid);

  /* colour scale / data bar points: each end carries a rule type, a value and
     (for scales) a colour */
  c.min_rule_type = (uint8_t) payload_int(p, "min_rule_type");
  c.mid_rule_type = (uint8_t) payload_int(p, "mid_rule_type");
  c.max_rule_type = (uint8_t) payload_int(p, "max_rule_type");
  if(payload_has(p, "mid_value")) c.mid_value = payload_dbl(p, "mid_value");
  c.mid_value_string = payload_str(p, "mid_value_string");
  if(payload_has(p, "min_color")) c.min_color = (lxw_color_t) payload_int(p, "min_color");
  if(payload_has(p, "mid_color")) c.mid_color = (lxw_color_t) payload_int(p, "mid_color");
  if(payload_has(p, "max_color")) c.max_color = (lxw_color_t) payload_int(p, "max_color");

  /* data bars */
  if(payload_has(p, "bar_color"))
    c.bar_color = (lxw_color_t) payload_int(p, "bar_color");
  if(payload_has(p, "bar_negative_color"))
    c.bar_negative_color = (lxw_color_t) payload_int(p, "bar_negative_color");
  if(payload_has(p, "bar_border_color"))
    c.bar_border_color = (lxw_color_t) payload_int(p, "bar_border_color");
  if(payload_has(p, "bar_negative_border_color"))
    c.bar_negative_border_color = (lxw_color_t) payload_int(p, "bar_negative_border_color");
  if(payload_has(p, "bar_axis_color"))
    c.bar_axis_color = (lxw_color_t) payload_int(p, "bar_axis_color");
  c.bar_solid         = (uint8_t) payload_int(p, "bar_solid");
  c.bar_no_border     = (uint8_t) payload_int(p, "bar_no_border");
  c.bar_only          = (uint8_t) payload_int(p, "bar_only");
  c.bar_direction     = (uint8_t) payload_int(p, "bar_direction");
  c.bar_axis_position = (uint8_t) payload_int(p, "bar_axis_position");

  /* icon sets: Excel's own icons, selected by index -- no image is embedded */
  c.icon_style    = (uint8_t) payload_int(p, "icon_style");
  c.reverse_icons = (uint8_t) payload_int(p, "reverse_icons");
  c.icons_only    = (uint8_t) payload_int(p, "icons_only");

  if(payload_has(p, "stop_if_true"))
    c.stop_if_true = (uint8_t) payload_int(p, "stop_if_true");
  c.multi_range = payload_str(p, "multi_range");

  assert_lxw(worksheet_conditional_format_range(ctx->sheet, (lxw_row_t) r[0],
                                                (lxw_col_t) r[1], (lxw_row_t) r[2],
                                                (lxw_col_t) r[3], &c));
}

/* Fill one lxw_filter_rule from a payload, using the given key suffix. */
static void filter_rule(SEXP p, const char *crit_key, const char *num_key,
                        const char *str_key, lxw_filter_rule *r){
  memset(r, 0, sizeof(*r));
  r->criteria = (uint8_t) payload_int(p, crit_key);
  if(payload_has(p, num_key)) r->value = payload_dbl(p, num_key);
  r->value_string = payload_str(p, str_key);
}

/*
 * Add the worksheet tables.  Applied after the row loop, beside the merges,
 * because worksheet_add_table() writes cells of its own: a header string per
 * column and, with a total row, the total cells.  Running it before the rows
 * would let the row loop overwrite them and lose any header format.
 *
 * Every column of the range carries a header from R -- the data frame's own
 * column name unless the caller overrode it -- because libxlsxwriter would
 * otherwise default them to "Column 1", "Column 2", ... and Excel rejects a
 * file whose table definition and header cells disagree.
 */
static void apply_tables(cell_write_ctx *ctx, SEXP opts){
  SEXP ts = list_get(opts, "tables");
  if(ts == R_NilValue || !Rf_isVectorList(ts)) return;
  for(R_xlen_t k = 0; k < Rf_length(ts); k++){
    SEXP t = VECTOR_ELT(ts, k);
    int r[4];
    if(!overlay_range(t, "range", r))
      Rf_errorcall(R_NilValue, "Error in writexl: table needs a length-4 range");

    lxw_table_options o = {0};
    o.name           = (char *) payload_str(t, "name");
    o.style_type     = (uint8_t) payload_int(t, "style_type");
    o.style_type_number = (uint8_t) payload_int(t, "style_number");
    /* R sends the positive sense of each flag; libxlsxwriter wants the
       negative one for three of them. */
    o.no_header_row  = (uint8_t) !payload_int(t, "header_row");
    o.no_autofilter  = (uint8_t) !payload_int(t, "autofilter");
    o.no_banded_rows = (uint8_t) !payload_int(t, "banded_rows");
    o.banded_columns = (uint8_t) payload_int(t, "banded_columns");
    o.first_column   = (uint8_t) payload_int(t, "first_column");
    o.last_column    = (uint8_t) payload_int(t, "last_column");
    o.total_row      = (uint8_t) payload_int(t, "total_row");

    SEXP cols = list_get(t, "columns");
    R_xlen_t ncols = (cols != R_NilValue && Rf_isVectorList(cols))
                     ? Rf_length(cols) : 0;
    lxw_table_column *entries = NULL;
    lxw_table_column **colptr = NULL;
    if(ncols > 0){
      entries = (lxw_table_column *) R_alloc((size_t) ncols, sizeof(lxw_table_column));
      colptr  = (lxw_table_column **) R_alloc((size_t) ncols + 1, sizeof(lxw_table_column *));
      memset(entries, 0, (size_t) ncols * sizeof(lxw_table_column));
      for(R_xlen_t i = 0; i < ncols; i++){
        SEXP c = VECTOR_ELT(cols, i);
        entries[i].header        = (char *) payload_str(c, "header");
        entries[i].formula       = (char *) payload_str(c, "formula");
        entries[i].total_string  = (char *) payload_str(c, "total_string");
        entries[i].total_function = (uint8_t) payload_int(c, "total_function");
        if(payload_has(c, "format_id")){
          int fid = payload_int(c, "format_id");
          note_protection(ctx, fid);
          entries[i].format = ctx_format(ctx, fid);
        }
        if(payload_has(c, "header_format_id")){
          int hid = payload_int(c, "header_format_id");
          note_protection(ctx, hid);
          entries[i].header_format = ctx_format(ctx, hid);
        }
        colptr[i] = &entries[i];
      }
      colptr[ncols] = NULL;
      o.columns = colptr;
    }

    assert_lxw(worksheet_add_table(ctx->sheet, (lxw_row_t) r[0], (lxw_col_t) r[1],
                                   (lxw_row_t) r[2], (lxw_col_t) r[3], &o));
  }
}

/*
 * Apply one column's autofilter criteria.  The rows this excludes are hidden
 * on the R side, because Excel stores the criteria and the hidden rows
 * independently and applies neither when the file is opened.
 */
static void apply_filter(cell_write_ctx *ctx, SEXP p){
  lxw_col_t col = (lxw_col_t) payload_int(p, "col");

  SEXP lst = list_get(p, "value_list");
  if(lst != R_NilValue && TYPEOF(lst) == STRSXP && Rf_length(lst) > 0){
    R_xlen_t n = Rf_length(lst);
    const char **items = (const char **) R_alloc((size_t) n + 1, sizeof(char *));
    for(R_xlen_t k = 0; k < n; k++)
      items[k] = Rf_translateCharUTF8(STRING_ELT(lst, k));
    items[n] = NULL;
    assert_lxw(worksheet_filter_list(ctx->sheet, col, items));
    return;
  }

  lxw_filter_rule r1;
  filter_rule(p, "criteria", "value", "value_string", &r1);
  if(payload_has(p, "criteria2")){
    lxw_filter_rule r2;
    filter_rule(p, "criteria2", "value2", "value_string2", &r2);
    assert_lxw(worksheet_filter_column2(ctx->sheet, col, &r1, &r2,
                                        (uint8_t) payload_int(p, "and_or")));
  } else {
    assert_lxw(worksheet_filter_column(ctx->sheet, col, &r1));
  }
}

static void apply_overlay(cell_write_ctx *ctx, SEXP p){
  const char *kind = payload_str(p, "kind");
  bail_if(kind == NULL, "sheet overlay payload is missing a 'kind'");
  if(strcmp(kind, "autofilter") == 0){
    int r[4];
    bail_if(!overlay_range(p, "range", r),
            "autofilter overlay needs a length-4 range");
    if(r[0] >= 0 && r[2] >= r[0] && r[3] >= r[1])
      assert_lxw(worksheet_autofilter(ctx->sheet, (lxw_row_t) r[0], (lxw_col_t) r[1],
                                      (lxw_row_t) r[2], (lxw_col_t) r[3]));
  } else if(strcmp(kind, "validation") == 0){
    int r[4];
    bail_if(!overlay_range(p, "range", r),
            "validation overlay needs a length-4 range");
    apply_validation(ctx, p, r);
  } else if(strcmp(kind, "conditional") == 0){
    int r[4];
    bail_if(!overlay_range(p, "range", r),
            "conditional overlay needs a length-4 range");
    apply_conditional(ctx, p, r);
  } else if(strcmp(kind, "filter") == 0){
    apply_filter(ctx, p);
  } else if(strcmp(kind, "ignore_errors") == 0){
    /* libxlsxwriter takes the range as an A1 string, so build it back from the
       0-based quad the shared range resolver produced on the R side. */
    int r[4];
    char range[LXW_MAX_CELL_RANGE_LENGTH];
    bail_if(!overlay_range(p, "range", r),
            "ignore_errors overlay needs a length-4 range");
    lxw_rowcol_to_range(range, (lxw_row_t) r[0], (lxw_col_t) r[1],
                        (lxw_row_t) r[2], (lxw_col_t) r[3]);
    assert_lxw(worksheet_ignore_errors(ctx->sheet,
                                       (uint8_t) payload_int(p, "type"),
                                       range));
  } else {
    Rf_errorcall(R_NilValue,
                 "Error in writexl: unknown sheet overlay kind '%s'", kind);
  }
}

static void apply_sheet_overlays(cell_write_ctx *ctx, SEXP opts){
  if(opts == R_NilValue) return;
  SEXP ov = list_get(opts, "overlay");
  if(ov == R_NilValue || !Rf_isVectorList(ov)) return;
  R_xlen_t n = Rf_length(ov);
  for(R_xlen_t k = 0; k < n; k++)
    apply_overlay(ctx, VECTOR_ELT(ov, k));
}

/* --- Individual per-type cell writers ------------------------------------ */
/*
 * Each writer returns 1 if it wrote a cell and 0 if the value had no
 * representation in xlsx (NA, NaN).  Callers that care -- write_cell_general()
 * -- use that to fall back to a formatted blank rather than losing the format.
 */

static int write_cell_date(cell_write_ctx *ctx, lxw_row_t row, lxw_col_t col,
                            SEXP col_data, lxw_row_t i, lxw_format *fmt){
  double val = Rf_isReal(col_data) ? REAL(col_data)[i] : INTEGER(col_data)[i];
  if(Rf_isReal(col_data) ? R_FINITE(val) : val != NA_INTEGER){
    /* Same arithmetic as write_cell_posixct() below: R days since 1970-01-01
       map to Excel serial 25568 + val, then +1 for every date on or after
       1900-03-01 to skip Excel's phantom 1900-02-29 (Excel treats 1900 as a
       leap year).  Without the correction 1900-01-01 .. 1900-02-28 are written
       one day too late, and 1900-02-28 lands on the nonexistent serial 60. */
    double serial = 25568.0 + val;
    if(serial >= 60.0)
      serial = serial + 1.0;
    assert_lxw(worksheet_write_number(ctx->sheet, row, col, serial, fmt));
    return 1;
  }
  return 0;
}

static int write_cell_posixct(cell_write_ctx *ctx, lxw_row_t row, lxw_col_t col,
                               SEXP col_data, lxw_row_t i, lxw_format *fmt){
  double val = REAL(col_data)[i];
  if(R_FINITE(val)){
    val = 25568.0 + val / (24*60*60);
    if(val >= 60.0)
      val = val + 1.0;
    assert_lxw(worksheet_write_number(ctx->sheet, row, col, val, fmt));
    return 1;
  }
  return 0;
}

static int write_cell_string(cell_write_ctx *ctx, lxw_row_t row, lxw_col_t col,
                              SEXP col_data, lxw_row_t i, lxw_format *fmt){
  SEXP val = STRING_ELT(col_data, i);
  // NB: xlsx does distinguish between empty string and NA
  if(val != NA_STRING && Rf_length(val)){
    assert_lxw(worksheet_write_string(ctx->sheet, row, col, Rf_translateCharUTF8(val), fmt));
    return 1;
  }
  return 0;
}

static int write_cell_real(cell_write_ctx *ctx, lxw_row_t row, lxw_col_t col,
                            SEXP col_data, lxw_row_t i, lxw_format *fmt){
  double val = REAL(col_data)[i];
  if(val == R_PosInf)
    assert_lxw(worksheet_write_string(ctx->sheet, row, col, "Inf", fmt));
  else if(val == R_NegInf)
    assert_lxw(worksheet_write_string(ctx->sheet, row, col, "-Inf", fmt));
  else if(R_FINITE(val)) // skips NA and NAN
    assert_lxw(worksheet_write_number(ctx->sheet, row, col, val, fmt));
  else
    return 0;
  return 1;
}

static int write_cell_integer(cell_write_ctx *ctx, lxw_row_t row, lxw_col_t col,
                               SEXP col_data, lxw_row_t i, lxw_format *fmt){
  int val = INTEGER(col_data)[i];
  if(val != NA_INTEGER){
    assert_lxw(worksheet_write_number(ctx->sheet, row, col, val, fmt));
    return 1;
  }
  return 0;
}

static int write_cell_logical(cell_write_ctx *ctx, lxw_row_t row, lxw_col_t col,
                               SEXP col_data, lxw_row_t i, lxw_format *fmt){
  int val = LOGICAL(col_data)[i];
  if(val != NA_LOGICAL){
    assert_lxw(worksheet_write_boolean(ctx->sheet, row, col, val, fmt));
    return 1;
  }
  return 0;
}

/* --- Atomic value dispatcher (all non-xl_cell_general types) ------------- */

static int write_atomic_value(cell_write_ctx *ctx, lxw_row_t row, lxw_col_t col,
                               SEXP col_data, R_COL_TYPE type, lxw_row_t i,
                               lxw_format *fmt){
  switch(type){
  case COL_DATE:
    return write_cell_date(ctx, row, col, col_data, i, fmt);
  case COL_POSIXCT:
    return write_cell_posixct(ctx, row, col, col_data, i, fmt);
  case COL_STRING:
    return write_cell_string(ctx, row, col, col_data, i, fmt);
  case COL_REAL:
    return write_cell_real(ctx, row, col, col_data, i, fmt);
  case COL_INTEGER:
    return write_cell_integer(ctx, row, col, col_data, i, fmt);
  case COL_LOGICAL:
    return write_cell_logical(ctx, row, col, col_data, i, fmt);
  default:
    return 0;
  }
}

/* --- xl_cell_general field accessors ------------------------------------- */

/* A per-cell logical field, absent or NA reading as FALSE. */
static int cell_flag(SEXP cell, const char *key){
  SEXP v = list_get(cell, key);
  if(v == R_NilValue || Rf_length(v) < 1) return 0;
  int f = Rf_asLogical(v);
  return f == NA_LOGICAL ? 0 : f;
}

/*
 * A formula's pre-calculated numeric result, if it has one.  Both double and
 * integer values count: an integer result used to fall through to the
 * no-result writer, silently dropping the value the caller supplied.
 */
static int formula_result(SEXP value, double *out){
  if(value == R_NilValue || Rf_length(value) < 1) return 0;
  if(TYPEOF(value) == REALSXP && R_FINITE(REAL(value)[0])){
    *out = REAL(value)[0];
    return 1;
  }
  if(TYPEOF(value) == INTSXP && INTEGER(value)[0] != NA_INTEGER){
    *out = (double) INTEGER(value)[0];
    return 1;
  }
  return 0;
}

/*
 * The array range resolved in R ("array_quad", 0-based first/last row/col).
 * Falls back to the anchor cell so a missing quad degenerates to a single-cell
 * array formula rather than to an invalid range.
 */
static void array_quad(SEXP cell, lxw_row_t *r1, lxw_col_t *c1,
                       lxw_row_t *r2, lxw_col_t *c2,
                       lxw_row_t row, lxw_col_t col){
  *r1 = row; *c1 = col; *r2 = row; *c2 = col;
  SEXP q = list_get(cell, "array_quad");
  if(q == R_NilValue || Rf_length(q) < 4) return;
  SEXP qi = PROTECT(Rf_coerceVector(q, INTSXP));
  int *a = INTEGER(qi);
  if(a[0] != NA_INTEGER && a[1] != NA_INTEGER &&
     a[2] != NA_INTEGER && a[3] != NA_INTEGER){
    *r1 = (lxw_row_t) a[0]; *c1 = (lxw_col_t) a[1];
    *r2 = (lxw_row_t) a[2]; *c2 = (lxw_col_t) a[3];
  }
  UNPROTECT(1);
}

/*
 * Write a rich string: one cell whose text is split into runs, each with its
 * own font.  R has already flattened the runs into the parallel "rich_text" and
 * "rich_format_id" vectors and guaranteed at least two of them (libxlsxwriter
 * rejects fewer).  The tuple array libxlsxwriter takes is NULL-terminated.
 */
static void write_rich_string_cell(cell_write_ctx *ctx, lxw_row_t row,
                                    lxw_col_t col, SEXP txt, SEXP ids,
                                    lxw_format *fmt){
  R_xlen_t n = Rf_length(txt);
  lxw_rich_string_tuple **tuples = (lxw_rich_string_tuple **)
    R_alloc((size_t) n + 1, sizeof(lxw_rich_string_tuple *));
  lxw_rich_string_tuple *runs = (lxw_rich_string_tuple *)
    R_alloc((size_t) n, sizeof(lxw_rich_string_tuple));
  for(R_xlen_t k = 0; k < n; k++){
    runs[k].string = Rf_translateCharUTF8(STRING_ELT(txt, k));
    int fid = (ids != R_NilValue && TYPEOF(ids) == INTSXP &&
               Rf_length(ids) > k) ? INTEGER(ids)[k] : 0;
    runs[k].format = ctx_format(ctx, fid);
    note_protection(ctx, fid);
    tuples[k] = &runs[k];
  }
  tuples[n] = NULL;
  assert_lxw(worksheet_write_rich_string(ctx->sheet, row, col, tuples, fmt));
}

/*
 * write_cell_general: write the i-th cell of an xl_cell_general column.
 * col is the full xl_cell_general list-column; VECTOR_ELT(col, i) is one
 * per-cell named list with elements: value, formula, hyperlink.  The
 * per-cell effective format id is carried in the integer attribute
 * "writexl_format_ids" (resolved entirely in R); 0 means no format.
 *
 * Priority: hyperlink > formula > value.
 */
static void write_cell_general(cell_write_ctx *ctx,
                                lxw_row_t row, lxw_col_t col_idx,
                                SEXP col, lxw_row_t i){
  SEXP cell      = VECTOR_ELT(col, (R_xlen_t) i);
  SEXP value     = list_get(cell, "value");
  SEXP formula   = list_get(cell, "formula");
  SEXP hyperlink = list_get(cell, "hyperlink");

  /* per-cell effective format (resolved in R) */
  int fid = 0;
  SEXP ids = Rf_getAttrib(col, Rf_install("writexl_format_ids"));
  if(ids != R_NilValue && TYPEOF(ids) == INTSXP && Rf_length(ids) > (R_xlen_t) i)
    fid = INTEGER(ids)[i];
  lxw_format *fmt = ctx_format(ctx, fid);
  note_protection(ctx, fid);

  /* --- comment overlay (independent of value/formula/hyperlink) ----------- */
  SEXP comment = list_get(cell, "comment");
  if(comment != R_NilValue && Rf_isVectorList(comment)){
    const char *ctext = payload_str(comment, "text");
    if(ctext){
      lxw_comment_options copts;
      memset(&copts, 0, sizeof(copts));
      const char *s;
      if((s = payload_str(comment, "author")))    copts.author = s;
      if((s = payload_str(comment, "font_name"))) copts.font_name = s;
      if(payload_has(comment, "visible"))     copts.visible     = (uint8_t)  payload_int(comment, "visible");
      if(payload_has(comment, "width"))       copts.width       = (uint16_t) payload_int(comment, "width");
      if(payload_has(comment, "height"))      copts.height      = (uint16_t) payload_int(comment, "height");
      if(payload_has(comment, "x_scale"))     copts.x_scale     = payload_dbl(comment, "x_scale");
      if(payload_has(comment, "y_scale"))     copts.y_scale     = payload_dbl(comment, "y_scale");
      if(payload_has(comment, "color"))       copts.color       = (lxw_color_t) payload_int(comment, "color");
      if(payload_has(comment, "font_size"))   copts.font_size   = payload_dbl(comment, "font_size");
      if(payload_has(comment, "font_family")) copts.font_family = (uint8_t)  payload_int(comment, "font_family");
      if(payload_has(comment, "start_row"))   copts.start_row   = (lxw_row_t) payload_int(comment, "start_row");
      if(payload_has(comment, "start_col"))   copts.start_col   = (lxw_col_t) payload_int(comment, "start_col");
      if(payload_has(comment, "x_offset"))    copts.x_offset    = payload_int(comment, "x_offset");
      if(payload_has(comment, "y_offset"))    copts.y_offset    = payload_int(comment, "y_offset");
      assert_lxw(worksheet_write_comment_opt(ctx->sheet, row, col_idx, ctext, &copts));
    }
  }

  /* --- hyperlink (character value provides the optional display string) --- */
  const char *display = NULL;
  if(value != R_NilValue && TYPEOF(value) == STRSXP &&
     Rf_length(value) > 0 && STRING_ELT(value, 0) != NA_STRING)
    display = Rf_translateCharUTF8(STRING_ELT(value, 0));

  if(hyperlink != R_NilValue && !Rf_isNull(hyperlink)){
    /* the hyperlink style (default + hyperlink_format + cell) is resolved
       into `fmt` on the R side; NULL only if no styling at all applies */
    if(TYPEOF(hyperlink) == STRSXP && STRING_ELT(hyperlink, 0) != NA_STRING){
      assert_lxw(worksheet_write_url_opt(ctx->sheet, row, col_idx,
                   Rf_translateCharUTF8(STRING_ELT(hyperlink, 0)),
                   fmt, display, NULL));
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
                     fmt, display, tooltip));
        return;
      }
    }
    /* hyperlink is NA or invalid — fall through to formula / value */
  }

  /* --- formula (with optional pre-calculated value) ----------------------- */
  if(formula != R_NilValue && TYPEOF(formula) == STRSXP &&
     STRING_ELT(formula, 0) != NA_STRING && Rf_length(formula) > 0){
    const char *fstr = Rf_translateCharUTF8(STRING_ELT(formula, 0));
    double result;
    int have_num = formula_result(value, &result);
    int is_array   = cell_flag(cell, "array");
    int is_dynamic = cell_flag(cell, "dynamic");

    if(is_array || is_dynamic){
      /* The range is resolved in R: absent it is the cell itself, which is the
         modern dynamic-array spelling (Excel spills the result).  There is no
         write_array_formula_str(), so a character result cannot reach here --
         xl_cell_general() rejects that combination. */
      lxw_row_t r1 = row, r2 = row;
      lxw_col_t c1 = col_idx, c2 = col_idx;
      array_quad(cell, &r1, &c1, &r2, &c2, row, col_idx);
      int single = (r1 == r2 && c1 == c2);

      if(is_dynamic && single){
        if(have_num)
          assert_lxw(worksheet_write_dynamic_formula_num(ctx->sheet, r1, c1,
                       fstr, fmt, result));
        else
          assert_lxw(worksheet_write_dynamic_formula(ctx->sheet, r1, c1,
                       fstr, fmt));
      } else if(is_dynamic){
        if(have_num)
          assert_lxw(worksheet_write_dynamic_array_formula_num(ctx->sheet,
                       r1, c1, r2, c2, fstr, fmt, result));
        else
          assert_lxw(worksheet_write_dynamic_array_formula(ctx->sheet,
                       r1, c1, r2, c2, fstr, fmt));
      } else {
        if(have_num)
          assert_lxw(worksheet_write_array_formula_num(ctx->sheet,
                       r1, c1, r2, c2, fstr, fmt, result));
        else
          assert_lxw(worksheet_write_array_formula(ctx->sheet,
                       r1, c1, r2, c2, fstr, fmt));
      }
      return;
    }

    if(have_num){
      assert_lxw(worksheet_write_formula_num(ctx->sheet, row, col_idx,
                                              fstr, fmt, result));
    } else if(value != R_NilValue && TYPEOF(value) == STRSXP &&
              Rf_length(value) > 0 && STRING_ELT(value, 0) != NA_STRING){
      assert_lxw(worksheet_write_formula_str(ctx->sheet, row, col_idx, fstr, fmt,
                   Rf_translateCharUTF8(STRING_ELT(value, 0))));
    } else {
      assert_lxw(worksheet_write_formula(ctx->sheet, row, col_idx, fstr, fmt));
    }
    return;
  }

  /* --- rich string: the cell's text, in per-run fonts --------------------- */
  /* R refuses a rich string alongside a formula or hyperlink, so reaching here
     after those branches means the runs are the cell's whole content. */
  SEXP rtext = list_get(cell, "rich_text");
  if(rtext != R_NilValue && TYPEOF(rtext) == STRSXP && Rf_length(rtext) > 0){
    write_rich_string_cell(ctx, row, col_idx, rtext,
                           list_get(cell, "rich_format_id"), fmt);
    return;
  }

  /* --- value only: dispatch through the atomic writer --------------------- */
  /* A cell whose value has no xlsx representation (absent, NA, NaN) still needs
     to exist when a format applies, or the format is silently lost.  Note that
     worksheet_write_blank() is a no-op on a NULL format -- Excel ignores
     unformatted blanks -- so the unformatted case costs nothing and adds no
     cell.  Plain (non-xl_cell_general) columns are deliberately left alone:
     their NAs are covered by the column format Excel already applies. */
  if(value == R_NilValue || Rf_isNull(value) || Rf_length(value) == 0 ||
     !write_atomic_value(ctx, row, col_idx, value, get_type(value), 0, fmt))
    assert_lxw(worksheet_write_blank(ctx->sheet, row, col_idx, fmt));
}

/* --- Top-level cell dispatcher ------------------------------------------- */

static void write_cell(cell_write_ctx *ctx, lxw_row_t row, lxw_col_t col,
                        SEXP col_data, R_COL_TYPE type, lxw_row_t i){
  if(type == COL_CELL_GENERAL)
    write_cell_general(ctx, row, col, col_data, i);
  else
    write_atomic_value(ctx, row, col, col_data, type, i, NULL);
}

/* --- Workbook document properties ---------------------------------------- */

static void apply_properties(lxw_workbook *wb, SEXP props){
  if(props == R_NilValue) return;

  /* document metadata */
  lxw_doc_properties dp;
  memset(&dp, 0, sizeof(dp));
  const char *keys[] = {"title","subject","author","manager","company",
                        "category","keywords","comments","status","hyperlink_base"};
  const char **fields[] = {&dp.title,&dp.subject,&dp.author,&dp.manager,&dp.company,
                           &dp.category,&dp.keywords,&dp.comments,&dp.status,&dp.hyperlink_base};
  int any_meta = 0;
  for(int i = 0; i < 10; i++){
    const char *v = payload_str(props, keys[i]);
    if(v){ *fields[i] = v; any_meta = 1; }
  }
  if(any_meta)
    workbook_set_properties(wb, &dp);

  /* custom properties (typed) */
  SEXP custom = list_get(props, "custom");
  if(custom != R_NilValue && Rf_isVectorList(custom)){
    for(R_xlen_t i = 0; i < Rf_length(custom); i++){
      SEXP e = VECTOR_ELT(custom, i);
      const char *nm = payload_str(e, "name");
      const char *ty = payload_str(e, "type");
      SEXP val = list_get(e, "value");
      if(!nm || !ty || val == R_NilValue) continue;
      if(!strcmp(ty, "string"))
        workbook_set_custom_property_string(wb, nm, payload_str(e, "value"));
      else if(!strcmp(ty, "integer"))
        workbook_set_custom_property_integer(wb, nm, Rf_asInteger(val));
      else if(!strcmp(ty, "number"))
        workbook_set_custom_property_number(wb, nm, Rf_asReal(val));
      else if(!strcmp(ty, "boolean"))
        workbook_set_custom_property_boolean(wb, nm, Rf_asLogical(val) == TRUE);
      else if(!strcmp(ty, "datetime")){
        /* broken out into fields on the R side, after the workbook-wide time
           zone decision has been applied */
        lxw_datetime dt;
        memset(&dt, 0, sizeof(dt));
        dt.year  = payload_int(val, "year");
        dt.month = payload_int(val, "month");
        dt.day   = payload_int(val, "day");
        dt.hour  = payload_int(val, "hour");
        dt.min   = payload_int(val, "min");
        dt.sec   = payload_dbl(val, "sec");
        workbook_set_custom_property_datetime(wb, nm, &dt);
      }
    }
  }

  /* read-only recommended */
  if(opt_scalar_int(props, "read_only", 0))
    workbook_read_only_recommended(wb);

  /* hyperlink styling opt-out.  worksheet_write_url_opt() falls back to the
     workbook's default_url_format when it is handed a NULL format, so an
     unstyled hyperlink is only reachable by clearing that default. */
  if(opt_scalar_int(props, "unset_url_format", 0))
    workbook_unset_default_url_format(wb);

  /* window size */
  SEXP ws = list_get(props, "window_size");
  if(ws != R_NilValue && Rf_length(ws) >= 2){
    SEXP wi = PROTECT(Rf_coerceVector(ws, INTSXP));
    workbook_set_size(wb, (uint16_t) INTEGER(wi)[0], (uint16_t) INTEGER(wi)[1]);
    UNPROTECT(1);
  }

  /* defined names */
  SEXP names = list_get(props, "names");
  if(names != R_NilValue && Rf_isVectorList(names)){
    for(R_xlen_t i = 0; i < Rf_length(names); i++){
      SEXP e = VECTOR_ELT(names, i);
      const char *nm = payload_str(e, "name");
      const char *fm = payload_str(e, "formula");
      if(nm && fm) workbook_define_name(wb, nm, fm);
    }
  }
}

/* --- Main entry point ---------------------------------------------------- */

SEXP C_write_data_frame_list(SEXP df_list, SEXP file, SEXP col_names,
                             SEXP format_headers, SEXP use_zip64, SEXP formats,
                             SEXP sheets, SEXP header_id, SEXP properties){
  assert_that(Rf_isVectorList(df_list), "Object is not a list");
  assert_that(Rf_isString(file) && Rf_length(file), "Invalid file path");
  assert_that(Rf_isLogical(col_names), "col_names must be logical");
  assert_that(Rf_isLogical(format_headers), "format_headers must be logical");
  bail_if(formats != R_NilValue && !Rf_isVectorList(formats), "formats must be a list");
  bail_if(sheets != R_NilValue && !Rf_isVectorList(sheets), "sheets must be a list");

  //create workbook.  constant_memory (libxlsxwriter's row-streaming mode) is
  //resolved on the R side by .resolve_constant_memory() and arrives inside the
  //properties payload; it defaults to on.
  lxw_workbook_options options = {
    .constant_memory = (uint8_t) (opt_scalar_int(properties, "constant_memory", 1) ? 1 : 0),
    .tmpdir = TEMPDIR,
    .use_zip64 = Rf_asLogical(use_zip64)
  };
  lxw_workbook *workbook = workbook_new_opt(Rf_translateChar(STRING_ELT(file, 0)), &options);
  assert_that(workbook, "failed to create workbook");

  //build the user-format table (1-based; index 0 is the NULL format).  All
  //formats -- header, hyperlink, date, and user cell formats -- come from R.
  int nfmts = (formats == R_NilValue) ? 0 : (int) Rf_length(formats);
  lxw_format **fmts = (lxw_format **) R_alloc((size_t) nfmts + 1, sizeof(lxw_format *));
  fmts[0] = NULL;
  int *fmt_prot = (int *) R_alloc((size_t) nfmts + 1, sizeof(int));
  fmt_prot[0] = 0;
  for(int k = 0; k < nfmts; k++){
    SEXP payload = VECTOR_ELT(formats, k);
    fmts[k + 1] = build_lxw_format(workbook, payload);
    fmt_prot[k + 1] = payload_has(payload, "unlocked") || payload_has(payload, "hidden");
  }

  //header row format id (0 = none) and height, resolved on the R side
  int hdr_id = (header_id == R_NilValue) ? 0 : Rf_asInteger(header_id);
  double hdr_height = opt_scalar_dbl(properties, "header_row_height", LXW_DEF_ROW_HEIGHT);

  //workbook document properties
  apply_properties(workbook, properties);

  //iterate over sheets
  SEXP df_names = PROTECT(Rf_getAttrib(df_list, R_NamesSymbol));
  R_xlen_t nsheets = Rf_length(df_list);
  for(R_xlen_t s = 0; s < nsheets; s++){

    //create sheet
    const char * sheet_name = Rf_length(df_names) > s && Rf_length(STRING_ELT(df_names, s)) ? \
      Rf_translateCharUTF8(STRING_ELT(df_names, s)) : NULL;
    /* .resolve_sheet_names() has already repaired the name on the R side.  This
       is a backstop so an R-side regression cannot silently produce a workbook
       Excel refuses to open; a NULL name is libxlsxwriter's own auto-naming and
       is not ours to validate. */
    if(sheet_name)
      assert_lxw(workbook_validate_sheet_name(workbook, sheet_name));
    lxw_worksheet *sheet = workbook_add_worksheet(workbook, sheet_name);
    assert_that(sheet, "failed to create workbook");

    //per-sheet worksheet options (column/row geometry, freeze panes, ...)
    SEXP opts = (sheets != R_NilValue && Rf_length(sheets) > s) ? VECTOR_ELT(sheets, s) : R_NilValue;

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
      lxw_format *hdr_fmt = (hdr_id > 0 && hdr_id <= nfmts) ? fmts[hdr_id] : NULL;
      for(lxw_col_t i = 0; i < cols; i++)
        assert_lxw(worksheet_write_string(sheet, cursor, i, Rf_translateCharUTF8(STRING_ELT(names, i)), NULL));
      if(Rf_asLogical(format_headers))
        assert_lxw(worksheet_set_row(sheet, cursor, hdr_height, hdr_fmt));
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
      if(coltypes[i] == COL_UNKNOWN)
        Rf_warning("Column '%s' has unrecognized data type.", CHAR(STRING_ELT(names, i)));
    }
    UNPROTECT(1); //names

    // validate row count against xlsx limit
    lxw_row_t max_rows = (lxw_row_t)(LXW_ROW_MAX - (Rf_asLogical(col_names) ? 1 : 0));
    bail_if(rows > max_rows, "data frame has too many rows for xlsx (max 1048576)");

    // Build context for cell writers
    int sheet_protected = opt_scalar_int(opts, "protect", 0);
    int warn_unlocked = 0;
    cell_write_ctx ctx = {workbook, sheet, fmts, nfmts, fmt_prot,
                          sheet_protected, &warn_unlocked};

    // Apply column geometry/formats and sheet-level options (freeze, etc.),
    // then the range-scoped overlays.  All of this must precede the row loop
    // because constant_memory flushes each row as it is written.
    apply_columns(&ctx, opts, cols);
    apply_sheet_scalars(&ctx, opts);
    apply_page_setup(&ctx, opts);
    apply_sheet_view(&ctx, opts);
    apply_outline(&ctx, opts);
    apply_sheet_overlays(&ctx, opts);

    // Need to iterate by row first for performance
    for (lxw_row_t i = 0; i < rows; i++) {
      // Row options must be set before the row is flushed (constant_memory).
      apply_row(&ctx, opts, cursor);
      for(lxw_col_t j = 0; j < cols; j++){
        SEXP col = VECTOR_ELT(df, j);
        if(Rf_length(col) <= (R_xlen_t) i) continue;
        write_cell(&ctx, cursor, j, col, coltypes[j], i);
      }
      cursor++;
    }

    /* Merges run after the rows.  worksheet_merge_range() writes the top-left
       text and blanks the rest of the range, so doing it here means a merge
       over cells the data frame filled keeps only the merged value -- exactly
       what merging does in Excel.  Writing back over emitted rows is why any
       merge turns constant memory off on the R side. */
    apply_merges(&ctx, opts);
    apply_tables(&ctx, opts);
    apply_images(&ctx, opts);

    if(warn_unlocked)
      Rf_warning("Worksheet '%s' uses cell protection formatting (locked = FALSE "
                 "or hidden = TRUE) but the worksheet is not protected, so it has "
                 "no effect. Set protect= in xl_sheet() to protect the worksheet.",
                 sheet_name ? sheet_name : "");
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
  {"C_write_data_frame_list", (DL_FUNC) &C_write_data_frame_list, 9},
  {NULL, NULL, 0}
};

attribute_visible void R_init_writexl(DllInfo *dll) {
  R_registerRoutines(dll, NULL, CallEntries, NULL, NULL);
  R_useDynamicSymbols(dll, FALSE);
  R_forceSymbols(dll, TRUE);
}
