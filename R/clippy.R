img_gif <- function( file, Rd = FALSE, alt = "image" ){
  input <- normalizePath(file, mustWork = TRUE)
  buf <- readBin(input, raw(), file.info(input)$size)
  base64 <- jsonlite::base64_enc(buf)
  sprintf('%s<img src="data:image/gif;base64,\n%s" alt="%s" />%s',
          if( Rd ) "\\out{" else "", base64, alt, if( Rd ) "}" else ""
  )
}
