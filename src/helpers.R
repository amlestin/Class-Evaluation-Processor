get_worksheet_entries <- function(wb, sheet) {
  #  https://github.com/awalker89/openxlsx/pull/382/commits/7e9cfd09934ac61491b71e91785c415a22b3dc31
  # get worksheet data
  dat <- wb$worksheets[[sheet]]$sheet_data
  # get vector of entries
  val <- dat$v
  # get boolean vector of text entries
  typ <- (dat$t == 1) & !is.na(dat$t)
  # get text entry strings
  str <- unlist(wb$sharedStrings[as.integer(val)[typ] + 1])
  # remove xml tags
  str <- gsub("<.*?>", "", str)
  # write strings to vector of entries
  val[typ] <- str
  # return vector of entries
  val
}

auto_heights <-
  function(wb,
           sheet,
           selected,
           fontsize = NULL,
           factor = 1.0,
           base_height = 15,
           extra_height = 12) {
    #    https://github.com/awalker89/openxlsx/pull/382/commits/7e9cfd09934ac61491b71e91785c415a22b3dc31
    # get base font size
    if (is.null(fontsize)) {
      fontsize <- as.integer(openxlsx::getBaseFont(wb)$size$val)
    }
    # set factor to adjust font width (empiricially found scale factor 4 here)
    factor <- 4 * factor / fontsize
    # get worksheet data
    dat <- wb$worksheets[[sheet]]$sheet_data
    # get columns widths
    colWidths <- wb$colWidths[[sheet]]
    # select fixed (non-auto) and visible (non-hidden) columns only
    specified <-
      (colWidths != "auto") & (attr(colWidths, "hidden") == "0")
    # return default row heights if no column widths are fixed
    if (length(specified) == 0) {
      message("No column widths specified, returning default row heights.")
      cols <- integer(0)
      heights <- rep(base_height, length(selected))
      return(list(cols, heights))
    }
    # get fixed column indices
    cols <- as.integer(names(specified)[specified])
    # get fixed column widths
    widths <- as.numeric(colWidths[specified])
    # get all worksheet entries
    val <- get_worksheet_entries(wb, sheet)
    # compute optimal height per selected row
    heights <- sapply(selected, function(row) {
      # select entries in given row and columns of fixed widths
      index <- (dat$rows == row) & (dat$cols %in% cols)
      # remove line break characters
      chr <- gsub("\\r|\\n", "", val[index])
      # measure width of entry (in pixels)
      wdt <-
        strwidth(chr, unit = "in") * 20 / 1.43 # 20 px = 1.43 in
      # compute optimal height
      if (length(wdt) == 0) {
        base_height
      } else {
        base_height + extra_height * as.integer(max(wdt / widths * factor))
      }
    })
    # return list of indices of columns with fixed widths and optimal row heights
    list(cols, heights)
  }