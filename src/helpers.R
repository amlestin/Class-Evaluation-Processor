evals.to.average <- function(reviews) {
  # counts frequencies of each possible review response
  if (any(reviews == "Excellent") |
      any(reviews == "Very Good") |
      any(reviews == "Good") |
      any(reviews == "Fair") | any(reviews == "Poor")) {
    ecount <- length(which(reviews == "Excellent"))
    
    vgcount <- length(which(reviews == "Very Good"))
    
    gcount <- length(which(reviews == "Good"))
    
    fcount <- length(which(reviews == "Fair"))
    
    pcount <- length(which(reviews == "Poor"))
  } else {
    ecount <- length(which(reviews == "5"))
    
    vgcount <- length(which(reviews == "4"))
    
    gcount <- length(which(reviews == "3"))
    
    fcount <- length(which(reviews == "2"))
    
    pcount <- length(which(reviews == "1"))
  }
  
  # vector of all rating counts
  freqs <- c(ecount, vgcount, gcount, fcount, pcount)
  
  # sums frequencies of all reviews in the section, i.e., total number of ratings the professor received
  num.ratings <- sum(freqs)
  
  # calculates a weighted average
  ratings.prod <-
    (5 * ecount + 4 * vgcount + 3 * gcount + 2 * fcount + 1 * pcount)
  
  average <- ratings.prod / num.ratings
  average <- round(average, digits = 2)
  average <- sprintf("%0.2f", average)
  
  return (average)
}

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