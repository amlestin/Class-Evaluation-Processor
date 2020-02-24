# HIM Simplified Reports


inputs <- choose.files()

for (f in inputs) {
  #input.file.name <- file.choose()
  input.file.name <- f
  dist <- read.csv(input.file.name, stringsAsFactors = FALSE)
  
  library(openxlsx)
  
  # sets the directory to output reports files to
  setwd("~/Class-Evaluation-Processor/reports")
  
  wb <- createWorkbook()
  options("openxlsx.numFmt" = "0.00")
  addWorksheet(wb, 1)
  
  # find the columns corresponding to questions
  response.cols <- grep("X[0-9]", colnames(dist))
  # remove comments column
  response.cols <- response.cols[-length(response.cols)]
  
  # collect the questions and their averages
  q.a <- c()
  for (col.num in 1:length(response.cols)) {
    question.col <- as.vector(dist[, response.cols[col.num]])
    question.col <- as.numeric(question.col[-c(1:2)])
    
    
    question.average <- round(mean(question.col), 2)
    question.average <- format(question.average, nsmall = 2)
    
    q.a <-
      rbind(q.a, c(dist[1, response.cols[col.num]], question.average, length(question.col)))
  }
  names(q.a) <- NULL
  
  # add column names
  q.a <-
    rbind(c("", "", ""), c("Question", "Average", "# Responses"), q.a)
  q.a <- rbind(c("HIMS Course Evaluations", "", ""), q.a)
  
  # extract and format comments
  comments <-  dist[3:nrow(dist), "X13"]
  comments <- comments[-which(comments == "")]
  comments <- paste(">>>", comments)
  comments <- cbind(comments, rep("", length(comments)))
  comments <- cbind(comments, rep("", length(comments)))
  
  # append comments to report
  q.a <-
    rbind(q.a, c("", "", ""), c("Comments", "", ""), c("", "", ""), comments)
  
  
  # define formatting of each report section
  headerStyle <-
    createStyle(fontSize = 20,
                fontColour = "#333333",
                halign = "center")
  tableHeaderStyle <-
    createStyle(
      fontSize = 13,
      fontColour = "#000000",
      wrapText = TRUE,
      textDecoration = "bold",
      border = "TopBottomLeftRight",
      borderStyle = "thin"
    )
  tableStyle <-
    createStyle(
      fontSize = 12,
      fontColour = "#000000",
      border = "TopBottomLeftRight",
      borderColour = "#000000",
      borderStyle = "thin",
      halign = "left"
    )
  textStyle <-
    createStyle(
      fontSize = 12,
      fontColour = "#000000",
      wrapText = TRUE,
      valign = "top"
    )
  commentHeaderStyle <-
    createStyle(
      fontSize = 13,
      fontColour = "#000000",
      wrapText = TRUE,
      textDecoration = "bold",
      border = "bottom",
      borderStyle = "thick"
    )
  
  # set the report column widths
  setColWidths(wb, 1, cols = 1, widths = 90)
  setColWidths(wb, 1, cols = 2, widths = 13)
  setColWidths(wb, 1, cols = 3, widths = 13)
  
  # format the report header text
  addStyle(wb, 1, headerStyle, cols = 1, rows = 1)
  
  # merge the cells of the report header
  mergeCells(wb, 1, cols = 1:3, rows = 1)
  
  # format the comments header
  addStyle(wb, 1, commentHeaderStyle, cols = 1, rows = 17)
  addStyle(wb, 1, commentHeaderStyle, cols = 2, rows = 17)
  addStyle(wb, 1, commentHeaderStyle, cols = 3, rows = 17)
  
  # apply formatting to the table
  Map(function(i) {
    addStyle(wb, 1, tableStyle, cols = 1, rows = i)
    addStyle(wb, 1, tableStyle, cols = 2, rows = i)
    addStyle(wb, 1, tableStyle, cols = 3, rows = i)
  }, 3:15)
  
  # apply formatting to all comments
  Map(function(r) {
    addStyle(wb, 1, textStyle, cols = 1, rows = r)
    mergeCells(wb, 1, cols = 1:3, rows = r)
  }, 18:(14 + length(comments)))
  
  # apply tableHeaderStyle to each column of the table header
  addStyle(wb, 1, tableHeaderStyle, cols = 1, rows = 3)
  addStyle(wb, 1, tableHeaderStyle, cols = 2, rows = 3)
  addStyle(wb, 1, tableHeaderStyle, cols = 3, rows = 3)
  
  # makes sure sheet fits on one printable page
  pageSetup(wb,
            1,
            fitToWidth = TRUE,
            fitToHeight = FALSE)
  
  # calculate the vertical heights of the comments
  comments <- comments[-which(comments == "")]
  heights <- strwidth(comments, units = "in", font = 1) * 10 / 1.43
  heights <- heights * 1.2
  setRowHeights(wb, 1, rows = 19:nrow(q.a), heights)
  
  # write to the workbook then write the xlsx file to disk
  writeData(wb, 1, q.a, colNames = FALSE)
  
  report.file.name <- basename(input.file.name)
  report.file.name <-
    substr(report.file.name, 1, (nchar(report.file.name) - 4))
  report.file.name <- make.names(report.file.name)
  report.file.name <- gsub("\\.", " ", report.file.name)
  report.file.name <-
    gsub("(?<=[\\s])\\s*|^\\s+|\\s+$", "", report.file.name, perl = TRUE)
  
  saveWorkbook(wb, paste(report.file.name, "HIMS eval.xlsx"), overwrite = TRUE)
}
