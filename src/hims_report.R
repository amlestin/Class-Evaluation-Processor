# HIM Simplified Reports

## Shows a data frame of how often each score occurs within the responses
## Function loops through score response column and gets the frequency associated with that score
countScoreResponse <- function(scoreResponses)
{
  scoreTable <- as.data.frame(table(scoreResponses))
  scoreVector <- 1:5
  scoreOccurences <- c()
  
  for (score in scoreVector)
  {
    found <- FALSE
    
    for (index in 1:length(scoreTable$scoreResponses))
    {
      if (score == scoreTable$scoreResponses[index])
      {
        found <- TRUE
        scoreOccurences <- append(scoreOccurences, scoreTable$Freq[index])
      }
    }
    ## When the current iterated score is not found, it's saying there are 0 responses that match that score
    if (!found) scoreOccurences <- append(scoreOccurences, 0)
  }
  print(scoreOccurences)
  names(scoreOccurences) <- c("Poor", "Fair", "Good", "Very Good", "Excellent")
  return(scoreOccurences)
}

inputs <- choose.files()

for (f in inputs) {
  #input.file.name <- file.choose()
  input.file.name <- f
  dist <- read.csv(input.file.name, stringsAsFactors = FALSE)
  
  library(openxlsx)
  
  # sets the directory to output reports files to
  work_dir <- "~/Evals/"
  setwd(work_dir)
  
  
  wb <- createWorkbook()
  options("openxlsx.numFmt" = "0.00")
  addWorksheet(wb, 1)
  
  # Report name
  report.file.name <- basename(input.file.name)
  report.file.name <- substr(report.file.name, 1, (nchar(report.file.name) - 4))
  report.file.name <- make.names(report.file.name)
  report.file.name <- gsub("\\.", " ", report.file.name)
  report.file.name <- gsub("(?<=[\\s])\\s*|^\\s+|\\s+$", "", report.file.name, perl = TRUE)
  test.name <-  report.file.name
  
  # find the columns corresponding to questions
  response.cols <- grep("X[0-9]", colnames(dist))
  # remove comments column
  response.cols <- response.cols[-length(response.cols)]
  
  # collect the questions and their averages
  q.a <- c()
  for (col.num in 1:length(response.cols)) {
    question.col <- as.vector(dist[, response.cols[col.num]])
    question.col <- as.numeric(question.col[-c(1:2)])
    
    
    question.average <- round(mean(question.col, na.rm = TRUE),  2)
    question.average <- format(question.average, nsmall = 2)
    
    ## stores the number of responses for each question column being iterated through.
    numberOfResponses <- NROW(na.omit(question.col))
    
    ## Stores the frequency of each score
    scoreOccurences <- countScoreResponse(question.col)
    
    q.a <-
      rbind(q.a, c(dist[1, response.cols[col.num]], question.average, numberOfResponses))
  }
  names(q.a) <- NULL
  
  # add column names
  q.a <-
    rbind(c("", "", ""), c("Question", "Average", "# Responses"), q.a)
	q.a <- rbind(c("", "", ""), q.a)
  q.a <- rbind(c("Course Evals", "", ""), q.a)
    q.a <- rbind(c(test.name, "", ""), q.a)
  
  # extract and format comments
  comments <-  dist[3:nrow(dist), "X13"]
  comments <- comments[-which(comments == "")]
  comments <- paste(">>>", comments)
  comments <- cbind(comments, rep("", length(comments)))
  comments <- cbind(comments, rep("", length(comments)))
  
   
  # append comments to report
  q.a <- rbind(q.a, c("", "", ""), c("Comments", "", ""), comments)
 
  
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
      valign = "top",
	  border = "top",
      borderColour = "#000000",
      borderStyle = "thin"
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
  addStyle(wb, 1, headerStyle, cols = 1, rows = 2)
  addStyle(wb, 1, headerStyle, cols = 1, rows = 3)
  
  # merge the cells of the report header
  mergeCells(wb, 1, cols = 1:3, rows = 1)
  mergeCells(wb, 1, cols = 1:3, rows = 2)
  mergeCells(wb, 1, cols = 1:3, rows = 3)
  
  # format the comments header
  addStyle(wb, 1, commentHeaderStyle, cols = 1, rows = 19)
  addStyle(wb, 1, commentHeaderStyle, cols = 2, rows = 19)
  addStyle(wb, 1, commentHeaderStyle, cols = 3, rows = 19)
  
  # apply formatting to the table
  Map(function(i) {
    addStyle(wb, 1, tableStyle, cols = 1, rows = i)
    addStyle(wb, 1, tableStyle, cols = 2, rows = i)
    addStyle(wb, 1, tableStyle, cols = 3, rows = i)
  }, 5:17)
  
  # apply formatting to all comments
  Map(function(r) {
    
		mergeCells(wb, 1, cols = 1:3, rows = r)
		addStyle(wb, 1, textStyle, cols = 1:3, rows = r)
	 
	  }, 20:nrow(q.a))
  
  
  # apply tableHeaderStyle to each column of the table header
  addStyle(wb, 1, tableHeaderStyle, cols = 1, rows = 5)
  addStyle(wb, 1, tableHeaderStyle, cols = 2, rows = 5)
  addStyle(wb, 1, tableHeaderStyle, cols = 3, rows = 5)
  
  # makes sure sheet fits on one printable page
  pageSetup(wb,
            1,
            fitToWidth = TRUE,
            fitToHeight = FALSE)
  
  # calculate the vertical heights of the comments
  # comments <- comments[-which(comments == "")]
 # heights <- strwidth(comments, units = "in", font = 1) * 10 / 1.43
 # heights <- heights 
  setRowHeights(wb, 1, rows = 20:nrow(q.a), heights = 50)
  
  # write to the workbook then write the xlsx file to disk
  writeData(wb, 1, q.a, colNames = FALSE)
  

  
  saveWorkbook(wb, paste(report.file.name, "course eval.xlsx"), overwrite = TRUE)
}
